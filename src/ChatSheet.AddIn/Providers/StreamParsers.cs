using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Providers
{
    /// <summary>
    /// 流式响应解析器。把各协议的增量事件归一成 ChatEvent。
    ///
    /// 工具调用是最容易出错的部分：多数协议按字符增量下发参数 JSON，
    /// 必须按索引或标识累积拼接，等本轮结束才算完整。
    /// 中途解析会得到残缺 JSON。
    /// </summary>
    internal abstract class StreamParser
    {
        /// <summary>按索引累积的工具调用。</summary>
        protected readonly Dictionary<int, ToolCall> PendingCalls = new Dictionary<int, ToolCall>();
        protected bool Completed;

        internal static StreamParser Create(ProtocolKind kind)
        {
            switch (kind)
            {
                case ProtocolKind.AnthropicMessages:
                    return new AnthropicStreamParser();
                case ProtocolKind.GoogleGemini:
                    return new GeminiStreamParser();
                case ProtocolKind.OpenAiResponses:
                    return new OpenAiResponsesStreamParser();
                default:
                    return new OpenAiChatStreamParser();
            }
        }

        /// <summary>解析一帧，产出零个或多个归一化事件。</summary>
        internal abstract IEnumerable<ChatEvent> Parse(SseFrame frame);

        /// <summary>流结束时交付累积完成的工具调用。</summary>
        internal virtual IEnumerable<ChatEvent> Flush()
        {
            foreach (var index in new SortedSet<int>(PendingCalls.Keys))
            {
                var call = PendingCalls[index];
                if (!string.IsNullOrEmpty(call.Name))
                {
                    yield return new ChatEvent { Kind = ChatEventKind.ToolCall, Call = call };
                }
            }

            PendingCalls.Clear();
        }

        protected static bool IsDone(SseFrame frame)
        {
            return string.Equals(frame.Data?.Trim(), "[DONE]", StringComparison.Ordinal);
        }

        protected static JObject TryParse(string data)
        {
            if (string.IsNullOrWhiteSpace(data))
            {
                return null;
            }

            try
            {
                return JObject.Parse(data);
            }
            catch
            {
                // 个别服务端会插入非 JSON 的调试行，跳过而不是中断整个流。
                return null;
            }
        }
    }

    /// <summary>OpenAI Chat Completions 的 SSE 解析。</summary>
    internal sealed class OpenAiChatStreamParser : StreamParser
    {
        internal override IEnumerable<ChatEvent> Parse(SseFrame frame)
        {
            if (IsDone(frame))
            {
                if (Completed) { yield break; }
                Completed = true;
                foreach (var e in Flush()) { yield return e; }
                yield return new ChatEvent { Kind = ChatEventKind.Completed };
                yield break;
            }

            var root = TryParse(frame.Data);
            if (root == null) { yield break; }

            // 体内错误。有的网关以 200 开流，再把错误放进帧里；此前这里没有这一支，
            // 于是那种错误表现为「正常返回但一个字都没有」——真实对话里是一轮空回复，
            // 而按需确认会因此把模型标成可用。Chat Completions 是默认协议，
            // 另外三个协议的解析器早就有这一支了。
            if (root["error"] is JObject error)
            {
                var message = error.Value<string>("message") ?? "服务端返回错误";
                var code = error.Value<string>("code") ?? error.Value<string>("type");
                throw new ProviderException(
                    "STREAM_ERROR",
                    message,
                    null)
                {
                    // 原文留给判据：它要认「这条错误在说谁」，而 Message 会被拼提示。
                    Detail = string.IsNullOrEmpty(code) ? message : code + "：" + message,
                };
            }

            // usage 可能单独出现在最后一帧（stream_options.include_usage）。
            var usage = root["usage"] as JObject;
            if (usage != null && usage.HasValues)
            {
                yield return new ChatEvent
                {
                    Kind = ChatEventKind.Usage,
                    PromptTokens = usage.Value<int?>("prompt_tokens") ?? 0,
                    CompletionTokens = usage.Value<int?>("completion_tokens") ?? 0,
                };
            }

            // finish_reason 后仍可能有独立 usage 帧，但不能再次交付内容或完成事件。
            if (Completed || !(root["choices"] is JArray choices) || choices.Count == 0)
            {
                yield break;
            }

            var choice = choices[0] as JObject;
            var delta = choice?["delta"] as JObject;

            if (delta != null)
            {
                var content = delta.Value<string>("content");
                if (!string.IsNullOrEmpty(content))
                {
                    yield return new ChatEvent { Kind = ChatEventKind.TextDelta, Text = content };
                }

                // 部分兼容服务把思考内容放在 reasoning_content。
                var reasoning = delta.Value<string>("reasoning_content") ?? delta.Value<string>("reasoning");
                if (!string.IsNullOrEmpty(reasoning))
                {
                    yield return new ChatEvent { Kind = ChatEventKind.ThinkingDelta, Text = reasoning };
                }

                if (delta["tool_calls"] is JArray toolCalls)
                {
                    foreach (var item in toolCalls)
                    {
                        if (!(item is JObject call)) { continue; }

                        // index 是拼接的依据：同一次调用的参数分多帧下发。
                        var index = call.Value<int?>("index") ?? 0;
                        if (!PendingCalls.TryGetValue(index, out var pending))
                        {
                            pending = new ToolCall { ArgumentsJson = string.Empty };
                            PendingCalls[index] = pending;
                        }

                        var id = call.Value<string>("id");
                        if (!string.IsNullOrEmpty(id)) { pending.Id = id; }

                        var function = call["function"] as JObject;
                        var name = function?.Value<string>("name");
                        if (!string.IsNullOrEmpty(name)) { pending.Name = name; }

                        var args = function?.Value<string>("arguments");
                        if (!string.IsNullOrEmpty(args)) { pending.ArgumentsJson += args; }
                    }
                }
            }

            var finish = choice?.Value<string>("finish_reason");
            if (!string.IsNullOrEmpty(finish))
            {
                Completed = true;
                foreach (var e in Flush()) { yield return e; }
                yield return new ChatEvent { Kind = ChatEventKind.Completed, FinishReason = finish };
            }
        }
    }

    /// <summary>OpenAI Responses 的 SSE 解析。事件名承载语义。</summary>
    internal sealed class OpenAiResponsesStreamParser : StreamParser
    {
        private readonly Dictionary<string, ToolCall> _byItemId = new Dictionary<string, ToolCall>(StringComparer.Ordinal);
        private readonly HashSet<string> _deliveredItems = new HashSet<string>(StringComparer.Ordinal);

        internal override IEnumerable<ChatEvent> Parse(SseFrame frame)
        {
            if (Completed || IsDone(frame)) { yield break; }

            var root = TryParse(frame.Data);
            if (root == null) { yield break; }

            var type = root.Value<string>("type") ?? frame.EventName ?? string.Empty;

            switch (type)
            {
                case "response.output_text.delta":
                {
                    var delta = root.Value<string>("delta");
                    if (!string.IsNullOrEmpty(delta))
                    {
                        yield return new ChatEvent { Kind = ChatEventKind.TextDelta, Text = delta };
                    }

                    break;
                }

                case "response.reasoning_summary_text.delta":
                {
                    var delta = root.Value<string>("delta");
                    if (!string.IsNullOrEmpty(delta))
                    {
                        yield return new ChatEvent { Kind = ChatEventKind.ThinkingDelta, Text = delta };
                    }

                    break;
                }

                case "response.output_item.added":
                {
                    UpdateFunctionItem(root["item"] as JObject);
                    break;
                }

                case "response.function_call_arguments.delta":
                {
                    var itemId = root.Value<string>("item_id");
                    var delta = root.Value<string>("delta");
                    if (itemId != null && _byItemId.TryGetValue(itemId, out var call) && !string.IsNullOrEmpty(delta))
                    {
                        call.ArgumentsJson += delta;
                    }

                    break;
                }

                case "response.function_call_arguments.done":
                {
                    var itemId = root.Value<string>("item_id");
                    if (itemId != null && _byItemId.TryGetValue(itemId, out var call))
                    {
                        var arguments = root.Value<string>("arguments");
                        if (arguments != null) { call.ArgumentsJson = arguments; }
                    }
                    break;
                }

                case "response.output_item.done":
                {
                    var item = root["item"] as JObject;
                    var call = UpdateFunctionItem(item);
                    var itemId = item?.Value<string>("id");
                    if (call != null && !string.IsNullOrEmpty(call.Name) && _deliveredItems.Add(itemId))
                    {
                        _byItemId.Remove(itemId);
                        yield return new ChatEvent { Kind = ChatEventKind.ToolCall, Call = call };
                    }
                    break;
                }

                case "response.completed":
                case "response.incomplete":
                {
                    Completed = true;
                    var response = root["response"] as JObject;
                    if (response?["output"] is JArray output)
                    {
                        foreach (var item in output) { UpdateFunctionItem(item as JObject); }
                    }
                    var usage = response?["usage"] as JObject;
                    if (usage != null)
                    {
                        yield return new ChatEvent
                        {
                            Kind = ChatEventKind.Usage,
                            PromptTokens = usage.Value<int?>("input_tokens") ?? 0,
                            CompletionTokens = usage.Value<int?>("output_tokens") ?? 0,
                        };
                    }

                    foreach (var e in Flush()) { yield return e; }
                    yield return new ChatEvent { Kind = ChatEventKind.Completed, FinishReason = type };
                    break;
                }

                case "response.failed":
                case "error":
                {
                    _byItemId.Clear();
                    var error = (root["response"] as JObject)?["error"] as JObject ?? root["error"] as JObject;
                    var message = error?.Value<string>("message") ?? root.Value<string>("message") ?? "服务端返回错误";
                    var code = error?.Value<string>("code") ?? root.Value<string>("code");
                    throw new ProviderException("STREAM_ERROR", message)
                    {
                        Detail = string.IsNullOrEmpty(code) ? message : code + "：" + message,
                    };
                }
            }
        }

        private ToolCall UpdateFunctionItem(JObject item)
        {
            var itemId = item?.Value<string>("id");
            if (string.IsNullOrEmpty(itemId) || _deliveredItems.Contains(itemId)) { return null; }
            if (!_byItemId.TryGetValue(itemId, out var call))
            {
                if (item.Value<string>("type") != "function_call") { return null; }
                call = new ToolCall { Id = itemId, ArgumentsJson = string.Empty };
                _byItemId[itemId] = call;
            }
            call.Id = item.Value<string>("call_id") ?? call.Id;
            call.Name = item.Value<string>("name") ?? call.Name;
            // done 事件携带完整快照，必须替换增量，不能再次拼接。
            call.ArgumentsJson = item.Value<string>("arguments") ?? call.ArgumentsJson;
            return call;
        }

        internal override IEnumerable<ChatEvent> Flush()
        {
            foreach (var item in _byItemId)
            {
                var call = item.Value;
                if (!string.IsNullOrEmpty(call.Name) && _deliveredItems.Add(item.Key))
                {
                    yield return new ChatEvent { Kind = ChatEventKind.ToolCall, Call = call };
                }
            }

            _byItemId.Clear();
        }
    }
}
