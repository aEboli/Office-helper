using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Providers
{
    /// <summary>
    /// Anthropic Messages 的 SSE 解析。
    ///
    /// 其流式结构以内容块为单位：content_block_start 声明块类型，
    /// content_block_delta 下发增量，content_block_stop 结束该块。
    /// 工具调用的参数走 input_json_delta，按块索引累积。
    /// </summary>
    internal sealed class AnthropicStreamParser : StreamParser
    {
        private readonly Dictionary<int, string> _blockTypes = new Dictionary<int, string>();
        private readonly HashSet<int> _inputDeltas = new HashSet<int>();
        private int _promptTokens;

        /// <summary>
        /// message_delta 里的 stop_reason，留到 message_stop 时一并上报。
        ///
        /// 两个事件是分开的：结束原因只在 message_delta 出现，而结束事件在
        /// message_stop。不记住的话上层永远收到空的结束原因，
        /// 也就分不出「说完了」和「max_tokens 截断」。
        /// </summary>
        private string _stopReason;

        internal override IEnumerable<ChatEvent> Parse(SseFrame frame)
        {
            if (Completed) { yield break; }
            var root = TryParse(frame.Data);
            if (root == null) { yield break; }

            var type = root.Value<string>("type") ?? frame.EventName ?? string.Empty;

            switch (type)
            {
                case "message_start":
                {
                    // 输入 token 只在开始事件里出现一次，需要记住到结束时一并上报。
                    var usage = (root["message"] as JObject)?["usage"] as JObject;
                    _promptTokens = usage?.Value<int?>("input_tokens") ?? 0;
                    break;
                }

                case "content_block_start":
                {
                    var index = root.Value<int?>("index") ?? 0;
                    var block = root["content_block"] as JObject;
                    var blockType = block?.Value<string>("type") ?? string.Empty;
                    _blockTypes[index] = blockType;

                    if (string.Equals(blockType, "tool_use", StringComparison.Ordinal))
                    {
                        _inputDeltas.Remove(index);
                        PendingCalls[index] = new ToolCall
                        {
                            Id = block.Value<string>("id"),
                            Name = block.Value<string>("name"),
                            ArgumentsJson = (block["input"] ?? new JObject()).ToString(Newtonsoft.Json.Formatting.None),
                        };
                    }

                    break;
                }

                case "content_block_delta":
                {
                    var index = root.Value<int?>("index") ?? 0;
                    var delta = root["delta"] as JObject;
                    var deltaType = delta?.Value<string>("type") ?? string.Empty;

                    if (string.Equals(deltaType, "text_delta", StringComparison.Ordinal))
                    {
                        var text = delta.Value<string>("text");
                        if (!string.IsNullOrEmpty(text))
                        {
                            yield return new ChatEvent { Kind = ChatEventKind.TextDelta, Text = text };
                        }
                    }
                    else if (string.Equals(deltaType, "thinking_delta", StringComparison.Ordinal))
                    {
                        var thinking = delta.Value<string>("thinking");
                        if (!string.IsNullOrEmpty(thinking))
                        {
                            yield return new ChatEvent { Kind = ChatEventKind.ThinkingDelta, Text = thinking };
                        }
                    }
                    else if (string.Equals(deltaType, "input_json_delta", StringComparison.Ordinal))
                    {
                        var partial = delta.Value<string>("partial_json");
                        if (!string.IsNullOrEmpty(partial) && PendingCalls.TryGetValue(index, out var call))
                        {
                            // start 中的 input 是初始对象；首个增量从头构造参数，不能拼在 {} 后面。
                            call.ArgumentsJson = _inputDeltas.Add(index) ? partial : call.ArgumentsJson + partial;
                        }
                    }

                    break;
                }

                case "content_block_stop":
                {
                    var index = root.Value<int?>("index") ?? 0;
                    if (PendingCalls.TryGetValue(index, out var call))
                    {
                        PendingCalls.Remove(index);
                        if (!string.IsNullOrEmpty(call.Name))
                        {
                            yield return new ChatEvent { Kind = ChatEventKind.ToolCall, Call = call };
                        }
                    }

                    _blockTypes.Remove(index);
                    _inputDeltas.Remove(index);
                    break;
                }

                case "message_delta":
                {
                    var stop = (root["delta"] as JObject)?.Value<string>("stop_reason");
                    if (!string.IsNullOrEmpty(stop)) { _stopReason = stop; }

                    var usage = root["usage"] as JObject;
                    if (usage != null)
                    {
                        yield return new ChatEvent
                        {
                            Kind = ChatEventKind.Usage,
                            PromptTokens = _promptTokens,
                            CompletionTokens = usage.Value<int?>("output_tokens") ?? 0,
                        };
                    }

                    break;
                }

                case "message_stop":
                {
                    Completed = true;
                    foreach (var e in Flush()) { yield return e; }
                    yield return new ChatEvent { Kind = ChatEventKind.Completed, FinishReason = _stopReason };
                    break;
                }

                case "error":
                {
                    var error = root["error"] as JObject;
                    var message = error?.Value<string>("message") ?? "服务端返回错误";
                    throw new ProviderException("STREAM_ERROR", message);
                }
            }
        }
    }

    /// <summary>
    /// Gemini 的流式解析。
    ///
    /// Gemini 一次下发完整的候选片段（不是字符级增量），
    /// 工具调用也是完整对象，因此无需累积拼接。
    /// 思考内容通过 part.thought 标记区分。
    /// </summary>
    internal sealed class GeminiStreamParser : StreamParser
    {
        internal override IEnumerable<ChatEvent> Parse(SseFrame frame)
        {
            var root = TryParse(frame.Data);
            if (root == null) { yield break; }

            var error = root["error"] as JObject;
            if (error != null)
            {
                throw new ProviderException("STREAM_ERROR", error.Value<string>("message") ?? "服务端返回错误");
            }

            var usage = root["usageMetadata"] as JObject;
            if (usage != null)
            {
                yield return new ChatEvent
                {
                    Kind = ChatEventKind.Usage,
                    PromptTokens = usage.Value<int?>("promptTokenCount") ?? 0,
                    CompletionTokens = usage.Value<int?>("candidatesTokenCount") ?? 0,
                };
            }

            if (Completed || !(root["candidates"] is JArray candidates) || candidates.Count == 0)
            {
                yield break;
            }

            var candidate = candidates[0] as JObject;
            var parts = (candidate?["content"] as JObject)?["parts"] as JArray;

            if (parts != null)
            {
                foreach (var item in parts)
                {
                    if (!(item is JObject part)) { continue; }

                    var functionCall = part["functionCall"] as JObject;
                    if (functionCall != null)
                    {
                        yield return new ChatEvent
                        {
                            Kind = ChatEventKind.ToolCall,
                            Call = new ToolCall
                            {
                                // 优先保留服务端标识，只有旧响应未提供时才生成本地标识。
                                Id = functionCall.Value<string>("id") ??
                                    (functionCall.Value<string>("name") ?? "call") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8),
                                Name = functionCall.Value<string>("name"),
                                ArgumentsJson = (functionCall["args"] ?? new JObject()).ToString(Newtonsoft.Json.Formatting.None),
                            },
                        };
                        continue;
                    }

                    var text = part.Value<string>("text");
                    if (string.IsNullOrEmpty(text)) { continue; }

                    // thought 为真表示这段是思考过程而非正文。
                    var isThought = part.Value<bool?>("thought") ?? false;
                    yield return new ChatEvent
                    {
                        Kind = isThought ? ChatEventKind.ThinkingDelta : ChatEventKind.TextDelta,
                        Text = text,
                    };
                }
            }

            var finish = candidate?.Value<string>("finishReason");
            if (!string.IsNullOrEmpty(finish))
            {
                Completed = true;
                yield return new ChatEvent { Kind = ChatEventKind.Completed, FinishReason = finish };
            }
        }
    }
}
