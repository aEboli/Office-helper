using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn.Providers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class ProviderToolRegressionTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            Check(report, "SSE 表情跨读取边界后仍可容纳整批字符", () =>
            {
                var text = new string('a', 8183) + "\U0001F680" + new string('b', 8191);
                var frames = new List<SseFrame>();
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("data: " + text + "\n\n")))
                {
                    SseReader.ReadAsync(stream, frame =>
                    {
                        frames.Add(frame);
                        return Task.FromResult(true);
                    }, CancellationToken.None).GetAwaiter().GetResult();
                }
                return frames.Count == 1 && frames[0].Data == text;
            });

            Check(report, "Chat finish_reason 与 DONE 只完成一次且保留末帧用量", () =>
            {
                var parser = StreamParser.Create(ProtocolKind.OpenAiChatCompletions);
                var events = Parse(parser, ChatDelta(new JObject(), "length"),
                    JObject.Parse("{'choices':[],'usage':{'prompt_tokens':9,'completion_tokens':4}}"));
                events.AddRange(parser.Parse(new SseFrame { Data = "[DONE]" }));
                events.AddRange(parser.Parse(new SseFrame { Data = "[DONE]" }));
                events.AddRange(parser.Flush());
                var ends = events.Where(e => e.Kind == ChatEventKind.Completed).ToList();
                return ends.Count == 1 && ends[0].FinishReason == "length"
                    && events.Any(e => e.Kind == ChatEventKind.Usage && e.PromptTokens == 9 && e.CompletionTokens == 4);
            });

            Check(report, "Chat 交错工具增量按索引交付且参数不串台", () =>
            {
                var calls = Calls(Parse(StreamParser.Create(ProtocolKind.OpenAiChatCompletions),
                    ChatCall(1, "second", "read_range", "{\"range\":"),
                    ChatCall(0, "first", "read_range", "{\"range\":\"A1\"}"),
                    ChatCall(1, null, null, "\"B2\"}"), ChatDelta(new JObject(), "tool_calls")));
                return calls.Count == 2 && calls[0].Id == "first" && calls[1].Id == "second"
                    && calls[0].ArgumentsJson == "{\"range\":\"A1\"}"
                    && calls[1].ArgumentsJson == "{\"range\":\"B2\"}";
            });

            Check(report, "Responses 重复终态不重复完成或交付工具", () =>
            {
                var done = JObject.Parse("{'type':'response.completed','response':{'output':[{'type':'function_call','id':'item_1','call_id':'call_1','name':'get_selection','arguments':'{}'}]}}");
                var parser = StreamParser.Create(ProtocolKind.OpenAiResponses);
                var events = Parse(parser, done, done);
                events.AddRange(parser.Flush());
                return Calls(events).Count == 1 && events.Count(e => e.Kind == ChatEventKind.Completed) == 1;
            });

            Check(report, "Anthropic 无增量时保留 content_block_start 参数", () =>
            {
                var calls = Calls(Parse(StreamParser.Create(ProtocolKind.AnthropicMessages),
                    AnthropicStart(JObject.Parse("{'range':'A1'}")), AnthropicStop()));
                return calls.Count == 1 && JObject.Parse(calls[0].ArgumentsJson).Value<string>("range") == "A1";
            });
            Check(report, "Anthropic 首个参数增量替换起始对象并继续拼接", () =>
            {
                var calls = Calls(Parse(StreamParser.Create(ProtocolKind.AnthropicMessages),
                    AnthropicStart(new JObject()),
                    AnthropicDelta("{\"range\":"), AnthropicDelta("\"B2\"}"), AnthropicStop()));
                return calls.Count == 1 && calls[0].ArgumentsJson == "{\"range\":\"B2\"}";
            });
            Check(report, "Anthropic 空参数对象在流结束时保留", () =>
            {
                var parser = StreamParser.Create(ProtocolKind.AnthropicMessages);
                Parse(parser, AnthropicStart(new JObject()));
                var calls = Calls(parser.Flush());
                return calls.Count == 1 && calls[0].ArgumentsJson == "{}" && !parser.Flush().Any();
            });
            Check(report, "Anthropic 重复 message_stop 只完成一次", () =>
            {
                var stop = JObject.Parse("{'type':'message_stop'}");
                return Parse(StreamParser.Create(ProtocolKind.AnthropicMessages), stop, stop)
                    .Count(e => e.Kind == ChatEventKind.Completed) == 1;
            });

            Check(report, "Gemini 保留服务端调用 ID 且重复终态不重交工具", () =>
            {
                var frame = JObject.Parse("{'candidates':[{'content':{'parts':[{'functionCall':{'id':'server_call_1','name':'read_range','args':{'range':'A1'}}}]},'finishReason':'STOP'}]}");
                var events = Parse(StreamParser.Create(ProtocolKind.GoogleGemini), frame, frame);
                var calls = Calls(events);
                return calls.Count == 1 && calls[0].Id == "server_call_1"
                    && events.Count(e => e.Kind == ChatEventKind.Completed) == 1;
            });

            Check(report, "Anthropic 并行结果合并为紧邻调用的一条 user 消息", () =>
            {
                var request = ParallelRequest(ProtocolKind.AnthropicMessages);
                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];
                var results = (JArray)messages[2]["content"];
                return messages.Count == 3 && results.Count == 2
                    && results[0].Value<string>("tool_use_id") == "call_a"
                    && results[1].Value<string>("tool_use_id") == "call_b"
                    && request.Messages.Count == 4 && request.Messages[1].ToolCalls.Count == 2;
            });
            Check(report, "Gemini 并行结果同轮返回并保留两侧调用 ID", () =>
            {
                var request = ParallelRequest(ProtocolKind.GoogleGemini);
                var body = RequestBuilder.Build(request, false);
                var contents = (JArray)body["contents"];
                var calls = (JArray)contents[1]["parts"];
                var results = (JArray)contents[2]["parts"];
                return contents.Count == 3 && calls.Count == 2 && results.Count == 2
                    && calls[0]["functionCall"].Value<string>("id") == "call_a"
                    && calls[1]["functionCall"].Value<string>("id") == "call_b"
                    && results[0]["functionResponse"].Value<string>("id") == "call_a"
                    && results[1]["functionResponse"].Value<string>("id") == "call_b"
                    && results[0]["functionResponse"].Value<string>("name") == "read_range"
                    && request.Messages.Count == 4 && request.Messages[1].ToolCalls[0].Id == "call_a";
            });

            Check(report, "文本闸门在每个分段位置都保留行内示例", () =>
                CheckEverySplit("这是示例：```chatsheet:tool\n{\"tool\":\"get_selection\",\"args\":{}}\n```\n结束。", 0));
            Check(report, "文本闸门在每个分段位置都吞掉完整指令和闭合行", () =>
                CheckEverySplit("前文\n```chatsheet:tool\n{\"tool\":\"get_selection\",\"args\":{}}\n```\n后文", 1, "前文\n后文"));
            Check(report, "文本闸门的长围栏示例不执行内层指令", () =>
                CheckEverySplit("````text\n```chatsheet:tool\n{\"tool\":\"get_selection\",\"args\":{}}\n```\n````\n", 0));
            Check(report, "文本闸门不把带正文的反引号行当作闭合", () =>
                CheckEverySplit("```text\n普通代码\n```这不是闭合行\n后续代码\n```\n", 0));
        }

        private static bool CheckEverySplit(string source, int expectedCalls, string expectedText = null)
        {
            for (var split = 0; split <= source.Length; split++)
            {
                var gate = new TextToolGate();
                var text = gate.Push(source.Substring(0, split));
                text += gate.Push(source.Substring(split));
                text += gate.Flush();
                if (text != (expectedText ?? source) || gate.Calls.Count != expectedCalls)
                {
                    throw new InvalidOperationException($"分段位置 {split}，调用 {gate.Calls.Count}，正文 {JsonConvert.SerializeObject(text)}");
                }
            }
            return true;
        }

        private static ChatRequest ParallelRequest(ProtocolKind protocol)
        {
            var request = new ChatRequest { Protocol = protocol, Model = "fixture", IncludeTools = true, SuppressThinking = true };
            request.Messages.Add(ChatMessage.FromUser("读取两处"));
            var assistant = ChatMessage.FromAssistant(string.Empty);
            assistant.ToolCalls.Add(new ToolCall { Id = "call_a", Name = "read_range", ArgumentsJson = "{\"range\":\"A1\"}" });
            assistant.ToolCalls.Add(new ToolCall { Id = "call_b", Name = "read_range", ArgumentsJson = "{\"range\":\"B2\"}" });
            request.Messages.Add(assistant);
            request.Messages.Add(ChatMessage.FromToolResult("call_a", "read_range", "{\"value\":1}"));
            request.Messages.Add(ChatMessage.FromToolResult("call_b", "read_range", "{\"value\":2}"));
            return request;
        }

        private static List<ChatEvent> Parse(StreamParser parser, params JObject[] frames)
        {
            var events = new List<ChatEvent>();
            foreach (var frame in frames)
            {
                events.AddRange(parser.Parse(new SseFrame { Data = frame.ToString(Formatting.None) }));
            }
            return events;
        }

        private static List<ToolCall> Calls(IEnumerable<ChatEvent> events) =>
            events.Where(e => e.Kind == ChatEventKind.ToolCall).Select(e => e.Call).ToList();

        private static JObject ChatDelta(JObject delta, string finish = null) => new JObject
        {
            ["choices"] = new JArray(new JObject { ["delta"] = delta, ["finish_reason"] = finish }),
        };

        private static JObject ChatCall(int index, string id, string name, string arguments) => ChatDelta(new JObject
        {
            ["tool_calls"] = new JArray(new JObject
            {
                ["index"] = index, ["id"] = id,
                ["function"] = new JObject { ["name"] = name, ["arguments"] = arguments },
            }),
        });

        private static JObject AnthropicStart(JObject input) => new JObject
        {
            ["type"] = "content_block_start", ["index"] = 0,
            ["content_block"] = new JObject { ["type"] = "tool_use", ["id"] = "tool_1", ["name"] = "read_range", ["input"] = input },
        };

        private static JObject AnthropicDelta(string partial) => new JObject
        {
            ["type"] = "content_block_delta", ["index"] = 0,
            ["delta"] = new JObject { ["type"] = "input_json_delta", ["partial_json"] = partial },
        };

        private static JObject AnthropicStop() => JObject.Parse("{'type':'content_block_stop','index':0}");

        private static void Check(Action<string, bool, string> report, string name, Func<bool> assertion)
        {
            try { report(name, assertion(), string.Empty); }
            catch (Exception ex) { report(name, false, ex.Message); }
        }
    }
}
