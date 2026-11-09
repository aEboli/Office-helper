using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ChatSheet.AddIn.Agent;
using ChatSheet.AddIn.Providers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    /// <summary>
    /// 请求结构配平验证。
    ///
    /// 实测故障（v0.10.3.18）：文本指令协议下，助手历史消息同时带着正文里的
    /// 指令块和一份原生 tool_calls，而结果只能以 user 消息回灌。第二步请求于是
    /// 变成「有调用、没结果」——OpenAI 系回 400
    /// "An assistant message with 'tool_calls' must be followed by tool messages…"，
    /// 网关回 502 "No tool output found for function call fc_txt1-739998."。
    /// 坏消息留在历史里，之后每轮都带着它重发，换模型也照样失败。
    /// </summary>
    internal static class RequestShapeTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            Check(report, "文本指令协议的请求不出现 tool_calls", () =>
            {
                var request = NewRequest();
                var assistant = ChatMessage.FromAssistant(
                    "我先读一下表头。\n\n```" + TextToolProtocol.BlockTag + "\n" +
                    "{\"tool\": \"read_range\", \"args\": {\"range\": \"A1:B1\"}}\n" + "```");
                var call = new ToolCall { Id = "txt1-739998", Name = "read_range", ArgumentsJson = "{\"range\":\"A1:B1\"}" };
                assistant.ToolCalls.Add(call);
                request.Messages.Add(assistant);
                request.Messages.Add(ChatMessage.FromTextProtocolToolResult("read_range", "（工具执行结果 · read_range）\n{}"));

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];

                // 历史本身不能被改写：面板与撤销登记还按原对象走。
                return messages.Count == 2
                    && messages.All(m => m["tool_calls"] == null)
                    && messages.All(m => (string)m["role"] != "tool")
                    && assistant.ToolCalls.Count == 1
                    && ((string)messages[0]["content"]).Contains("chatsheet:tool");
            });

            // 这三条验的是「原生配对原样保留」，因此必须走声明了 tools 的请求：
            // NewRequest 走 IncludeTools = false，会先被历史降级命中，
            // 断言看到的是降级结果而不是 MessageRepair 的配对结果。
            Check(report, "原生调用与结果成对时原样保留", () =>
            {
                var request = NewNativeRequest();
                var assistant = ChatMessage.FromAssistant(string.Empty);
                assistant.ToolCalls.Add(new ToolCall { Id = "call_1", Name = "read_range", ArgumentsJson = "{\"range\":\"A1\"}" });
                request.Messages.Add(assistant);
                request.Messages.Add(ChatMessage.FromToolResult("call_1", "read_range", "{\"ok\":true}"));

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];
                var calls = messages[0]["tool_calls"] as JArray;
                return calls != null && calls.Count == 1
                    && (string)calls[0]["id"] == "call_1"
                    && (string)messages[1]["role"] == "tool"
                    && (string)messages[1]["tool_call_id"] == "call_1";
            });

            Check(report, "缺少结果的原生调用被摘掉", () =>
            {
                var request = NewRequest();
                var assistant = ChatMessage.FromAssistant("我读一下。");
                assistant.ToolCalls.Add(new ToolCall { Id = "txt1-739998", Name = "read_range", ArgumentsJson = "{}" });
                request.Messages.Add(assistant);

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];
                return messages.Count == 1
                    && messages[0]["tool_calls"] == null
                    && (string)messages[0]["content"] == "我读一下。";
            });

            Check(report, "缺少调用的工具结果被丢弃", () =>
            {
                var request = NewRequest();
                request.Messages.Add(ChatMessage.FromUser("你好"));
                request.Messages.Add(ChatMessage.FromToolResult("call_9", "read_range", "{}"));

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];
                return messages.Count == 1 && (string)messages[0]["role"] == "user";
            });

            Check(report, "部分配对时只保留配上结果的调用", () =>
            {
                var request = NewNativeRequest();
                var assistant = ChatMessage.FromAssistant(string.Empty);
                assistant.ToolCalls.Add(new ToolCall { Id = "call_1", Name = "read_range", ArgumentsJson = "{}" });
                assistant.ToolCalls.Add(new ToolCall { Id = "call_2", Name = "read_range", ArgumentsJson = "{}" });
                request.Messages.Add(assistant);
                request.Messages.Add(ChatMessage.FromToolResult("call_2", "read_range", "{}"));

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];
                var calls = messages[0]["tool_calls"] as JArray;
                return calls != null && calls.Count == 1
                    && (string)calls[0]["id"] == "call_2"
                    && messages.Count == 2
                    && (string)messages[1]["tool_call_id"] == "call_2";
            });

            Check(report, "调用标识为空时按工具名配对", () =>
            {
                var request = NewNativeRequest();
                var assistant = ChatMessage.FromAssistant(string.Empty);
                assistant.ToolCalls.Add(new ToolCall { Name = "read_range", ArgumentsJson = "{}" });
                request.Messages.Add(assistant);
                request.Messages.Add(ChatMessage.FromToolResult(null, "read_range", "{}"));

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];
                var calls = messages[0]["tool_calls"] as JArray;
                return calls != null && calls.Count == 1 && messages.Count == 2;
            });

            Check(report, "Anthropic 下被摘空的助手消息不留空正文块", () =>
            {
                var request = NewRequest();
                request.Protocol = ProtocolKind.AnthropicMessages;
                var assistant = ChatMessage.FromAssistant(string.Empty);
                assistant.ToolCalls.Add(new ToolCall { Id = "txt1-739998", Name = "read_range", ArgumentsJson = "{}" });
                request.Messages.Add(assistant);

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];

                // 正文可能以字符串或内容块数组两种形态发出，两种都要非空：
                // Anthropic 对空的 assistant content 直接回 400。
                var content = messages[0]["content"];
                var text = content is JArray blocks
                    ? string.Concat(blocks.Select(block => (string)block["text"] ?? string.Empty))
                    : (string)content;
                return !string.IsNullOrWhiteSpace(text)
                    && !body.ToString().Contains("tool_use");
            });

            Check(report, "修复不影响其余历史消息", () =>
            {
                var request = NewRequest();
                request.Messages.Add(ChatMessage.FromUser("原始提问"));
                var assistant = ChatMessage.FromAssistant("我读一下。");
                assistant.ToolCalls.Add(new ToolCall { Id = "call_1", Name = "read_range", ArgumentsJson = "{}" });
                request.Messages.Add(assistant);
                request.Messages.Add(ChatMessage.FromUser("继续"));

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];
                return messages.Count == 3
                    && (string)messages[0]["content"] == "原始提问"
                    && (string)messages[2]["content"] == "继续"
                    && messages[1]["tool_calls"] == null;
            });

            // ---- P0-1 回归：全配对里混进孤儿结果 ----
            //
            // 这是最常走、也最容易漏设防的分支：调用全都配得上时，修复逻辑
            // 只需把跟随的结果整段放行；孤儿一旦混在这段里就会被一起带进请求体，
            // 而 OpenAI 要求每条 role:"tool" 都回应紧邻那条 assistant 的某个
            // tool_call_id——多出来的那条让整轮 400。
            Check(report, "全部配对时混入的孤儿结果被丢弃", () =>
            {
                var request = NewNativeRequest();
                request.Messages.Add(ChatMessage.FromUser("读一下 A1"));
                var assistant = ChatMessage.FromAssistant(string.Empty);
                assistant.ToolCalls.Add(new ToolCall { Id = "call_1", Name = "read_range", ArgumentsJson = "{\"range\":\"A1\"}" });
                request.Messages.Add(assistant);
                request.Messages.Add(ChatMessage.FromToolResult("call_1", "read_range", "{\"ok\":true}"));
                // 重放的历史、被裁掉调用的残留、并发分支互相覆盖后的剩饭都长这样。
                request.Messages.Add(ChatMessage.FromToolResult("call_orphan", "read_range", "{\"ok\":true}"));

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];

                // 用户提问 + 助手调用 + 唯一配对的结果，与配对后的条数一致。
                return messages.Count == 3
                    && messages.Count(m => (string)m["role"] == "tool") == 1
                    && (string)messages[2]["tool_call_id"] == "call_1"
                    && messages.All(m => (string)m["tool_call_id"] != "call_orphan")
                    && (messages[1]["tool_calls"] as JArray)?.Count == 1
                    // 历史原对象不能被改写：面板的操作卡片与撤销按钮还按它登记。
                    && assistant.ToolCalls.Count == 1
                    && request.Messages.Count == 4;
            });

            // 孤儿结果的计数要如实报出来：日志里少报一条，排查时就会往别处找。
            Check(report, "孤儿结果计入 DroppedResults 且不改写历史", () =>
            {
                var request = NewNativeRequest();
                var assistant = ChatMessage.FromAssistant(string.Empty);
                assistant.ToolCalls.Add(new ToolCall { Id = "call_1", Name = "read_range", ArgumentsJson = "{}" });
                request.Messages.Add(assistant);
                request.Messages.Add(ChatMessage.FromToolResult("call_1", "read_range", "{}"));
                request.Messages.Add(ChatMessage.FromToolResult("call_orphan", "read_range", "{}"));

                var repaired = MessageRepair.Repair(request.Messages);
                return repaired.DroppedResults == 1
                    && repaired.DroppedCalls == 0
                    && repaired.Messages.Count == 2
                    && ReferenceEquals(repaired.Messages[0], assistant)
                    && assistant.ToolCalls.Count == 1;
            });

            // 直接盯住「记进历史」这一步：文本形态一旦把合成调用挂到助手消息上，
            // 下一条请求就带着有调用没结果的半边——这正是 502 的起点。
            // 出口处的修复能兜住它，但历史本身不该被写坏。
            foreach (var mode in new[] { ToolProtocolMode.Text, ToolProtocolMode.Native })
            {
                Check(report, mode + " 形态记录助手消息时的调用数", () =>
                {
                    var agent = new AgentRunner(() => null);
                    typeof(AgentRunner)
                        .GetField("_toolMode", BindingFlags.Instance | BindingFlags.NonPublic)
                        .SetValue(agent, mode);

                    var method = typeof(AgentRunner).GetMethod(
                        "BuildAssistantMessage", BindingFlags.Instance | BindingFlags.NonPublic);
                    var calls = new List<ToolCall>
                    {
                        new ToolCall { Id = "txt1-739998", Name = "read_range", ArgumentsJson = "{}" },
                    };
                    var message = (ChatMessage)method.Invoke(
                        agent,
                        new object[] { StepStall.None, "我读一下。", "我读一下。", calls });

                    var expected = mode == ToolProtocolMode.Text ? 0 : 1;
                    return message.ToolCalls.Count == expected
                        && message.Content == "我读一下。";
                });
            }

            RunProtocolSwitch(report);
        }

        private static ChatRequest NewRequest()
        {
            var request = new ChatRequest
            {
                Protocol = ProtocolKind.OpenAiChatCompletions,
                Model = "fixture",
                IncludeTools = false,
            };
            return request;
        }

        /// <summary>
        /// 声明了 tools 的请求：工具形态仍是原生。
        ///
        /// 需要单独一个构造：IncludeTools 为假时 RequestBuilder 会先把历史里的
        /// 原生调用与结果降级成纯文本，于是「配对原样保留」这类断言看到的是
        /// 降级结果，而它要验的是 MessageRepair 的配对本身。
        /// </summary>
        private static ChatRequest NewNativeRequest()
        {
            var request = NewRequest();
            request.IncludeTools = true;
            return request;
        }

        /// <summary>
        /// P2-8：「原生模式跑过两步后切 Text」的出口形状。
        ///
        /// 四种协议的写法各不相同，因此逐个协议验一遍：降级必须发生在序列化之前，
        /// 只对某一种协议生效是不够的。断言同时盯住两件事——请求体里一种工具痕迹
        /// 都不能剩，历史原对象一个字段都不能变。
        /// </summary>
        private static void RunProtocolSwitch(Action<string, bool, string> report)
        {
            foreach (var protocol in new[]
            {
                ProtocolKind.OpenAiChatCompletions,
                ProtocolKind.OpenAiResponses,
                ProtocolKind.AnthropicMessages,
                ProtocolKind.GoogleGemini,
            })
            {
                Check(report, protocol + " 下原生跑过两步后切 Text 不留工具痕迹", () =>
                {
                    // NewRequest 的 IncludeTools 为假，正是「切到 Text」的样子。
                    var request = NewRequest();
                    request.Protocol = protocol;
                    request.Messages.Add(ChatMessage.FromUser("把表头写一下"));

                    // 第一步（原生）：写表头并拿到结果。
                    var first = ChatMessage.FromAssistant("我先写表头。");
                    first.ToolCalls.Add(new ToolCall
                    {
                        Id = "call_step1",
                        Name = "write_values",
                        ArgumentsJson = "{\"range\":\"A1:B1\",\"values\":[[\"名称\",\"数量\"]]}",
                    });
                    request.Messages.Add(first);
                    var firstResult = ChatMessage.FromToolResult("call_step1", "write_values", "{\"cells_written\":2}");
                    request.Messages.Add(firstResult);

                    // 第二步（原生）：读回确认并拿到结果。历史里现在有两对。
                    var second = ChatMessage.FromAssistant("再读回来确认。");
                    second.ToolCalls.Add(new ToolCall
                    {
                        Id = "call_step2",
                        Name = "read_range",
                        ArgumentsJson = "{\"range\":\"A1:B1\"}",
                    });
                    request.Messages.Add(second);
                    var secondResult = ChatMessage.FromToolResult("call_step2", "read_range", "{\"values\":[[\"名称\",\"数量\"]]}");
                    request.Messages.Add(secondResult);

                    var json = RequestBuilder.Build(request, false).ToString(Formatting.None);
                    bool Has(string token) => json.IndexOf(token, StringComparison.Ordinal) >= 0;

                    // 六种字面量分属四种协议：role:"tool" 与 tool_calls 是
                    // Chat Completions 的，function_call(_output) 是 Responses 的，
                    // tool_use / tool_result 是 Anthropic 的，functionResponse 是
                    // Gemini 的。任一个漏出来都会被真实网关整轮拒绝。
                    var clean = !Has("\"tool\"")
                        && !Has("tool_calls")
                        && !Has("function_call")
                        && !Has("tool_use")
                        && !Has("tool_result")
                        && !Has("functionResponse")
                        // 切到 Text 之后本来就不该有 tools 声明，顺带确认前提成立。
                        && !Has("\"tools\"");

                    // 历史原对象不许被就地改写：面板的操作卡片与撤销按钮还按它们登记。
                    return clean
                        && ReferenceEquals(request.Messages[1], first)
                        && ReferenceEquals(request.Messages[2], firstResult)
                        && first.ToolCalls.Count == 1
                        && second.ToolCalls.Count == 1
                        && firstResult.Role == ChatRole.Tool
                        && secondResult.Role == ChatRole.Tool
                        && firstResult.ToolCallId == "call_step1"
                        && request.Messages.Count == 5
                        // 信息也不能丢：降级后的正文里仍能读到做过什么、拿到什么。
                        && Has("write_values")
                        && Has("read_range")
                        && Has("cells_written");
                });
            }

            // 降级只产出副本：Tool 角色的消息要变成普通 user 文本，原件不动。
            Check(report, "降级把结果消息改成 user 文本而不动原件", () =>
            {
                var request = NewRequest();
                var assistant = ChatMessage.FromAssistant("我读一下。");
                assistant.ToolCalls.Add(new ToolCall { Id = "call_1", Name = "read_range", ArgumentsJson = "{}" });
                request.Messages.Add(assistant);
                var result = ChatMessage.FromToolResult("call_1", "read_range", "{\"ok\":true}");
                request.Messages.Add(result);

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];

                return messages.Count == 2
                    && messages.All(m => (string)m["role"] != "tool")
                    && messages.All(m => m["tool_calls"] == null)
                    && result.Role == ChatRole.Tool
                    && result.Content == "{\"ok\":true}"
                    && assistant.ToolCalls.Count == 1;
            });

            // 相邻两次 Build 必须是同一结果：降级若就地改了历史，第二次会看到
            // 已经降级过的正文再被降级一遍，正文里叠出两段「此前的原生调用记录」。
            Check(report, "连续两次 Build 结果一致，降级不累积", () =>
            {
                var request = NewRequest();
                var assistant = ChatMessage.FromAssistant("我读一下。");
                assistant.ToolCalls.Add(new ToolCall { Id = "call_1", Name = "read_range", ArgumentsJson = "{}" });
                request.Messages.Add(assistant);
                request.Messages.Add(ChatMessage.FromToolResult("call_1", "read_range", "{}"));

                var once = RequestBuilder.Build(request, false).ToString(Formatting.None);
                var twice = RequestBuilder.Build(request, false).ToString(Formatting.None);
                return once == twice && request.Messages.Count == 2;
            });

            // mock 端的场景也是这条链路的验收物：场景被删掉或措辞改了，
            // 端到端验证会静默地什么都验不到，这里先把它钉住。
            Check(report, "mock 的 native-to-text 场景按网关口径拒绝裸工具痕迹", () =>
            {
                var root = FindRepoRoot();
                if (root == null)
                {
                    // 装机副本里没有仓库可读，这项无从核对，不算失败。
                    return true;
                }

                var mock = System.IO.File.ReadAllText(
                    System.IO.Path.Combine(root, "tests", "mock-provider", "server.mjs"));
                var e2e = System.IO.File.ReadAllText(
                    System.IO.Path.Combine(root, "scripts", "verify-chat-e2e.ps1"));

                // 措辞在 mock 里是拼接出来的，只取一段连续的片段来认。
                return mock.Contains("scenario === 'native-to-text'")
                    && mock.Contains("followed by tool messages responding to each 'tool_call_id'")
                    && mock.Contains("native tool trace")
                    && mock.Contains("the top-level 'tools' parameter")
                    && e2e.Contains("'native-to-text'");
            });
        }

        /// <summary>从测试产物目录向上找仓库根（认 ChatSheet.sln）。找不到返回 null。</summary>
        private static string FindRepoRoot()
        {
            var directory = new System.IO.DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            for (var depth = 0; depth < 8 && directory != null; depth++)
            {
                if (System.IO.File.Exists(System.IO.Path.Combine(directory.FullName, "ChatSheet.sln")))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            return null;
        }

        private static void Check(Action<string, bool, string> report, string name, Func<bool> check)
        {
            try { report(name, check(), null); }
            catch (Exception ex) { report(name, false, ex.GetBaseException().Message); }
        }
    }
}
