using System;
using System.Collections.Generic;
using System.Linq;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Tools;
using ChatWord.AddIn.Agent;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    /// <summary>
    /// Word 侧的请求形状验证。
    ///
    /// Excel 在 v0.10.3.20 修过「文本指令协议下助手消息同时带正文指令块和一份原生
    /// tool_calls」的故障，Word 当时没跟上：它无条件把调用挂到助手消息上，结果又恒用
    /// tool 角色回灌。WorkBuddy 走 ACP（纯文本协议），于是同一次调用在历史里出现两遍，
    /// 而真实网关遇到「有调用、没结果」会整轮拒绝，坏历史还会被每一轮重发。
    ///
    /// 这里直接盯两个构造入口，不碰 Word 宿主：形态分流一旦回退，断言立刻红。
    /// </summary>
    internal static class WordRequestShapeTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            Check(report, "Word 文本形态的助手消息不挂工具调用", () =>
            {
                var raw = "我先读一下。" + TextToolProtocol.BlockTag + " {tool: word_read, args: {}}";
                var message = WordAgentRunner.BuildStepMessage(raw, OneCall(), textProtocol: true);
                return message.ToolCalls.Count == 0
                    && message.Content == raw
                    && message.Content.Contains("chatsheet:tool");
            });

            Check(report, "Word 原生形态的助手消息保留工具调用", () =>
            {
                var message = WordAgentRunner.BuildStepMessage("我读一下。", OneCall(), textProtocol: false);
                return message.ToolCalls.Count == 1
                    && message.ToolCalls[0].Id == "word-call-1"
                    && message.ToolCalls[0].Name == "word_read";
            });

            // 文本协议没有 tool 角色可用：结果只能是 user 消息，但仍要保住「这是工具结果」
            // 的身份，否则上下文压缩会先压别的东西、模型也认不出这是自己那次调用的回执。
            Check(report, "Word 文本形态的工具结果以 user 消息回灌", () =>
            {
                var json = @"{""ok"": true}";
                var message = WordAgentRunner.BuildToolResultMessage(OneCall()[0], json, textProtocol: true);
                return message.Role == ChatRole.User
                    && message.IsTextProtocolToolResult
                    && message.ToolName == "word_read"
                    && message.Content.Contains(json)
                    && message.Content.Contains(@"（工具执行结果 · word_read）");
            });

            Check(report, "Word 原生形态的工具结果带调用标识", () =>
            {
                var message = WordAgentRunner.BuildToolResultMessage(OneCall()[0], @"{""ok"": true}", textProtocol: false);
                return message.Role == ChatRole.Tool
                    && message.ToolCallId == "word-call-1"
                    && message.ToolName == "word_read";
            });

            // 单看两个构造入口还不够：真发出去的是 RequestBuilder 序列化后的请求体。
            // 历史按文本协议攒一条完整往返，再走一遍请求构造，确认出口没有任何工具痕迹。
            Check(report, "Word 文本形态的整条历史过请求构造后不留工具痕迹", () =>
            {
                var request = new ChatRequest
                {
                    Protocol = ProtocolKind.OpenAiChatCompletions,
                    Model = "fixture",
                    IncludeTools = false,
                };
                var call = OneCall()[0];
                var raw = "我先读一下。" + TextToolProtocol.BlockTag + " {tool: word_read, args: {}}";
                request.Messages.Add(ChatMessage.FromUser("把开头改一下"));
                request.Messages.Add(WordAgentRunner.BuildStepMessage(raw, new[] { call }, textProtocol: true));
                request.Messages.Add(WordAgentRunner.BuildToolResultMessage(call, @"{""ok"": true}", textProtocol: true));

                var body = RequestBuilder.Build(request, false);
                var messages = (JArray)body["messages"];
                return messages.Count == 3
                    && messages.All(m => m["tool_calls"] == null)
                    && messages.All(m => (string)m["role"] != "tool")
                    && messages[1]["content"].ToString().Contains("chatsheet:tool");
            });

            // Schema 改成构造期缓存后，校验读的是同一份共享实例。行为必须与每次重新
            // 序列化完全一致，而且校验本身不能把这份实例改脏——它被所有工具、所有轮次共用。
            Check(report, "Schema 缓存后校验语义不变且实例不被改写", () =>
            {
                var definition = ToolCatalog.Find("read_range");
                if (definition == null || definition.Schema == null) { return false; }

                var before = definition.Schema.ToString(Formatting.None);

                var ok = definition.ValidateArguments(new JObject { ["range"] = "A1:B2" });
                var okAgain = definition.ValidateArguments(new JObject { ["range"] = "A1:B2" });

                var missing = definition.ValidateArguments(new JObject());
                var missingAgain = definition.ValidateArguments(new JObject());

                // 字符串数字与整数值浮点在执行层读得出来，校验层也一直放行；
                // 这条同时盯住 MatchesType 没有被顺手改严。
                var loose = definition.ValidateArguments(new JObject { ["range"] = "A1", ["offset"] = "12" });

                var wrongType = definition.ValidateArguments(new JObject { ["range"] = 1 });
                var wrongTypeAgain = definition.ValidateArguments(new JObject { ["range"] = 1 });

                var after = definition.Schema.ToString(Formatting.None);

                return ok == null && okAgain == null
                    && missing != null && missing.ErrorCode == "ARG_MISSING"
                    && missingAgain != null && missingAgain.ErrorCode == "ARG_MISSING"
                    && loose == null
                    && wrongType != null && wrongType.ErrorCode == "ARG_INVALID"
                    && wrongTypeAgain != null && wrongTypeAgain.ErrorCode == "ARG_INVALID"
                    && string.Equals(before, after, StringComparison.Ordinal);
            });
        }

        private static IReadOnlyList<ToolCall> OneCall()
        {
            return new List<ToolCall>
            {
                new ToolCall { Id = "word-call-1", Name = "word_read", ArgumentsJson = "{}" },
            };
        }

        private static void Check(Action<string, bool, string> report, string name, Func<bool> check)
        {
            try { report(name, check(), null); }
            catch (Exception ex) { report(name, false, ex.GetBaseException().Message); }
        }
    }
}
