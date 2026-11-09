using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using ChatSheet.AddIn.Agent;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;
using ChatSheet.AddIn.Tools;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class AgentToolTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            Check(report, "Responses 使用 arguments.done 的完整参数", () =>
                Calls(Added(""), Delta("{\"range\":"), ArgumentsDone(), ItemDone(null))
                    .Single().ArgumentsJson == "{\"range\":\"A1\"}");
            Check(report, "Responses 使用 output_item.done 的完整参数", () =>
                Calls(Added(""), ItemDone("{\"range\":\"A1\"}"))
                    .Single().ArgumentsJson == "{\"range\":\"A1\"}");
            Check(report, "Responses 保留 added 中已有的参数", () =>
                Calls(Added("{\"range\":\"A1\"}"), ItemDone(null))
                    .Single().ArgumentsJson == "{\"range\":\"A1\"}");
            Check(report, "Responses 完整 done 快照无需 added 也能恢复调用", () =>
                Calls(ItemDone("{\"range\":\"A1\"}"))
                    .Single().Id == "call_1");
            Check(report, "Responses 增量和完整快照不会重复调用", () =>
                Calls(Added(""), Delta("{\"range\":\"A1\"}"), ArgumentsDone(),
                    ItemDone("{\"range\":\"A1\"}"), ItemDone("{\"range\":\"A1\"}"))
                    .Count == 1);
            Check(report, "Responses completed 快照恢复参数且不重复交付", () =>
                Calls(Added(""), new JObject { ["type"] = "response.completed", ["response"] = new JObject
                { ["output"] = new JArray(Item("{\"range\":\"A1\"}")) } })
                    .Single().ArgumentsJson == "{\"range\":\"A1\"}");
            Check(report, "Responses failed 返回真实错误且不交付待执行调用", () =>
            {
                var parser = StreamParser.Create(ProtocolKind.OpenAiResponses);
                parser.Parse(Frame(Added("{\"range\":\"A1\"}"))).ToList();
                try
                {
                    parser.Parse(Frame(new JObject
                    {
                        ["type"] = "response.failed",
                        ["response"] = new JObject { ["error"] = new JObject
                        { ["message"] = "fixture failure", ["code"] = "server_error" } },
                    })).ToList();
                    return false;
                }
                catch (ProviderException ex)
                {
                    return ex.Message.Contains("fixture failure") && !parser.Flush().Any();
                }
            });

            foreach (var body in new[] { "{\"tool\":{}}", "{\"args\":{}}", "{\"tool\":[]}" })
            {
                Check(report, "错误工具名作为无效指令回传 " + body, () =>
                    TextToolProtocol.TryParseBlockBody(TextToolProtocol.BlockTag, body, out var call)
                    && string.IsNullOrEmpty(call.Name));
                Check(report, "普通 JSON 的非字符串名称不触发工具 " + body, () =>
                    !TextToolProtocol.TryParseBlockBody("json", body, out _));
            }
            Check(report, "文本工具清单包含对象路径参数约定", () =>
                TextToolProtocol.CatalogText().Contains("kind=get/call")
                && TextToolProtocol.CatalogText().Contains("path:array"));
            Check(report, "合法可选空值保持有效", () =>
                ToolCatalog.Find("fit_range").ValidateArguments(JObject.Parse("{\"range\":null,\"sheet\":null}")) == null);
            Check(report, "合法二维写入和布尔参数保持有效", () =>
                ToolCatalog.Find("write_values").ValidateArguments(JObject.Parse("{\"range\":\"A1\",\"values\":[[1]]}")) == null
                && ToolCatalog.Find("read_range").ValidateArguments(JObject.Parse("{\"range\":\"A1\",\"include_formulas\":true,\"offset\":0}")) == null);

            foreach (var mode in new[] { ToolProtocolMode.Native, ToolProtocolMode.Text })
            {
                Check(report, mode + " 缺少 range 时回传参数定义", () =>
                    InvalidCall(mode, "read_range", "{}", "ARG_MISSING"));
                Check(report, mode + " range 类型错误不使任务崩溃", () =>
                    InvalidCall(mode, "fit_range", "{\"range\":{}}", "ARG_INVALID"));
                Check(report, mode + " root 类型错误不使任务崩溃", () =>
                    InvalidCall(mode, "excel_object", "{\"root\":{},\"path\":[],\"action\":\"get\",\"member\":\"Name\"}", "ARG_INVALID"));
                Check(report, mode + " path 类型错误在审批前返回", () =>
                    InvalidCall(mode, "excel_object", "{\"root\":\"worksheet\",\"path\":\"Range\",\"action\":\"get\",\"member\":\"Name\"}", "ARG_INVALID"));
                Check(report, mode + " 非对象参数可纠正", () =>
                    InvalidCall(mode, "read_range", "[]", "ARGS_INVALID"));
            }

            // ---- P0-2 回归：参数校验的宽容度必须与执行层同一口径 ----
            //
            // 校验层比执行层更严的代价是整轮白烧：工具里读的是 args.Value<long?>()
            // 与 Value<bool>()，Newtonsoft 会先做一次字符串转换，"0" 与 "true"
            // 在执行层本来就读得出来。校验却按 schema 的 type 字面比对，
            // 于是模型只是换了个等价写法就被 ARG_INVALID 弹回去，而错误里说的
            // 「必须为 integer」又和工具实际行为不符——它只会把同一个值反复改写。

            // 形状正确的一律放行。
            Check(report, "read_range 的 offset 接受字符串 \"0\"", () =>
                Accepts("read_range", "{\"range\":\"A1\",\"offset\":\"0\"}"));
            Check(report, "read_range 的 include_formulas 接受 null", () =>
                Accepts("read_range", "{\"range\":\"A1\",\"include_formulas\":null}"));
            Check(report, "sort_range 的 ascending 接受字符串 \"true\"", () =>
                Accepts("sort_range", "{\"range\":\"A1:B2\",\"key_column\":\"A\",\"ascending\":\"true\"}"));
            Check(report, "read_range 的 offset 接受无损浮点 12.0", () =>
                Accepts("read_range", "{\"range\":\"A1\",\"offset\":12.0}"));

            // 反例：放宽容错不能把真正的类型错误也吞掉。丢精度的浮点
            // 与不是数字的字符串在执行层读出来是 0 / 异常，必须在入口挡住。
            Check(report, "read_range 的 offset 传 \"abc\" 仍报 ARG_INVALID", () =>
                Rejects("read_range", "{\"range\":\"A1\",\"offset\":\"abc\"}", "ARG_INVALID"));
            Check(report, "read_range 的 offset 传 12.5 仍报 ARG_INVALID", () =>
                Rejects("read_range", "{\"range\":\"A1\",\"offset\":12.5}", "ARG_INVALID"));
            Check(report, "write_values 的 values 传字符串仍报 ARG_INVALID", () =>
                Rejects("write_values", "{\"range\":\"A1\",\"values\":\"wrong\"}", "ARG_INVALID"));
            Check(report, "sort_range 的 ascending 传 \"yes\" 仍报 ARG_INVALID", () =>
                Rejects("sort_range", "{\"range\":\"A1:B2\",\"key_column\":\"A\",\"ascending\":\"yes\"}", "ARG_INVALID"));

            // 必填项写 null 仍算缺失：这条是「非必填 null 视为未提供」的边界。
            // 放行必填 null 会让工具拿到 null 后在更深的地方崩掉，报出的错误
            // 与真实原因毫无关系。
            Check(report, "read_range 的 range 传 null 仍报 ARG_MISSING", () =>
                Rejects("read_range", "{\"range\":null}", "ARG_MISSING"));
            Check(report, "write_values 的 values 传 null 仍报 ARG_MISSING", () =>
                Rejects("write_values", "{\"range\":\"A1\",\"values\":null}", "ARG_MISSING"));

            // ---- excel_object 的两段式校验 ----
            //
            // root 缺省时要先由宿主弹窗补上，因此第一遍带 skip 集合只验其余参数：
            // 模型少写 path 却先弹了窗，用户白点一次才收到 ARG_MISSING，
            // 而那次点击换不来任何结果。补完 root 再全量验一次。
            Check(report, "excel_object 跳过 root 时仍校验 path", () =>
            {
                var definition = ToolCatalog.Find("excel_object");
                var partial = definition.ValidateArguments(
                    JObject.Parse("{\"action\":\"get\",\"member\":\"Name\"}"),
                    new[] { "root" });
                return partial != null && partial.ErrorCode == "ARG_MISSING" && partial.Error.Contains("path");
            });
            Check(report, "excel_object 补上 root 后全量校验通过", () =>
                Accepts("excel_object",
                    "{\"root\":\"worksheet\",\"path\":[],\"action\":\"get\",\"member\":\"Name\"}"));
            Check(report, "excel_object 缺 root 时全量校验报 ARG_MISSING", () =>
                Rejects("excel_object", "{\"path\":[],\"action\":\"get\",\"member\":\"Name\"}", "ARG_MISSING"));
        }

        internal static void RunInHost(object excel, Action<string, bool, string> report)
        {
            foreach (var mode in new[] { ToolProtocolMode.Native, ToolProtocolMode.Text })
            {
                Check(report, mode + " 参数纠错后可实际写入并读回", () =>
                {
                    var agent = new AgentRunner(() => excel);
                    var settings = new Settings { Approval = ApprovalPolicy.Automatic };
                    typeof(AgentRunner).GetField("_toolMode", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(agent, mode);
                    typeof(AgentRunner).GetField("_liveSettings", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(agent, settings);
                    var method = typeof(AgentRunner).GetMethod("ExecuteOneAsync", BindingFlags.Instance | BindingFlags.NonPublic);
                    Func<AgentUpdate, Task> push = _ => Task.CompletedTask;
                    var call = new ToolCall { Id = "fixture-invalid", Name = "write_values", ArgumentsJson = "{\"range\":\"Z1\",\"values\":\"wrong\"}" };
                    try
                    {
                        ((Task)method.Invoke(agent, new object[] { call, settings, push, null, null })).GetAwaiter().GetResult();
                        if (!agent.Conversation.Messages.Last().Content.Contains("ARG_INVALID")) { return false; }
                        call = new ToolCall { Id = "fixture-repaired", Name = "write_values", ArgumentsJson = "{\"range\":\"Z1\",\"values\":[[73]]}" };
                        ((Task)method.Invoke(agent, new object[] { call, settings, push, null, null })).GetAwaiter().GetResult();
                        var read = agent.Tools.Execute("read_range", JObject.Parse("{\"range\":\"Z1\"}"));
                        return read.Ok && JObject.FromObject(read.Data)["values"][0][0].Value<int>() == 73;
                    }
                    finally { agent.Tools.Undo.Clear(); }
                });
            }
        }

        private static void Check(Action<string, bool, string> report, string name, Func<bool> check)
        {
            try { report(name, check(), null); }
            catch (Exception ex) { report(name, false, ex.GetBaseException().Message); }
        }

        /// <summary>校验通过：返回 null 而不是失败结果。</summary>
        private static bool Accepts(string tool, string args)
        {
            return ToolCatalog.Find(tool).ValidateArguments(JObject.Parse(args)) == null;
        }

        /// <summary>校验失败且错误码相符。</summary>
        private static bool Rejects(string tool, string args, string code)
        {
            var result = ToolCatalog.Find(tool).ValidateArguments(JObject.Parse(args));
            return result != null && !result.Ok && result.ErrorCode == code;
        }

        private static bool InvalidCall(ToolProtocolMode mode, string name, string args, string code)
        {
            var agent = new AgentRunner(() => null);
            typeof(AgentRunner).GetField("_toolMode", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(agent, mode);
            var call = new ToolCall { Id = "fixture", Name = name, ArgumentsJson = args };
            var method = typeof(AgentRunner).GetMethod("ExecuteOneCoreAsync", BindingFlags.Instance | BindingFlags.NonPublic);
            var updates = new List<AgentUpdate>();
            ((Task)method.Invoke(agent, new object[]
            {
                call, new Settings(),
                new Func<AgentUpdate, Task>(update => { updates.Add(update); return Task.CompletedTask; }),
                new Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>>(
                    (definition, parameters, impact) => throw new Exception("不应进入审批")), null,
            })).GetAwaiter().GetResult();
            var result = agent.Conversation.Messages.Last().Content;
            var payload = JObject.Parse(result.Substring(result.IndexOf('{')));
            return payload.Value<string>("error_code") == code
                && payload["parameters"] is JObject
                && updates.Any(u => u.Kind == "tool-result");
        }

        private static SseFrame Frame(JObject data) => new SseFrame { Data = data.ToString(Formatting.None) };

        private static List<ToolCall> Calls(params JObject[] frames)
        {
            var parser = StreamParser.Create(ProtocolKind.OpenAiResponses);
            var events = frames.SelectMany(frame => parser.Parse(Frame(frame))).ToList();
            events.AddRange(parser.Flush());
            return events.Where(e => e.Kind == ChatEventKind.ToolCall).Select(e => e.Call).ToList();
        }

        private static JObject Item(string args)
        {
            var item = new JObject { ["type"] = "function_call", ["id"] = "item_1", ["call_id"] = "call_1", ["name"] = "read_range" };
            if (args != null) { item["arguments"] = args; }
            return item;
        }

        private static JObject Added(string args) => new JObject { ["type"] = "response.output_item.added", ["item"] = Item(args) };
        private static JObject ItemDone(string args) => new JObject { ["type"] = "response.output_item.done", ["item"] = Item(args) };
        private static JObject Delta(string args) => new JObject { ["type"] = "response.function_call_arguments.delta", ["item_id"] = "item_1", ["delta"] = args };
        private static JObject ArgumentsDone() => new JObject { ["type"] = "response.function_call_arguments.done", ["item_id"] = "item_1", ["arguments"] = "{\"range\":\"A1\"}" };
    }
}
