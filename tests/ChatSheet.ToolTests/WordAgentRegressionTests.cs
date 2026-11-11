using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn.Agent;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;
using ChatSheet.AddIn.Tools;
using ChatWord.AddIn.Agent;
using ChatWord.AddIn.Bridge;
using ChatWord.AddIn.Hosts;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class WordAgentRegressionTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            Check(report, "Word 非法 JSON 不进入宿主并回传参数错误", () =>
                RejectsArguments("get_document_info", "not-json", "ARGS_INVALID"));
            Check(report, "Word 截断参数不按空参数执行", () =>
                RejectsArguments("get_document_info", "{\"target\":", "ARGS_TRUNCATED"));
            Check(report, "Word 数组参数不按空对象执行", () =>
                RejectsArguments("get_document_info", "[]", "ARGS_INVALID"));
            Check(report, "Word 草稿同样拒绝非法 JSON", () =>
                RejectsArguments("draft_document", "not-json", "ARGS_INVALID"));
            Check(report, "Word 缺失工具名不导致整轮异常", () =>
                RejectsArguments(null, "{}", "BLOCK_INVALID"));
            Check(report, "Word 缺失替换正文在审批前拒绝", () =>
                RejectsArguments("replace_text", "{\"target\":{\"story\":\"main_text\",\"start\":0,\"end\":1}}", "ARG_MISSING"));
            Check(report, "Word 错误目标类型在审批前拒绝", () =>
                RejectsArguments("insert_text", "{\"target\":[],\"text\":\"正文\"}", "ARG_INVALID"));

            foreach (var block in new[] { "{\"type\":\"paragraph\",\"text\":{}}", "{\"type\":\"heading\",\"text\":\"标题\",\"level\":\"abc\"}" })
            {
                Check(report, "Word 草稿嵌套参数错误可纠正：" + block, () =>
                {
                    var summary = new WordDocumentSummary { HasDocument = true, Name = "fixture" };
                    var runner = new WordAgentRunner(() => null, "Word.Application", work => Task.FromResult<object>(summary));
                    var result = Execute(runner, "draft_document", "{\"operation\":\"insert\",\"blocks\":[" + block + "]}", CancellationToken.None,
                        documentKey: WordDocumentContext.DocumentKey(summary),
                        draftTarget: new WordSelectionInfo { HasSelection = true, Story = "main_text", Start = 0, End = 0 });
                    return !result.Ok && result.ErrorCode == "DRAFT_INVALID";
                });
            }

            Check(report, "Word 已停止的调用不触发审批或宿主", () =>
            {
                using (var cancellation = new CancellationTokenSource())
                {
                    cancellation.Cancel();
                    var dispatches = 0;
                    var runner = new WordAgentRunner(() => null, "Word.Application", work =>
                    {
                        dispatches++;
                        return Task.FromResult<object>(ToolResult.Success(new { }));
                    });
                    return IsCancelled(() => Execute(runner, "get_document_info", "{}", cancellation.Token))
                        && dispatches == 0;
                }
            });

            Check(report, "Word 审批返回时已停止不会继续写入", () =>
            {
                using (var cancellation = new CancellationTokenSource())
                {
                    var dispatches = 0;
                    var approvals = 0;
                    var runner = new WordAgentRunner(() => null, "Word.Application", work =>
                    {
                        dispatches++;
                        if (dispatches != 1) { return Task.FromResult<object>(ToolResult.Success(new { })); }
                        var previewType = typeof(WordAgentRunner).GetNestedType("ApprovalPreview", BindingFlags.NonPublic);
                        var preview = Activator.CreateInstance(previewType, true);
                        previewType.GetField("DocumentKey", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(preview, "fixture-document");
                        return Task.FromResult(preview);
                    });
                    var cancelled = IsCancelled(() => Execute(runner, "insert_text",
                        "{\"target\":{\"story\":\"main_text\",\"start\":0,\"end\":0},\"text\":\"不可写入\"}",
                        cancellation.Token, ApprovalPolicy.PerWrite, (definition, args, impact) =>
                        {
                            approvals++;
                            cancellation.Cancel();
                            return Task.FromResult(new ApprovalDecision { Approved = true });
                        }));
                    return cancelled && approvals == 1 && dispatches == 1;
                }
            });

            Check(report, "Word 排队的 UI 调用在停止后不访问宿主", () =>
            {
                using (var cancellation = new CancellationTokenSource())
                {
                    var hostAccesses = 0;
                    var runner = new WordAgentRunner(() => { hostAccesses++; return null; }, "Word.Application", work =>
                    {
                        cancellation.Cancel();
                        return Task.FromResult(work());
                    });
                    return IsCancelled(() => Execute(runner, "get_document_info", "{}", cancellation.Token))
                        && hostAccesses == 0;
                }
            });

            Check(report, "Word UI 关闭产生的取消不伪装为宿主错误", () =>
            {
                var runner = new WordAgentRunner(() => null, "Word.Application", work =>
                    Task.FromException<object>(new OperationCanceledException("面板已关闭")));
                return IsCancelled(() => Execute(runner, "get_document_info", "{}", CancellationToken.None));
            });

            Check(report, "Word 有效空参数仍正常执行", () =>
            {
                var dispatches = 0;
                var runner = new WordAgentRunner(() => null, "Word.Application", work =>
                {
                    dispatches++;
                    return Task.FromResult<object>(ToolResult.Success(new { }));
                });
                var result = Execute(runner, "get_document_info", "{}", CancellationToken.None);
                return result.Ok && dispatches == 1;
            });

            Check(report, "Word 面板关闭后拒绝已排队的宿主回调", () =>
            {
                var type = typeof(WordBridge);
                var bridge = (WordBridge)FormatterServices.GetUninitializedObject(type);
                var context = new QueuedContext();
                type.GetField("_uiContext", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(bridge, context);
                var accessed = false;
                var method = type.GetMethod("InvokeOnUiAsync", BindingFlags.Instance | BindingFlags.NonPublic);
                var pending = (Task<object>)method.Invoke(bridge, new object[]
                {
                    new Func<object>(() => { accessed = true; return null; }),
                });
                type.GetField("_disposed", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(bridge, true);
                context.Run();
                return IsCancelled(() => pending.GetAwaiter().GetResult()) && !accessed;
            });

            Check(report, "Word 等待审批时停止能收束当前任务", () => ApprovalStops(false));
            Check(report, "Word 停止与审批创建相撞也不会永久等待", () => ApprovalStops(true));
            Check(report, "Word 折叠光标的 null 文本视为空文本", () =>
                WordCom.ReadText(new NullTextRange { Start = 3, End = 3 }) == string.Empty);
            Check(report, "Word 非空范围的 null 文本仍视为读回失败", () =>
            {
                try { WordCom.ReadText(new NullTextRange { Start = 3, End = 4 }); return false; }
                catch (WordToolException ex) { return ex.Code == "READBACK_FAILED"; }
            });
        }

        private sealed class NullTextRange
        {
            public string Text => null;
            public int Start { get; set; }
            public int End { get; set; }
        }

        private static bool ApprovalStops(bool cancelledBeforeRequest)
        {
            using (var cancellation = new CancellationTokenSource())
            {
                var type = typeof(WordBridge);
                var bridge = (WordBridge)FormatterServices.GetUninitializedObject(type);
                var approvals = new ConcurrentDictionary<string, TaskCompletionSource<ApprovalDecision>>();
                type.GetField("_approvals", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(bridge, approvals);
                type.GetField("_run", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(bridge, cancellation);
                if (cancelledBeforeRequest) { cancellation.Cancel(); }
                var method = type.GetMethod("RequestApprovalAsync", BindingFlags.Instance | BindingFlags.NonPublic);
                var pending = (Task<ApprovalDecision>)method.Invoke(bridge, new object[]
                {
                    new ToolDefinition("test", "测试审批", ToolRisk.Write, new { type = "object" }), new JObject(), null,
                });
                if (!cancelledBeforeRequest) { type.GetMethod("Stop", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(bridge, null); }
                if (Task.WhenAny(pending, Task.Delay(1000)).GetAwaiter().GetResult() != pending)
                {
                    foreach (var source in approvals.Values) { source.TrySetCanceled(); }
                    return false;
                }
                return IsCancelled(() => pending.GetAwaiter().GetResult()) && approvals.IsEmpty;
            }
        }

        private sealed class QueuedContext : SynchronizationContext
        {
            private SendOrPostCallback _callback;
            private object _state;
            public override void Post(SendOrPostCallback callback, object state) { _callback = callback; _state = state; }
            internal void Run() { _callback(_state); }
        }

        private static bool RejectsArguments(string name, string json, string code)
        {
            var dispatches = 0;
            var updates = new List<AgentUpdate>();
            var runner = new WordAgentRunner(() => null, "Word.Application", work =>
            {
                dispatches++;
                return Task.FromResult<object>(ToolResult.Success(new { }));
            });
            var result = Execute(runner, name, json, CancellationToken.None, updates: updates);
            return !result.Ok && result.ErrorCode == code && dispatches == 0
                && updates.Any(update => update.Kind == "tool-result")
                && runner.Conversation.Messages.Last().Content.Contains(code);
        }

        private static ToolResult Execute(WordAgentRunner runner, string name, string json, CancellationToken token,
            ApprovalPolicy policy = ApprovalPolicy.Automatic,
            Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> approve = null,
            List<AgentUpdate> updates = null, string documentKey = "fixture-document", WordSelectionInfo draftTarget = null)
        {
            var method = typeof(WordAgentRunner).GetMethod("ExecuteCallAsync", BindingFlags.Instance | BindingFlags.NonPublic);
            var task = (Task<ToolResult>)method.Invoke(runner, new object[]
            {
                new ToolCall { Id = "word-regression", Name = name, ArgumentsJson = json },
                new Settings { Approval = policy }, documentKey, draftTarget,
                new Func<AgentUpdate, Task>(update => { updates?.Add(update); return Task.CompletedTask; }),
                approve ?? new Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>>(
                    (definition, args, impact) => throw new Exception("不应进入审批")),
                token, null, true,
            });
            return task.GetAwaiter().GetResult();
        }

        private static bool IsCancelled(Action action)
        {
            try { action(); return false; }
            catch (OperationCanceledException) { return true; }
        }

        private static void Check(Action<string, bool, string> report, string name, Func<bool> check)
        {
            try { report(name, check(), null); }
            catch (Exception ex) { report(name, false, ex.GetBaseException().Message); }
        }
    }
}
