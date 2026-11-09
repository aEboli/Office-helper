using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn.Agent;
using ChatSheet.AddIn.Hosts;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;
using ChatSheet.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class ProjectAuditTests
    {
        internal static void Run(object excel, Action<string, bool, string> report)
        {
            using (var full = new ResolvedRange(null, null, "test", "$A$1:$XFD$1048576", 1048576, 16384))
            {
                RangeResolver.AssertCellLimit(full, long.MaxValue, "测试");
                report("完整工作表计数不溢出且可执行", (long)full.CellCount == 17179869184L, full.CellCount.ToString());
            }

            object books = null, first = null, second = null;
            var wasVisible = Com.Get(excel, "Visible");
            var executor = new ToolExecutor(() => excel);
            try
            {
                Com.Set(excel, "Visible", true);
                books = Com.Get(excel, "Workbooks");
                first = Com.Call(books, "Add");
                RenameActive(excel);
                executor.Execute("write_values", JObject.Parse("{\"range\":\"A1:B1\",\"values\":[[11,12]]}"));
                var union = executor.Execute("clear_range", JObject.Parse("{\"range\":\"A1,B1\",\"scope\":\"contents\"}"));
                report("不连续范围被拒绝且两处内容均保留", !union.Ok && Read(executor, "A1") == "11" && Read(executor, "B1") == "12", union.ErrorCode);
                foreach (var tool in new[] { "create_table", "create_chart" })
                {
                    var oversized = executor.Execute(tool, JObject.Parse("{\"range\":\"A1:A5001\",\"chart_type\":\"line\"}"));
                    report(tool + " 不再因固定范围阈值拒绝", oversized.ErrorCode != "RANGE_TOO_LARGE", oversized.ErrorCode);
                }

                executor.Execute("write_values", JObject.Parse("{\"range\":\"C1\",\"values\":[[21]]}"));
                executor.Execute("write_values", JObject.Parse("{\"range\":\"C1\",\"values\":[[22]]}"), "audit-undo");
                second = Com.Call(books, "Add");
                RenameActive(excel);
                executor.Execute("write_values", JObject.Parse("{\"range\":\"C1\",\"values\":[[99]]}"), "audit-other");
                var undo = executor.Undo.Undo("audit-undo", true);
                report("跨工作簿撤销不覆盖同名工作表", !undo.Ok && Read(executor, "C1") == "99", undo.ErrorCode);
                Com.Call(first, "Activate");
                var originalUndo = executor.Undo.Undo("audit-undo");
                report("回原工作簿仍可撤销", originalUndo.Ok && Read(executor, "C1") == "21", originalUndo.ErrorCode + " " + originalUndo.Message + " value=" + Read(executor, "C1"));
                Com.Call(second, "Activate");
                var redo = executor.Undo.Redo("audit-undo");
                report("跨工作簿恢复不覆盖同名工作表", !redo.Ok && Read(executor, "C1") == "99", redo.ErrorCode);

                Com.Call(first, "Activate");
                var agent = new AgentRunner(() => excel);
                typeof(AgentRunner).GetField("_liveSettings", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(agent, new Settings { Approval = ApprovalPolicy.PerWrite });
                var call = new ToolCall { Id = "audit-approval", Name = "write_values", ArgumentsJson = "{\"range\":\"C1\",\"values\":[[88]]}" };
                var method = typeof(AgentRunner).GetMethod("ExecuteOneAsync", BindingFlags.Instance | BindingFlags.NonPublic);
                Func<AgentUpdate, Task> push = _ => Task.CompletedTask;
                Func<ToolDefinition, JObject, ImpactEstimate, Task<ApprovalDecision>> approve = (definition, args, impact) =>
                {
                    Com.Call(second, "Activate");
                    return Task.FromResult(new ApprovalDecision { Approved = true });
                };
                try
                {
                    var args = new object[] { call, new Settings { Approval = ApprovalPolicy.PerWrite }, push, approve, null };
                    ((Task)method.Invoke(agent, args)).GetAwaiter().GetResult();
                }
                catch (ProviderException ex) when (ex.Code == "WORKBOOK_CHANGED") { }
                report("审批期间切换工作簿不会误写新工作簿", Read(executor, "C1") == "99", Read(executor, "C1"));
                using (var cancellation = new CancellationTokenSource())
                {
                    typeof(AgentRunner).GetField("_turnCancellation", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(agent, cancellation.Token);
                    approve = (definition, args, impact) =>
                    {
                        cancellation.Cancel();
                        return Task.FromResult(new ApprovalDecision { Approved = true });
                    };
                    var cancelled = false;
                    try { ((Task)method.Invoke(agent, new object[] { call, new Settings(), push, approve, null })).GetAwaiter().GetResult(); }
                    catch (OperationCanceledException) { cancelled = true; }
                    report("审批返回时已停止，不再执行迟到写入", cancelled && Read(executor, "C1") == "99", null);
                }
                agent.Tools.Undo.Clear();
                var automatic = new AgentRunner(() => excel);
                var defaults = new Settings();
                typeof(AgentRunner).GetField("_liveSettings", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(automatic, defaults);
                var approvalCount = 0;
                approve = (definition, arguments, impact) => {
                    approvalCount++;
                    return Task.FromResult(new ApprovalDecision { Approved = false });
                };
                var objectCall = new ToolCall { Id = "audit-object-auto", Name = "excel_object",
                    ArgumentsJson = "{root:'worksheet',path:[{member:'Range',args:['D1']}],action:'set',member:'Value2',args:[77]}" };
                ((Task)method.Invoke(automatic, new object[] { objectCall, defaults, push, approve, null })).GetAwaiter().GetResult();
                report("默认自动执行通用对象工具且不请求审批", approvalCount == 0 && Read(executor, "D1") == "77", null);
                objectCall.ArgumentsJson = "{root:'worksheet',path:[],action:'call',member:'NonexistentMethod'}";
                var stoppedLoop = false;
                for (var i = 0; i < 8; i++)
                {
                    try { ((Task)method.Invoke(automatic, new object[] { objectCall, defaults, push, approve, null })).GetAwaiter().GetResult(); }
                    catch (ProviderException ex) when (ex.Code == "NO_PROGRESS") { stoppedLoop = i == 7; }
                }
                report("持续执行仍终止八次相同失败", stoppedLoop, null);
                automatic.Tools.Undo.Clear();
            }
            finally
            {
                executor.Undo.Clear();
                if (second != null) { Com.Call(second, "Close", false); }
                if (first != null) { Com.Call(first, "Close", false); }
                Com.Release(second);
                Com.Release(first);
                Com.Release(books);
                Com.Set(excel, "Visible", wasVisible);
            }
        }

        private static void RenameActive(object excel)
        {
            var sheet = Com.Get(excel, "ActiveSheet");
            try { Com.Set(sheet, "Name", "AuditTarget"); }
            finally { Com.Release(sheet); }
        }

        private static string Read(ToolExecutor executor, string address)
        {
            var read = executor.Execute("read_range", new JObject { ["range"] = address });
            return read.Ok ? JObject.FromObject(read.Data)["values"]?[0]?[0]?.ToString() : read.Error;
        }
    }
}
