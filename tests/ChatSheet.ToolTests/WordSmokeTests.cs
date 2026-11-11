using System;
using System.Globalization;
using System.Reflection;
using ChatWord.AddIn.Hosts;
using ChatWord.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class WordSmokeTests
    {
        internal static int Run(Action<string, bool, string> report, string progId)
        {
            object application = null; object documents = null; object document = null; object content = null; object selection = null; object range = null;
            try
            {
                var type = Type.GetTypeFromProgID(progId, false);
                if (type == null) { Console.WriteLine($"未执行 {progId}：未注册该 Word/WPS ProgID。"); return 2; }
                application = Activator.CreateInstance(type);
                Set(application, "Visible", false);
                documents = Get(application, "Documents");
                document = Call(documents, "Add");
                content = Get(document, "Content");
                try { Call(content, "InsertAfter", "smoke text"); }
                catch { Set(content, "Text", "smoke text\r"); }
                var executor = new WordToolExecutor(() => application, progId);
                var info = executor.Execute("get_document_info", new JObject());
                report(progId + " 读取文档信息", info.Ok, info.Error ?? "读取失败");
                var structure = executor.Execute("read_document_structure", new JObject());
                var structureData = structure.Ok ? JObject.FromObject(structure.Data) : new JObject();
                report(progId + " 读取可用文档结构", structure.Ok && structureData.Value<bool?>("ok") == true,
                    structure.Error ?? "结构读取失败");

                selection = Get(application, "Selection");
                Call(selection, "SetRange", 0, 10);
                var selected = executor.Execute("get_selection", new JObject());
                report(progId + " 读取当前选区", selected.Ok, selected.Error ?? "读取失败");

                range = Get(document, "Content");
                var start = Convert.ToInt32(Get(range, "Start"), CultureInfo.InvariantCulture);
                var end = Convert.ToInt32(Get(range, "End"), CultureInfo.InvariantCulture);
                var read = executor.Execute("read_range", new JObject
                {
                    ["target"] = new JObject { ["story"] = "main_text", ["start"] = start, ["end"] = end },
                });
                report(progId + " 读取正文范围", read.Ok, read.Error ?? "读取失败");
                var replace = executor.Execute("replace_text", new JObject
                {
                    ["target"] = new JObject { ["story"] = "main_text", ["start"] = start, ["end"] = end },
                    ["replace"] = "smoke replacement",
                }, "smoke-undo");
                report(progId + " 修改并读回", replace.Ok && JObject.FromObject(replace.Data).Value<string>("text") != null, replace.Error ?? "修改失败");
                var undone = executor.Undo.TryUndo("smoke-undo", document, out var undoMessage);
                report(progId + " 快照撤销", undone, undoMessage);

                var insertion = executor.Execute("insert_text", new JObject
                {
                    ["target"] = new JObject
                    {
                        ["story"] = "main_text", ["start"] = 0, ["end"] = 0,
                        ["paragraph_index"] = 0, ["table_index"] = 0, ["row"] = 0, ["column"] = 0,
                        ["bookmark"] = "", ["content_control"] = "",
                    },
                    ["text"] = "smoke insertion",
                }, "smoke-insert");
                report(progId + " 折叠光标插入并读回", insertion.Ok &&
                    JObject.FromObject(insertion.Data).Value<string>("text") == "smoke insertion", insertion.Error ?? "插入失败");
                var insertionUndone = executor.Undo.TryUndo("smoke-insert", document, out var insertionUndoMessage);
                report(progId + " 撤销折叠光标插入", insertionUndone, insertionUndoMessage);
                ToolRegressions(document, executor, progId, report);
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"未执行 {progId}：{ex}");
                return 2;
            }
            finally
            {
                try { if (document != null) { Call(document, "Close", false); } } catch { }
                try { if (application != null) { Call(application, "Quit", false); } } catch { }
                Release(range); Release(selection); Release(content); Release(document); Release(documents); Release(application);
            }
        }

        private static void ToolRegressions(object document, WordToolExecutor executor, string progId, Action<string, bool, string> report)
        {
            object tables = null; object table = null; object cell = null; object cellRange = null; object position = null;
            try
            {
                tables = Get(document, "Tables");
                position = Call(document, "Range", 0, 0);
                table = Call(tables, "Add", position, 1, 1);
                cell = Call(table, "Cell", 1, 1);
                cellRange = Get(cell, "Range");
                Set(cellRange, "Text", "原内容");
                var args = new JObject
                {
                    ["target"] = new JObject { ["story"] = "main_text", ["table_index"] = 1, ["row"] = 1, ["column"] = 1 },
                    ["text"] = "更新内容",
                };
                var edited = executor.Execute("edit_table", args, "smoke-table");
                var editedText = CellText(cell);
                report(progId + " 嵌套表格定位替换已有内容", edited.Ok &&
                    editedText == "更新内容", edited.Error ?? "实际内容：" + editedText);
                var undone = executor.Undo.TryUndo("smoke-table", document, out var undoMessage);
                report(progId + " 表格文字撤销恢复原内容", undone &&
                    CellText(cell) == "原内容", undoMessage);

                args.Remove("text");
                var missing = executor.Execute("edit_table", args);
                report(progId + " 表格缺少正文不清空内容", !missing.Ok && missing.ErrorCode == "ARG_MISSING" &&
                    CellText(cell) == "原内容", missing.Error);
                args["text"] = string.Empty;
                var cleared = executor.Execute("edit_table", args);
                var clearedText = CellText(cell);
                report(progId + " 表格显式空文本可以清空", cleared.Ok &&
                    clearedText == string.Empty, cleared.Error ?? "实际内容：" + clearedText);
            }
            finally { Release(position); Release(cellRange); Release(cell); Release(table); Release(tables); }

            var pageBreak = executor.Execute("insert_page_break", new JObject
            {
                ["target"] = new JObject { ["story"] = "main_text", ["start"] = 0, ["end"] = 0 },
            });
            report(progId + " 分页符插入后准确读回", pageBreak.Ok &&
                JObject.FromObject(pageBreak.Data).Value<string>("readback") == "\f", pageBreak.Error);
        }

        private static string CellText(object cell)
        {
            var range = Get(cell, "Range");
            try { return Convert.ToString(Get(range, "Text")).TrimEnd('\r', '\a'); }
            finally { Release(range); }
        }

        private static object Get(object target, string name, params object[] args) => target.GetType().InvokeMember(name, BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public, null, target, args, CultureInfo.GetCultureInfo("en-US"));
        private static object Call(object target, string name, params object[] args) => target.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public, null, target, args, CultureInfo.GetCultureInfo("en-US"));
        private static void Set(object target, string name, object value) => target.GetType().InvokeMember(name, BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public, null, target, new[] { value }, CultureInfo.GetCultureInfo("en-US"));
        private static void Release(object target) { try { if (target != null && System.Runtime.InteropServices.Marshal.IsComObject(target)) { System.Runtime.InteropServices.Marshal.ReleaseComObject(target); } } catch { } }
    }
}
