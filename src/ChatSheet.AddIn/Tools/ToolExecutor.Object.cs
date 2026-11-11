using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ChatSheet.AddIn.Hosts;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Tools
{
    internal sealed partial class ToolExecutor
    {
        // 通用入口仅遍历当前工作簿对象，不提供应用程序、宏和外部文件入口。
        private static readonly HashSet<string> ExternalMembers = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Application", "Parent", "Creator", "VBProject", "VBE", "Run", "ExecuteExcel4Macro",
            "Evaluate", "_Evaluate", "DDEInitiate", "DDEExecute", "DDEPoke", "DDERequest",
            "OnAction", "OnTime", "OnKey", "Save", "SaveAs", "SaveCopyAs", "Open", "Close",
            "Quit", "Export", "ExportAsFixedFormat", "PrintOut", "SendMail", "FollowHyperlink",
            "Connections", "QueryTables", "Queries", "PublishObjects", "OLEObjects", "RTD",
            "UpdateLink", "ChangeLink", "XmlImport", "XmlImportXml", "WebOptions"
        };

        private ToolResult ExcelObject(JObject args)
        {
            var held = new List<object>();
            var started = false;
            try
            {
                // 在任何可能修改对象的路径调用前，验证所有成员和对象参数。
                ValidateObjectRequest(args);
                started = true;
                var target = ResolveObject(args, held);
                var action = RequireString(args, "action").ToLowerInvariant();
                var member = RequireString(args, "member");
                var parameters = ObjectArguments(args["args"] as JArray, held);
                object value = null;
                WithoutDisplayAlerts(() =>
                {
                    if (action == "get") { value = Com.Get(target, member, parameters); }
                    else if (action == "set") { Com.Set(target, member, parameters); }
                    else { value = Com.Call(target, member, parameters); }
                });
                if (value != null && Marshal.IsComObject(value)) { held.Add(value); }
                return ToolResult.Success(new Dictionary<string, object>
                {
                    ["root"] = args.Value<string>("root"), ["sheet"] = args.Value<string>("sheet"),
                    ["action"] = action, ["member"] = member,
                    ["value"] = ObjectValue(value),
                    ["verification_required"] = action != "get",
                    ["undo_note"] = "通用对象操作不提供自动撤销；修改后请读取目标状态确认。"
                });
            }
            catch (Exception ex)
            {
                var error = Unwrap(ex);
                return ToolResult.Failure(error is ToolException te ? te.Code : "HOST_ERROR", error.Message,
                    started ? new { operation_started = true,
                        note = "对象路径或方法可能已修改部分状态，请先读回确认，不要盲目重试。" } : null);
            }
            finally
            {
                for (var i = held.Count - 1; i >= 0; i--) { Com.Release(held[i]); }
            }
        }

        private static void ValidateObjectRequest(JObject request, bool reference = false)
        {
            var root = RequireString(request, "root");
            if (root != "workbook" && root != "worksheet" && root != "window")
                throw new ToolException("ARG_INVALID", "root 应为 workbook、worksheet 或 window。");
            if (!(request["path"] is JArray path))
                throw new ToolException("ARG_INVALID", "path 必须为数组，可为空。");
            foreach (var item in path)
            {
                if (!(item is JObject step)) { throw new ToolException("ARG_INVALID", "path 元素必须为对象。"); }
                CheckObjectMember(RequireString(step, "member"));
                var kind = OptionalString(step, "kind") ?? "get";
                if (kind != "get" && kind != "call") { throw new ToolException("ARG_INVALID", "path kind 应为 get 或 call。"); }
                ValidateObjectArguments(step["args"]);
            }
            if (!reference)
            {
                CheckObjectMember(RequireString(request, "member"));
                var action = RequireString(request, "action");
                if (action != "get" && action != "set" && action != "call")
                    throw new ToolException("ARG_INVALID", "action 应为 get、set 或 call。");
                ValidateObjectArguments(request["args"]);
            }
        }

        private static void CheckObjectMember(string member)
        {
            if (ExternalMembers.Contains(member) || !System.Text.RegularExpressions.Regex.IsMatch(member, "^[A-Za-z][A-Za-z0-9_]*$"))
                throw new ToolException("WORKBOOK_SCOPE", "该入口仅操作当前工作簿的表格对象，不能调用成员 " + member + "。");
        }

        private static void ValidateObjectArguments(JToken args)
        {
            if (args == null || args.Type == JTokenType.Null) { return; }
            if (!(args is JArray array)) { throw new ToolException("ARG_INVALID", "args 必须为数组。"); }
            foreach (var token in array)
            {
                if (token is JArray nested) { ValidateObjectArguments(nested); }
                else if (token is JObject obj)
                {
                    if (OptionalBool(obj, "missing") == true) { continue; }
                    if (obj["ref"] is JObject reference) { ValidateObjectRequest(reference, true); }
                    else if (obj["matrix"] is JArray matrix)
                    {
                        var columns = (matrix.First as JArray)?.Count ?? 0;
                        foreach (var item in matrix)
                        {
                            if (!(item is JArray row) || row.Count != columns)
                                throw new ToolException("SHAPE_MISMATCH", "matrix 行列必须一致。");
                            foreach (var cell in row) { ObjectScalar(cell); }
                        }
                    }
                    else { throw new ToolException("ARG_INVALID", "对象参数应为 ref、matrix 或 missing。"); }
                }
                else { ObjectScalar(token); }
            }
        }

        private object ResolveObject(JObject request, List<object> held)
        {
            object current;
            switch (RequireString(request, "root"))
            {
                case "worksheet": current = _resolver.ResolveWorksheet(OptionalString(request, "sheet")); break;
                case "window": current = Com.Get(Application, "ActiveWindow"); break;
                default: current = Com.Get(Application, "ActiveWorkbook"); break;
            }
            if (current == null) { throw new ToolException("NO_WORKBOOK", "没有可操作的工作簿或窗口。"); }
            held.Add(current);
            foreach (JObject step in (JArray)request["path"])
            {
                var parameters = ObjectArguments(step["args"] as JArray, held);
                current = OptionalString(step, "kind") == "call"
                    ? Com.Call(current, RequireString(step, "member"), parameters)
                    : Com.Get(current, RequireString(step, "member"), parameters);
                if (current == null) { throw new ToolException("OBJECT_NOT_FOUND", "对象路径没有返回对象。"); }
                if (Marshal.IsComObject(current)) { held.Add(current); }
            }
            return current;
        }

        private object[] ObjectArguments(JArray args, List<object> held)
        {
            if (args == null) { return new object[0]; }
            var values = new object[args.Count];
            for (var i = 0; i < args.Count; i++)
            {
                var token = args[i];
                if (token is JObject obj)
                {
                    if (OptionalBool(obj, "missing") == true) { values[i] = Type.Missing; }
                    else if (obj["ref"] is JObject reference) { values[i] = new DispatchWrapper(ResolveObject(reference, held)); }
                    else if (obj["matrix"] is JArray matrix)
                    {
                        var cols = (matrix.Count > 0 ? matrix[0] as JArray : null)?.Count ?? 0;
                        var buffer = new object[matrix.Count, cols];
                        for (var r = 0; r < matrix.Count; r++)
                        {
                            if (!(matrix[r] is JArray row) || row.Count != cols) { throw new ToolException("SHAPE_MISMATCH", "matrix 行列必须一致。"); }
                            for (var c = 0; c < cols; c++) { buffer[r, c] = ObjectScalar(row[c]); }
                        }
                        values[i] = buffer;
                    }
                    else { throw new ToolException("ARG_INVALID", "对象参数应为 ref、matrix 或 missing。"); }
                }
                else if (token is JArray array) { values[i] = ObjectArguments(array, held); }
                else { values[i] = ObjectScalar(token); }
            }
            return values;
        }

        private static object ObjectScalar(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null) { return null; }
            if (token is JContainer) { throw new ToolException("ARG_INVALID", "单元格参数必须是标量。"); }
            if (token.Type == JTokenType.Integer)
            {
                try
                {
                    var value = token.Value<long>();
                    return value >= int.MinValue && value <= int.MaxValue ? (object)(int)value : value;
                }
                catch (OverflowException)
                {
                    throw new ToolException("ARG_INVALID", "整数参数超出 Int64 范围：" + token + "。");
                }
            }
            if (token.Type == JTokenType.Float) { return token.Value<double>(); }
            if (token.Type == JTokenType.Boolean) { return token.Value<bool>(); }
            return token.Value<string>();
        }

        private static object ObjectValue(object value)
        {
            if (value != null && Marshal.IsComObject(value))
                return new { object_returned = true, name = Com.GetString(value, "Name"), address = Com.GetString(value, "Address") };
            if (value is Array array)
            {
                var result = new List<object>();
                foreach (var cell in array)
                {
                    if (result.Count == ToolLimits.ReadPageCells) { break; }
                    result.Add(Normalize(cell));
                }
                return new { values = result, total = array.LongLength, truncated = array.LongLength > result.Count };
            }
            return Normalize(value);
        }
    }
}
