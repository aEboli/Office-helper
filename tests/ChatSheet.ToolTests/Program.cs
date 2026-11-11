using System;
using System.Runtime.InteropServices;
using System.Text;
using ChatSheet.AddIn.Tools;
using ChatSheet.AddIn.Storage;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    /// <summary>
    /// 工具层真实验证。用 COM 起一个 Excel 实例，逐个执行工具并检查结果。
    /// 构建通过只能证明类型正确，唯有真实调用才能验证 COM 序列正确。
    /// </summary>
    internal static class Program
    {
        private static int _passed;
        private static int _failed;

        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = new UTF8Encoding(false);
            // ACP 标准输入必须明确使用无 BOM 的 UTF-8；Windows 默认输入编码会把
            // BOM 写进重定向管道，真实 CLI 通常会容忍，但协议 fixture 应严格校验。
            Console.InputEncoding = new UTF8Encoding(false);
            if (Array.IndexOf(args, "--agent-tool-tests") >= 0)
            {
                StreamTests.Run(ReportProvider);
                CapabilityTests.Run(ReportProvider);
                AgentToolTests.Run(ReportProvider);
                RequestShapeTests.Run(ReportProvider);
                WordRequestShapeTests.Run(ReportProvider);
                WordAgentRegressionTests.Run(ReportProvider);
                ExcelToolRegressionTests.Run(ReportProvider);
                WordToolRegressionTests.Run(ReportProvider);
                ProviderToolRegressionTests.Run(ReportProvider);
                Console.WriteLine($"=== Agent 工具回归：通过 {_passed}，失败 {_failed} ===");
                return _failed == 0 ? 0 : 1;
            }
            if (Array.IndexOf(args, "--reliability-tests") >= 0)
            {
                ReliabilityTests.Run(ReportProvider);
                Console.WriteLine($"=== 可靠性回归：通过 {_passed}，失败 {_failed} ===");
                return _failed == 0 ? 0 : 1;
            }
            if (Array.IndexOf(args, "--codex-config-tests") >= 0)
            {
                return CodexConfigTests.Run();
            }
            if (Array.IndexOf(args, "--codex-config-live") >= 0)
            {
                return CodexConfigTests.RunLive();
            }
            if (args.Length > 0 && (args[0] == "--workbuddy-install" || args[0] == "--workbuddy-login-live" || args[0] == "--workbuddy-login-live-intl" || args[0] == "--workbuddy-fallback-live" || args[0] == "--workbuddy-fallback-live-intl" || args[0] == "--workbuddy-account-live" || args[0] == "--workbuddy-account-live-intl" || args[0] == "--workbuddy-account-tests"))
            {
                return WorkBuddyAccountTests.Run(args[0]);
            }
            if (Array.IndexOf(args, "--acp") >= 0)
            {
                if (System.IO.Path.GetFileName(args[0]).StartsWith("login-", StringComparison.Ordinal)) { return WorkBuddyLoginTests.RunFixture(args[0]); }
                return WorkBuddyChatTests.RunFixture(args[0]);
            }
            if (Array.IndexOf(args, "--workbuddy-login-tests") >= 0) { return WorkBuddyLoginTests.Run(); }
            if (Array.IndexOf(args, "--workbuddy-live") >= 0)
            {
                var status = WorkBuddyChatTests.RunLive(ReportProvider);
                return status == 2 ? 2 : (_failed == 0 && status == 0 ? 0 : 1);
            }
            if (Array.IndexOf(args, "--workbuddy-live-intl") >= 0)
            {
                var status = WorkBuddyChatTests.RunLive(ConnectionMode.AuthorizedInternational, ReportProvider);
                // 2 = 组件存在但尚未登录；1 = 实机回归失败；0 = 全部通过。
                return status == 2 ? 2 : (_failed == 0 && status == 0 ? 0 : 1);
            }
            if (Array.IndexOf(args, "--workbuddy-tests") >= 0)
            {
                WorkBuddyChatTests.Run(ReportProvider);
                return _failed == 0 ? 0 : 1;
            }

            if (Array.IndexOf(args, "--word-tests") >= 0)
            {
                WordTests.Run(ReportProvider);
                WordRequestShapeTests.Run(ReportProvider);
                WordAgentRegressionTests.Run(ReportProvider);
                WordToolRegressionTests.Run(ReportProvider);
                Console.WriteLine($"=== Word 工具测试：通过 {_passed}，失败 {_failed} ===");
                return _failed == 0 ? 0 : 1;
            }
            if (Array.IndexOf(args, "--word-smoke") >= 0 || Array.IndexOf(args, "--wps-word-smoke") >= 0)
            {
                var progId = Array.IndexOf(args, "--wps-word-smoke") >= 0 ? "kwps.Application" : "Word.Application";
                var status = WordSmokeTests.Run(ReportProvider, progId);
                return status == 2 ? 2 : (_failed == 0 ? 0 : 1);
            }

            object excel = null;
            object workbooks = null;
            object workbook = null;

            try
            {
                Console.WriteLine("启动 Excel 实例…");
                var type = Type.GetTypeFromProgID(Array.IndexOf(args, "--wps") >= 0 ? "ket.Application" : "Excel.Application", throwOnError: true);
                excel = Activator.CreateInstance(type);
                Set(excel, "Visible", false);
                Set(excel, "DisplayAlerts", false);

                workbooks = Get(excel, "Workbooks");
                workbook = Call(workbooks, "Add");

                var executor = new ToolExecutor(() => excel);
                if (Array.IndexOf(args, "--open-operations-tests") >= 0)
                {
                    OpenOperationsTests.Run(executor, ReportProvider);
                    Console.WriteLine($"=== 开放操作：通过 {_passed}，失败 {_failed} ===");
                    return _failed == 0 ? 0 : 1;
                }
            if (Array.IndexOf(args, "--project-audit-tests") >= 0)
                {
                    AgentToolTests.RunInHost(excel, ReportProvider);
                    ProjectAuditTests.Run(excel, ReportProvider);
                    Console.WriteLine($"=== 项目审查：通过 {_passed}，失败 {_failed} ===");
                return _failed == 0 ? 0 : 1;
            }
            RunAll(executor);

                Console.WriteLine();
                Console.WriteLine("=== 撤销与恢复 ===");
                UndoTests.Run(excel, executor, ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 撤销不说谎 ===");
                HonestUndoTests.Run(excel, executor, ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 接入层 ===");
                ProviderTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 思考参数映射 ===");
                ThinkingTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 图片输入 ===");
                ImageTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 文件附件 ===");
                FileTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 上下文管理 ===");
                ContextTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 系统提示里的当前时间 ===");
                SystemPromptTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 请求结构配平 ===");
                RequestShapeTests.Run(ReportProvider);
                WordRequestShapeTests.Run(ReportProvider);
                WordAgentRegressionTests.Run(ReportProvider);
                ExcelToolRegressionTests.Run(ReportProvider);
                WordToolRegressionTests.Run(ReportProvider);
                ProviderToolRegressionTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 流式解析 ===");
                StreamTests.Run(ReportProvider);

                AgentToolTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 截断与续跑 ===");
                StallTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 失败重试 ===");
                RetryTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 重试计数归零 ===");
                RetryResetTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 模型能力回退 ===");
                CapabilityTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 模型可用性判定 ===");
                AvailabilityTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 按需确认的请求形态 ===");
                ProbeTests.Run(ReportProvider);
                BulkTestTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 授权类别分档 ===");
                ApprovalClassTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 地址区间与相交 ===");
                AddressSpanTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 审批卡对照 ===");
                PreviewTests.Run(ReportProvider);
                PreviewTests.RunDisplayReads(excel, executor, ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 面板宽度换算 ===");
                PaneWidthTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine("=== 面板打不开的成因判定 ===");
                PaneOpenDiagnosisTests.Run(ReportProvider);

                Console.WriteLine("=== 项目审查回归 ===");
                ProjectAuditTests.Run(excel, ReportProvider);
                executor.Undo.Clear();
                ReliabilityTests.Run(ReportProvider);

                Console.WriteLine();
                Console.WriteLine($"=== 结果：通过 {_passed}，失败 {_failed} ===");
                return _failed == 0 ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("测试宿主异常：" + ex);
                return 2;
            }
            finally
            {
                try
                {
                    if (workbook != null) { Call(workbook, "Close", false); }
                    if (excel != null) { Call(excel, "Quit"); }
                }
                catch
                {
                }

                Release(workbook);
                Release(workbooks);
                Release(excel);
            }
        }

        private static void RunAll(ToolExecutor executor)
        {
            // 读取类
            Expect(executor, "get_workbook_info", "{}", r => r.Ok && r.Data != null);
            Expect(executor, "get_selection", "{}", r => r.Ok);

            // 写入值，并验证读回内容与写入一致
            Expect(
                executor,
                "write_values",
                @"{""range"":""A1:C2"",""values"":[[""名称"",""数量"",""单价""],[""铅笔"",10,1.5]]}",
                r => r.Ok && Json(r).Contains("\"cells_written\": 6") && Json(r).Contains("铅笔"));

            Expect(
                executor,
                "read_range",
                @"{""range"":""A1:C2""}",
                r => r.Ok && Json(r).Contains("铅笔") && Json(r).Contains("1.5"));

            // 公式：写入后读回应得到计算结果而非公式文本
            Expect(
                executor,
                "write_formulas",
                @"{""range"":""D2"",""formulas"":[[""=B2*C2""]]}",
                r => r.Ok && Json(r).Contains("15"));

            Expect(
                executor,
                "read_range",
                @"{""range"":""D2"",""include_formulas"":true}",
                r => r.Ok && Json(r).Contains("=B2*C2"));

            // 形状不匹配必须被拦截
            Expect(
                executor,
                "write_values",
                @"{""range"":""A1:C2"",""values"":[[1,2]]}",
                r => !r.Ok && r.ErrorCode == "SHAPE_MISMATCH");

            // 公式缺少 = 必须被拦截
            Expect(
                executor,
                "write_formulas",
                @"{""range"":""F1"",""formulas"":[[""B2*C2""]]}",
                r => !r.Ok && r.ErrorCode == "FORMULA_INVALID");

            // 整列读取改为分页，首屏应正常返回而不是被固定阈值拦截。
            Expect(
                executor,
                "read_range",
                @"{""range"":""A:A""}",
                r => r.Ok && JObject.FromObject(r.Data)["next_offset"] != null);

            // 读取超过 5000 格时自动分页，不能返回固定范围错误。
            Expect(
                executor,
                "read_range",
                @"{""range"":""A1:E1000""}",
                r => r.Ok);

            Expect(
                executor,
                "read_range",
                @"{""range"":""A1:E1001""}",
                r => r.Ok && JObject.FromObject(r.Data)["next_offset"] != null);

            // 非法范围地址
            Expect(
                executor,
                "read_range",
                @"{""range"":""这不是范围""}",
                r => !r.Ok && r.ErrorCode == "RANGE_INVALID");

            // 不存在的工作表，错误信息应列出现有工作表
            Expect(
                executor,
                "read_range",
                @"{""range"":""A1"",""sheet"":""不存在的表""}",
                r => !r.Ok && r.ErrorCode == "SHEET_NOT_FOUND" && r.Error.Contains("现有工作表"));

            // 格式
            Expect(
                executor,
                "format_range",
                @"{""range"":""A1:C1"",""bold"":true,""fill_color"":""#FFFF00"",""horizontal_alignment"":""center""}",
                r => r.Ok && Json(r).Contains("bold") && Json(r).Contains("fill_color"));

            // 未提供任何格式属性应报错
            Expect(
                executor,
                "format_range",
                @"{""range"":""A1""}",
                r => !r.Ok && r.ErrorCode == "NO_CHANGES");

            // 颜色格式非法
            Expect(
                executor,
                "format_range",
                @"{""range"":""A1"",""fill_color"":""yellow""}",
                r => !r.Ok && r.ErrorCode == "ARG_INVALID");

            // 数字格式
            Expect(
                executor,
                "set_number_format",
                @"{""range"":""C2:D2"",""format_code"":""0.00""}",
                r => r.Ok && Json(r).Contains("actual_format"));

            // 自动列宽
            Expect(
                executor,
                "autofit_range",
                @"{""range"":""A1:D2"",""target"":""columns""}",
                r => r.Ok && Json(r).Contains("dimensions_adjusted"));

            Expect(
                executor,
                "autofit_range",
                @"{""range"":""A1"",""target"":""diagonal""}",
                r => !r.Ok && r.ErrorCode == "ARG_INVALID");

            // 垂直对齐
            Expect(
                executor,
                "format_range",
                @"{""range"":""A1:C1"",""vertical_alignment"":""center""}",
                r => r.Ok && Json(r).Contains("vertical_alignment"));

            Expect(
                executor,
                "format_range",
                @"{""range"":""A1"",""vertical_alignment"":""middle""}",
                r => !r.Ok && r.ErrorCode == "ARG_INVALID");

            // 适配：一次调用要同时报告对齐、列宽上限、换行和行高改动
            Expect(
                executor,
                "fit_range",
                @"{""range"":""A1:C2""}",
                r => r.Ok
                    && Json(r).Contains("vertical_alignment")
                    && Json(r).Contains("row_height")
                    && Json(r).Contains("column_width")
                    && Json(r).Contains("wrap_text")
                    && Json(r).Contains("columns_capped")
                    && Json(r).Contains("rows_adjusted"));

            // 适配刻意不设单元格上限：整列约百万单元格也应放行。
            // 这条与 read_range 的 A:A 被拦形成对照，两者的约束理由不同。
            Expect(
                executor,
                "fit_range",
                @"{""range"":""A:A""}",
                r => r.Ok);

            // 省略 range 时自行取已用范围
            Expect(
                executor,
                "fit_range",
                @"{}",
                r => r.Ok && Json(r).Contains("address"));

            // 水平对齐三选一：省略即 center
            Expect(
                executor,
                "fit_range",
                @"{""range"":""A1:C2""}",
                r => r.Ok && Json(r).Contains("\"horizontal_alignment\": \"center\""));

            Expect(
                executor,
                "fit_range",
                @"{""range"":""A1:C2"",""horizontal_alignment"":""left""}",
                r => r.Ok && Json(r).Contains("\"horizontal_alignment\": \"left\""));

            Expect(
                executor,
                "fit_range",
                @"{""range"":""A1:C2"",""horizontal_alignment"":""right""}",
                r => r.Ok && Json(r).Contains("\"horizontal_alignment\": \"right\""));

            // 垂直方向不开放选项，非法水平值要拒绝
            Expect(
                executor,
                "fit_range",
                @"{""range"":""A1:C2"",""horizontal_alignment"":""justify""}",
                r => !r.Ok && r.ErrorCode == "ARG_INVALID");

            // 合并：单格无从合并
            Expect(
                executor,
                "merge_cells",
                @"{""range"":""H1""}",
                r => !r.Ok && r.ErrorCode == "NOTHING_TO_MERGE");

            // across 为真表示逐行合并，单列时每行都无从合并
            Expect(
                executor,
                "merge_cells",
                @"{""range"":""H1:H3"",""across"":true}",
                r => !r.Ok && r.ErrorCode == "NOTHING_TO_MERGE");

            // 非法对齐值必须在动手前拒绝：合并已生效再报错，丢掉的值就回不来了
            Expect(
                executor,
                "merge_cells",
                @"{""range"":""H1:J1"",""horizontal_alignment"":""justify""}",
                r => !r.Ok && r.ErrorCode == "ARG_INVALID");

            // 合并前范围里有值，要如实回报会丢掉几个
            Expect(
                executor,
                "write_values",
                @"{""range"":""H1:J1"",""values"":[[""标题"",""乙"",""丙""]]}",
                r => r.Ok);

            Expect(
                executor,
                "merge_cells",
                @"{""range"":""H1:J1"",""horizontal_alignment"":""center""}",
                r => r.Ok
                    && Json(r).Contains("\"merged_areas\": 1")
                    && Json(r).Contains("\"discarded_values\": 2")
                    && Json(r).Contains("horizontal_alignment"));

            // 取消合并：拆回独立单元格
            Expect(
                executor,
                "unmerge_cells",
                @"{""range"":""H1:J1""}",
                r => r.Ok && Json(r).Contains("\"areas_unmerged\": 1"));

            // 没有合并可拆时要明确说没有，而不是「成功但什么也没做」
            Expect(
                executor,
                "unmerge_cells",
                @"{""range"":""H1:J1""}",
                r => !r.Ok && r.ErrorCode == "NO_MERGED_CELLS");

            // 逐行合并：三行各成一格，共三个合并区域
            Expect(
                executor,
                "merge_cells",
                @"{""range"":""H5:J7"",""across"":true}",
                r => r.Ok && Json(r).Contains("\"merged_areas\": 3"));

            Expect(
                executor,
                "unmerge_cells",
                @"{""range"":""H5:J7""}",
                r => r.Ok && Json(r).Contains("\"areas_unmerged\": 3"));

            // 大范围合并可执行；超过快照预算时只是不承诺自动撤销。
            Expect(
                executor,
                "merge_cells",
                @"{""range"":""A1:E1001""}",
                r => r.Ok && Json(r).Contains("\"merged_areas\": 1"));

            // 工作表结构
            Expect(
                executor,
                "add_worksheet",
                @"{""name"":""测试表""}",
                r => r.Ok && Json(r).Contains("测试表"));

            Expect(
                executor,
                "add_worksheet",
                @"{""name"":""测试表""}",
                r => !r.Ok && r.ErrorCode == "SHEET_EXISTS");

            Expect(
                executor,
                "add_worksheet",
                @"{""name"":""非法:名称""}",
                r => !r.Ok && r.ErrorCode == "NAME_INVALID");

            Expect(
                executor,
                "rename_worksheet",
                @"{""old_name"":""测试表"",""new_name"":""改名后""}",
                r => r.Ok && Json(r).Contains("改名后"));

            // 排序：键列超出范围必须被拦截
            Expect(
                executor,
                "sort_range",
                @"{""range"":""A1:D2"",""key_column"":""Z""}",
                r => !r.Ok && r.ErrorCode == "KEY_OUT_OF_RANGE");

            Expect(
                executor,
                "sort_range",
                @"{""range"":""A1:D2"",""key_column"":""A"",""has_header"":true}",
                r => r.Ok);

            // 清除
            Expect(
                executor,
                "clear_range",
                @"{""range"":""D2"",""scope"":""contents""}",
                r => r.Ok && Json(r).Contains("contents"));

            Expect(
                executor,
                "clear_range",
                @"{""range"":""D2"",""scope"":""everything""}",
                r => !r.Ok && r.ErrorCode == "ARG_INVALID");

            // 日期格式：中文环境下容易因区域设置而失败，需单独覆盖
            Expect(
                executor,
                "set_number_format",
                @"{""range"":""E1"",""format_code"":""yyyy-mm-dd""}",
                r => r.Ok);

            // 表格与图表
            Expect(
                executor,
                "create_table",
                @"{""range"":""A1:C2"",""has_header"":true,""name"":""测试表格""}",
                r => r.Ok && Json(r).Contains("table_name"));

            Expect(
                executor,
                "create_chart",
                @"{""range"":""B1:C2"",""chart_type"":""column"",""title"":""测试图表""}",
                r => r.Ok && Json(r).Contains("column"));

            Expect(
                executor,
                "create_chart",
                @"{""range"":""B1:C2"",""chart_type"":""donut""}",
                r => !r.Ok && r.ErrorCode == "ARG_INVALID");

            // 未知工具
            Expect(executor, "no_such_tool", "{}", r => !r.Ok && r.ErrorCode == "UNKNOWN_TOOL");

            // 缺少必需参数
            Expect(executor, "read_range", "{}", r => !r.Ok && r.ErrorCode == "ARG_MISSING");
        }

        private static void Expect(ToolExecutor executor, string tool, string argsJson, Func<ToolResult, bool> check)
        {
            var label = $"{tool} {Compact(argsJson)}";
            try
            {
                var args = JObject.Parse(argsJson);
                var result = executor.Execute(tool, args);
                if (check(result))
                {
                    _passed++;
                    Console.WriteLine($"  通过  {label}");
                }
                else
                {
                    _failed++;
                    Console.WriteLine($"  失败  {label}");
                    Console.WriteLine($"        实际: ok={result.Ok} code={result.ErrorCode} error={result.Error}");
                    Console.WriteLine($"        数据: {Compact(Json(result))}");
                }
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine($"  异常  {label}");
                Console.WriteLine($"        {ex.GetType().Name}: {ex.Message}");
            }
        }

        private static void ReportProvider(string label, bool ok, string detail)
        {
            if (ok)
            {
                _passed++;
                Console.WriteLine($"  通过  {label}");
            }
            else
            {
                _failed++;
                Console.WriteLine($"  失败  {label}");
                if (!string.IsNullOrEmpty(detail))
                {
                    Console.WriteLine($"        {detail}");
                }
            }
        }

        private static string Json(ToolResult result)
        {
            try
            {
                return JsonConvert.SerializeObject(result.ToPayload(), Formatting.Indented);
            }
            catch
            {
                return "<无法序列化>";
            }
        }

        private static string Compact(string text)
        {
            if (text == null) { return string.Empty; }
            var flat = text.Replace("\r", " ").Replace("\n", " ").Replace("  ", " ");
            return flat.Length > 150 ? flat.Substring(0, 150) + "…" : flat;
        }

        private static object Get(object target, string name, params object[] args) =>
            target.GetType().InvokeMember(name, System.Reflection.BindingFlags.GetProperty, null, target, args);

        private static void Set(object target, string name, params object[] args) =>
            target.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, target, args);

        private static object Call(object target, string name, params object[] args) =>
            target.GetType().InvokeMember(name, System.Reflection.BindingFlags.InvokeMethod, null, target, args);

        private static void Release(object target)
        {
            try
            {
                if (target != null && Marshal.IsComObject(target))
                {
                    Marshal.ReleaseComObject(target);
                }
            }
            catch
            {
            }
        }
    }
}
