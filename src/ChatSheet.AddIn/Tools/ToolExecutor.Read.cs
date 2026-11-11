using System;
using System.Collections.Generic;
using System.Globalization;
using ChatSheet.AddIn.Hosts;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Tools
{
    /// <summary>
    /// 工具执行器。把模型的工具调用翻译成宿主 COM 操作。
    /// 全程后期绑定，不引用 Office PIA。
    /// </summary>
    internal sealed partial class ToolExecutor
    {
        private readonly Func<object> _applicationAccessor;
        private readonly RangeResolver _resolver;
        private readonly WorkbookContext _context;

        private readonly UndoStore _undo;

        internal ToolExecutor(Func<object> applicationAccessor)
        {
            _applicationAccessor = applicationAccessor ?? throw new ArgumentNullException(nameof(applicationAccessor));
            _resolver = new RangeResolver(applicationAccessor);
            _context = new WorkbookContext(applicationAccessor);
            _undo = new UndoStore(applicationAccessor, _resolver);
        }

        internal UndoStore Undo => _undo;
        internal Action BeforeWriteChunk { get; set; }

        /// <summary>
        /// 这片范围的外观属性是不是逐项都不一致。
        ///
        /// 给审批卡用：格式类操作的参数已经说清「要改成什么」，缺的是「现在是什么」。
        /// 整片一致时那个问题有答案（卡片可以不提），逐项都不一致时只能如实说
        /// 「当前格式不统一」——而这也正是撤销还原不回来的那种范围。
        ///
        /// 只读范围级属性，是 O(1) 次 COM 调用，不逐格问。
        /// 读不出来返回 null：分不清「一致」与「读失败」时不要替用户下结论。
        /// </summary>
        internal bool? IsFormattingMixed(string address, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return null;
            }

            try
            {
                using (var range = _resolver.Resolve(address, sheetName))
                {
                    return SnapshotCapture.CaptureFormatForProbe(range.Range)?.IsAllNull;
                }
            }
            catch (Exception ex)
            {
                Log.Warn("探测格式是否统一失败：" + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 按单元格**显示出来的样子**读一片范围，只给审批卡的对照用。
        ///
        /// 不能复用 read_range：那个工具读 Value2，返回的是底层值——日期是序列号
        /// （45900）、时间是小数（0.75）、百分比是 0.1234、千分位是 1234.5。
        /// 模型需要的正是这种未经格式化的值，所以工具那一侧不能改。
        ///
        /// 但对照表的用途是「和你在格子里看到的对上」。把 45900 摆在
        /// 「将改为 2025-08-31」旁边，看起来像换了一种东西，而不是改了一个值。
        /// Range.Text 给的就是屏幕上那一串，四种格式一次全对。
        ///
        /// 代价：Text 受列宽影响——列太窄时 Excel 显示 ####，Text 也跟着给 ####。
        /// 因此拿不到内容时退回 Value2 的读法，不让对照变成一片井号。
        /// </summary>
        internal List<List<object>> ReadDisplayMatrix(string address, string sheetName, int maxRows, int maxColumns)
        {
            if (string.IsNullOrWhiteSpace(address) || maxRows <= 0 || maxColumns <= 0)
            {
                return null;
            }

            try
            {
                using (var range = _resolver.Resolve(address, sheetName))
                {
                    // 只读对照窗口那一块，不读整片。
                    //
                    // Range.Text 与 Value2 不同：对多格范围它不返回二维数组，
                    // 只给一个值（各格不一致时是 Null），所以必须逐格问宿主。
                    // 逐格对大范围本来不可接受，但卡上最多画 8×6，
                    // 读窗口之外的格子没有任何用处。
                    var rows = Math.Min(range.Rows, maxRows);
                    var columns = Math.Min(range.Columns, maxColumns);

                    object cells = null;
                    try
                    {
                        cells = Com.Get(range.Range, "Cells");
                        var result = new List<List<object>>();

                        for (var r = 1; r <= rows; r++)
                        {
                            var row = new List<object>();
                            for (var c = 1; c <= columns; c++)
                            {
                                row.Add(ReadOneDisplayCell(cells, r, c));
                            }

                            result.Add(row);
                        }

                        return result;
                    }
                    finally
                    {
                        Com.Release(cells);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warn("读取显示文本失败，对照退回底层值：" + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 读一格显示出来的样子。
        ///
        /// 列宽不足时 Excel 把内容显示成 ####，Text 如实返回那几个井号——
        /// 那种串没有信息量，用户要判断的是内容而不是当前列宽，因此退回底层值。
        /// </summary>
        private static object ReadOneDisplayCell(object cells, int row, int column)
        {
            object cell = null;
            try
            {
                cell = Com.Get(cells, "Item", row, column);
                var shown = Com.GetString(cell, "Text");
                if (shown.Length > 0 && !LooksTooNarrow(shown))
                {
                    return shown;
                }

                return Normalize(Com.Get(cell, "Value2"));
            }
            catch
            {
                return null;
            }
            finally
            {
                Com.Release(cell);
            }
        }

        /// <summary>整串都是 # 说明列宽不够，显示串没有信息量。</summary>
        private static bool LooksTooNarrow(string shown)
        {
            if (shown.Length == 0)
            {
                return false;
            }

            foreach (var ch in shown)
            {
                if (ch != '#')
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// 把 Excel 的选区跳到一个已解析范围。
        ///
        /// 面板卡片上写出地址却不能点进去，用户要核对只能自己在 Excel 里翻。
        /// 这条不是给模型的工具，所以不进 ToolCatalog；它只是让面板借宿主的
        /// `Goto` 走到已经展示出来的那块范围。
        ///
        /// 刻意不保存再还原旧选区：Office 对象每次读取都是新代理，不能用 COM
        /// 标识判断是不是同一个范围；误还原会静默跳到错误位置。悬停会明说此操作
        /// 改变当前选区，和用户在表上自己点格子是同一种显式动作。
        /// </summary>
        internal ToolResult GotoRange(string address, string sheetName)
        {
            using (var range = _resolver.Resolve(address, sheetName))
            {
                Com.Call(Application, "Goto", range.Range);
                return ToolResult.Success(new Dictionary<string, object>
                {
                    ["sheet"] = range.SheetName,
                    ["address"] = range.Address,
                });
            }
        }

        private object Application =>
            _applicationAccessor() ?? throw new InvalidOperationException("尚未连接到宿主应用程序。");

        /// <summary>
        /// 执行一个工具调用。
        /// 可预期的错误一律转成结构化结果回传模型，不向上抛出。
        /// </summary>
        internal ToolResult Execute(string name, JObject args)
        {
            return Execute(name, args, undoId: null);
        }

        /// <summary>
        /// 执行工具，并在提供 undoId 时登记撤销记录。
        ///
        /// 快照在工具执行前后各采集一次：撤销还原「前」，恢复还原「后」。
        /// 两个方向对称，不必重跑工具——重跑对写入类操作不安全，
        /// 因为期间用户可能已手工改过同一片区域。
        /// </summary>
        internal ToolResult Execute(string name, JObject args, string undoId)
        {
            args = args ?? new JObject();
            var definition = ToolCatalog.Find(name);
            var tracking = undoId != null
                && definition != null
                && definition.Risk != ToolRisk.Read;

            // fit_range 允许省略 range 并在执行时取 UsedRange；但撤销快照必须先于
            // 执行采集，因此有撤销标识时先把隐式范围解析回参数。解析失败交给
            // ExecuteCore 按原路径转换为结构化错误，不让预处理改变错误契约。
            if (tracking
                && string.Equals(name, "fit_range", StringComparison.Ordinal)
                && string.IsNullOrWhiteSpace(args.Value<string>("range")))
            {
                try
                {
                    args["range"] = UsedRangeAddress(args.Value<string>("sheet"));
                }
                catch
                {
                    return ExecuteCore(name, args);
                }
            }

            RangeSnapshot before = null;
            var detail = SnapshotDetail.None;

            if (tracking)
            {
                detail = UndoStore.DetailFor(name);
                if (detail != SnapshotDetail.None)
                {
                    before = TryCapture(name, args, detail);
                    // 采集失败就不登记撤销，而不是让整个操作失败——
                    // 用户宁可少一个撤销按钮，也不愿操作被拒。
                    if (before == null)
                    {
                        tracking = false;
                    }
                }
            }

            var result = ExecuteCore(name, args);

            if (tracking && result.Ok)
            {
                TryRegisterUndo(undoId, name, args, before, detail, result);
            }

            return result;
        }

        private RangeSnapshot TryCapture(string toolName, JObject args, SnapshotDetail detail)
        {
            try
            {
                var address = args?.Value<string>("range");
                if (string.IsNullOrWhiteSpace(address))
                {
                    return null;
                }

                using (var range = _resolver.Resolve(address, args.Value<string>("sheet")))
                {
                    // 内容或完整格式始终按单元格存储，超过上限就不登记撤销。
                    // 适配的统一对齐可保持范围级快照；混合对齐则只有在这个上限内
                    // 才能逐格保存。因此把是否允许逐格对齐传给快照层，由它在需要时
                    // 选择安全地放弃撤销记录，而不影响适配操作本身。
                    // 合并快照也算逐格：判断哪些格属于哪个合并区域只能逐格问宿主。
                    var cellwise = (detail
                        & (SnapshotDetail.Content | SnapshotDetail.Format | SnapshotDetail.Merge)) != 0;

                    if (cellwise)
                    {
                        if (range.CellCount > ToolLimits.MaxSnapshotCells)
                        {
                            return null;
                        }
                    }
                    else if (range.Rows + range.Columns > ToolLimits.MaxSnapshotDimensions)
                    {
                        return null;
                    }

                    var snapshot = SnapshotCapture.Capture(
                        range,
                        detail,
                        allowCellwiseAlignment: range.CellCount <= ToolLimits.MaxSnapshotCells);

                    // 清除格式会连边框一起抹掉，而边框不在采集范围内。
                    // 标出来，好让卡片如实说明撤销还原不回什么。
                    if (snapshot != null && ClearsFormats(toolName, args))
                    {
                        snapshot.ClearsUncapturedFormats = true;
                    }

                    return snapshot;
                }
            }
            catch (Exception ex)
            {
                Log.Warn("采集撤销快照失败：" + ex.Message);
                return null;
            }
        }

        /// <summary>这次调用会不会抹掉快照覆盖不到的格式维度（边框等）。</summary>
        private static bool ClearsFormats(string toolName, JObject args)
        {
            if (!string.Equals(toolName, "clear_range", StringComparison.Ordinal))
            {
                return false;
            }

            var scope = args?.Value<string>("scope");
            return string.Equals(scope, "formats", StringComparison.OrdinalIgnoreCase)
                || string.Equals(scope, "all", StringComparison.OrdinalIgnoreCase);
        }

        private void TryRegisterUndo(
            string undoId,
            string name,
            JObject args,
            RangeSnapshot before,
            SnapshotDetail detail,
            ToolResult result)
        {
            try
            {
                var record = new UndoRecord
                {
                    Workbook = WorkbookTarget.Capture(Application),
                    Id = undoId,
                    ToolName = name,
                    At = DateTime.Now,
                    Summary = DescribeForUndo(name, args, result),
                    ArgumentsJson = args?.ToString(Newtonsoft.Json.Formatting.None),
                };

                var structureKind = UndoStore.StructureKindFor(name);
                if (structureKind != StructureKind.None)
                {
                    record.Structure = BuildStructureAction(structureKind, args, result);
                }
                else
                {
                    record.Before = before;
                    // 操作后再采一次作为恢复依据。
                    record.After = TryCapture(name, args, detail);
                }

                _undo.Add(record);
            }
            catch (Exception ex)
            {
                // 撤销登记失败不应影响已成功的操作。
                Log.Warn("登记撤销记录失败：" + ex.Message);
            }
        }

        private static StructureAction BuildStructureAction(StructureKind kind, JObject args, ToolResult result)
        {
            var data = result.Data == null ? null : JObject.FromObject(result.Data);

            switch (kind)
            {
                case StructureKind.AddedWorksheet:
                    return new StructureAction
                    {
                        Kind = kind,
                        Name = data?.Value<string>("created_sheet") ?? args.Value<string>("name"),
                        AfterSheet = args.Value<string>("after_sheet"),
                    };

                case StructureKind.RenamedWorksheet:
                    return new StructureAction
                    {
                        Kind = kind,
                        Name = data?.Value<string>("new_name") ?? args.Value<string>("new_name"),
                        PreviousName = args.Value<string>("old_name"),
                    };

                case StructureKind.CreatedTable:
                    return new StructureAction
                    {
                        Kind = kind,
                        Name = data?.Value<string>("table_name"),
                        SheetName = data?.Value<string>("sheet") ?? args.Value<string>("sheet"),
                    };

                case StructureKind.CreatedChart:
                    {
                        // 没有 Shape 名字就不登记这条记录。
                        //
                        // CanUndo 只看「Structure 不为空」，所以登记一条名字为空的记录
                        // 必然做出一个点下去只能得到「找不到图表「」」的按钮。宁可不给
                        // 按钮并说明原因——这与面板「适配」已经立下的规矩是同一句话：
                        // 保不住足以完整还原的依据，就不承诺可以撤销。
                        var chartName = data?.Value<string>("chart_name");
                        if (string.IsNullOrWhiteSpace(chartName))
                        {
                            Log.Warn("图表未回报 Shape 名称，本次不登记撤销记录。");
                            return null;
                        }

                        return new StructureAction
                        {
                            Kind = kind,
                            Name = chartName,
                            SheetName = data?.Value<string>("sheet") ?? args.Value<string>("sheet"),

                            // 图表删除后无法自动重建，因此撤销之后不得再显示「恢复」。
                            // 只修撤销不修恢复，等于把同一个谎言从一个按钮挪到另一个。
                            CanRestore = false,
                        };
                    }

                default:
                    return null;
            }
        }

        /// <summary>撤销条目的一句话描述。要让用户看得懂自己在撤销什么。</summary>
        private static string DescribeForUndo(string name, JObject args, ToolResult result)
        {
            var data = result.Data == null ? null : JObject.FromObject(result.Data);
            var sheet = data?.Value<string>("sheet") ?? args?.Value<string>("sheet");
            var address = data?.Value<string>("address") ?? args?.Value<string>("range");
            var where = string.IsNullOrEmpty(sheet) ? address : $"{sheet}!{address}";

            switch (name)
            {
                case "write_values":
                    return $"写入值 {where}";
                case "write_formulas":
                    return $"写入公式 {where}";
                case "format_range":
                    return $"设置格式 {where}";
                case "set_number_format":
                    return $"设置数字格式 {where}";
                case "autofit_range":
                    return $"自动调整 {where}";
                case "fit_range":
                    return $"适配 {where}";
                case "merge_cells":
                    return $"合并单元格 {where}";
                case "unmerge_cells":
                    return $"取消合并 {where}";
                case "clear_range":
                    return $"清除 {where}";
                case "sort_range":
                    return $"排序 {where}";
                case "add_worksheet":
                    return $"新增工作表 {data?.Value<string>("created_sheet") ?? args?.Value<string>("name")}";
                case "rename_worksheet":
                    return $"重命名工作表 {args?.Value<string>("old_name")} → {data?.Value<string>("new_name")}";
                case "create_table":
                    return $"创建表格 {data?.Value<string>("table_name")}";
                case "create_chart":
                    return $"创建图表（{where}）";
                default:
                    return name;
            }
        }

        private ToolResult ExecuteCore(string name, JObject args)
        {
            try
            {
                switch (name)
                {
                    case "get_workbook_info":
                        return GetWorkbookInfo();
                    case "excel_object":
                        return ExcelObject(args);
                    case "get_selection":
                        return GetSelection();
                    case "read_range":
                        return ReadRange(args);
                    case "write_values":
                        return WriteMatrix(args, formulas: false);
                    case "write_formulas":
                        return WriteMatrix(args, formulas: true);
                    case "format_range":
                        return FormatRange(args);
                    case "set_number_format":
                        return SetNumberFormat(args);
                    case "autofit_range":
                        return AutofitRange(args);
                    case "fit_range":
                        return FitRange(args);
                    case "merge_cells":
                        return MergeCells(args);
                    case "unmerge_cells":
                        return UnmergeCells(args);
                    case "clear_range":
                        return ClearRange(args);
                    case "add_worksheet":
                        return AddWorksheet(args);
                    case "rename_worksheet":
                        return RenameWorksheet(args);
                    case "sort_range":
                        return SortRange(args);
                    case "create_table":
                        return CreateTable(args);
                    case "create_chart":
                        return CreateChart(args);
                    default:
                        return ToolResult.Failure("UNKNOWN_TOOL", $"未知工具：{name}。");
                }
            }
            catch (ToolException ex)
            {
                Log.Warn($"工具 {name} 返回可预期错误：{ex.Code} {ex.Message}");
                return ToolResult.Failure(ex.Code, ex.Message);
            }
            catch (Exception ex)
            {
                // 必须剥开 TargetInvocationException：后期绑定调用失败时，
                // 外层异常只会说「调用的目标发生了异常」，真正原因在内层。
                // 不剥开的话错误信息对模型和排查都毫无价值。
                var root = Unwrap(ex);
                if (root is ToolException toolError)
                {
                    Log.Warn($"工具 {name} 返回可预期错误：{toolError.Code} {toolError.Message}");
                    return ToolResult.Failure(toolError.Code, toolError.Message);
                }

                Log.Error($"工具 {name} 执行失败", root);
                return ToolResult.Failure("HOST_ERROR", $"宿主操作失败：{root.Message}");
            }
        }

        /// <summary>剥离反射调用的包装异常，取出真正的失败原因。</summary>
        private static Exception Unwrap(Exception ex)
        {
            var current = ex;
            var depth = 0;
            while (current is System.Reflection.TargetInvocationException && current.InnerException != null && depth < 5)
            {
                current = current.InnerException;
                depth++;
            }

            return current;
        }

        private ToolResult GetWorkbookInfo()
        {
            var summary = _context.GetSummary();
            if (!summary.HasWorkbook)
            {
                return ToolResult.Failure("NO_WORKBOOK", "当前没有打开的工作簿。");
            }

            var sheets = new List<object>();
            foreach (var sheet in summary.Sheets)
            {
                sheets.Add(new Dictionary<string, object>
                {
                    ["name"] = sheet.Name,
                    ["used_range"] = sheet.UsedRange,
                    ["rows"] = sheet.RowCount,
                    ["columns"] = sheet.ColumnCount,
                });
            }

            return ToolResult.Success(new Dictionary<string, object>
            {
                ["workbook"] = summary.Name,
                ["saved"] = summary.Saved,
                ["active_sheet"] = summary.ActiveSheet,
                ["sheet_count"] = summary.SheetCount,
                ["sheets"] = sheets,
                ["host"] = Com.GetString(Application, "Name"),
                ["host_version"] = Com.GetString(Application, "Version"),
                ["object_operations"] = new {
                    tool = "excel_object", actions = new[] { "get", "set", "call" },
                    roots = new[] { "workbook", "worksheet", "window" },
                    families = new[] { "范围和行列", "工作表", "字体边框与对齐", "筛选排序", "条件格式与数据验证",
                        "命名区域", "表格与图表", "透视表", "窗格与页面设置" },
                    compatibility = "入口已开放；具体成员是否可用以当前宿主返回结果为准，修改后需要读取验证。"
                },
            });
        }

        private ToolResult GetSelection()
        {
            var selection = _context.GetSelection();
            if (!selection.HasSelection)
            {
                return ToolResult.Failure("NO_SELECTION", "当前没有选区。");
            }

            return ToolResult.Success(new Dictionary<string, object>
            {
                ["sheet"] = selection.SheetName,
                ["address"] = selection.Address,
                ["rows"] = selection.RowCount,
                ["columns"] = selection.ColumnCount,
            });
        }

        private ToolResult ReadRange(JObject args)
        {
            var address = RequireString(args, "range");
            var sheet = OptionalString(args, "sheet");

            using (var range = _resolver.Resolve(address, sheet))
            {
                return ReadRange(range, args);
            }
        }

        private ToolResult ReadRange(ResolvedRange range, JObject args)
        {
            var includeFormulas = OptionalBool(args, "include_formulas") ?? false;
            RangeResolver.AssertCellLimit(range, ToolLimits.MaxReadCells, "读取");

            var offset = args.Value<long?>("offset") ?? 0;
            if (offset < 0 || offset >= range.CellCount)
            {
                throw new ToolException("ARG_INVALID", "offset 必须位于范围内。");
            }
            var rowOffset = offset / range.Columns;
            var colOffset = (int)(offset % range.Columns);
            var columns = Math.Min(range.Columns - colOffset, ToolLimits.ReadPageCells);
            var rows = colOffset == 0 && columns == range.Columns
                ? Math.Min(range.Rows - (int)rowOffset, Math.Max(1, ToolLimits.ReadPageCells / columns)) : 1;
            var firstRow = Convert.ToInt32(Com.Get(range.Range, "Row")) + (int)rowOffset;
            var firstCol = Convert.ToInt32(Com.Get(range.Range, "Column")) + colOffset;
            var pageAddress = ColumnName(firstCol) + firstRow + ":" +
                ColumnName(firstCol + columns - 1) + (firstRow + rows - 1);
            object pageRange = null;
            try
            {
                // 宿主事件可能切换活动工作簿，分页与验证必须沿用最初解析的工作表。
                pageRange = Com.Get(range.Worksheet, "Range", pageAddress);
                var page = new ResolvedRange(null, pageRange, range.SheetName,
                    Com.GetString(pageRange, "Address"), rows, columns);
                var values = ReadMatrix(page, "Value2");
                var payload = new Dictionary<string, object>
                {
                    ["sheet"] = range.SheetName,
                    ["address"] = range.Address,
                    ["rows"] = rows,
                    ["columns"] = columns,
                    ["page_address"] = page.Address,
                    ["total_cells"] = range.CellCount,
                    ["total_rows"] = range.Rows,
                    ["total_columns"] = range.Columns,
                    ["next_offset"] = offset + (long)rows * columns < range.CellCount
                        ? (object)(offset + (long)rows * columns) : null,
                    ["values"] = values,
                };

                if (includeFormulas)
                {
                    payload["formulas"] = ReadMatrix(page, "Formula");
                }

                return ToolResult.Success(payload);
            }
            finally { Com.Release(pageRange); }
        }

        private static string ColumnName(int column)
        {
            var name = string.Empty;
            while (column > 0) { column--; name = (char)('A' + column % 26) + name; column /= 26; }
            return name;
        }

        /// <summary>
        /// 读取矩阵。单个单元格时宿主返回标量而非二维数组，
        /// 这里统一成二维结构，避免模型面对两种形态。
        /// </summary>
        private static List<List<object>> ReadMatrix(ResolvedRange range, string property)
        {
            var raw = Com.Get(range.Range, property);
            var result = new List<List<object>>();

            if (raw is Array array && array.Rank == 2)
            {
                var lowerRow = array.GetLowerBound(0);
                var upperRow = array.GetUpperBound(0);
                var lowerCol = array.GetLowerBound(1);
                var upperCol = array.GetUpperBound(1);

                for (var r = lowerRow; r <= upperRow; r++)
                {
                    var row = new List<object>();
                    for (var c = lowerCol; c <= upperCol; c++)
                    {
                        row.Add(Normalize(array.GetValue(r, c)));
                    }

                    result.Add(row);
                }

                return result;
            }

            result.Add(new List<object> { Normalize(raw) });
            return result;
        }

        /// <summary>把 COM 值规范成 JSON 友好形态，并截断超长文本。</summary>
        private static object Normalize(object value)
        {
            if (value == null || value is DBNull)
            {
                return null;
            }

            if (value is string text)
            {
                return text.Length > ToolLimits.MaxCellTextLength
                    ? text.Substring(0, ToolLimits.MaxCellTextLength) + "…（已截断）"
                    : text;
            }

            if (value is DateTime dt)
            {
                return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }

            if (value is bool || value is double || value is int || value is long || value is decimal || value is float)
            {
                return value;
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
    }
}
