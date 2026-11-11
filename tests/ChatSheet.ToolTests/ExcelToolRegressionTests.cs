using System;
using System.Runtime.CompilerServices;
using ChatSheet.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class ExcelToolRegressionTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            ObjectArguments(report);
            MatrixShape(report);
            FormatValidation(report);
            BoundRanges(report);
        }

        private static void ObjectArguments(Action<string, bool, string> report)
        {
            Check(report, "对象工具允许省略或置空可选 args", () =>
            {
                foreach (var suffix in new[] { "", ",args:null" })
                {
                    var app = new FakeApplication();
                    var result = Execute(app, "excel_object", "{root:'workbook',path:[],action:'get',member:'Name'" + suffix + "}");
                    Require(result.Ok && Data(result)["value"].Value<string>() == "原工作簿", result.Error);
                    Require(app.DisplayAlerts, "没有恢复 DisplayAlerts");
                }
            });
            Check(report, "对象路径允许 args:null 并统一处理 kind 的首尾空白", () =>
            {
                foreach (var kind in new[] { "call", " call " })
                {
                    var app = new FakeApplication();
                    var result = Execute(app, "excel_object", "{root:'workbook',path:[{member:'Next',kind:'" + kind + "',args:null}],action:'get',member:'Name'}");
                    Require(result.Ok && app.ActiveWorkbook.PathCalls == 1, result.Error);
                }
            });
            Check(report, "对象工具保留 Int32 枚举和 Int64 数值", () =>
            {
                foreach (var value in new[] { -4108L, 2147483648L, -2147483649L })
                {
                    var app = new FakeApplication();
                    var args = JObject.Parse("{root:'workbook',path:[],action:'set',member:'Payload'}");
                    args["args"] = new JArray(value);
                    var result = new ToolExecutor(() => app).Execute("excel_object", args);
                    Require(result.Ok && Convert.ToInt64(app.ActiveWorkbook.Payload) == value, result.Error);
                    Require(value != -4108L || app.ActiveWorkbook.Payload is int, "枚举参数不应变为 Int64");
                }
            });
            Check(report, "对象矩阵和普通数组接受 Int64 单元格", () =>
            {
                var app = new FakeApplication();
                var matrix = Execute(app, "excel_object", "{root:'workbook',path:[],action:'set',member:'Payload',args:[{matrix:[[2147483648]]}]}");
                Require(matrix.Ok && Convert.ToInt64(((object[,])app.ActiveWorkbook.Payload)[0, 0]) == 2147483648L, matrix.Error);
                var array = Execute(app, "excel_object", "{root:'workbook',path:[],action:'set',member:'Payload',args:[[2147483648]]}");
                Require(array.Ok && Convert.ToInt64(((object[])app.ActiveWorkbook.Payload)[0]) == 2147483648L, array.Error);
            });
            Check(report, "对象参数错误在可能修改状态的路径调用前拒绝", () =>
            {
                var app = new FakeApplication();
                var result = Execute(app, "excel_object", "{root:'workbook',path:[{member:'Next',kind:'call'}],action:'set',member:'Payload',args:[{matrix:[[1]],missing:'invalid'}]}");
                Require(!result.Ok && result.ErrorCode == "ARG_INVALID" && app.ActiveWorkbook.PathCalls == 0, result.Error);
                Require(result.Data == null, "未开始的操作不应回报 operation_started");
            });
            Check(report, "对象工具仍在路径执行前拒绝外部成员和错误形状", () =>
            {
                foreach (var json in new[]
                {
                    "{root:'workbook',path:[{member:'Next',kind:'call'}],action:'get',member:'Application'}",
                    "{root:'workbook',path:[{member:'Next',kind:'call'}],action:'set',member:'Payload',args:{}}",
                    "{root:'workbook',path:[{member:'Next',kind:'call'}],action:'set',member:'Payload',args:[{matrix:[[1],[2,3]]}]}",
                })
                {
                    var app = new FakeApplication();
                    var result = Execute(app, "excel_object", json);
                    Require(!result.Ok && app.ActiveWorkbook.PathCalls == 0, result.Error);
                }
            });
            Check(report, "对象调用失败后恢复宿主提示设置", () =>
            {
                foreach (var initial in new[] { true, false })
                {
                    var app = new FakeApplication { DisplayAlerts = initial };
                    var result = Execute(app, "excel_object", "{root:'workbook',path:[],action:'call',member:'Fail'}");
                    Require(!result.Ok && result.ErrorCode == "HOST_ERROR" && app.DisplayAlerts == initial, result.Error);
                }
            });
        }

        private static void MatrixShape(Action<string, bool, string> report)
        {
            foreach (var formulas in new[] { false, true })
            {
                var tool = formulas ? "write_formulas" : "write_values";
                var key = formulas ? "formulas" : "values";
                foreach (var data in formulas
                    ? new[] { "[['=1','=2'],['=3']]", "[[],['=1','=2']]" }
                    : new[] { "[[1,2],[3]]", "[[],[1,2]]" })
                {
                    Check(report, tool + " 拒绝不等长行 " + data, () =>
                    {
                        var app = new FakeApplication();
                        var result = Execute(app, tool, "{range:'A1:B2'," + key + ":" + data + "}");
                        Require(!result.Ok && result.ErrorCode == "SHAPE_MISMATCH", result.Error);
                        Require(app.ActiveSheet.WriteCalls == 0 && Equals(app.ActiveSheet.Values[1, 1], "原值"), "错误数据仍修改了单元格");
                    });
                }
            }
            Check(report, "合法矩形写入保留显式 null 并完整读回", () =>
            {
                var app = new FakeApplication();
                var result = Execute(app, "write_values", "{range:'A1:B2',values:[[1,2],[3,null]]}");
                Require(result.Ok && app.ActiveSheet.WriteCalls == 1, result.Error);
                Require(app.ActiveSheet.Values[1, 1] == null && Data(result)["verification"][0][1].Value<int>() == 2, "写入或读回不一致");
            });
            Check(report, "合法公式矩形继续执行", () =>
            {
                var app = new FakeApplication();
                var result = Execute(app, "write_formulas", "{range:'A1:B2',formulas:[['=1','=2'],['=3','=4']]}");
                Require(result.Ok && Equals(app.ActiveSheet.Values[1, 1], "=4"), result.Error);
            });
            Check(report, "单元格对象或数组返回参数错误且不写入", () =>
            {
                foreach (var cell in new[] { "{}", "[1]" })
                {
                    var app = new FakeApplication();
                    var result = Execute(app, "write_values", "{range:'A1',values:[[" + cell + "]]}" );
                    Require(!result.Ok && result.ErrorCode == "ARG_INVALID" && app.ActiveSheet.WriteCalls == 0, result.Error);
                }
            });
        }

        private static void FormatValidation(Action<string, bool, string> report)
        {
            foreach (var invalid in new[]
            {
                "font_color:'wrong'", "fill_color:'wrong'", "font_size:0", "font_size:'NaN'",
                "horizontal_alignment:'wrong'", "vertical_alignment:'wrong'", "wrap_text:'wrong'",
            })
            {
                Check(report, "无效格式不留下提前写入的加粗状态：" + invalid, () =>
                {
                    var app = new FakeApplication();
                    var result = Execute(app, "format_range", "{range:'A1',bold:true," + invalid + "}");
                    Require(!result.Ok && result.ErrorCode == "ARG_INVALID", result.Error);
                    Require(!app.ActiveSheet.Font.Bold && app.ActiveSheet.Font.Size == 11, "拒绝参数前已经改了格式");
                });
            }
            Check(report, "合法格式保留字符串标量和可选 null 的兼容性", () =>
            {
                var app = new FakeApplication();
                var result = Execute(app, "format_range", "{range:'A1',bold:'true',italic:null,font_size:'12.5',font_color:'#123456',horizontal_alignment:'center',wrap_text:'false'}");
                Require(result.Ok && app.ActiveSheet.Font.Bold && app.ActiveSheet.Font.Size == 12.5, result.Error);
                Require(app.ActiveSheet.Font.Color == 0x563412, "颜色没有转换为 BGR");
            });
        }

        private static void BoundRanges(Action<string, bool, string> report)
        {
            Check(report, "分页读取使用已解析工作表", () =>
            {
                var app = new FakeApplication();
                var original = app.ActiveSheet;
                var other = new FakeWorkbook("另一工作簿", 2, 2, "其他数据");
                original.AfterResolve = () => app.ActiveWorkbook = other;
                var result = Execute(app, "read_range", "{range:'A1:B2'}");
                Require(result.Ok && Data(result)["values"][0][0].Value<string>() == "原值", result.Error);
            });
            Check(report, "分块写入在宿主切换工作簿后仍写入原工作表并从原表验证", () =>
            {
                var app = new FakeApplication(5001, 1);
                var original = app.ActiveSheet;
                var other = new FakeWorkbook("另一工作簿", 5001, 1, "其他数据");
                original.AfterWrite = () => app.ActiveWorkbook = other;
                var result = new ToolExecutor(() => app).Execute("write_values", LargeWrite());
                Require(result.Ok && Convert.ToInt32(original.Values[5000, 0]) == 5000, result.Error);
                Require(other.Sheet.WriteCalls == 0 && Equals(other.Sheet.Values[5000, 0], "其他数据"), "误写了另一份同名工作表");
                Require(Data(result)["verification_page"]["values"][0][0].Value<int>() == 0, "验证读取了另一份工作簿");
            });
            Check(report, "分块取消回报准确进度且不执行后续块", () =>
            {
                var app = new FakeApplication(5001, 1);
                var executor = new ToolExecutor(() => app);
                var chunks = 0;
                executor.BeforeWriteChunk = () => { if (++chunks == 2) { throw new OperationCanceledException("测试取消"); } };
                var result = executor.Execute("write_values", LargeWrite());
                Require(!result.Ok && result.ErrorCode == "WRITE_PARTIAL" && app.ActiveSheet.WriteCalls == 1, result.Error);
                Require(Data(result)["cells_written"].Value<int>() == 5000 && Data(result)["next_offset"].Value<int>() == 5000, "进度不准确");
                Require(Equals(app.ActiveSheet.Values[5000, 0], "原值"), "取消后仍写入尾格");
            });
            Check(report, "分块宿主错误保留已完成范围并停止", () =>
            {
                var app = new FakeApplication(5001, 1);
                app.ActiveSheet.ThrowOnWrite = 2;
                var result = new ToolExecutor(() => app).Execute("write_values", LargeWrite());
                Require(!result.Ok && result.ErrorCode == "WRITE_PARTIAL" && Data(result)["cells_written"].Value<int>() == 5000, result.Error);
                Require(Equals(app.ActiveSheet.Values[5000, 0], "原值"), "失败后仍写入尾格");
            });
        }

        private static JObject LargeWrite()
        {
            var values = new JArray();
            for (var row = 0; row < 5001; row++) { values.Add(new JArray(row)); }
            return new JObject { ["range"] = "A1:A5001", ["values"] = values };
        }

        private static ToolResult Execute(FakeApplication app, string tool, string json)
        {
            return new ToolExecutor(() => app).Execute(tool, JObject.Parse(json));
        }

        private static JObject Data(ToolResult result) => JObject.FromObject(result.Data);

        private static void Require(bool condition, string detail)
        {
            if (!condition) { throw new InvalidOperationException(detail ?? "断言失败"); }
        }

        private static void Check(Action<string, bool, string> report, string name, Action action)
        {
            try { action(); report(name, true, null); }
            catch (Exception ex) { report(name, false, ex.Message); }
        }

        // 仅提供被测工具实际调用的后期绑定成员，不创建或修改 Office 文档。
        private sealed class FakeApplication
        {
            public FakeApplication(int rows = 2, int columns = 2)
            {
                ActiveWorkbook = new FakeWorkbook("原工作簿", rows, columns, "原值");
            }

            public FakeWorkbook ActiveWorkbook { get; set; }
            public FakeSheet ActiveSheet => ActiveWorkbook.Sheet;
            public bool DisplayAlerts { get; set; } = true;
        }

        private sealed class FakeWorkbook
        {
            public FakeWorkbook(string name, int rows, int columns, object initial)
            {
                Name = name;
                Sheet = new FakeSheet(rows, columns, initial);
                Worksheets = new FakeSheets(Sheet);
            }

            public string Name { get; }
            public FakeSheet Sheet { get; }
            public FakeSheets Worksheets { get; }
            public object Payload { get; set; }
            public int PathCalls { get; private set; }
            public FakeWorkbook Next() { PathCalls++; return this; }
            public void Fail() { throw new InvalidOperationException("模拟宿主失败"); }
        }

        private sealed class FakeSheets
        {
            private readonly FakeSheet _sheet;
            public FakeSheets(FakeSheet sheet) { _sheet = sheet; }
            public int Count => 1;
            public FakeSheet this[object index] => _sheet;
        }

        private sealed class FakeSheet
        {
            public FakeSheet(int rows, int columns, object initial)
            {
                Values = new object[rows, columns];
                for (var row = 0; row < rows; row++)
                    for (var column = 0; column < columns; column++) { Values[row, column] = initial; }
            }

            public string Name => "Sheet1";
            public object[,] Values { get; }
            public FakeFont Font { get; } = new FakeFont();
            public int WriteCalls { get; set; }
            public int ThrowOnWrite { get; set; }
            public Action AfterWrite { get; set; }
            public Action AfterResolve { get; set; }

            [IndexerName("Range")]
            public FakeRange this[string address]
            {
                get
                {
                    if (!AddressSpan.TryParse(address, out var span)) { throw new ArgumentException(address); }
                    var range = new FakeRange(this, address, span);
                    var after = AfterResolve;
                    AfterResolve = null;
                    after?.Invoke();
                    return range;
                }
            }
        }

        private sealed class FakeCount
        {
            public FakeCount(int count) { Count = count; }
            public int Count { get; }
        }

        private sealed class FakeFont
        {
            public bool Bold { get; set; }
            public bool Italic { get; set; }
            public double Size { get; set; } = 11;
            public int Color { get; set; }
        }

        private sealed class FakeRange
        {
            private readonly FakeSheet _sheet;
            private readonly AddressSpan _span;
            public FakeRange(FakeSheet sheet, string address, AddressSpan span)
            {
                _sheet = sheet;
                _span = span;
                Address = address;
            }

            public string Address { get; }
            public int Row => _span.FirstRow;
            public int Column => _span.FirstColumn;
            public FakeCount Areas => new FakeCount(1);
            public FakeCount Rows => new FakeCount(_span.LastRow - Row + 1);
            public FakeCount Columns => new FakeCount(_span.LastColumn - Column + 1);
            public FakeFont Font => _sheet.Font;
            public FakeFont Interior { get; } = new FakeFont();
            public int HorizontalAlignment { get; set; }
            public int VerticalAlignment { get; set; }
            public bool WrapText { get; set; }
            public object Formula { get => Value2; set => Value2 = value; }

            public object Value2
            {
                get
                {
                    var values = new object[Rows.Count, Columns.Count];
                    for (var row = 0; row < Rows.Count; row++)
                        for (var column = 0; column < Columns.Count; column++)
                            values[row, column] = _sheet.Values[Row - 1 + row, Column - 1 + column];
                    return values;
                }
                set
                {
                    if (++_sheet.WriteCalls == _sheet.ThrowOnWrite) { throw new InvalidOperationException("模拟分块写入失败"); }
                    var values = (object[,])value;
                    for (var row = 0; row < Rows.Count; row++)
                        for (var column = 0; column < Columns.Count; column++)
                            _sheet.Values[Row - 1 + row, Column - 1 + column] = values[row, column];
                    _sheet.AfterWrite?.Invoke();
                }
            }
        }
    }
}
