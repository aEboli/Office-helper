using System;
using System.Collections.Generic;
using System.Linq;
using ChatSheet.AddIn.Agent;
using ChatSheet.AddIn.Tools;
using ChatWord.AddIn.Agent;
using ChatWord.AddIn.Hosts;
using ChatWord.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class WordTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            TestCatalog(report);
            TestPrompt(report);
            TestTargetAndRead(report);
            TestHostIdentification(report);
            TestRibbonTerminology(report);
            TestApprovalContract(report);
        }

        private static void TestCatalog(Action<string, bool, string> report)
        {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var tool in WordToolCatalog.All)
            {
                names.Add(tool.Name);
                report("Word 工具参数有 Schema：" + tool.Name, tool.Parameters != null, "缺少参数 Schema");
                report("Word 工具不暴露 Excel 对象：" + tool.Name,
                    !tool.Description.Contains("Workbook") && !tool.Description.Contains("Worksheet") && !tool.Description.Contains("A1"),
                    "描述包含 Excel 术语");
            }
            report("包含文档信息工具", names.Contains("get_document_info"), "缺少 get_document_info");
            report("包含明确目标工具", names.Contains("read_range") && names.Contains("replace_text"), "缺少读写工具");
            var paragraphTool = WordToolCatalog.All.FirstOrDefault(tool => tool.Name == "format_paragraph");
            var indentSchema = paragraphTool == null ? null : JObject.FromObject(paragraphTool.Parameters)["properties"]?["first_line_indent_chars"];
            report("段落格式提供字符单位的首行缩进", indentSchema?["type"]?.Value<string>() == "number", "缺少 first_line_indent_chars 数值参数");
        }

        private static void TestPrompt(Action<string, bool, string> report)
        {
            var prompt = WordSystemPrompt.Build(
                new WordDocumentSummary { HasDocument = true, Name = "测试文档.docx", Host = "Microsoft Word", Version = "16.0", Pages = 2 },
                new WordSelectionInfo { HasSelection = true, Story = "main_text", Start = 4, End = 8, Text = "摘要" });
            report("提示词使用 Word 身份", prompt.Contains("Word 文档助手"), "缺少身份");
            report("提示词要求先读后写", prompt.Contains("先读取文档结构"), "缺少先读后写边界");
            report("提示词说明顾问模式", prompt.Contains("顾问模式"), "缺少顾问模式边界");
            report("提示词限定折叠光标定位", prompt.Contains("target 只传 story、start、end"), "插入目标未限定为单一字符范围");
            report("格式规范要求用户明确点名", prompt.Contains("格式规范只在用户明确提到文种、格式或标准时应用"), "没有限制格式规范触发条件");
            report("格式规范不得按正文语言推断", prompt.Contains("不能仅凭正文语言推断版式"), "语言被用作隐式格式触发条件");
            report("未指定书信版式变体时先澄清", prompt.Contains("具体变体会改变布局且未指定时，先澄清"), "缺少书信格式变体澄清规则");
            report("格式仅作用于本轮新增或指定段落", prompt.Contains("仅格式化本轮新增或用户明确指定的段落"), "格式范围可能扩展到整篇文档");
            report("提示词区分 full block 与 modified block", prompt.Contains("full block 时，信函各部分左对齐") && prompt.Contains("modified block 时，正文和地址仍左对齐"), "英文商务信函版式细节不完整");
            report("提示词写明中文书信常见惯例", prompt.Contains("称呼另起一行顶格") && prompt.Contains("首行缩进两字"), "缺少中文书信格式细节");
            report("提示词区分中英文称呼与开场", prompt.Contains("中文称谓可用") && prompt.Contains("英文商务称呼通常用 Dear") && prompt.Contains("正文首段自然说明目的"), "缺少中英文书信措辞差异");
            report("书信称呼开头避免逐词直译", prompt.Contains("不要将“尊敬的领导/您好/冒昧打扰”逐词翻译成英文"), "缺少中英文称呼和开头差异规则");
            report("提示词没有工作簿术语", !prompt.Contains("工作簿") && !prompt.Contains("工作表"), "混入 Excel 术语");
        }

        private static void TestTargetAndRead(Action<string, bool, string> report)
        {
            var document = new FakeDocument();
            var target = WordTargetResolver.Resolve(document, new JObject
            {
                ["target"] = new JObject { ["story"] = "main_text", ["start"] = 2, ["end"] = 7 },
            });
            report("Story Start/End 目标解析", target.Start == 2 && target.End == 7 && target.Story == "main_text", "目标位置不正确");
            target.Dispose();

            var app = new FakeApplication { Name = "Microsoft Word", Version = "16.0", ActiveDocument = document };
            var executor = new WordToolExecutor(() => app, "Word.Application");
            var result = executor.Execute("read_range", new JObject
            {
                ["target"] = new JObject { ["story"] = "main_text", ["start"] = 0, ["end"] = 20 },
            });
            report("读取范围去除段落/单元格标记", result.Ok && result.Data != null && !result.Data.ToString().Contains("\a"), result.Error ?? "读取失败");

            var formatted = executor.Execute("format_paragraph", new JObject
            {
                ["target"] = new JObject { ["story"] = "main_text", ["start"] = 0, ["end"] = 4 },
                ["alignment"] = "left",
                ["first_line_indent_chars"] = 2,
            });
            var formattedData = formatted.Ok ? JObject.FromObject(formatted.Data) : new JObject();
            report("段落首行缩进按字符写入并读回", formatted.Ok &&
                formattedData["readback"]?.Value<string>("firstLineIndentChars") == "2",
                formatted.Error ?? "首行缩进读回值不匹配");

            var alignmentDocument = new FakeDocument();
            var alignmentRange = (FakeRange)alignmentDocument.StoryRanges.Item(1);
            var alignmentExecutor = new WordToolExecutor(() => new FakeApplication
            {
                Name = "Microsoft Word", Version = "16.0", ActiveDocument = alignmentDocument,
            }, "Word.Application");
            var invalidAlignment = alignmentExecutor.Execute("format_paragraph", new JObject
            {
                ["target"] = new JObject { ["story"] = "main_text", ["start"] = 0, ["end"] = 4 },
                ["alignment"] = "diagonal",
                ["first_line_indent_chars"] = 2,
            });
            report("无效对齐值先验证且不写入缩进", !invalidAlignment.Ok &&
                alignmentRange.ParagraphFormat.CharacterUnitFirstLineIndent == 0,
                invalidAlignment.Error ?? "无效对齐值导致部分格式被写入");

            var defaultedLocators = WordTargetResolver.Resolve(document, new JObject
            {
                ["target"] = new JObject
                {
                    ["story"] = "main_text", ["start"] = 0, ["end"] = 20,
                    ["paragraph_index"] = 0, ["table_index"] = 0, ["row"] = 0, ["column"] = 0,
                },
            });
            report("忽略模型补出的零值定位字段", defaultedLocators.Start == 0 && defaultedLocators.End > 0, "零值定位字段导致目标歧义");
            defaultedLocators.Dispose();

            var insertion = executor.Execute("insert_text", new JObject
            {
                ["target"] = new JObject
                {
                    ["story"] = "main_text", ["start"] = 0, ["end"] = 0,
                    ["paragraph_index"] = 0, ["table_index"] = 0, ["row"] = 0, ["column"] = 0,
                    ["bookmark"] = "", ["content_control"] = "",
                },
                ["text"] = "插入成功",
            });
            var insertionData = insertion.Ok ? JObject.FromObject(insertion.Data) : new JObject();
            report("折叠光标插入忽略模型补出的空定位项", insertion.Ok && insertionData.Value<string>("text") == "插入成功",
                insertion.Error ?? "折叠光标未能写入");

            var conflictingTarget = executor.Execute("insert_text", new JObject
            {
                ["target"] = new JObject { ["story"] = "main_text", ["start"] = 0, ["end"] = 4, ["paragraph_index"] = 1 },
                ["text"] = "不应写入",
            });
            report("拒绝真正冲突的定位字段", !conflictingTarget.Ok && conflictingTarget.ErrorCode == "TARGET_AMBIGUOUS",
                conflictingTarget.Error ?? "冲突目标未被拒绝");

            var partialDocument = new FakeDocument(throwForMissingStories: true);
            var partialExecutor = new WordToolExecutor(() => new FakeApplication
            {
                Name = "Microsoft Word", Version = "16.0", ActiveDocument = partialDocument,
            }, "Word.Application");
            var structure = partialExecutor.Execute("read_document_structure", new JObject());
            var structureData = structure.Ok ? JObject.FromObject(structure.Data) : new JObject();
            var availableStories = structureData["stories"] as JArray;
            var storyErrors = structureData["storyErrors"] as JArray;
            report("单个缺失 Story 不阻断结构读取", structure.Ok && structureData.Value<bool?>("partial") == true &&
                availableStories?.Any(item => item.Value<string>("story") == "main_text") == true &&
                storyErrors?.Any(item => item.Value<string>("story") == "footnotes") == true,
                structure.Error ?? "缺少可用正文或 Story 错误详情");

            var paragraphs = executor.Execute("read_paragraphs", new JObject { ["story"] = "main_text", ["limit"] = 10 });
            report("按 Story 读取段落", paragraphs.Ok, paragraphs.Error ?? "读取失败");

            var table = executor.Execute("read_table", new JObject { ["story"] = "main_text", ["table_index"] = 1, ["offset"] = 1, ["limit"] = 1 });
            var tableData = table.Ok ? JObject.FromObject(table.Data) : new JObject();
            report("读取表格尺寸与行分页", table.Ok && tableData.Value<int?>("rows") == 3 && tableData.Value<int?>("columns") == 2 &&
                tableData.Value<int?>("next_offset") == 2 && (tableData["data"] as JArray)?.Count == 1,
                table.Error ?? "尺寸或分页结果不正确");

            var preview = JObject.FromObject(executor.BuildPreview("replace_text", new JObject
            {
                ["target"] = new JObject { ["story"] = "main_text", ["start"] = 2, ["end"] = 7 },
                ["find"] = "目标", ["replace"] = "预览",
            }));
            report("Word 审批预览只读计算前后文本", preview.Value<bool?>("supported") == true &&
                preview.Value<string>("before") == "目标文本" && preview.Value<string>("after") == "预览文本", "预览不正确");
        }

        private static void TestHostIdentification(Action<string, bool, string> report)
        {
            var word = new FakeApplication { Name = "Microsoft Word", Version = "16.0" };
            var wps = new FakeApplication { Name = "Microsoft Word", Version = "12.0" };
            report("Word ProgID 识别", WordHostProbe.Detect(word, "Word.Application") == WordHostKind.MicrosoftWord, "错误识别 Word");
            report("WPS ProgID 优先于 Name", WordHostProbe.Detect(wps, "kwps.Application") == WordHostKind.WpsWriter, "WPS 被误识别为 Word");
        }

        private static void TestRibbonTerminology(Action<string, bool, string> report)
        {
            var assembly = typeof(ChatWord.AddIn.WordComAddIn).Assembly;
            using (var stream = assembly.GetManifestResourceStream("ChatWord.AddIn.Resources.Ribbon.xml"))
            using (var reader = new System.IO.StreamReader(stream))
            {
                var ribbon = reader.ReadToEnd();
                report("Word Ribbon 使用文档术语", ribbon.Contains("文档") && ribbon.Contains("Word/Writer"), "缺少 Word/Writer 文案");
                report("Word Ribbon 不含 Excel 专属按钮", !ribbon.Contains("工作簿") && !ribbon.Contains("工作表") && !ribbon.Contains("适配当前表"), "混入 Excel 专属术语");
            }
        }

        private static void TestApprovalContract(Action<string, bool, string> report)
        {
            var writes = 0;
            var reads = 0;
            var safe = true;
            foreach (var tool in WordToolCatalog.All)
            {
                if (tool.Risk == ToolRisk.Read)
                {
                    reads++;
                    safe &= !tool.RequiresApproval;
                }
                else
                {
                    writes++;
                    safe &= tool.RequiresApproval;
                }
            }
            report("Word 读工具不弹审批、写工具必须审批", safe && writes > 0 && reads > 0, "审批风险分类不完整");
        }

        private sealed class FakeApplication
        {
            public string Name { get; set; }
            public string Version { get; set; }
            public FakeDocument ActiveDocument { get; set; }
        }

        private sealed class FakeDocument
        {
            public FakeDocument(bool throwForMissingStories = false)
            {
                var table = new FakeTable(new[,] { { "类别", "具体内容" }, { "消费品", "商品" }, { "服务", "乐园" } });
                StoryRanges = new FakeCollection(new Dictionary<int, object>
                {
                    [1] = new FakeRange("第一段\r", 0, 20, new FakeCollection(new Dictionary<int, object> { [1] = table })),
                }, throwForMissingStories);
            }

            public string Name => "测试文档.docx";
            public string FullName => "";
            public bool Saved => true;
            public bool ReadOnly => false;
            public int ProtectionType => 0;
            public bool TrackRevisions => false;
            public FakeCollection StoryRanges { get; }
            public FakeCollection Paragraphs { get; } = new FakeCollection();
            public FakeCollection Tables { get; } = new FakeCollection();
            public FakeCollection Sections { get; } = new FakeCollection();
            public FakeCollection Bookmarks { get; } = new FakeCollection();
            public FakeCollection ContentControls { get; } = new FakeCollection();
            public FakeCollection Fields { get; } = new FakeCollection();
            public FakeCollection Comments { get; } = new FakeCollection();
            public FakeCollection Footnotes { get; } = new FakeCollection();
            public FakeCollection Endnotes { get; } = new FakeCollection();
            public int ComputeStatistics(int statistic) => 1;
            public FakeRange Range(int start, int end) => new FakeRange("目标文本", start, end);
        }

        private sealed class FakeCollection
        {
            private readonly Dictionary<int, object> _items;
            private readonly bool _throwWhenMissing;
            public FakeCollection() : this(new Dictionary<int, object>()) { }
            public FakeCollection(int count)
            {
                _items = new Dictionary<int, object>();
                for (var i = 1; i <= count; i++) { _items[i] = new object(); }
            }
            public FakeCollection(Dictionary<int, object> items) : this(items, false) { }
            public FakeCollection(Dictionary<int, object> items, bool throwWhenMissing) { _items = items; _throwWhenMissing = throwWhenMissing; }
            public int Count => _items.Count;
            public object Item(object index)
            {
                if (_items.TryGetValue(Convert.ToInt32(index), out var item)) { return item; }
                if (_throwWhenMissing) { throw new InvalidOperationException("Story " + index + " does not exist."); }
                return null;
            }
        }

        private sealed class FakeRange
        {
            private string _text;
            private readonly FakeCollection _tables;
            private readonly FakeParagraphFormat _paragraphFormat;
            public FakeRange(string text, int start, int end, FakeCollection tables = null, FakeParagraphFormat paragraphFormat = null)
            { _text = text; Start = start; End = end; _tables = tables ?? new FakeCollection(); _paragraphFormat = paragraphFormat ?? new FakeParagraphFormat(); }
            public int Start { get; private set; }
            public int End { get; private set; }
            public string Text { get { return _text; } set { _text = value; End = Start + (value ?? string.Empty).Length; } }
            public FakeRange Duplicate() { return new FakeRange(_text, Start, End, _tables, _paragraphFormat); }
            public void SetRange(int start, int end) { Start = start; End = end; }
            public int StoryType => 1;
            public FakeCollection Paragraphs => new FakeCollection();
            public FakeParagraphFormat ParagraphFormat => _paragraphFormat;
            public FakeCollection Tables => _tables;
            public bool Hidden => false;
            public FakeCollection Fields => new FakeCollection();
        }

        private sealed class FakeParagraphFormat
        {
            public int Alignment { get; set; }
            public double CharacterUnitFirstLineIndent { get; set; }
            public double SpaceBefore { get; set; }
            public double SpaceAfter { get; set; }
            public bool KeepWithNext { get; set; }
        }

        private sealed class FakeTable
        {
            private readonly string[,] _cells;
            public FakeTable(string[,] cells) { _cells = cells; }
            public FakeCollection Rows => new FakeCollection(_cells.GetLength(0));
            public FakeCollection Columns => new FakeCollection(_cells.GetLength(1));
            public FakeCell Cell(int row, int column) => new FakeCell(_cells[row - 1, column - 1]);
        }

        private sealed class FakeCell
        {
            public FakeCell(string text) { Range = new FakeRange(text + "\a", 0, text.Length + 1); }
            public FakeRange Range { get; }
        }
    }
}
