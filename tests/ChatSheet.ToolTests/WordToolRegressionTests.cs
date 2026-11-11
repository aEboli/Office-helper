using System;
using System.Collections.Generic;
using ChatSheet.AddIn.Tools;
using ChatWord.AddIn.Hosts;
using ChatWord.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class WordToolRegressionTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            TableEdits(report);
            TextArguments(report);
            StoryTargets(report);
            Readback(report);
            PageBreaks(report);
        }

        private static void TableEdits(Action<string, bool, string> report)
        {
            Check(report, "Word 表格编辑保留嵌套定位并替换已有文字", () =>
            {
                var document = new FakeDocument();
                var args = CellArgs("新内容");
                var original = args.DeepClone();
                var executor = Executor(document);
                var preview = JObject.FromObject(executor.BuildPreview("edit_table", args));
                var result = executor.Execute("edit_table", args);
                return preview.Value<bool?>("supported") == true && preview.Value<string>("before") == "单元格"
                    && result.Ok && document.Main.Text == "前缀新内容\r\a后缀\r" && JToken.DeepEquals(args, original);
            });
            Check(report, "Word 表格编辑兼容顶层行列定位", () =>
            {
                var document = new FakeDocument();
                var args = CellArgs("新内容");
                foreach (var key in new[] { "table_index", "row", "column" })
                {
                    args[key] = args["target"][key].DeepClone();
                    ((JObject)args["target"]).Remove(key);
                }
                return Executor(document).Execute("edit_table", args).Ok && document.Main.Text == "前缀新内容\r\a后缀\r";
            });
            Check(report, "Word 表格定位冲突时拒绝修改", () =>
            {
                var document = new FakeDocument();
                var args = CellArgs("不应写入");
                args["row"] = 2;
                var result = Executor(document).Execute("edit_table", args);
                return !result.Ok && result.ErrorCode == "TARGET_AMBIGUOUS" && document.Main.Writes == 0;
            });
            Check(report, "Word 表格工具不接受普通字符区间", () =>
            {
                var document = new FakeDocument();
                var args = RangeArgs(0, 0);
                args["text"] = "不应写入";
                return !Executor(document).Execute("edit_table", args).Ok && document.Main.Writes == 0;
            });
        }

        private static void TextArguments(Action<string, bool, string> report)
        {
            foreach (var name in new[] { "replace_text", "insert_text", "edit_table" })
            {
                foreach (var invalid in new[] { "missing", "null", "number" })
                {
                    Check(report, "Word " + name + " 拒绝 " + invalid + " 文本且不写入", () =>
                    {
                        var document = new FakeDocument();
                        var args = name == "edit_table" ? CellArgs("临时") : RangeArgs(0, name == "insert_text" ? 0 : 2);
                        var key = name == "replace_text" ? "replace" : "text";
                        args.Remove(key);
                        if (invalid == "null") { args[key] = JValue.CreateNull(); }
                        if (invalid == "number") { args[key] = 123; }
                        var result = Executor(document).Execute(name, args);
                        return !result.Ok && result.ErrorCode == (invalid == "number" ? "ARG_INVALID" : "ARG_MISSING")
                            && document.Main.Writes == 0;
                    });
                }
                foreach (var text in new[] { string.Empty, "  " })
                {
                    Check(report, "Word " + name + " 保留显式空白文本语义 " + text.Length, () =>
                    {
                        var document = new FakeDocument();
                        var args = name == "edit_table" ? CellArgs(text) : RangeArgs(0, name == "insert_text" ? 0 : 2);
                        args[name == "replace_text" ? "replace" : "text"] = text;
                        var result = Executor(document).Execute(name, args);
                        return result.Ok && (name == "edit_table"
                            ? document.Main.Text == "前缀" + text + "\r\a后缀\r"
                            : document.Main.Text.StartsWith(text + (name == "insert_text" ? "前缀" : "单元格"), StringComparison.Ordinal));
                    });
                }
            }
        }

        private static void StoryTargets(Action<string, bool, string> report)
        {
            foreach (var locator in new[] { "bookmark", "content_control" })
            {
                Check(report, "Word " + locator + " 使用实际 Story 且不承诺正文撤销", () =>
                {
                    var document = new FakeDocument();
                    var args = new JObject { ["target"] = new JObject { [locator] = "页眉目标" }, ["replace"] = "新页眉" };
                    var result = Executor(document).Execute("replace_text", args, "header-undo");
                    var data = result.Ok ? JObject.FromObject(result.Data) : new JObject();
                    return result.Ok && data.Value<string>("story") == "primary_header" && data.Value<bool?>("canUndo") == false
                        && document.Header.Text == "新页眉\r" && document.Main.Writes == 0;
                });
                Check(report, "Word " + locator + " 拒绝与声明 Story 不符的目标", () =>
                {
                    var document = new FakeDocument();
                    var args = new JObject
                    {
                        ["target"] = new JObject { ["story"] = "main_text", [locator] = "页眉目标" },
                        ["replace"] = "不应写入",
                    };
                    var result = Executor(document).Execute("replace_text", args);
                    return !result.Ok && result.ErrorCode == "TARGET_INVALID" && document.Header.Writes == 0 && document.Main.Writes == 0;
                });
            }
        }

        private static void Readback(Action<string, bool, string> report)
        {
            Check(report, "Word 删除后读回失败不能报告成功", () =>
            {
                var document = new FakeDocument();
                document.Main.FailReadsAfterWrite = true;
                var executor = Executor(document);
                var result = executor.Execute("delete_range", RangeArgs(0, 2), "failed-delete");
                return !result.Ok && result.ErrorCode == "READBACK_FAILED" && document.Main.Writes == 1
                    && !executor.Undo.TryUndo("failed-delete", document, out _);
            });
            Check(report, "Word 修改前读取失败不得以空文本继续写入", () =>
            {
                var document = new FakeDocument();
                document.Main.FailReads = true;
                var args = RangeArgs(0, 2);
                args["replace"] = "新内容";
                return !Executor(document).Execute("replace_text", args).Ok && document.Main.Writes == 0;
            });
            Check(report, "Word 读取失败不得返回成功的空文本", () =>
            {
                var document = new FakeDocument();
                document.Main.FailReads = true;
                return !Executor(document).Execute("read_range", RangeArgs(0, 2)).Ok;
            });
            Check(report, "Word 写入结果返回实际更新后的范围", () =>
            {
                var document = new FakeDocument();
                var args = RangeArgs(2, 2);
                args["text"] = "新增文字";
                var result = Executor(document).Execute("insert_text", args);
                var data = result.Ok ? JObject.FromObject(result.Data) : new JObject();
                return result.Ok && data.Value<int?>("start") == 2 && data.Value<int?>("end") == 6;
            });
            Check(report, "Word 撤销恢复空文本时读回失败不能报告成功", () =>
            {
                var document = new FakeDocument();
                var executor = Executor(document);
                var args = RangeArgs(0, 0);
                args["text"] = "插入";
                var inserted = executor.Execute("insert_text", args, "undo-readback");
                document.Main.FailReadsAtWrite = 2;
                return inserted.Ok && !executor.Undo.TryUndo("undo-readback", document, out _);
            });
            Check(report, "Word 缺少文档时结构读取返回失败", () =>
            {
                var result = Executor(null).Execute("read_document_structure", new JObject());
                return !result.Ok && result.ErrorCode == "DOCUMENT_REQUIRED";
            });
        }

        private static void PageBreaks(Action<string, bool, string> report)
        {
            Check(report, "Word 插入分页符不能替换非空选区", () =>
            {
                var document = new FakeDocument();
                var result = Executor(document).Execute("insert_page_break", RangeArgs(0, 2));
                return !result.Ok && result.ErrorCode == "TARGET_INVALID" && document.Main.Writes == 0;
            });
            Check(report, "Word 分页符未实际写入时不能报告成功", () =>
            {
                var document = new FakeDocument();
                document.Main.IgnoreWrites = true;
                var result = Executor(document).Execute("insert_page_break", RangeArgs(0, 0));
                return !result.Ok && result.ErrorCode == "READBACK_FAILED";
            });
            Check(report, "Word 折叠光标插入分页符后读回", () =>
            {
                var document = new FakeDocument();
                var result = Executor(document).Execute("insert_page_break", RangeArgs(0, 0));
                return result.Ok && document.Main.Text.StartsWith("\f前缀", StringComparison.Ordinal);
            });
        }

        private static JObject RangeArgs(int start, int end) => new JObject
        { ["target"] = new JObject { ["story"] = "main_text", ["start"] = start, ["end"] = end } };

        private static JObject CellArgs(string text) => new JObject
        {
            ["target"] = new JObject { ["story"] = "main_text", ["table_index"] = 1, ["row"] = 1, ["column"] = 1 },
            ["text"] = text,
        };

        private static WordToolExecutor Executor(FakeDocument document) => new WordToolExecutor(
            () => new FakeApplication { ActiveDocument = document }, "Word.Application");

        private static void Check(Action<string, bool, string> report, string name, Func<bool> test)
        {
            try { report(name, test(), "回归断言失败"); }
            catch (Exception ex) { report(name, false, ex.GetBaseException().Message); }
        }

        private sealed class FakeApplication
        {
            public FakeDocument ActiveDocument { get; set; }
        }

        private sealed class FakeDocument
        {
            internal readonly FakeStory Main = new FakeStory(1, "前缀单元格\r\a后缀\r");
            internal readonly FakeStory Header = new FakeStory(7, "页眉文字\r");
            public int ProtectionType => -1;
            public bool ReadOnly => false;
            public bool TrackRevisions => false;
            public FakeStories StoryRanges => new FakeStories(Main, Header);
            public FakeCollection Bookmarks => new FakeCollection(new Dictionary<object, object>
            { ["页眉目标"] = new FakeNamedRange(Header) });
            public FakeCollection ContentControls => new FakeCollection(new Dictionary<object, object>
            { [1] = new FakeNamedRange(Header) });
            public FakeRange Range(int start, int end) => new FakeRange(Main, start, end);
        }

        private sealed class FakeStory
        {
            internal FakeStory(int type, string text) { Type = type; Text = text; }
            internal readonly int Type;
            internal string Text;
            internal int Writes;
            internal bool FailReads;
            internal bool FailReadsAfterWrite;
            internal int FailReadsAtWrite = int.MaxValue;
            internal bool IgnoreWrites;
        }

        private sealed class FakeStories
        {
            private readonly FakeStory _main;
            private readonly FakeStory _header;
            internal FakeStories(FakeStory main, FakeStory header) { _main = main; _header = header; }
            public FakeRange Item(int index)
            {
                var story = index == 1 ? _main : index == 7 ? _header : null;
                return story == null ? null : new FakeRange(story, 0, story.Text.Length);
            }
        }

        private sealed class FakeCollection
        {
            private readonly Dictionary<object, object> _items;
            internal FakeCollection(Dictionary<object, object> items = null) { _items = items ?? new Dictionary<object, object>(); }
            public int Count => _items.Count;
            public object Item(object key) => _items.TryGetValue(key, out var item) ? item : null;
        }

        private sealed class FakeNamedRange
        {
            private readonly FakeStory _story;
            internal FakeNamedRange(FakeStory story) { _story = story; }
            public string Title => "页眉目标";
            public string Tag => "页眉目标";
            public FakeRange Range => new FakeRange(_story, 0, _story.Text.Length - 1);
        }

        private sealed class FakeRange
        {
            private readonly FakeStory _story;
            internal FakeRange(FakeStory story, int start, int end) { _story = story; SetRange(start, end); }
            public int Start { get; private set; }
            public int End { get; private set; }
            public int StoryType => _story.Type;
            public FakeRange Duplicate => new FakeRange(_story, Start, End);
            public FakeCollection Fields => new FakeCollection();
            public FakeCollection Bookmarks => new FakeCollection();
            public FakeCollection ContentControls => new FakeCollection();
            public FakeCollection Revisions => new FakeCollection();
            public FakeCollection Tables => new FakeCollection(new Dictionary<object, object> { [1] = new FakeTable(_story) });
            public string Text
            {
                get
                {
                    if (_story.FailReads || (_story.FailReadsAfterWrite && _story.Writes > 0) || _story.Writes >= _story.FailReadsAtWrite)
                    { throw new InvalidOperationException("模拟宿主无法读取 Text。"); }
                    return _story.Text.Substring(Start, End - Start);
                }
                set
                {
                    _story.Writes++;
                    if (_story.IgnoreWrites) { return; }
                    _story.Text = _story.Text.Substring(0, Start) + value + _story.Text.Substring(End);
                    End = Start + value.Length;
                }
            }
            public void SetRange(int start, int end)
            {
                if (start < 0 || end < start || end > _story.Text.Length) { throw new ArgumentOutOfRangeException(); }
                Start = start; End = end;
            }
            public void InsertBreak(int type) { Text = "\f"; }
        }

        private sealed class FakeTable
        {
            private readonly FakeStory _story;
            internal FakeTable(FakeStory story) { _story = story; }
            public FakeNamedCell Cell(int row, int column)
            {
                if (row != 1 || column != 1) { throw new ArgumentOutOfRangeException(); }
                return new FakeNamedCell(_story);
            }
        }

        private sealed class FakeNamedCell
        {
            private readonly FakeStory _story;
            internal FakeNamedCell(FakeStory story) { _story = story; }
            public FakeRange Range => new FakeRange(_story, 2, _story.Text.IndexOf('\a') + 1);
        }
    }
}
