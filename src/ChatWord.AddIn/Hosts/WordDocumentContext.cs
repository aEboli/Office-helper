using System;
using System.Collections.Generic;
using System.Text;
using ChatSheet.AddIn;
using Newtonsoft.Json.Linq;

namespace ChatWord.AddIn.Hosts
{
    internal sealed class WordDocumentSummary
    {
        internal bool HasDocument { get; set; }
        internal string Name { get; set; }
        internal string Path { get; set; }
        internal bool Saved { get; set; }
        internal bool ReadOnly { get; set; }
        internal string Host { get; set; }
        internal string Version { get; set; }
        internal string Build { get; set; }
        internal int Pages { get; set; }
        internal int Sections { get; set; }
        internal int Paragraphs { get; set; }
        internal int Tables { get; set; }
        internal int ProtectionType { get; set; }
        internal bool TrackRevisions { get; set; }
        internal object Capabilities { get; set; }

        internal string ToPromptText()
        {
            if (!HasDocument) { return "当前没有打开的 Word 文档。"; }
            return string.Format(
                "宿主：{0} {1}（Build {2}）；文档：{3}；路径：{4}；保存状态：{5}；只读：{6}；页数：{7}；节数：{8}；段落：{9}；表格：{10}；保护类型：{11}；修订跟踪：{12}。",
                Host, Version, Build, Name, string.IsNullOrEmpty(Path) ? "未保存" : Path,
                Saved ? "已保存" : "未保存", ReadOnly ? "是" : "否", Pages, Sections,
                Paragraphs, Tables, ProtectionType, TrackRevisions ? "开启" : "关闭");
        }
    }

    internal sealed class WordSelectionInfo
    {
        internal bool HasSelection { get; set; }
        internal string Story { get; set; }
        internal int StoryType { get; set; }
        internal int Start { get; set; }
        internal int End { get; set; }
        internal string Text { get; set; }
        internal bool TextIncluded { get; set; } = true;
        internal string ParagraphStyle { get; set; }
        internal int HeadingLevel { get; set; }
        internal int? TableIndex { get; set; }
        internal int? Row { get; set; }
        internal int? Column { get; set; }

        internal string ToPromptText()
        {
            return HasSelection
                ? string.Format("当前选区：Story={0}，Start={1}，End={2}，段落样式={3}，标题层级={4}，表格位置={5}，{6}",
                    Story, Start, End, ParagraphStyle ?? "未知", HeadingLevel, TableIndex.HasValue ?
                        string.Format("表 {0} 行 {1} 列 {2}", TableIndex, Row, Column) : "不在表格中",
                    TextIncluded ? "摘要：" + Text : "选区文字未纳入上下文")
                : "当前没有可用选区。";
        }
    }

    internal sealed class WordDocumentContext
    {
        private readonly Func<object> _applicationAccessor;
        private readonly string _registeredProgId;

        internal WordDocumentContext(Func<object> applicationAccessor, string registeredProgId = null)
        {
            _applicationAccessor = applicationAccessor ?? throw new ArgumentNullException(nameof(applicationAccessor));
            _registeredProgId = registeredProgId;
        }

        internal object Application => _applicationAccessor();

        internal object ActiveDocument()
        {
            var application = Application;
            return application == null || !WordCom.TryGet(application, "ActiveDocument", out var document)
                ? null : document;
        }

        internal static string DocumentKey(WordDocumentSummary summary)
        {
            if (summary == null || !summary.HasDocument) { return null; }
            var identity = !string.IsNullOrWhiteSpace(summary.Path) ? "path:" + summary.Path : "name:" + summary.Name;
            return string.IsNullOrWhiteSpace(identity) ? null : identity.Trim().ToUpperInvariant();
        }

        internal WordDocumentSummary GetSummary()
        {
            object document = null;
            try
            {
                document = ActiveDocument();
                var application = Application;
                if (document == null) { return new WordDocumentSummary { Host = WordHostProbe.DisplayName(WordHostKind.Unknown) }; }
                var kind = WordHostProbe.Detect(application, _registeredProgId);
                object paragraphCollection = null; object tableCollection = null; object sectionCollection = null;
                var paragraphs = WordCom.TryGet(document, "Paragraphs", out paragraphCollection)
                    ? WordCom.Int(paragraphCollection, "Count") : 0;
                var tables = WordCom.TryGet(document, "Tables", out tableCollection)
                    ? WordCom.Int(tableCollection, "Count") : 0;
                var sections = WordCom.TryGet(document, "Sections", out sectionCollection)
                    ? WordCom.Int(sectionCollection, "Count") : 0;
                var pages = 0;
                if (WordCom.TryCall(document, "ComputeStatistics", out var statistic, 2) && statistic != null)
                {
                    pages = Convert.ToInt32(statistic);
                }

                var summary = new WordDocumentSummary
                {
                    HasDocument = true,
                    Name = WordCom.String(document, "Name"),
                    Path = WordCom.String(document, "FullName"),
                    Saved = WordCom.Bool(document, "Saved"),
                    ReadOnly = WordCom.Bool(document, "ReadOnly"),
                    Host = WordHostProbe.DisplayName(kind),
                    Version = WordCom.String(application, "Version"),
                    Build = WordCom.String(application, "Build"),
                    Pages = pages,
                    Sections = sections,
                    Paragraphs = paragraphs,
                    Tables = tables,
                    ProtectionType = WordCom.Int(document, "ProtectionType", -1),
                    TrackRevisions = WordCom.Bool(document, "TrackRevisions"),
                    Capabilities = CapabilityPayload(document),
                };
                WordCom.Release(paragraphCollection); WordCom.Release(tableCollection); WordCom.Release(sectionCollection);
                return summary;
            }
            finally { WordCom.Release(document); }
        }

        internal WordSelectionInfo GetSelection(bool includeText = true)
        {
            object selection = null;
            object paragraph = null;
            object paragraphs = null;
            try
            {
                var application = Application;
                if (application == null || !WordCom.TryGet(application, "Selection", out selection) || selection == null)
                {
                    return new WordSelectionInfo();
                }

                var text = includeText ? Sanitize(WordCom.String(selection, "Text")) : string.Empty;
                var storyType = WordCom.Int(selection, "StoryType");
                var result = new WordSelectionInfo
                {
                    HasSelection = true,
                    StoryType = storyType,
                    Story = StoryName(storyType),
                    Start = WordCom.Int(selection, "Start"),
                    End = WordCom.Int(selection, "End"),
                    Text = includeText ? Limit(text, 1200) : string.Empty,
                    TextIncluded = includeText,
                };

                if (WordCom.TryGet(selection, "Paragraphs", out paragraphs) && paragraphs != null && WordCom.Int(paragraphs, "Count") > 0)
                {
                    paragraph = WordCom.Item(paragraphs, 1);
                    if (paragraph != null)
                    {
                        object style = null;
                        if (WordCom.TryGet(paragraph, "Style", out style) && style != null)
                        {
                            result.ParagraphStyle = WordCom.String(style, "NameLocal", WordCom.String(style, "Name"));
                        }
                        result.HeadingLevel = WordCom.Int(paragraph, "OutlineLevel");
                        result.TableIndex = FindTablePosition(selection, out var row, out var column);
                        result.Row = row;
                        result.Column = column;
                        WordCom.Release(style);
                    }
                }
                return result;
            }
            finally
            {
                WordCom.Release(paragraph);
                WordCom.Release(paragraphs);
                WordCom.Release(selection);
            }
        }

        internal object ReadStructure(int maxStories = 32)
        {
            object document = null;
            try
            {
                document = ActiveDocument();
                if (document == null) { return new { ok = false, errorCode = "DOCUMENT_REQUIRED", error = "当前没有打开的文档。" }; }
                var stories = new List<object>();
                var storyErrors = new List<object>();
                foreach (var pair in KnownStories)
                {
                    if (stories.Count >= maxStories) { break; }
                    object ranges = null;
                    object range = null;
                    try
                    {
                        if (!WordCom.TryGet(document, "StoryRanges", out ranges) || ranges == null) { continue; }
                        range = WordCom.Item(ranges, pair.Key);
                        if (range == null) { continue; }
                        stories.Add(new
                        {
                            story = pair.Value,
                            storyType = pair.Key,
                            start = WordCom.Int(range, "Start"),
                            end = WordCom.Int(range, "End"),
                            characters = Math.Max(0, WordCom.Int(range, "End") - WordCom.Int(range, "Start")),
                        });
                    }
                    catch (Exception ex)
                    {
                        var cause = ex.GetBaseException();
                        storyErrors.Add(new
                        {
                            story = pair.Value,
                            storyType = pair.Key,
                            member = "StoryRanges.Item",
                            errorCode = cause is MissingMemberException ? "UNSUPPORTED_MEMBER" : "HOST_ERROR",
                            error = cause.Message,
                        });
                    }
                    finally { WordCom.Release(range); WordCom.Release(ranges); }
                }
                return new
                {
                    ok = true,
                    partial = storyErrors.Count > 0,
                    document = GetSummary(),
                    stories,
                    storyErrors,
                    bookmarks = Count(document, "Bookmarks"),
                    contentControls = Count(document, "ContentControls"),
                    fields = Count(document, "Fields"),
                    comments = Count(document, "Comments"),
                    footnotes = Count(document, "Footnotes"),
                    endnotes = Count(document, "Endnotes"),
                    capabilities = CapabilityPayload(document),
                };
            }
            finally { WordCom.Release(document); }
        }

        internal static string StoryName(int storyType)
        {
            foreach (var item in KnownStories) { if (item.Key == storyType) { return item.Value; } }
            return "story_" + storyType;
        }

        internal static string Sanitize(string value)
        {
            if (string.IsNullOrEmpty(value)) { return string.Empty; }
            return value.Replace("\a", string.Empty).Replace("\r", "\n").TrimEnd('\n');
        }

        internal static string Limit(string value, int max)
        {
            value = value ?? string.Empty;
            return value.Length <= max ? value : value.Substring(0, max) + "…（已截断）";
        }

        internal static IReadOnlyList<KeyValuePair<int, string>> KnownStories { get; } = new[]
        {
            new KeyValuePair<int, string>(1, "main_text"), new KeyValuePair<int, string>(2, "footnotes"),
            new KeyValuePair<int, string>(3, "endnotes"), new KeyValuePair<int, string>(4, "comments"),
            new KeyValuePair<int, string>(5, "text_frames"), new KeyValuePair<int, string>(6, "even_pages_header"),
            new KeyValuePair<int, string>(7, "primary_header"), new KeyValuePair<int, string>(8, "even_pages_footer"),
            new KeyValuePair<int, string>(9, "primary_footer"), new KeyValuePair<int, string>(10, "first_page_header"),
            new KeyValuePair<int, string>(11, "first_page_footer"), new KeyValuePair<int, string>(12, "footnote_separator"),
            new KeyValuePair<int, string>(13, "footnote_continuation_separator"), new KeyValuePair<int, string>(14, "footnote_continuation_notice"),
            new KeyValuePair<int, string>(15, "endnote_separator"),
        };

        private static int Count(object document, string collectionName)
        {
            object collection = null;
            try { return WordCom.TryGet(document, collectionName, out collection) ? WordCom.Int(collection, "Count") : 0; }
            finally { WordCom.Release(collection); }
        }

        private static object CapabilityPayload(object document)
        {
            var unsupported = new List<string>();
            foreach (var member in new[] { "StoryRanges", "Bookmarks", "ContentControls", "Fields", "Comments", "Revisions", "Undo", "Sections" })
            {
                if (!Has(document, member)) { unsupported.Add(member); }
            }
            return new
            {
                storyRanges = Has(document, "StoryRanges"), bookmarks = Has(document, "Bookmarks"),
                contentControls = Has(document, "ContentControls"), fields = Has(document, "Fields"),
                comments = Has(document, "Comments"), revisions = Has(document, "Revisions"),
                undo = Has(document, "Undo"), pageSetup = Has(document, "Sections"),
                unsupportedMembers = unsupported,
            };
        }

        private static bool Has(object target, string member)
        {
            return WordCom.TryGet(target, member, out var value) && value != null;
        }

        private static int? FindTablePosition(object selection, out int? row, out int? column)
        {
            row = null; column = null;
            object tables = null; object table = null; object cell = null;
            try
            {
                if (!WordCom.TryGet(selection, "Tables", out tables) || tables == null || WordCom.Int(tables, "Count") == 0) { return null; }
                var tableCount = WordCom.Int(tables, "Count");
                for (var tableNumber = 1; tableNumber <= tableCount; tableNumber++)
                {
                    table = WordCom.Item(tables, tableNumber);
                    object tableRange = null;
                    try
                    {
                        if (!WordCom.TryGet(table, "Range", out tableRange) || tableRange == null) { continue; }
                        var selectionStart = WordCom.Int(selection, "Start");
                        if (selectionStart < WordCom.Int(tableRange, "Start") || selectionStart > WordCom.Int(tableRange, "End")) { continue; }
                        var rows = WordCom.Int(table, "Rows"); var columns = WordCom.Int(table, "Columns");
                        for (var r = 1; r <= rows; r++)
                        {
                            for (var c = 1; c <= columns; c++)
                            {
                                object currentCell = null; object cellRange = null;
                                try
                                {
                                    currentCell = WordCom.Call(table, "Cell", r, c); cellRange = WordCom.Get(currentCell, "Range");
                                    if (selectionStart >= WordCom.Int(cellRange, "Start") && selectionStart <= WordCom.Int(cellRange, "End"))
                                    { row = r; column = c; return tableNumber; }
                                }
                                finally { WordCom.Release(cellRange); WordCom.Release(currentCell); }
                            }
                        }
                        return tableNumber;
                    }
                    finally { WordCom.Release(tableRange); WordCom.Release(table); table = null; }
                }
                return null;
            }
            catch { return null; }
            finally { WordCom.Release(cell); WordCom.Release(table); WordCom.Release(tables); }
        }
    }
}
