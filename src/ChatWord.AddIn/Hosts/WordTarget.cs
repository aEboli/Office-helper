using System;
using Newtonsoft.Json.Linq;

namespace ChatWord.AddIn.Hosts
{
    internal sealed class WordTarget
    {
        internal object Range { get; set; }
        internal string Story { get; set; }
        internal int StoryType { get; set; }
        internal int Start { get; set; }
        internal int End { get; set; }
        internal string Locator { get; set; }
        internal int? TableIndex { get; set; }
        internal int? Row { get; set; }
        internal int? Column { get; set; }
        internal bool IsCollapsed => Start == End;

        internal string Label => string.Format("{0} Start={1} End={2}{3}",
            Story ?? "story", Start, End, TableIndex.HasValue ? string.Format("（表 {0} 行 {1} 列 {2}）", TableIndex, Row, Column) : string.Empty);

        internal void Dispose()
        {
            WordCom.Release(Range);
            Range = null;
        }
    }

    internal static class WordTargetResolver
    {
        internal static object ResolveStory(object document, string storyName)
        {
            var type = StoryType(storyName);
            return StoryRange(document, type);
        }

        internal static WordTarget Resolve(object document, JObject args, bool required = true, bool allowStoryOnly = false)
        {
            if (document == null) { throw new WordToolException("DOCUMENT_REQUIRED", "当前没有打开的文档。"); }
            args = args ?? new JObject();
            var target = args["target"] as JObject ?? args;
            var hasRangeFields = target["start"] != null || target["end"] != null;
            var hasParagraph = target.Value<int?>("paragraph_index").GetValueOrDefault() != 0;
            var hasTable = target.Value<int?>("table_index").GetValueOrDefault() != 0 ||
                target.Value<int?>("row").GetValueOrDefault() != 0 || target.Value<int?>("column").GetValueOrDefault() != 0;
            var hasBookmark = !string.IsNullOrWhiteSpace(target.Value<string>("bookmark"));
            var hasContentControl = !string.IsNullOrWhiteSpace(target.Value<string>("content_control"));
            var hasOtherLocator = hasParagraph || hasTable || hasBookmark || hasContentControl;
            var hasRange = hasRangeFields &&
                (target.Value<int?>("start").GetValueOrDefault() != 0 || target.Value<int?>("end").GetValueOrDefault() != 0 || (!hasOtherLocator && !allowStoryOnly));
            var locators = (hasRange ? 1 : 0) + (hasParagraph ? 1 : 0) + (hasTable ? 1 : 0) +
                (hasBookmark ? 1 : 0) + (hasContentControl ? 1 : 0);
            if (locators == 0)
            {
                if (allowStoryOnly)
                {
                    var fullStoryName = (target.Value<string>("story") ?? "main_text").Trim();
                    var fullStoryType = StoryType(fullStoryName);
                    var storyRange = StoryRange(document, fullStoryType);
                    if (storyRange == null) { throw new WordToolException("TARGET_NOT_FOUND", "宿主不提供 Story：" + fullStoryName); }
                    try
                    {
                        TrimDocumentMarkers(storyRange);
                        var storyTarget = CreateTarget(storyRange, fullStoryType, "story", null);
                        storyRange = null;
                        return storyTarget;
                    }
                    finally { WordCom.Release(storyRange); }
                }
                if (required) { throw new WordToolException("TARGET_REQUIRED", "必须提供明确的文档目标（Story、Start/End、段落、表格、书签或内容控件）。"); }
                return null;
            }
            if (locators > 1) { throw new WordToolException("TARGET_AMBIGUOUS", "一个请求只能使用一种主要文档定位方式。"); }

            var storyName = (target.Value<string>("story") ?? "main_text").Trim();
            if (target["story"] == null && !hasBookmark && !hasContentControl)
            {
                throw new WordToolException("STORY_REQUIRED", "字符区间、段落或表格目标必须明确指定 story。");
            }
            var storyType = StoryType(storyName);
            object range = null;
            try
            {
                range = StoryRange(document, storyType);
                if (range == null) { throw new WordToolException("TARGET_NOT_FOUND", "宿主不提供 Story：" + storyName); }

                var locator = "story";
                if (hasRange)
                {
                    var start = target.Value<int?>("start");
                    var end = target.Value<int?>("end");
                    if (!start.HasValue || !end.HasValue || start.Value < 0 || end.Value < start.Value)
                    { throw new WordToolException("TARGET_INVALID", "Start/End 必须是同一 Story 内的非负区间。"); }
                    var selected = storyType == 1
                        ? WordCom.Call(document, "Range", start.Value, end.Value)
                        : DuplicateAndSet(range, start.Value, end.Value);
                    WordCom.Release(range);
                    range = selected;
                }
                else if (hasParagraph)
                {
                    var index = target.Value<int?>("paragraph_index");
                    if (!index.HasValue || index.Value < 1) { throw new WordToolException("TARGET_INVALID", "段落索引从 1 开始。"); }
                    object paragraphs = null; object paragraph = null; object selected = null;
                    try
                    {
                        paragraphs = WordCom.Get(range, "Paragraphs");
                        paragraph = WordCom.Item(paragraphs, index.Value);
                        if (paragraph == null) { throw new WordToolException("TARGET_NOT_FOUND", "找不到指定段落。"); }
                        selected = WordCom.Get(paragraph, "Range");
                        WordCom.Release(range); range = selected; selected = null; locator = "paragraph";
                    }
                    finally { WordCom.Release(selected); WordCom.Release(paragraph); WordCom.Release(paragraphs); }
                }
                else if (hasTable)
                {
                    var tableIndex = target.Value<int?>("table_index");
                    var row = target.Value<int?>("row");
                    var column = target.Value<int?>("column");
                    if (!tableIndex.HasValue || !row.HasValue || !column.HasValue || tableIndex < 1 || row < 1 || column < 1)
                    { throw new WordToolException("TARGET_INVALID", "表格目标必须包含从 1 开始的 table_index、row 和 column。"); }
                    object tables = null; object table = null; object cell = null; object selected = null;
                    try
                    {
                        tables = WordCom.Get(range, "Tables"); table = WordCom.Item(tables, tableIndex.Value);
                        if (table == null) { throw new WordToolException("TARGET_NOT_FOUND", "找不到指定表格。"); }
                        cell = WordCom.Call(table, "Cell", row.Value, column.Value);
                        selected = WordCom.Get(cell, "Range");
                        // Cell.Range includes the paragraph and end-of-cell markers.
                        WordCom.Call(selected, "SetRange", WordCom.Int(selected, "Start"), WordCom.Int(selected, "End") - 2);
                        WordCom.Release(range); range = selected; selected = null; locator = "table_cell";
                    }
                    finally { WordCom.Release(selected); WordCom.Release(cell); WordCom.Release(table); WordCom.Release(tables); }
                }
                else if (hasBookmark)
                {
                    var name = target.Value<string>("bookmark");
                    object bookmarks = null; object bookmark = null; object selected = null;
                    try
                    {
                        bookmarks = WordCom.Get(document, "Bookmarks"); bookmark = WordCom.Item(bookmarks, name);
                        if (bookmark == null) { throw new WordToolException("TARGET_NOT_FOUND", "找不到书签：" + name); }
                        selected = WordCom.Get(bookmark, "Range");
                        WordCom.Release(range); range = selected; selected = null; locator = "bookmark:" + name;
                    }
                    finally { WordCom.Release(selected); WordCom.Release(bookmark); WordCom.Release(bookmarks); }
                }
                else
                {
                    var name = target.Value<string>("content_control");
                    object controls = null; object control = null; object selected = null;
                    try
                    {
                        controls = WordCom.Get(document, "ContentControls");
                        var count = WordCom.Int(controls, "Count");
                        for (var i = 1; i <= count; i++)
                        {
                            var candidate = WordCom.Item(controls, i);
                            var title = WordCom.String(candidate, "Title"); var tag = WordCom.String(candidate, "Tag");
                            if (string.Equals(name, title, StringComparison.OrdinalIgnoreCase) || string.Equals(name, tag, StringComparison.OrdinalIgnoreCase) || string.Equals(name, i.ToString(), StringComparison.Ordinal))
                            { control = candidate; break; }
                            WordCom.Release(candidate);
                        }
                        if (control == null) { throw new WordToolException("TARGET_NOT_FOUND", "找不到内容控件：" + name); }
                        selected = WordCom.Get(control, "Range"); WordCom.Release(range); range = selected; selected = null; locator = "content_control:" + name;
                    }
                    finally { WordCom.Release(selected); WordCom.Release(control); WordCom.Release(controls); }
                }

                TrimDocumentMarkers(range);
                return CreateTarget(range, storyType, locator, target);
            }
            catch { WordCom.Release(range); throw; }
        }

        private static WordTarget CreateTarget(object range, int storyType, string locator, JObject target = null)
        {
            return new WordTarget
            {
                Range = range,
                Story = WordDocumentContext.StoryName(storyType),
                StoryType = storyType,
                Start = WordCom.Int(range, "Start"),
                End = WordCom.Int(range, "End"),
                Locator = locator,
                TableIndex = locator == "table_cell" ? target?.Value<int?>("table_index") : null,
                Row = locator == "table_cell" ? target?.Value<int?>("row") : null,
                Column = locator == "table_cell" ? target?.Value<int?>("column") : null,
            };
        }

        private static object StoryRange(object document, int storyType)
        {
            object stories = null;
            try
            {
                if (!WordCom.TryGet(document, "StoryRanges", out stories) || stories == null) { return null; }
                return WordCom.Item(stories, storyType);
            }
            catch { WordCom.Release(stories); return null; }
            finally { WordCom.Release(stories); }
        }

        private static object DuplicateAndSet(object source, int start, int end)
        {
            object copy;
            try { copy = WordCom.Get(source, "Duplicate"); }
            catch { copy = WordCom.Call(source, "Duplicate"); }
            try { WordCom.Call(copy, "SetRange", start, end); return copy; }
            catch { WordCom.Release(copy); throw; }
        }

        private static void TrimDocumentMarkers(object range)
        {
            var raw = WordCom.String(range, "Text");
            if (string.IsNullOrEmpty(raw)) { return; }
            var visibleLength = raw.Length;
            while (visibleLength > 0 && (raw[visibleLength - 1] == '\r' || raw[visibleLength - 1] == '\a')) { visibleLength--; }
            var trim = raw.Length - visibleLength;
            if (trim > 0)
            {
                var start = WordCom.Int(range, "Start");
                var end = WordCom.Int(range, "End");
                try { WordCom.Call(range, "SetRange", start, Math.Max(start, end - trim)); } catch { }
            }
        }

        internal static int StoryType(string name)
        {
            name = (name ?? "main_text").Trim().ToLowerInvariant();
            foreach (var item in WordDocumentContext.KnownStories)
            {
                if (string.Equals(item.Value, name, StringComparison.OrdinalIgnoreCase)) { return item.Key; }
            }
            if (int.TryParse(name, out var value) && value > 0) { return value; }
            throw new WordToolException("STORY_INVALID", "未知 Story：" + name);
        }
    }

    internal sealed class WordToolException : Exception
    {
        internal WordToolException(string code, string message) : base(message) { Code = code; }
        internal string Code { get; }
    }
}
