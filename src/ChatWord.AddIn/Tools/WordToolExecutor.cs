using System;
using System.Collections.Generic;
using ChatSheet.AddIn.Tools;
using ChatWord.AddIn.Hosts;
using Newtonsoft.Json.Linq;

namespace ChatWord.AddIn.Tools
{
    internal sealed class WordToolExecutor
    {
        private readonly Func<object> _applicationAccessor;
        private readonly WordDocumentContext _context;
        private readonly WordUndoStore _undo = new WordUndoStore();

        internal WordToolExecutor(Func<object> applicationAccessor, string progId = null)
        {
            _applicationAccessor = applicationAccessor ?? throw new ArgumentNullException(nameof(applicationAccessor));
            _context = new WordDocumentContext(applicationAccessor, progId);
        }

        internal WordDocumentContext Context => _context;
        internal WordUndoStore Undo => _undo;

        internal object BuildPreview(string name, JObject args)
        {
            object document = null; WordTarget target = null;
            try
            {
                document = ActiveDocument();
                if (document == null) { return new { supported = false, error = "当前没有打开的文档。" }; }
                var effectiveArgs = args;
                if (name == "edit_table")
                {
                    var tableTarget = args?["target"] as JObject != null ? (JObject)(args["target"] as JObject).DeepClone() : new JObject();
                    tableTarget["table_index"] = args?.Value<int?>("table_index"); tableTarget["row"] = args?.Value<int?>("row"); tableTarget["column"] = args?.Value<int?>("column");
                    effectiveArgs = new JObject { ["target"] = tableTarget, ["text"] = args?.Value<string>("text") ?? string.Empty };
                }
                target = WordTargetResolver.Resolve(document, effectiveArgs);
                var before = SafeTextTarget(target);
                string after;
                switch (name)
                {
                    case "replace_text":
                        var find = args?.Value<string>("find"); var replace = args?.Value<string>("replace") ?? string.Empty;
                        after = string.IsNullOrEmpty(find) ? replace : before.Replace(find, replace); break;
                    case "insert_text": after = args?.Value<string>("text") ?? string.Empty; break;
                    case "delete_range": after = string.Empty; break;
                    case "edit_table": after = args?.Value<string>("text") ?? string.Empty; break;
                    default: return new { supported = false, target = target.Label, note = "该操作不是文本替换，执行后将通过宿主读回实际格式或结构。" };
                }
                return new
                {
                    supported = true, story = target.Story, start = target.Start, end = target.End,
                    target = target.Label, before = WordDocumentContext.Limit(before, 2000), after = WordDocumentContext.Limit(after, 2000),
                };
            }
            catch (WordToolException ex) { return new { supported = false, errorCode = ex.Code, error = ex.Message }; }
            catch (Exception ex) { return new { supported = false, errorCode = "PREVIEW_FAILED", error = ex.Message }; }
            finally { target?.Dispose(); WordCom.Release(document); }
        }

        internal ToolResult Execute(string name, JObject args, string undoId = null)
        {
            try
            {
                switch (name)
                {
                    case "get_document_info": return ToolResult.Success(_context.GetSummary());
                    case "get_selection": return ToolResult.Success(_context.GetSelection());
                    case "read_document_structure": return ToolResult.Success(_context.ReadStructure(args?.Value<int?>("max_stories") ?? 32));
                    case "read_range": return ReadRange(args);
                    case "read_paragraphs": return ReadParagraphs(args);
                    case "read_table": return ReadTable(args);
                    case "find_text": return FindText(args);
                    case "replace_text": return ReplaceText(args, undoId);
                    case "insert_text": return InsertText(args, undoId);
                    case "delete_range": return DeleteRange(args, undoId);
                    case "format_text": return FormatText(args);
                    case "format_paragraph": return FormatParagraph(args);
                    case "apply_style": return ApplyStyle(args);
                    case "edit_table": return EditTable(args, undoId);
                    case "set_page_setup": return SetPageSetup(args);
                    case "insert_page_break": return InsertPageBreak(args, undoId);
                    default: return ToolResult.Failure("UNKNOWN_TOOL", "Word 工具不存在：" + name);
                }
            }
            catch (WordToolException ex) { return ToolResult.Failure(ex.Code, ex.Message); }
            catch (MissingMemberException ex) { return ToolResult.Failure("UNSUPPORTED_MEMBER", "当前 Word/WPS Writer 不支持所需成员：" + ex.Message); }
            catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.IndexOf("member", StringComparison.OrdinalIgnoreCase) >= 0 || ex.Message.IndexOf("不支持", StringComparison.Ordinal) >= 0)
            { return ToolResult.Failure("UNSUPPORTED_MEMBER", "当前 Word/WPS Writer 不支持所需成员：" + ex.Message); }
            catch (Exception ex) { return ToolResult.Failure(MapErrorCode(ex), "Word/WPS Writer 操作失败：" + ex.Message); }
        }

        private ToolResult ReadRange(JObject args)
        {
            object document = ActiveDocument(); WordTarget target = null;
            try
            {
                target = WordTargetResolver.Resolve(document, args);
                var raw = WordCom.String(target.Range, "Text");
                var text = WordDocumentContext.Sanitize(raw);
                var offset = Math.Max(0, args?.Value<int?>("offset") ?? 0); var max = Math.Min(20000, Math.Max(1, args?.Value<int?>("max_chars") ?? 8000));
                var page = offset >= text.Length ? string.Empty : text.Substring(offset, Math.Min(max, text.Length - offset));
                return ToolResult.Success(new { story = target.Story, start = target.Start, end = target.End, text = page, next_offset = offset + page.Length < text.Length ? (int?)(offset + page.Length) : null, truncated = offset + page.Length < text.Length, fields = Count(target.Range, "Fields"), hidden = WordCom.Bool(target.Range, "Hidden") });
            }
            finally { target?.Dispose(); WordCom.Release(document); }
        }

        private ToolResult ReadParagraphs(JObject args)
        {
            object document = ActiveDocument(); WordTarget target = null; object paragraphs = null;
            try
            {
                target = WordTargetResolver.Resolve(document, args, allowStoryOnly: true); paragraphs = WordCom.Get(target.Range, "Paragraphs");
                var offset = Math.Max(0, args?.Value<int?>("offset") ?? 0); var limit = Math.Min(100, Math.Max(1, args?.Value<int?>("limit") ?? 50)); var count = WordCom.Int(paragraphs, "Count"); var list = new List<object>();
                for (var i = offset + 1; i <= Math.Min(count, offset + limit); i++)
                {
                    object paragraph = null; object range = null; object style = null;
                    try
                    {
                        paragraph = WordCom.Item(paragraphs, i); range = WordCom.Get(paragraph, "Range"); WordCom.TryGet(paragraph, "Style", out style);
                        list.Add(new { index = i, start = WordCom.Int(range, "Start"), end = WordCom.Int(range, "End"), text = WordDocumentContext.Limit(WordDocumentContext.Sanitize(WordCom.String(range, "Text")), 2000), style = WordCom.String(style, "NameLocal", WordCom.String(style, "Name")), heading_level = WordCom.Int(paragraph, "OutlineLevel") });
                    }
                    finally { WordCom.Release(style); WordCom.Release(range); WordCom.Release(paragraph); }
                }
                return ToolResult.Success(new { story = target.Story, paragraphs = list, next_offset = offset + list.Count < count ? (int?)(offset + list.Count) : null });
            }
            finally { WordCom.Release(paragraphs); target?.Dispose(); WordCom.Release(document); }
        }

        private ToolResult ReadTable(JObject args)
        {
            object document = ActiveDocument(); object range = null; object tables = null; object table = null; object rowCollection = null; object columnCollection = null;
            try
            {
                var target = args?["target"] as JObject ?? args ?? new JObject(); var story = (target.Value<string>("story") ?? "main_text");
                range = WordTargetResolver.ResolveStory(document, story);
                if (range == null) { throw new WordToolException("TARGET_NOT_FOUND", "找不到指定 Story。"); }
                tables = WordCom.Get(range, "Tables");
                var tableIndex = args?.Value<int?>("table_index");
                if (!tableIndex.HasValue || tableIndex.Value == 0) { tableIndex = target.Value<int?>("table_index"); }
                if (!tableIndex.HasValue || tableIndex.Value == 0) { tableIndex = 1; }
                if (tableIndex.Value < 1) { throw new WordToolException("TARGET_INVALID", "表格索引从 1 开始。"); }
                if (tableIndex.Value > WordCom.Int(tables, "Count")) { throw new WordToolException("TARGET_NOT_FOUND", "找不到指定表格。"); }
                table = WordCom.Item(tables, tableIndex.Value); if (table == null) { throw new WordToolException("TARGET_NOT_FOUND", "找不到指定表格。"); }
                rowCollection = WordCom.Get(table, "Rows"); columnCollection = WordCom.Get(table, "Columns");
                var rows = WordCom.Int(rowCollection, "Count"); var columns = WordCom.Int(columnCollection, "Count");
                var offset = Math.Max(0, args?.Value<int?>("offset") ?? 0); var limit = Math.Min(200, Math.Max(1, args?.Value<int?>("limit") ?? 50));
                var result = new List<object>(); var lastRow = (int)Math.Min(rows, (long)offset + limit);
                for (var r = offset + 1; r <= lastRow; r++)
                {
                    var cells = new List<string>();
                    for (var c = 1; c <= columns && c <= 100; c++)
                    {
                        object cell = null; object cellRange = null;
                        try { cell = WordCom.Call(table, "Cell", r, c); cellRange = WordCom.Get(cell, "Range"); cells.Add(WordDocumentContext.Limit(WordDocumentContext.Sanitize(WordCom.String(cellRange, "Text")), 1000)); }
                        finally { WordCom.Release(cellRange); WordCom.Release(cell); }
                    }
                    result.Add(new { row = r, cells });
                }
                return ToolResult.Success(new { story, table_index = tableIndex.Value, rows, columns, data = result, next_offset = offset + result.Count < rows ? (int?)(offset + result.Count) : null });
            }
            finally { WordCom.Release(columnCollection); WordCom.Release(rowCollection); WordCom.Release(table); WordCom.Release(tables); WordCom.Release(range); WordCom.Release(document); }
        }

        private ToolResult FindText(JObject args)
        {
            object document = ActiveDocument(); WordTarget target = null; object duplicate = null; object find = null;
            try
            {
                target = WordTargetResolver.Resolve(document, args); var text = args?.Value<string>("text"); if (string.IsNullOrEmpty(text)) { throw new WordToolException("TEXT_REQUIRED", "缺少要查找的文字。"); }
                try { duplicate = WordCom.Get(target.Range, "Duplicate"); } catch { duplicate = WordCom.Call(target.Range, "Duplicate"); }
                find = WordCom.Get(duplicate, "Find"); WordCom.Set(find, "Text", text); WordCom.Set(find, "Forward", true); WordCom.Set(find, "Wrap", 0); WordCom.Set(find, "MatchCase", args?.Value<bool?>("match_case") == true); var executeResult = WordCom.Call(find, "Execute"); var found = executeResult is bool ? (bool)executeResult : WordCom.Bool(find, "Found", false);
                return ToolResult.Success(new { found, start = WordCom.Int(duplicate, "Start"), end = WordCom.Int(duplicate, "End"), text = found ? WordDocumentContext.Sanitize(WordCom.String(duplicate, "Text")) : string.Empty });
            }
            finally { WordCom.Release(find); WordCom.Release(duplicate); target?.Dispose(); WordCom.Release(document); }
        }

        private ToolResult ReplaceText(JObject args, string undoId)
        {
            object document = ActiveDocument(); WordTarget target = null;
            try
            {
                target = WordTargetResolver.Resolve(document, args); EnsureWritable(document);
                var before = SafeTextTarget(target); var find = args?.Value<string>("find");
                var replace = args?.Value<string>("replace") ?? string.Empty;
                if (ContainsMarkers(replace)) { throw new WordToolException("STRUCTURE_UNSUPPORTED", "替换内容不能包含文档控制字符。"); }
                if (!string.IsNullOrEmpty(find) && !before.Contains(find))
                { throw new WordToolException("TEXT_NOT_FOUND", "指定范围中没有要替换的文字，文档未修改。"); }
                var expected = string.IsNullOrEmpty(find) ? replace : before.Replace(find, replace);
                return WriteText(document, target, before, expected, undoId);
            }
            finally { target?.Dispose(); WordCom.Release(document); }
        }

        private ToolResult InsertText(JObject args, string undoId) { return SetText(args, undoId, false); }
        private ToolResult DeleteRange(JObject args, string undoId) { return SetText(args, undoId, true); }
        private ToolResult SetText(JObject args, string undoId, bool delete)
        {
            object document = ActiveDocument(); WordTarget target = null;
            try
            {
                target = WordTargetResolver.Resolve(document, args); EnsureWritable(document);
                var before = SafeTextTarget(target);
                if (delete && target.IsCollapsed) { throw new WordToolException("TARGET_INVALID", "删除目标不能是空区间。"); }
                if (!delete && !target.IsCollapsed) { throw new WordToolException("TARGET_INVALID", "插入位置必须是折叠的 Start/End 区间。"); }
                var expected = delete ? string.Empty : args?.Value<string>("text") ?? string.Empty;
                if (ContainsMarkers(expected)) { throw new WordToolException("STRUCTURE_UNSUPPORTED", "插入内容不能包含文档控制字符。"); }
                return WriteText(document, target, before, expected, undoId);
            }
            finally { target?.Dispose(); WordCom.Release(document); }
        }

        private ToolResult WriteText(object document, WordTarget target, string before, string expected, string undoId)
        {
            WordCom.Set(target.Range, "Text", expected);
            var actual = WordCom.String(target.Range, "Text");
            // Word/WPS 会把调用方传入的 LF 或 CRLF 段落换行存成 CR；按换行语义比较，
            // 避免多段落写入已经成功，却仅因宿主规范化换行而被误报为读回失败。
            if (!TextMatches(actual, expected))
            { return ToolResult.Failure("READBACK_FAILED", "文档已尝试写入，但目标读回与预期不一致；请重新读取文档确认实际状态。"); }
            var canUndo = target.StoryType == 1 && !string.IsNullOrEmpty(undoId);
            if (canUndo) { _undo.Add(undoId, document, target, before, actual); }
            return WriteResult(target, WordDocumentContext.Sanitize(actual), canUndo ? undoId : null);
        }

        private static bool TextMatches(string actual, string expected)
        {
            return string.Equals(NormalizeLineEndings(actual), NormalizeLineEndings(expected), StringComparison.Ordinal);
        }

        private static string NormalizeLineEndings(string value)
        {
            return (value ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n");
        }

        private static bool ContainsMarkers(string value)
        { return value != null && (value.IndexOf('\r') >= 0 || value.IndexOf('\a') >= 0 || value.IndexOf('\f') >= 0); }

        private static string SafeTextTarget(WordTarget target)
        {
            var raw = WordCom.String(target.Range, "Text");
            if (ContainsMarkers(raw))
            { throw new WordToolException("STRUCTURE_UNSUPPORTED", "目标包含段落、单元格或分页标记；请缩小到纯文本区间。"); }
            foreach (var member in new[] { "Fields", "Bookmarks", "ContentControls", "Revisions" })
            {
                object collection = null;
                try
                {
                    if (WordCom.TryGet(target.Range, member, out collection) && WordCom.Int(collection, "Count") > 0)
                    { throw new WordToolException("STRUCTURE_UNSUPPORTED", "目标包含 " + member + "，不能直接替换其内部文本。"); }
                }
                finally { WordCom.Release(collection); }
            }
            return raw;
        }

        private ToolResult FormatText(JObject args)
        {
            object document = ActiveDocument(); WordTarget target = null; object font = null;
            try { target = WordTargetResolver.Resolve(document, args); EnsureWritable(document); font = WordCom.Get(target.Range, "Font"); SetIf(args, "bold", font, "Bold"); SetIf(args, "italic", font, "Italic"); SetIf(args, "underline", font, "Underline"); SetIf(args, "font_size", font, "Size"); return ToolResult.Success(new { story = target.Story, start = target.Start, end = target.End, readback = new { bold = WordCom.String(font, "Bold"), italic = WordCom.String(font, "Italic"), size = WordCom.String(font, "Size") }, canUndo = false, undoNote = "字符格式未保存完整快照，不能提供虚假撤销。" }); }
            finally { WordCom.Release(font); target?.Dispose(); WordCom.Release(document); }
        }

        private ToolResult FormatParagraph(JObject args)
        {
            object document = ActiveDocument(); WordTarget target = null; object format = null;
            try
            {
                target = WordTargetResolver.Resolve(document, args);
                EnsureWritable(document);
                format = WordCom.Get(target.Range, "ParagraphFormat");

                var alignment = args?.Value<string>("alignment");
                int? alignmentValue = null;
                if (!string.IsNullOrWhiteSpace(alignment))
                {
                    var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                    { ["left"] = 0, ["center"] = 1, ["right"] = 2, ["justify"] = 3 };
                    if (!map.TryGetValue(alignment, out var value))
                    { throw new WordToolException("ARGUMENT_INVALID", "不支持的段落对齐方式。"); }
                    alignmentValue = value;
                }

                string firstLineIndentChars = null;
                var indentToken = args?["first_line_indent_chars"];
                if (indentToken != null && indentToken.Type != JTokenType.Null)
                {
                    if (indentToken.Type != JTokenType.Integer && indentToken.Type != JTokenType.Float)
                    { throw new WordToolException("ARGUMENT_INVALID", "首行缩进必须是字符数值。"); }
                    var requestedIndent = indentToken.Value<double>();
                    if (double.IsNaN(requestedIndent) || double.IsInfinity(requestedIndent))
                    { throw new WordToolException("ARGUMENT_INVALID", "首行缩进必须是有限数值。"); }

                    WordCom.Set(format, "CharacterUnitFirstLineIndent", requestedIndent);
                    var actualIndent = Convert.ToDouble(WordCom.Get(format, "CharacterUnitFirstLineIndent"), System.Globalization.CultureInfo.InvariantCulture);
                    if (Math.Abs(actualIndent - requestedIndent) > 0.01)
                    { throw new WordToolException("READBACK_FAILED", "首行缩进写入后的读回值与请求不一致。"); }
                    firstLineIndentChars = actualIndent.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (alignmentValue.HasValue) { WordCom.Set(format, "Alignment", alignmentValue.Value); }
                SetIf(args, "space_before", format, "SpaceBefore");
                SetIf(args, "space_after", format, "SpaceAfter");
                SetIf(args, "keep_with_next", format, "KeepWithNext");
                return ToolResult.Success(new
                {
                    story = target.Story,
                    start = target.Start,
                    end = target.End,
                    readback = new
                    {
                        alignment = WordCom.String(format, "Alignment"),
                        firstLineIndentChars,
                        spaceBefore = WordCom.String(format, "SpaceBefore"),
                        spaceAfter = WordCom.String(format, "SpaceAfter"),
                    },
                    canUndo = false,
                    undoNote = "段落格式未保存完整快照，不能提供虚假撤销。",
                });
            }
            finally { WordCom.Release(format); target?.Dispose(); WordCom.Release(document); }
        }

        private ToolResult ApplyStyle(JObject args)
        {
            object document = ActiveDocument(); WordTarget target = null;
            try { target = WordTargetResolver.Resolve(document, args); EnsureWritable(document); var style = (args.Value<string>("style") ?? string.Empty).Trim(); if (style.Length == 0) { throw new WordToolException("STYLE_REQUIRED", "缺少样式名称。"); } WordCom.Set(target.Range, "Style", style); return ToolResult.Success(new { story = target.Story, style, readback = WordCom.String(target.Range, "Style"), canUndo = false, undoNote = "样式操作未保存完整快照，不能提供虚假撤销。" }); }
            finally { target?.Dispose(); WordCom.Release(document); }
        }

        private ToolResult EditTable(JObject args, string undoId)
        {
            var target = args?["target"] as JObject ?? new JObject(); target["table_index"] = args.Value<int?>("table_index"); target["row"] = args.Value<int?>("row"); target["column"] = args.Value<int?>("column"); return SetText(new JObject { ["target"] = target, ["text"] = args.Value<string>("text") ?? string.Empty }, undoId, false);
        }

        private ToolResult SetPageSetup(JObject args)
        {
            object document = ActiveDocument(); object sections = null; object section = null; object setup = null;
            try { EnsureWritable(document); sections = WordCom.Get(document, "Sections"); section = WordCom.Item(sections, args.Value<int?>("section_index") ?? 1); if (section == null) { throw new WordToolException("TARGET_NOT_FOUND", "找不到指定节。"); } setup = WordCom.Get(section, "PageSetup"); var orientation = args.Value<string>("orientation"); if (!string.IsNullOrWhiteSpace(orientation)) { WordCom.Set(setup, "Orientation", orientation.Equals("landscape", StringComparison.OrdinalIgnoreCase) ? 1 : 0); } SetIf(args, "top_margin", setup, "TopMargin"); SetIf(args, "bottom_margin", setup, "BottomMargin"); SetIf(args, "left_margin", setup, "LeftMargin"); SetIf(args, "right_margin", setup, "RightMargin"); return ToolResult.Success(new { section = args.Value<int?>("section_index") ?? 1, orientation = WordCom.String(setup, "Orientation"), canUndo = false, undoNote = "页面设置未保存完整快照，不能提供虚假撤销。" }); }
            finally { WordCom.Release(setup); WordCom.Release(section); WordCom.Release(sections); WordCom.Release(document); }
        }

        private ToolResult InsertPageBreak(JObject args, string undoId)
        {
            object document = ActiveDocument(); WordTarget target = null;
            try { target = WordTargetResolver.Resolve(document, args); EnsureWritable(document); WordCom.Call(target.Range, "InsertBreak", 7); var after = WordDocumentContext.Sanitize(WordCom.String(target.Range, "Text")); return ToolResult.Success(new { story = target.Story, start = target.Start, end = target.End, inserted = true, readback = after, canUndo = false, undoNote = "分页符操作未保存完整快照，不能提供虚假撤销。" }); }
            finally { target?.Dispose(); WordCom.Release(document); }
        }

        private object ActiveDocument() { return _context.ActiveDocument(); }
        private static int Count(object target, string member) { object collection = null; try { return WordCom.TryGet(target, member, out collection) ? WordCom.Int(collection, "Count") : 0; } finally { WordCom.Release(collection); } }
        private static void EnsureWritable(object document)
        {
            if (document == null) { throw new WordToolException("DOCUMENT_REQUIRED", "当前没有打开的文档。"); }
            if (!WordCom.TryGet(document, "ProtectionType", out var protectionValue) || protectionValue == null)
            { throw new WordToolException("UNSUPPORTED_MEMBER", "宿主不提供文档保护状态，拒绝写入。"); }
            var protection = Convert.ToInt32(protectionValue);
            // Word uses wdNoProtection=-1; some WPS builds expose the same state as 0.
            if (protection != -1 && protection != 0) { throw new WordToolException("PROTECTED_DOCUMENT", "文档受保护，宿主拒绝写入。"); }
            if (!WordCom.TryGet(document, "ReadOnly", out var readOnly)) { throw new WordToolException("UNSUPPORTED_MEMBER", "宿主不提供文档只读状态，拒绝写入。"); }
            if (Convert.ToBoolean(readOnly)) { throw new WordToolException("READ_ONLY_DOCUMENT", "文档为只读，不能写入。"); }
            if (!WordCom.TryGet(document, "TrackRevisions", out var tracking)) { throw new WordToolException("UNSUPPORTED_MEMBER", "宿主不提供修订状态，拒绝写入。"); }
            if (Convert.ToBoolean(tracking)) { throw new WordToolException("TRACK_CHANGES_REQUIRED", "文档正在跟踪修订；当前写入/撤销不能可靠验证修订结果。"); }
        }
        private static void SetIf(JObject args, string key, object target, string member) { if (args?[key] != null && args[key].Type != JTokenType.Null) { WordCom.Set(target, member, args[key].ToObject<object>()); } }
        private static ToolResult WriteResult(WordTarget target, string after, string undoId) { return ToolResult.Success(new { story = target.Story, start = target.Start, end = target.End, text = after, canUndo = !string.IsNullOrEmpty(undoId), undoId, undoNote = string.IsNullOrEmpty(undoId) ? "未提供快照标识。" : "已保存文档目标快照。" }); }
        private static string MapErrorCode(Exception ex) { var text = ex.Message ?? string.Empty; return text.IndexOf("not supported", StringComparison.OrdinalIgnoreCase) >= 0 ? "UNSUPPORTED_MEMBER" : "HOST_ERROR"; }
    }
}
