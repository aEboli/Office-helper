using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ChatSheet.AddIn.Tools;
using ChatWord.AddIn.Agent;
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
                    case "insert_document_draft":
                        VerifyDraftVersion(args);
                        VerifyDraftTarget(args, target);
                        var draftBlocks = WordDocumentDraft.ParseBlocks(args?["blocks"] as JArray);
                        EnsureDraftElementsSupported(document, args, target, draftBlocks);
                        after = string.Join("\n", draftBlocks.Select(block => block.PlainText()));
                        break;
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
                    case "insert_document_draft": return InsertDocumentDraft(args);
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

        private ToolResult InsertDocumentDraft(JObject args)
        {
            object document = null;
            WordTarget target = null;
            var modified = false;
            var before = string.Empty;
            var originalStart = 0;
            var fullText = string.Empty;
            var tableLengthDelta = 0;
            try
            {
                document = ActiveDocument();
                VerifyDraftVersion(args);
                target = WordTargetResolver.Resolve(document, args);
                EnsureWritable(document);
                VerifyDraftTarget(args, target);
                var blocks = WordDocumentDraft.ParseBlocks(args?["blocks"] as JArray);
                EnsureDraftElementsSupported(document, args, target, blocks);
                var operation = args?.Value<string>("operation");
                if (operation != "insert" && operation != "replace")
                { throw new WordToolException("DRAFT_INVALID", "草稿操作必须是 insert 或 replace。"); }
                if ((operation == "insert") != target.IsCollapsed)
                { throw new WordToolException("TARGET_INVALID", "草稿操作与当前选区不匹配。"); }

                before = SafeTextTarget(target);
                originalStart = target.Start;
                var spans = ComposeDraftText(blocks, out fullText);
                WordCom.Set(target.Range, "Text", fullText);
                modified = true;

                for (var i = spans.Count - 1; i >= 0; i--)
                {
                    var span = spans[i];
                    var start = originalStart + span.Start;
                    var end = originalStart + span.End;
                    if (span.Block.Type == "table")
                    {
                        object tableTextRange = null;
                        object table = null;
                        object tableRange = null;
                        try
                        {
                            tableTextRange = WordCom.Call(document, "Range", start, originalStart + span.End + 1);
                            table = WordCom.Call(tableTextRange, "ConvertToTable", "\t", span.Block.Rows.Count + 1, span.Block.Headers.Count);
                            tableRange = WordCom.Get(table, "Range");
                            tableLengthDelta += WordCom.Int(tableRange, "End") - WordCom.Int(tableRange, "Start") - (span.End - span.Start + 1);
                            ApplyDraftTableStyle(table, args.Value<string>("style"));
                            VerifyDraftTable(table, span.Block);
                        }
                        finally { WordCom.Release(tableRange); WordCom.Release(table); WordCom.Release(tableTextRange); }
                        continue;
                    }

                    object blockRange = null;
                    try
                    {
                        blockRange = WordCom.Call(document, "Range", start, end);
                        ApplyDraftParagraphStyle(document, blockRange, span.Block, args.Value<string>("style"));
                        if (span.Block.Type == "bullets" || span.Block.Type == "numbered")
                        {
                            object listFormat = null;
                            try
                            {
                                listFormat = WordCom.Get(blockRange, "ListFormat");
                                WordCom.Call(listFormat, span.Block.Type == "bullets" ? "ApplyBulletDefault" : "ApplyNumberDefault");
                            }
                            finally { WordCom.Release(listFormat); }
                        }
                        else
                        {
                            ApplyDraftRuns(document, blockRange, span.Block, start);
                            var actual = WordDocumentContext.Sanitize(WordCom.String(blockRange, "Text"));
                            if (!TextMatches(actual, span.Block.PlainText()))
                            { throw new WordToolException("READBACK_FAILED", "草稿段落读回与预期不一致。"); }
                        }
                    }
                    finally { WordCom.Release(blockRange); }
                }

                var finalEnd = originalStart + fullText.Length + tableLengthDelta;
                WordCom.Call(target.Range, "SetRange", originalStart, finalEnd);
                target.End = finalEnd;
                var actualText = WordCom.String(target.Range, "Text");
                return ToolResult.Success(new
                {
                    draft_id = args.Value<string>("draft_id"), operation,
                    story = target.Story, start = originalStart, end = finalEnd,
                    text = WordDocumentContext.Sanitize(actualText),
                    canUndo = false, undoId = (string)null,
                    undoNote = "结构化草稿可能包含段落样式、列表或表格；当前文本快照不能可靠恢复这些格式，因此面板不提供撤销。",
                });
            }
            catch (Exception ex)
            {
                var errorCode = ex is WordToolException wordError ? wordError.Code : MapErrorCode(ex);
                var message = ex.Message;
                if (modified && target != null)
                {
                    try
                    {
                        WordCom.Set(target.Range, "Text", before);
                        var restored = WordCom.String(target.Range, "Text");
                        if (TextMatches(restored, before))
                        { message += " 草稿未完整应用，原目标文本已恢复。"; }
                        else
                        { errorCode = "DRAFT_PARTIAL_WRITE"; message += " 尝试恢复原文后读回仍不一致，请检查当前文档。"; }
                    }
                    catch (Exception rollbackError)
                    {
                        errorCode = "DRAFT_PARTIAL_WRITE";
                        message += " 无法确认已恢复原文，请检查当前文档：" + rollbackError.Message;
                    }
                }
                return ToolResult.Failure(errorCode, message);
            }
            finally { target?.Dispose(); WordCom.Release(document); }
        }

        private void VerifyDraftTarget(JObject args, WordTarget target)
        {
            var expectedDocumentKey = args?.Value<string>("document_key");
            if (!string.Equals(expectedDocumentKey, WordDocumentContext.DocumentKey(_context.GetSummary()), StringComparison.Ordinal))
            { throw new WordToolException("DOCUMENT_CHANGED", "草稿关联的文档已切换；请在原文档重新选择目标。"); }

            var targetArgs = args?["target"] as JObject;
            var selection = _context.GetSelection();
            if (selection == null || !selection.HasSelection ||
                !string.Equals(selection.Story, target.Story, StringComparison.OrdinalIgnoreCase) ||
                selection.Start != target.Start || selection.End != target.End ||
                string.IsNullOrWhiteSpace(targetArgs?.Value<string>("fingerprint")) ||
                !string.Equals(selection.TextFingerprint, targetArgs.Value<string>("fingerprint"), StringComparison.Ordinal))
            { throw new WordToolException("DRAFT_TARGET_CHANGED", "草稿原选区已变化；请重新选择原目标后再确认插入。"); }
        }

        private static void VerifyDraftVersion(JObject args)
        {
            if (args?.Value<int?>("version") != WordDocumentDraft.CurrentVersion)
            { throw new WordToolException("DRAFT_VERSION_UNSUPPORTED", "草稿版本不受当前 Word/WPS Writer 插入器支持。"); }
        }

        private static void EnsureDraftElementsSupported(object document, JObject args, WordTarget target, IList<WordDraftBlock> blocks)
        {
            var style = args?.Value<string>("style") ?? "follow_document";
            if (!WordDocumentDraft.IsStyle(style)) { throw new WordToolException("DRAFT_STYLE_INVALID", "草稿风格无效。"); }
            var hasTable = blocks.Any(block => block.Type == "table");
            if (hasTable && (!string.Equals(target.Story, "main_text", StringComparison.OrdinalIgnoreCase) || target.TableIndex.HasValue))
            { throw new WordToolException("STRUCTURE_UNSUPPORTED", "表格草稿只能插入正文中的非表格位置。"); }

            object styles = null;
            try
            {
                styles = WordCom.Get(document, "Styles");
                foreach (var block in blocks)
                {
                    var styleId = BuiltInStyleId(block);
                    if (!styleId.HasValue) { continue; }
                    object builtInStyle = null;
                    try
                    {
                        builtInStyle = WordCom.Item(styles, styleId.Value);
                        if (builtInStyle == null)
                        { throw new WordToolException("UNSUPPORTED_STYLE", "当前宿主缺少草稿需要的内置样式。"); }
                    }
                    finally { WordCom.Release(builtInStyle); }
                }
            }
            finally { WordCom.Release(styles); }
        }

        private static List<WordDraftBlockSpan> ComposeDraftText(IList<WordDraftBlock> blocks, out string text)
        {
            var builder = new StringBuilder();
            var spans = new List<WordDraftBlockSpan>();
            foreach (var block in blocks)
            {
                var blockText = block.PlainText().Replace("\n", "\r");
                var start = builder.Length;
                builder.Append(blockText);
                var end = builder.Length;
                builder.Append('\r');
                spans.Add(new WordDraftBlockSpan { Block = block, Start = start, End = end });
            }
            text = builder.ToString();
            return spans;
        }

        private static int? BuiltInStyleId(WordDraftBlock block)
        {
            switch (block.Type)
            {
                case "title": return -63;
                case "heading": return block.Level == 1 ? -2 : block.Level == 2 ? -3 : -4;
                case "quote": return -181;
                case "paragraph":
                case "bullets":
                case "numbered": return -1;
                default: return null;
            }
        }

        private static void ApplyDraftParagraphStyle(object document, object range, WordDraftBlock block, string style)
        {
            var styleId = BuiltInStyleId(block);
            if (styleId.HasValue) { WordCom.Set(range, "Style", styleId.Value); }
            object format = null;
            object font = null;
            try
            {
                format = WordCom.Get(range, "ParagraphFormat");
                if (block.Type == "heading" || block.Type == "title") { WordCom.Set(format, "KeepWithNext", true); }
                if (block.Type == "quote") { WordCom.Set(format, "LeftIndent", 18f); }
                if (style == "minimal" || style == "business" || style == "report")
                {
                    var after = style == "minimal" ? 4f : style == "business" ? 6f : 8f;
                    WordCom.Set(format, "SpaceAfter", after);
                    if (block.Type == "heading") { WordCom.Set(format, "SpaceBefore", block.Level == 1 ? 10f : 6f); }
                    if (block.Type == "title" && style != "minimal") { WordCom.Set(format, "Alignment", 1); }
                }

                var size = DraftFontSize(block, style);
                if (size.HasValue)
                {
                    font = WordCom.Get(range, "Font");
                    WordCom.Set(font, "Size", size.Value);
                }
            }
            finally { WordCom.Release(font); WordCom.Release(format); }
        }

        private static float? DraftFontSize(WordDraftBlock block, string style)
        {
            if (style == "follow_document") { return null; }
            if (block.Type == "title") { return style == "report" ? 24f : 20f; }
            if (block.Type == "heading")
            {
                if (block.Level == 1) { return style == "report" ? 16f : 15f; }
                if (block.Level == 2) { return style == "report" ? 14f : 13f; }
                return 12f;
            }
            return null;
        }

        private static void ApplyDraftRuns(object document, object blockRange, WordDraftBlock block, int start)
        {
            var offset = 0;
            foreach (var run in block.Runs)
            {
                if (!run.Bold && !run.Italic) { offset += run.Text.Length; continue; }
                object runRange = null;
                object font = null;
                try
                {
                    runRange = WordCom.Call(document, "Range", start + offset, start + offset + run.Text.Length);
                    font = WordCom.Get(runRange, "Font");
                    if (run.Bold) { WordCom.Set(font, "Bold", -1); }
                    if (run.Italic) { WordCom.Set(font, "Italic", -1); }
                }
                finally { WordCom.Release(font); WordCom.Release(runRange); }
                offset += run.Text.Length;
            }
        }

        private static void ApplyDraftTableStyle(object table, string style)
        {
            object rows = null;
            object headerRow = null;
            object headerRange = null;
            object font = null;
            try
            {
                WordCom.TryCall(table, "AutoFitBehavior", out var ignored, 2);
                WordCom.Release(ignored);
                WordCom.TryGet(table, "Rows", out rows);
                headerRow = WordCom.Item(rows, 1);
                headerRange = WordCom.Get(headerRow, "Range");
                font = WordCom.Get(headerRange, "Font");
                WordCom.Set(font, "Bold", -1);
                object shading = null;
                try
                {
                    shading = WordCom.Get(headerRow, "Shading");
                    var color = style == "business" ? 0xE5EFE7 : style == "report" ? 0xE5F0EE : 0xE8E8E8;
                    WordCom.Set(shading, "BackgroundPatternColor", color);
                }
                finally { WordCom.Release(shading); }
            }
            finally { WordCom.Release(font); WordCom.Release(headerRange); WordCom.Release(headerRow); WordCom.Release(rows); }
        }

        private static void VerifyDraftTable(object table, WordDraftBlock block)
        {
            for (var row = 0; row <= block.Rows.Count; row++)
            {
                var expected = row == 0 ? block.Headers : block.Rows[row - 1];
                for (var column = 0; column < expected.Count; column++)
                {
                    object cell = null;
                    object range = null;
                    try
                    {
                        cell = WordCom.Call(table, "Cell", row + 1, column + 1);
                        range = WordCom.Get(cell, "Range");
                        var actual = WordDocumentContext.Sanitize(WordCom.String(range, "Text"));
                        if (!string.Equals(actual, expected[column], StringComparison.Ordinal))
                        { throw new WordToolException("READBACK_FAILED", "草稿表格单元格读回与预期不一致。"); }
                    }
                    finally { WordCom.Release(range); WordCom.Release(cell); }
                }
            }
        }

        private sealed class WordDraftBlockSpan
        {
            internal WordDraftBlock Block;
            internal int Start;
            internal int End;
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
