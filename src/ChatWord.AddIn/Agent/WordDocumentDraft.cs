using System;
using System.Collections.Generic;
using System.Linq;
using ChatWord.AddIn.Hosts;
using ChatWord.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatWord.AddIn.Agent
{
    internal sealed class WordDocumentDraft
    {
        internal const int CurrentVersion = 1;
        private const int MaxBlocks = 100;
        private const int MaxTextLength = 50000;
        private const int MaxRows = 100;
        private const int MaxColumns = 20;
        private const int MaxItems = 100;

        internal string Id { get; private set; }
        internal int Version { get; private set; }
        internal string DocumentKey { get; private set; }
        internal string DocumentName { get; private set; }
        internal string Operation { get; private set; }
        internal string Style { get; private set; }
        internal string Story { get; private set; }
        internal int Start { get; private set; }
        internal int End { get; private set; }
        internal string SelectionText { get; private set; }
        internal string SelectionFingerprint { get; private set; }
        internal List<WordDraftBlock> Blocks { get; private set; }

        internal static WordDocumentDraft Create(
            string documentKey,
            string documentName,
            WordSelectionInfo selection,
            JObject args,
            WordDocumentDraft existing = null)
        {
            var operation = (args?.Value<string>("operation") ?? string.Empty).Trim().ToLowerInvariant();
            if (operation != "insert" && operation != "replace")
            { throw new WordToolException("DRAFT_INVALID", "草稿操作必须是 insert 或 replace。"); }

            if (existing != null)
            {
                if (!string.Equals(existing.DocumentKey, documentKey, StringComparison.Ordinal))
                { throw new WordToolException("DOCUMENT_CHANGED", "草稿关联的文档已切换；请重新生成草稿。"); }
                if (!string.Equals(existing.Operation, operation, StringComparison.Ordinal))
                { throw new WordToolException("DRAFT_TARGET_CHANGED", "不能在现有草稿上更改插入或替换目标；请重新发起草稿。"); }
            }
            else
            {
                if (selection == null || !selection.HasSelection)
                { throw new WordToolException("TARGET_REQUIRED", "无法确定草稿对应的文档位置。"); }
                if (operation == "replace" && selection.End <= selection.Start)
                { throw new WordToolException("TARGET_INVALID", "替换草稿需要一个非空选区。"); }
                if (operation == "insert" && selection.End != selection.Start)
                { throw new WordToolException("TARGET_AMBIGUOUS", "当前选区有内容；请先确认草稿是替换选区还是插入到其他位置。"); }
            }

            var style = (args?.Value<string>("style") ?? existing?.Style ?? "follow_document").Trim().ToLowerInvariant();
            if (!IsStyle(style)) { throw new WordToolException("DRAFT_STYLE_INVALID", "草稿风格必须是 follow_document、minimal、business 或 report。"); }

            return new WordDocumentDraft
            {
                Id = existing?.Id ?? Guid.NewGuid().ToString("N"),
                Version = existing?.Version ?? CurrentVersion,
                DocumentKey = existing?.DocumentKey ?? documentKey,
                DocumentName = existing?.DocumentName ?? documentName,
                Operation = existing?.Operation ?? operation,
                Style = style,
                Story = existing?.Story ?? selection.Story,
                Start = existing?.Start ?? selection.Start,
                End = existing?.End ?? selection.End,
                SelectionText = existing?.SelectionText ?? selection.Text ?? string.Empty,
                SelectionFingerprint = existing?.SelectionFingerprint ?? selection.TextFingerprint,
                Blocks = ParseBlocks(args?["blocks"] as JArray),
            };
        }

        internal WordDocumentDraft WithStyle(string style)
        {
            style = (style ?? string.Empty).Trim().ToLowerInvariant();
            if (!IsStyle(style)) { throw new WordToolException("DRAFT_STYLE_INVALID", "不支持的草稿风格。"); }
            Style = style;
            return this;
        }

        internal object ToPreviewPayload()
        {
            return new
            {
                id = Id,
                version = Version,
                documentName = DocumentName,
                operation = Operation,
                style = Style,
                target = new { story = Story, start = Start, end = End, text = Operation == "replace" ? SelectionText : string.Empty },
                blocks = Blocks.Select(block => block.ToPayload()).ToArray(),
            };
        }

        internal object ToModelResult()
        {
            return new { draft_id = Id, version = Version, status = "preview", operation = Operation, style = Style, document = DocumentName };
        }

        internal JObject ToExecutionArgs()
        {
            return new JObject
            {
                ["draft_id"] = Id,
                ["version"] = Version,
                ["document_key"] = DocumentKey,
                ["operation"] = Operation,
                ["style"] = Style,
                ["target"] = new JObject
                {
                    ["story"] = Story,
                    ["start"] = Start,
                    ["end"] = End,
                    ["fingerprint"] = SelectionFingerprint ?? string.Empty,
                    ["text"] = SelectionText,
                },
                ["blocks"] = new JArray(Blocks.Select(block => block.ToToken())),
            };
        }

        internal string PlainText()
        {
            return string.Join("\n", Blocks.Select(block => block.PlainText()));
        }

        internal static bool IsStyle(string value)
        {
            return value == "follow_document" || value == "minimal" || value == "business" || value == "report";
        }

        internal static List<WordDraftBlock> ParseBlocks(JArray source)
        {
            if (source == null || source.Count == 0 || source.Count > MaxBlocks)
            { throw new WordToolException("DRAFT_INVALID", "草稿必须包含 1 到 100 个内容块。"); }

            var blocks = new List<WordDraftBlock>();
            var totalLength = 0;
            foreach (var token in source)
            {
                var item = token as JObject;
                if (item == null) { throw new WordToolException("DRAFT_INVALID", "草稿内容块格式无效。"); }
                var type = (item.Value<string>("type") ?? string.Empty).Trim().ToLowerInvariant();
                var block = new WordDraftBlock { Type = type };
                switch (type)
                {
                    case "title":
                    case "paragraph":
                    case "heading":
                    case "quote":
                        block.Level = item.Value<int?>("level") ?? 1;
                        if (type == "heading" && (block.Level < 1 || block.Level > 3))
                        { throw new WordToolException("DRAFT_INVALID", "标题层级只支持 1 到 3。"); }
                        block.Runs = ParseRuns(item);
                        if (block.Runs.Count == 0) { throw new WordToolException("DRAFT_INVALID", "标题或段落内容不能为空。"); }
                        totalLength += block.Runs.Sum(run => run.Text.Length);
                        break;
                    case "bullets":
                    case "numbered":
                        block.Items = ReadTextArray(item["items"] as JArray, MaxItems, "列表");
                        totalLength += block.Items.Sum(value => value.Length);
                        break;
                    case "table":
                        block.Headers = ReadTextArray(item["headers"] as JArray, MaxColumns, "表头");
                        var rows = item["rows"] as JArray;
                        if (block.Headers.Count == 0 || rows == null || rows.Count == 0 || rows.Count > MaxRows)
                        { throw new WordToolException("DRAFT_INVALID", "表格必须包含表头和 1 到 100 行数据。"); }
                        block.Rows = new List<List<string>>();
                        foreach (var rowToken in rows)
                        {
                            var row = ReadTextArray(rowToken as JArray, MaxColumns, "表格行");
                            if (row.Count != block.Headers.Count) { throw new WordToolException("DRAFT_INVALID", "表格每行的列数必须与表头一致。"); }
                            block.Rows.Add(row);
                            totalLength += row.Sum(value => value.Length);
                        }
                        totalLength += block.Headers.Sum(value => value.Length);
                        break;
                    default:
                        throw new WordToolException("DRAFT_INVALID", "不支持的草稿内容块：" + type);
                }

                if (totalLength > MaxTextLength) { throw new WordToolException("DRAFT_TOO_LARGE", "草稿内容超过 50000 个字符。"); }
                blocks.Add(block);
            }
            return blocks;
        }

        private static List<WordDraftRun> ParseRuns(JObject item)
        {
            var source = item["runs"] as JArray;
            var runs = new List<WordDraftRun>();
            if (source != null)
            {
                if (source.Count > 100) { throw new WordToolException("DRAFT_INVALID", "单段强调片段不能超过 100 个。"); }
                foreach (var token in source)
                {
                    var run = token as JObject;
                    if (run == null) { throw new WordToolException("DRAFT_INVALID", "文本片段格式无效。"); }
                    var text = ReadText(run.Value<string>("text"), "文本片段");
                    if (text.Length > 8000) { throw new WordToolException("DRAFT_TOO_LARGE", "单个文本片段超过 8000 个字符。"); }
                    runs.Add(new WordDraftRun { Text = text, Bold = run.Value<bool?>("bold") == true, Italic = run.Value<bool?>("italic") == true });
                }
            }
            else
            {
                var text = ReadText(item.Value<string>("text"), "文本");
                if (text.Length > 8000) { throw new WordToolException("DRAFT_TOO_LARGE", "单段文本超过 8000 个字符。"); }
                runs.Add(new WordDraftRun { Text = text });
            }
            return runs;
        }

        private static List<string> ReadTextArray(JArray source, int max, string label)
        {
            if (source == null || source.Count == 0 || source.Count > max)
            { throw new WordToolException("DRAFT_INVALID", label + "数量超出允许范围。"); }
            return source.Select(token => token.Type == JTokenType.String
                ? ReadText(token.Value<string>(), label)
                : throw new WordToolException("DRAFT_INVALID", label + "必须是文本。"))
                .ToList();
        }

        private static string ReadText(string value, string label)
        {
            if (value == null) { throw new WordToolException("DRAFT_INVALID", label + "不能为空。"); }
            if (value.IndexOfAny(new[] { '\0', '\a', '\f' }) >= 0)
            { throw new WordToolException("DRAFT_INVALID", label + "包含不支持的控制字符。"); }
            value = value.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ').Trim();
            if (value.Length == 0) { throw new WordToolException("DRAFT_INVALID", label + "不能为空。"); }
            return value;
        }
    }

    internal sealed class WordDraftBlock
    {
        internal string Type { get; set; }
        internal int Level { get; set; }
        internal List<WordDraftRun> Runs { get; set; } = new List<WordDraftRun>();
        internal List<string> Items { get; set; } = new List<string>();
        internal List<string> Headers { get; set; } = new List<string>();
        internal List<List<string>> Rows { get; set; } = new List<List<string>>();

        internal string PlainText()
        {
            switch (Type)
            {
                case "bullets":
                case "numbered": return string.Join("\n", Items);
                case "table": return string.Join("\n", new[] { Headers }.Concat(Rows).Select(row => string.Join("\t", row)));
                default: return string.Concat(Runs.Select(run => run.Text));
            }
        }

        internal object ToPayload()
        {
            return new
            {
                type = Type,
                level = Level,
                runs = Runs.Select(run => new { text = run.Text, bold = run.Bold, italic = run.Italic }).ToArray(),
                items = Items,
                headers = Headers,
                rows = Rows,
            };
        }

        internal JObject ToToken()
        {
            return (JObject)JToken.FromObject(new
            {
                type = Type,
                level = Level,
                runs = Runs.Select(run => new { text = run.Text, bold = run.Bold, italic = run.Italic }).ToArray(),
                items = Items,
                headers = Headers,
                rows = Rows,
            });
        }
    }

    internal sealed class WordDraftRun
    {
        internal string Text { get; set; }
        internal bool Bold { get; set; }
        internal bool Italic { get; set; }
    }
}
