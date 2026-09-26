using System.Collections.Generic;
using System.Linq;
using System.Text;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatWord.AddIn.Tools
{
    internal static class WordToolCatalog
    {
        private static object Obj(object properties, params string[] required)
        { return new { type = "object", properties, required, additionalProperties = false }; }
        private static object Str(string description) { return new { type = "string", description }; }
        private static object Int(string description) { return new { type = "integer", description }; }
        private static object Bool(string description) { return new { type = "boolean", description }; }
        private static readonly object Target = new
        {
            type = "object", description = "明确的 Word 目标；每次只提供一种定位方式。插入光标位置仅传 story、start、end；段落、表格、书签和内容控件定位不要混传，也不要补零值或空字符串。正文、页眉、页脚、脚注和批注必须显式指定 story。",
            properties = new
            {
                story = Str("main_text、primary_header、primary_footer、footnotes、endnotes 或 comments。"),
                start = Int("目标 Story 内的字符起点。"), end = Int("目标 Story 内的字符终点。"),
                paragraph_index = Int("目标 Story 内从 1 开始的段落索引。"),
                table_index = Int("目标 Story 内从 1 开始的表格索引。"), row = Int("表格行，从 1 开始。"), column = Int("表格列，从 1 开始。"),
                bookmark = Str("书签名称。"), content_control = Str("内容控件的 Title、Tag 或序号。"),
            }, additionalProperties = false,
        };
        private static object WithTarget(object properties, params string[] required)
        {
            var fields = JObject.FromObject(properties);
            fields.AddFirst(new JProperty("target", JToken.FromObject(Target)));
            return new { type = "object", properties = fields, required, additionalProperties = false };
        }

        internal static readonly IReadOnlyList<ToolDefinition> All = new List<ToolDefinition>
        {
            new ToolDefinition("get_document_info", "读取当前 Word/WPS Writer 文档、宿主和能力摘要。", ToolRisk.Read, Obj(new { })),
            new ToolDefinition("get_selection", "读取当前选区的 Story、Start/End、段落样式、标题层级和表格位置。", ToolRisk.Read, Obj(new { })),
            new ToolDefinition("read_document_structure", "读取文档统计、Story 清单以及书签、字段、内容控件、批注和修订能力。", ToolRisk.Read, Obj(new { max_stories = Int("最多返回的 Story 数，默认 32。") })),
            new ToolDefinition("read_range", "读取明确 Story 范围的显示文本；自动去除段落和表格单元格控制字符并限制长度。", ToolRisk.Read, WithTarget(new { max_chars = Int("最多返回字符数，默认 8000。"), offset = Int("显示文本分页偏移，默认 0。") }, "target")),
            new ToolDefinition("read_paragraphs", "读取指定 Story 中的段落、样式和标题层级；省略 Story 时读取正文。", ToolRisk.Read, Obj(new { story = Str("Story 名称，默认 main_text。"), offset = Int("段落偏移，从 0 开始。"), limit = Int("最多返回段落数，默认 50。") })),
            new ToolDefinition("read_table", "读取指定 Story 内的表格单元格，隐藏单元格结束标记。", ToolRisk.Read, Obj(new { story = Str("Story 名称，默认 main_text。"), table_index = Int("表格索引，从 1 开始。"), offset = Int("行偏移，从 0 开始。"), limit = Int("最多返回行数，默认 50。") }, "table_index")),
            new ToolDefinition("find_text", "在明确 Story/范围内查找文字，不会默认搜索整个文档。", ToolRisk.Read, WithTarget(new { text = Str("要查找的文字。"), match_case = Bool("是否区分大小写。") }, "target", "text")),
            new ToolDefinition("replace_text", "在明确目标范围内替换文字；不提供目标时拒绝执行。", ToolRisk.Write, WithTarget(new { find = Str("要替换的文字；为空时将目标整体替换。"), replace = Str("替换后的文字。") }, "target", "replace")),
            new ToolDefinition("insert_text", "在明确的折叠位置插入文字；光标目标只使用 story、start、end。", ToolRisk.Write, WithTarget(new { text = Str("要插入的文字。") }, "target", "text")),
            new ToolDefinition("delete_range", "删除明确目标范围的文字。", ToolRisk.Write, WithTarget(new { } , "target")),
            new ToolDefinition("format_text", "设置明确范围的字符格式。", ToolRisk.Write, WithTarget(new { bold = Bool("是否加粗。"), italic = Bool("是否倾斜。"), underline = Bool("是否下划线。"), font_size = new { type = "number", description = "字号。" } }, "target")),
            new ToolDefinition("format_paragraph", "设置明确范围所在段落的缩进、段前段后和对齐方式。", ToolRisk.Write, WithTarget(new { alignment = Str("left、center、right 或 justify。"), first_line_indent_chars = new { type = "number", description = "首行缩进字符数；负值表示悬挂缩进。" }, space_before = new { type = "number" }, space_after = new { type = "number" }, keep_with_next = Bool("是否与下一段同页。") }, "target")),
            new ToolDefinition("apply_style", "将明确范围的段落应用现有 Word 样式名称。", ToolRisk.Write, WithTarget(new { style = Str("样式名称，不创建新样式。") }, "target", "style")),
            new ToolDefinition("edit_table", "修改明确表格单元格的显示文本。", ToolRisk.Write, Obj(new { target = Target, table_index = Int("表格索引。"), row = Int("行号。"), column = Int("列号。"), text = Str("单元格新文本。") }, "target", "text")),
            new ToolDefinition("set_page_setup", "修改明确节的页面设置。", ToolRisk.Structure, Obj(new { section_index = Int("节索引，从 1 开始。"), orientation = Str("portrait 或 landscape。"), top_margin = new { type = "number" }, bottom_margin = new { type = "number" }, left_margin = new { type = "number" }, right_margin = new { type = "number" } }, "section_index")),
            new ToolDefinition("insert_page_break", "在明确 Story 位置插入分页符。", ToolRisk.Structure, WithTarget(new { }, "target")),
        };

        internal static string PromptSection(bool textProtocol)
        {
            var builder = new StringBuilder();
            if (textProtocol)
            {
                builder.AppendLine("当前连接使用文本工具协议，没有原生函数调用。调用工具时只输出一个信息串为 `" + TextToolProtocol.BlockTag + "` 的围栏代码块，块内使用 `tool` 和 `args` JSON 字段；输出后等待工具结果，不要伪造执行结果。工具块对用户隐藏。");
                builder.AppendLine("示例：");
                builder.AppendLine("```" + TextToolProtocol.BlockTag);
                builder.AppendLine("{\"tool\":\"get_selection\",\"args\":{}}");
                builder.AppendLine("```");
            }

            builder.AppendLine("Word 工具清单：");
            foreach (var tool in All)
            {
                var prefix = textProtocol ? "- `" + tool.Name + "(" + ParameterNames(tool) + ")`：" : "- " + tool.Name + "：";
                builder.AppendLine(prefix + tool.Description);
            }
            return builder.ToString().TrimEnd();
        }

        private static string ParameterNames(ToolDefinition tool)
        {
            var schema = JObject.FromObject(tool.Parameters);
            var properties = schema["properties"] as JObject;
            if (properties == null) { return string.Empty; }
            var required = new HashSet<string>((schema["required"] as JArray ?? new JArray())
                .Select(item => item.Value<string>()).Where(name => !string.IsNullOrEmpty(name)), System.StringComparer.Ordinal);
            return string.Join(", ", properties.Properties()
                .Select(property => property.Name + (required.Contains(property.Name) ? string.Empty : "?")));
        }
    }
}
