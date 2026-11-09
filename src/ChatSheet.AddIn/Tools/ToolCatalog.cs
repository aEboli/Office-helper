using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Tools
{
    /// <summary>工具的风险级别，决定是否需要用户审批。</summary>
    internal enum ToolRisk
    {
        /// <summary>只读，自动执行。</summary>
        Read = 0,

        /// <summary>会修改工作簿内容，需要审批。</summary>
        Write = 1,

        /// <summary>会改变工作簿结构（增删工作表等），需要审批。</summary>
        Structure = 2,
    }

    internal sealed class ToolDefinition
    {
        internal ToolDefinition(string name, string description, ToolRisk risk, object parameters)
        {
            Name = name;
            Description = description;
            Risk = risk;
            Parameters = parameters;
        }

        internal string Name { get; }

        internal string Description { get; }

        internal ToolRisk Risk { get; }

        /// <summary>JSON Schema 形态的参数定义，直接作为函数声明发给模型。</summary>
        internal object Parameters { get; }

        internal bool RequiresApproval => Risk != ToolRisk.Read;

        // 在审批和快照预处理前校验必填项及顶层类型；具体内容仍由各工具校验。
        internal ToolResult ValidateArguments(JObject args)
        {
            return ValidateArguments(args, null);
        }

        /// <summary>
        /// 带跳过集合的校验：<paramref name="skip"/> 里的键完全不做检查（必填与类型都跳过）。
        ///
        /// 存在的理由是 excel_object 的 root：它是必填项，但缺省时会由宿主弹一次
        /// 「选哪一层」把值补上。必须在弹窗之前先把其余参数（path/action/member）
        /// 验完——否则模型少写 path，用户白点一次才收到 ARG_MISSING，
        /// 而那次点击换不来任何结果。
        /// </summary>
        internal ToolResult ValidateArguments(JObject args, IReadOnlyCollection<string> skip)
        {
            args = args ?? new JObject();
            var schema = JObject.FromObject(Parameters);

            var requiredKeys = new HashSet<string>(StringComparer.Ordinal);
            if (schema["required"] is JArray required)
            {
                foreach (var key in required.Values<string>())
                {
                    requiredKeys.Add(key);
                    if (IsSkipped(skip, key)) { continue; }

                    var value = args[key];
                    if (value == null || value.Type == JTokenType.Null
                        || (value.Type == JTokenType.String && string.IsNullOrWhiteSpace(value.Value<string>())))
                    {
                        return ToolResult.Failure("ARG_MISSING", $"缺少必需参数 {key}。请按工具参数定义补全后重试。");
                    }
                }
            }

            if (!(schema["properties"] is JObject properties))
            {
                return null;
            }

            foreach (var parameter in args.Properties())
            {
                if (IsSkipped(skip, parameter.Name)) { continue; }

                // schema 里只有对象形态的条目才有 type；形状不对就交给工具自己报错，
                // 不在这一层抛异常——校验层崩掉会让整轮对话失败。
                if (!(properties[parameter.Name] is JObject property)) { continue; }
                var type = property["type"];
                if (type == null) { continue; }

                var types = type is JArray array ? array.Values<string>().ToArray() : new[] { type.Value<string>() };

                // 非必填参数写 null 等于「没提供」。执行层本来就是这么读的
                // （OptionalBool 见 null 返回 null、args.Value<T?> 见 null 返回 null），
                // 这里若按 schema 的 type 严格拒绝，就会出现「校验比执行更严」：
                // 模型把可选项写成 null 被 ARG_INVALID 挡下，换个写法重发又能过，白烧一轮。
                if (parameter.Value.Type == JTokenType.Null && !requiredKeys.Contains(parameter.Name))
                {
                    continue;
                }

                if (!MatchesType(types, parameter.Value))
                {
                    return ToolResult.Failure(
                        "ARG_INVALID",
                        $"参数 {parameter.Name} 必须为 {string.Join(" 或 ", types)}，实际为 {DescribeValue(parameter.Value)}。请按工具参数定义重试。");
                }
            }

            return null;
        }

        private static bool IsSkipped(IReadOnlyCollection<string> skip, string name)
        {
            return skip != null && skip.Contains(name, StringComparer.Ordinal);
        }

        /// <summary>
        /// 顶层类型是否可接受。
        ///
        /// 标量放宽到执行层同一口径：工具里读的是 args.Value&lt;long?&gt;()、
        /// Value&lt;double?&gt;()、Value&lt;bool&gt;()，Newtonsoft 会先做一次字符串转换，
        /// 因此 "0"、"12.5"、"true"/"false" 在执行层是能读出来的。
        /// 校验若比执行更严，模型只是换了个等价写法就被拒，而错误里说的
        /// 「必须为 integer」又和工具实际行为不符，它只会把同一个值反复改写。
        ///
        /// 数组与对象不在此列：它们没有无歧义的字符串形态，
        /// 放行只会把错误推迟到更深的地方，且更难解释。
        /// </summary>
        private static bool MatchesType(string[] types, JToken value)
        {
            var valueType = value.Type;
            foreach (var t in types)
            {
                switch (t)
                {
                    case "string":
                        if (valueType == JTokenType.String) { return true; }
                        break;
                    case "array":
                        if (valueType == JTokenType.Array) { return true; }
                        break;
                    case "object":
                        if (valueType == JTokenType.Object) { return true; }
                        break;
                    case "null":
                        if (valueType == JTokenType.Null) { return true; }
                        break;
                    case "boolean":
                        if (valueType == JTokenType.Boolean) { return true; }
                        if (valueType == JTokenType.String && bool.TryParse(value.Value<string>(), out _)) { return true; }
                        break;
                    case "integer":
                        // 整数值的浮点（12.0）执行层也能读成 long，按「无损」放行；
                        // 12.5 会丢小数，仍按类型错误拒绝。
                        if (valueType == JTokenType.Integer) { return true; }
                        if (valueType == JTokenType.Float && IsIntegral(value.Value<double>())) { return true; }
                        if (valueType == JTokenType.String
                            && long.TryParse(value.Value<string>(), NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                        {
                            return true;
                        }
                        break;
                    case "number":
                        if (valueType == JTokenType.Integer || valueType == JTokenType.Float) { return true; }
                        if (valueType == JTokenType.String
                            && double.TryParse(value.Value<string>(), NumberStyles.Float | NumberStyles.AllowThousands,
                                CultureInfo.InvariantCulture, out _))
                        {
                            return true;
                        }
                        break;
                }
            }

            return false;
        }

        private static bool IsIntegral(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value) && Math.Truncate(value) == value;
        }

        /// <summary>错误信息里的实际取值。字符串要带出来，否则模型只知道「不对」。</summary>
        private static string DescribeValue(JToken value)
        {
            if (value == null) { return "null"; }
            if (value.Type == JTokenType.String)
            {
                var text = value.Value<string>() ?? string.Empty;
                if (text.Length > 40) { text = text.Substring(0, 40) + "…"; }
                return $"字符串 \"{text}\"";
            }

            return value.Type.ToString();
        }
    }

    /// <summary>
    /// 工具清单。
    ///
    /// 边界是刻意收紧的：只暴露表格操作，不提供文件系统、shell 或网络访问。
    /// 这既是安全约束，也让模型的选择空间足够小，从而更可靠。
    /// </summary>
    internal static class ToolCatalog
    {
        private static object Obj(object properties, params string[] required)
        {
            return new
            {
                type = "object",
                properties,
                required,
                additionalProperties = false,
            };
        }

        private static object Str(string description)
        {
            return new { type = "string", description };
        }

        private static object OptStr(string description)
        {
            return new { type = new[] { "string", "null" }, description };
        }

        // 可空布尔：OptionalBool 读 null 时按「未提供」处理，
        // schema 里也必须承认 null，否则校验层与执行层的口径就不一致
        // （模型按要求写 null 表达「不改这一项」，却被 ARG_INVALID 挡下）。
        private static object OptBool(string description)
        {
            return new { type = new[] { "boolean", "null" }, description };
        }

        private static readonly object SheetProp = OptStr("工作表名称。省略则使用当前活动工作表。");

        internal static readonly IReadOnlyList<ToolDefinition> All = new List<ToolDefinition>
        {
            new ToolDefinition("excel_object",
                "操作当前工作簿的 Excel COM 对象模型：行列增删、复制、筛选、边框、条件格式、验证、命名区域、图表、透视表、冻结窗格、页面设置。" +
                "root 为 workbook、worksheet、window；path 按序取属性或调用返回对象的方法（kind=get/call，member，args数组）。" +
                "最后 action=get/set/call 对 member 操作；args 用官方 COM 位置参数和枚举数值，省略参数用 {missing:true}。" +
                "对象参数用 {ref:{root,sheet,path}}；二维数据用 {matrix:[[...]]}。" +
                "例如隐藏行：root=worksheet,path=[{member:Range,args:[\"2:4\"]},{member:EntireRow}],action=set,member=Hidden,args=[true]。" +
                "修改后必须读回验证；无自动撤销。数据读取优先 read_range 分页。只支持当前工作簿表格对象，不提供文件、宏或应用程序入口。",
                ToolRisk.Structure,
                Obj(new {
                    root = Str("workbook、worksheet 或 window。"), sheet = SheetProp,
                    path = new { type = "array", items = new { type = "object", properties = new {
                        kind = OptStr("get（默认）或 call。"), member = Str("官方对象模型成员名称。"),
                        args = new { type = "array", items = new { } } } } },
                    action = Str("get、set 或 call。"), member = Str("最终成员名称。"),
                    args = new { type = "array", items = new { } }
                }, "root", "path", "action", "member")),
            new ToolDefinition(
                "get_workbook_info",
                "获取当前工作簿的结构摘要：文件名、工作表清单、各表已用范围与行列数。开始任何任务前应先调用它了解全局。",
                ToolRisk.Read,
                Obj(new { })),

            new ToolDefinition(
                "get_selection",
                "获取用户当前选中的范围地址与尺寸。用户说“这里”“这一列”时用它确定目标。",
                ToolRisk.Read,
                Obj(new { })),

            new ToolDefinition(
                "read_range",
                "读取指定范围，每页最多 5000 格。返回 next_offset 时用原范围和该 offset 继续，直到 null。",
                ToolRisk.Read,
                Obj(
                    new
                    {
                        range = Str("范围地址，例如 A1:D20。"),
                        sheet = SheetProp,
                        include_formulas = OptBool("为真时同时返回公式文本，而不只是计算结果。"),
                        offset = new { type = "integer", description = "按行遍历的单元格偏移，默认 0，续页使用 next_offset。" },
                    },
                    "range")),

            new ToolDefinition(
                "write_values",
                "向范围写入字面值。values 的行列数必须与范围尺寸一致。大范围自动分块写入；输入仍需适合模型的单次输出长度。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("目标范围地址，例如 A1:C3。"),
                        sheet = SheetProp,
                        values = new
                        {
                            type = "array",
                            description = "二维数组，外层为行、内层为列。元素可为字符串、数字、布尔或 null。",
                            items = new { type = "array", items = new { type = new[] { "string", "number", "boolean", "null" } } },
                        },
                    },
                    "range", "values")),

            new ToolDefinition(
                "write_formulas",
                "向范围写入公式。每个元素需以 = 开头。大范围自动分块写入；输入仍需适合模型的单次输出长度。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("目标范围地址。"),
                        sheet = SheetProp,
                        formulas = new
                        {
                            type = "array",
                            description = "二维公式数组，外层为行、内层为列，元素为以 = 开头的公式文本。",
                            items = new { type = "array", items = new { type = "string" } },
                        },
                    },
                    "range", "formulas")),

            new ToolDefinition(
                "format_range",
                "设置范围的字体与填充样式。只需提供要改变的属性，未提供的保持原样。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("目标范围地址。"),
                        sheet = SheetProp,
                        bold = new { type = new[] { "boolean", "null" }, description = "加粗。" },
                        italic = new { type = new[] { "boolean", "null" }, description = "倾斜。" },
                        font_size = new { type = new[] { "number", "null" }, description = "字号，单位磅。" },
                        font_color = OptStr("字体颜色，十六进制如 #FF0000。"),
                        fill_color = OptStr("填充色，十六进制如 #FFFF00。"),
                        horizontal_alignment = OptStr("水平对齐：left、center、right。"),
                        vertical_alignment = OptStr("垂直对齐：top、center、bottom。"),
                        wrap_text = new { type = new[] { "boolean", "null" }, description = "自动换行。" },
                    },
                    "range")),

            new ToolDefinition(
                "set_number_format",
                "设置范围的数字格式代码，例如 0.00、#,##0、yyyy-mm-dd、0.0%。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("目标范围地址。"),
                        sheet = SheetProp,
                        format_code = Str("Excel 数字格式代码。"),
                    },
                    "range", "format_code")),

            new ToolDefinition(
                "autofit_range",
                "按内容自动调整列宽或行高。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("目标范围地址，例如 A:D 或 A1:D20。"),
                        sheet = SheetProp,
                        target = Str("调整对象：columns 调列宽，rows 调行高。"),
                    },
                    "range", "target")),

            new ToolDefinition(
                "fit_range",
                "一次完成适配：设置水平与垂直对齐，按内容调整列宽后将超过 100 的列限制为 100，启用自动换行并调整行高。" +
                "用户说「适配」「排版整理一下」这类要求时优先用它，比分别调用 format_range 与 autofit_range 更省步数。" +
                "省略 range 表示适配整张表的已用范围，这也是面板「适配」按钮的行为。不受单元格数量限制。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = OptStr("目标范围地址。省略则取该表的已用范围。"),
                        sheet = SheetProp,
                        horizontal_alignment =
                            OptStr("水平对齐：left、center、right。省略为 center。垂直方向固定居中。"),
                    })),

            new ToolDefinition(
                "merge_cells",
                "把范围合并成一个单元格。用户说「合并单元格」「跨列居中」「把标题横过来」时用它。" +
                "只有左上角单元格的内容会保留，其余内容会被丢弃，因此合并前应先读一遍范围确认没有要保住的值。" +
                "大范围也可执行，但超出快照预算时不提供自动撤销。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("要合并的范围地址，例如 A1:D1。必须多于一个单元格。"),
                        sheet = SheetProp,
                        across = OptBool(
                            "为真时逐行分别合并（每行合成一格，行与行不相连），为假或省略时整片合成一格。"),
                        horizontal_alignment =
                            OptStr("同时设置水平对齐：left、center、right。省略则保持原有对齐。"),
                        vertical_alignment =
                            OptStr("同时设置垂直对齐：top、center、bottom。省略则保持原有对齐。"),
                    },
                    "range")),

            new ToolDefinition(
                "unmerge_cells",
                "取消范围内的单元格合并，把合并区域拆回独立单元格。" +
                "原合并区域的内容留在左上角单元格，其余单元格为空。" +
                "大范围也可执行，但超出快照预算时不提供自动撤销。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("范围地址。范围内与之相交的合并区域都会被拆开。"),
                        sheet = SheetProp,
                    },
                    "range")),

            new ToolDefinition(
                "clear_range",
                "清除范围内容、格式或两者。这是破坏性操作。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("目标范围地址。"),
                        sheet = SheetProp,
                        scope = Str("清除范围：contents 仅内容，formats 仅格式，all 全部。"),
                    },
                    "range", "scope")),

            new ToolDefinition(
                "add_worksheet",
                "新增一张工作表。",
                ToolRisk.Structure,
                Obj(
                    new
                    {
                        name = Str($"新工作表名称，最长 {ToolLimits.MaxSheetNameLength} 个字符。"),
                        after_sheet = OptStr("插入到该工作表之后。省略则置于末尾。"),
                    },
                    "name")),

            new ToolDefinition(
                "rename_worksheet",
                "重命名工作表。",
                ToolRisk.Structure,
                Obj(
                    new
                    {
                        old_name = Str("现有工作表名称。"),
                        new_name = Str($"新名称，最长 {ToolLimits.MaxSheetNameLength} 个字符。"),
                    },
                    "old_name", "new_name")),

            new ToolDefinition(
                "sort_range",
                "按指定列排序范围。",
                ToolRisk.Write,
                Obj(
                    new
                    {
                        range = Str("待排序范围地址，应包含数据区域。"),
                        sheet = SheetProp,
                        key_column = Str("排序依据的列字母，例如 B。必须落在 range 之内。"),
                        ascending = OptBool("为真升序，为假降序。"),
                        has_header = OptBool("范围首行是否为标题行。"),
                    },
                    "range", "key_column")),

            new ToolDefinition(
                "create_table",
                "把范围转换为表格（带筛选与样式）。",
                ToolRisk.Structure,
                Obj(
                    new
                    {
                        range = Str("表格范围地址，应包含标题行。"),
                        sheet = SheetProp,
                        name = OptStr("表格名称。省略则由宿主自动命名。"),
                        has_header = OptBool("首行是否为标题行。"),
                    },
                    "range")),

            new ToolDefinition(
                "create_chart",
                "基于数据范围创建图表。",
                ToolRisk.Structure,
                Obj(
                    new
                    {
                        range = Str("数据范围地址。"),
                        sheet = SheetProp,
                        chart_type = Str("图表类型：column、bar、line、pie、scatter、area。"),
                        title = OptStr("图表标题。"),
                    },
                    "range", "chart_type")),
        };

        internal static ToolDefinition Find(string name)
        {
            return All.FirstOrDefault(t => string.Equals(t.Name, name, StringComparison.Ordinal));
        }
    }
}
