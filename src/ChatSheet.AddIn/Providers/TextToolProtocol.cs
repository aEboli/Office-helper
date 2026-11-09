using System;
using System.Collections.Generic;
using System.Text;
using ChatSheet.AddIn.Tools;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Providers
{
    /// <summary>从模型正文里解析出的一次工具调用。</summary>
    internal sealed class TextToolCall
    {
        internal string Name { get; set; }

        internal string ArgumentsJson { get; set; }

        /// <summary>原始块在正文中的起止，供闸门决定吞掉哪一段。</summary>
        internal int Start { get; set; }

        internal int Length { get; set; }
    }

    /// <summary>
    /// 不支持原生函数调用时的替代协议：工具清单写进系统提示，
    /// 模型用围栏代码块发出调用。
    ///
    /// 为什么用围栏块而不是自定义标签（如 &lt;tool&gt;）：围栏块是模型最熟练的
    /// 结构，弱模型也很少写坏；自定义标签一旦少个尖括号就整块失效。
    /// 也因此必须容错——信息串写错、多写一层缩进都不该让调用丢失。
    /// </summary>
    internal static class TextToolProtocol
    {
        /// <summary>围栏的信息串。</summary>
        internal const string BlockTag = "chatsheet:tool";

        /// <summary>
        /// 工具清单的紧凑签名文本。
        ///
        /// 保留参数名、类型和必填/可选标记；通用对象入口另保留完整路径约定。
        /// </summary>
        internal static string CatalogText()
        {
            var builder = new StringBuilder();
            foreach (var tool in ToolCatalog.All)
            {
                builder.Append("- `").Append(tool.Name).Append('(');
                builder.Append(string.Join(", ", ParameterNames(tool)));
                builder.Append(")`");

                if (tool.RequiresApproval)
                {
                    // 标出来是为了让模型知道这些调用可能被用户拦下，
                    // 被拒绝时不要绕道重试。
                    builder.Append("（写操作，可能需要用户批准）");
                }

                builder.Append('：').Append(tool.Name == "excel_object" ? tool.Description : FirstSentence(tool.Description));
                builder.AppendLine();
            }

            return builder.ToString().TrimEnd();
        }

        /// <summary>参数名与类型清单，可选参数带 ? 后缀。</summary>
        private static IReadOnlyList<string> ParameterNames(ToolDefinition tool)
        {
            var names = new List<string>();
            try
            {
                var schema = JObject.FromObject(tool.Parameters);
                var properties = schema["properties"] as JObject;
                if (properties == null)
                {
                    return names;
                }

                var required = new HashSet<string>(StringComparer.Ordinal);
                if (schema["required"] is JArray requiredArray)
                {
                    foreach (var item in requiredArray)
                    {
                        var name = item.Value<string>();
                        if (!string.IsNullOrEmpty(name)) { required.Add(name); }
                    }
                }

                foreach (var property in properties.Properties())
                {
                    var type = property.Value["type"];
                    var typeName = type is JArray types ? string.Join("|", types.Values<string>()) : type?.Value<string>();
                    names.Add((required.Contains(property.Name) ? property.Name : property.Name + "?")
                        + (string.IsNullOrEmpty(typeName) ? string.Empty : ":" + typeName));
                }
            }
            catch
            {
                // 取不到参数名不该让整份清单失效：模型仍能从说明里知道这个工具做什么，
                // 参数写错会得到明确的工具错误。
            }

            return names;
        }

        /// <summary>取说明的第一句。清单要短，完整说明留给原生声明。</summary>
        private static string FirstSentence(string description)
        {
            if (string.IsNullOrEmpty(description))
            {
                return string.Empty;
            }

            var index = description.IndexOf('。');
            return index > 0 ? description.Substring(0, index + 1) : description;
        }

        /// <summary>系统提示里描述协议的那一段。</summary>
        internal static string PromptSection()
        {
            var builder = new StringBuilder();

            builder.AppendLine("## 如何操作表格");
            builder.AppendLine("你没有原生的函数调用通道，改用下面的方式动手。需要调用工具时，" +
                "输出一个围栏代码块，信息串写 " + BlockTag + "，块内是一个 JSON 对象：");
            builder.AppendLine();
            builder.AppendLine("```" + BlockTag);
            builder.AppendLine("{\"tool\": \"read_range\", \"args\": {\"range\": \"A1:D20\"}}");
            builder.AppendLine("```");
            builder.AppendLine();
            builder.AppendLine("规则：");
            builder.AppendLine("- 块内只放这一个 JSON 对象，不要加注释或额外文字。`tool` 是工具名，`args` 是参数对象（没有参数也要写 `{}`）。");
            builder.AppendLine("- 一次回复可以发多个块，会按先后顺序执行。");
            builder.AppendLine("- 发出块后就停下，等系统把执行结果回给你，再决定下一步。不要自己编造结果。");
            builder.AppendLine("- 用户看不到这些块，只会看到操作卡片。因此块外要用一句话说明你在做什么。");
            builder.AppendLine("- 普通代码示例照常用普通围栏块（例如 ```text），只有信息串是 " +
                BlockTag + " 的块会被当作调用。");
            builder.AppendLine();
            builder.AppendLine("可用工具：");
            builder.AppendLine(CatalogText());

            return builder.ToString().TrimEnd();
        }

        /// <summary>
        /// 判断一个围栏块的内容是不是工具调用，是则给出调用。
        ///
        /// 宽松匹配：信息串写错（甚至没写）也认，只要块内 JSON 同时有 tool 与 args、
        /// 且工具名确实存在于清单。这种巧合在普通代码块里不成立，
        /// 而弱模型漏写信息串是常态。
        /// </summary>
        internal static bool TryParseBlockBody(string infoString, string body, out TextToolCall call)
        {
            call = null;
            if (string.IsNullOrWhiteSpace(body))
            {
                return false;
            }

            var tagged = (infoString ?? string.Empty).Trim()
                .Equals(BlockTag, StringComparison.OrdinalIgnoreCase);

            JObject json;
            try
            {
                json = JObject.Parse(body.Trim());
            }
            catch
            {
                // 信息串明确标了工具块，内容却不是合法 JSON：这仍是一次调用意图，
                // 必须交上去，好让模型收到「参数不是合法 JSON」而不是被静默忽略。
                if (tagged)
                {
                    call = new TextToolCall { Name = null, ArgumentsJson = body.Trim() };
                    return true;
                }

                return false;
            }

            var nameToken = json["tool"] ?? json["name"];
            var name = nameToken?.Type == JTokenType.String ? nameToken.Value<string>() : null;
            var args = json["args"] ?? json["arguments"] ?? json["parameters"];

            if (string.IsNullOrWhiteSpace(name))
            {
                if (tagged)
                {
                    call = new TextToolCall { ArgumentsJson = body.Trim() };
                    return true;
                }
                return false;
            }

            // 没标信息串时要求工具名真实存在，避免把讲解用的示例 JSON 当成调用。
            if (!tagged && (args == null || ToolCatalog.Find(name) == null))
            {
                return false;
            }

            call = new TextToolCall
            {
                Name = name.Trim(),
                ArgumentsJson = args == null ? "{}" : args.ToString(Newtonsoft.Json.Formatting.None),
            };

            return true;
        }
    }
}
