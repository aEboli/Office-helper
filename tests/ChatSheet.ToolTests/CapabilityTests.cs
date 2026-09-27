using System;
using System.Linq;
using ChatSheet.AddIn.Providers;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    /// <summary>
    /// 模型能力回退验证。
    ///
    /// 这一块全是启发式：没有任何协议提供「该模型支持什么」的查询，只能从
    /// 服务端错误文本和模型正文里认。因此判据本身就是最容易出错的部分——
    /// 认宽了会把正常的 401 当成「不支持工具」而白白降级，认窄了则整轮失败。
    /// 两个方向都要有反例。
    /// </summary>
    internal static class CapabilityTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            TestErrorClassification(report);
            TestRefusalDetection(report);
            TestCatalogText(report);
            TestBlockParsing(report);
            TestGate(report);
            TestCapabilityStore(report);
        }

        private static ProviderException Http(int status, string message)
        {
            return new ProviderException("HTTP_" + status, message);
        }

        private static void TestErrorClassification(Action<string, bool, string> report)
        {
            report(
                "400 提到 tools 判为不支持工具",
                CapabilitySignals.LooksLikeToolUnsupported(
                    Http(400, "接口返回 400 Bad Request：model does not support tools")),
                "");

            report(
                "400 提到 function calling 判为不支持工具",
                CapabilitySignals.LooksLikeToolUnsupported(
                    Http(400, "This model does not support Function Calling")),
                "");

            report(
                "400 提到 tool_choice 判为不支持工具",
                CapabilitySignals.LooksLikeToolUnsupported(Http(400, "unknown parameter: tool_choice")),
                "");

            // 反例：配置类错误绝不能被当成能力缺失，否则用户填错密钥
            // 会得到「已改用文本指令」这种毫不相干的提示。
            report(
                "401 密钥错误不判为不支持工具",
                !CapabilitySignals.LooksLikeToolUnsupported(Http(401, "invalid api key")),
                "");

            // 反例：5xx 是服务端故障，重试同一请求可能就成功了。
            // 判成能力缺失会让一次网关抖动永久降级掉这个模型。
            report(
                "503 不判为不支持工具",
                !CapabilitySignals.LooksLikeToolUnsupported(Http(503, "tools upstream unavailable")),
                "");

            report(
                "400 提到 image_url 判为不支持视觉",
                CapabilitySignals.LooksLikeVisionUnsupported(
                    Http(400, "invalid content type image_url for this model")),
                "");

            report(
                "400 提到 vision 判为不支持视觉",
                CapabilitySignals.LooksLikeVisionUnsupported(Http(400, "model has no vision capability")),
                "");

            // 两条回退链必须互斥：混在一起会让「看不了图」把工具也降级掉。
            report(
                "图片类错误不落进工具判据",
                !CapabilitySignals.LooksLikeToolUnsupported(
                    Http(400, "this model does not support image input")),
                "");

            report(
                "工具类错误不落进视觉判据",
                !CapabilitySignals.LooksLikeVisionUnsupported(Http(400, "unknown parameter: tool_choice")),
                "");
        }

        private static void TestRefusalDetection(Action<string, bool, string> report)
        {
            report(
                "「无法访问你的表格」判为推辞",
                CapabilitySignals.LooksLikeToolRefusal("抱歉，我无法访问你的 Excel 表格，请把数据贴给我。"),
                "");

            report(
                "「看不到工作簿」判为推辞",
                CapabilitySignals.LooksLikeToolRefusal("我看不到你的工作簿内容。"),
                "");

            report(
                "英文 cannot access spreadsheet 判为推辞",
                CapabilitySignals.LooksLikeToolRefusal("I cannot access your spreadsheet directly."),
                "");

            report(
                "中文明确表示没有工具调用能力",
                CapabilitySignals.LooksLikeToolRefusal("我没有工具调用的能力。"),
                "");

            report(
                "中文明确表示不支持函数调用",
                CapabilitySignals.LooksLikeToolRefusal("当前模型不支持函数调用。"),
                "");

            report(
                "英文明确表示不能调用 tools",
                CapabilitySignals.LooksLikeToolRefusal("I don't have the ability to call tools."),
                "");

            report(
                "英文明确表示 function calling 不受支持",
                CapabilitySignals.LooksLikeToolRefusal("This model does not support function calling."),
                "");

            // 反例：普通安全拒绝不代表工具能力缺失。
            report(
                "拒绝无关请求不判为推辞",
                !CapabilitySignals.LooksLikeToolRefusal("我不能帮你写病毒程序。"),
                "");

            report(
                "英文安全拒绝不判为推辞",
                !CapabilitySignals.LooksLikeToolRefusal("I can't help create malware."),
                "");

            report(
                "只讨论工具调用不判为推辞",
                !CapabilitySignals.LooksLikeToolRefusal("I can use function calling to read the workbook."),
                "");

            // 反例：正常作答里会大量出现「表格」二字。
            report(
                "正常作答不判为推辞",
                !CapabilitySignals.LooksLikeToolRefusal("已读取表格，A1:D20 共 80 个单元格，其中 3 处为空。"),
                "");

            report(
                "空正文不判为推辞",
                !CapabilitySignals.LooksLikeToolRefusal(""),
                "");
        }

        private static void TestCatalogText(Action<string, bool, string> report)
        {
            var text = TextToolProtocol.CatalogText();

            report("清单含 read_range", text.Contains("read_range"), "");
            report("清单含 write_values", text.Contains("write_values"), "");
            report("清单标出必填参数 range", text.Contains("range"), "");
            report("清单用 ? 标出可选参数", text.Contains("sheet?"), "");
            report("清单标出写操作需批准", text.Contains("写操作"), "");

            // 紧凑签名的意义就在于比原生声明短得多。原生声明约 2100 token，
            // 这里若也涨到同一量级，弱模型仍会读不出重点。
            var nativeSize = ChatSheet.AddIn.Tools.ToolCatalog.All
                .Sum(t => JObject.FromObject(t.Parameters).ToString().Length);
            report(
                "清单显著短于原生 JSON Schema",
                text.Length * 2 < nativeSize,
                $"清单 {text.Length} 字符，schema 合计 {nativeSize} 字符");

            var prompt = TextToolProtocol.PromptSection();
            report("提示段含信息串", prompt.Contains(TextToolProtocol.BlockTag), "");
            report("提示段含示例", prompt.Contains("read_range"), "");
        }

        private static void TestBlockParsing(Action<string, bool, string> report)
        {
            var ok = TextToolProtocol.TryParseBlockBody(
                TextToolProtocol.BlockTag,
                "{\"tool\": \"read_range\", \"args\": {\"range\": \"A1:B2\"}}",
                out var call);
            report("解析标准指令块", ok && call.Name == "read_range", ok ? call.Name : "未解析");
            report("解析出参数 JSON", ok && call.ArgumentsJson.Contains("A1:B2"), ok ? call.ArgumentsJson : "");

            // 弱模型漏写信息串是常态，但块内同时有 tool 与 args、
            // 且工具名确实存在时不可能是巧合，必须认。
            var lenient = TextToolProtocol.TryParseBlockBody(
                "json",
                "{\"tool\": \"get_selection\", \"args\": {}}",
                out var lenientCall);
            report("漏写信息串仍能识别", lenient && lenientCall.Name == "get_selection", "");

            // 反例：讲解 JSON 结构时的示例不该被当成调用。
            var notACall = TextToolProtocol.TryParseBlockBody(
                "json",
                "{\"name\": \"张三\", \"age\": 30}",
                out _);
            report("普通 JSON 不被当作调用", !notACall, "");

            var unknownLenient = TextToolProtocol.TryParseBlockBody(
                "json",
                "{\"tool\": \"launch_missiles\", \"args\": {}}",
                out _);
            report("未标信息串且工具名不存在时不认", !unknownLenient, "");

            // 标了信息串的未知工具要认下来，好让模型收到「未知工具」这个明确错误，
            // 而不是被静默忽略后干等结果。
            var unknownTagged = TextToolProtocol.TryParseBlockBody(
                TextToolProtocol.BlockTag,
                "{\"tool\": \"launch_missiles\", \"args\": {}}",
                out var unknownCall);
            report("标了信息串的未知工具仍上交", unknownTagged && unknownCall.Name == "launch_missiles", "");

            // 被截断的块也要上交：模型需要收到「参数被截断」才会改为分批，
            // 静默丢弃只会让它干等一个永远不来的结果。
            var broken = TextToolProtocol.TryParseBlockBody(
                TextToolProtocol.BlockTag,
                "{\"tool\": \"write_values\", \"args\": {\"range\": \"A1:B2\", \"valu",
                out var brokenCall);
            report("坏 JSON 在标了信息串时仍上交", broken, "");
            report("坏 JSON 上交时没有工具名", broken && string.IsNullOrEmpty(brokenCall.Name), "");

            report(
                "空块不被当作调用",
                !TextToolProtocol.TryParseBlockBody(TextToolProtocol.BlockTag, "   ", out _),
                "");
        }

        private static void TestGate(Action<string, bool, string> report)
        {
            // 一次性喂完：最基本的分流。
            var gate = new TextToolGate();
            var visible = gate.Push(
                "我来读一下。\n```" + TextToolProtocol.BlockTag + "\n" +
                "{\"tool\": \"read_range\", \"args\": {\"range\": \"A1:B2\"}}\n```\n读完再说。\n");
            visible += gate.Flush();

            report("闸门放行块外正文", visible.Contains("我来读一下"), visible);
            report("闸门吞掉指令块", !visible.Contains("read_range"), visible);
            report("闸门放行块后正文", visible.Contains("读完再说"), visible);
            report("闸门收下调用", gate.Calls.Count == 1, gate.Calls.Count.ToString());
            report("闸门报告见过指令块", gate.SawToolBlock, "");

            // 逐字喂：真实流式就是这样来的，一次一两个字符。
            // 这条最容易出问题——围栏的三个反引号会被拆到不同的增量里。
            var streamed = new TextToolGate();
            var source = "先看一眼。\n```" + TextToolProtocol.BlockTag + "\n" +
                "{\"tool\": \"get_selection\", \"args\": {}}\n```\n好了。";
            var accumulated = string.Empty;
            foreach (var c in source)
            {
                accumulated += streamed.Push(c.ToString());
            }

            accumulated += streamed.Flush();

            report("逐字喂也能吞掉指令块", !accumulated.Contains("get_selection"), accumulated);
            report("逐字喂放行前后正文", accumulated.Contains("先看一眼") && accumulated.Contains("好了"), accumulated);
            report("逐字喂解析出一个调用", streamed.Calls.Count == 1, streamed.Calls.Count.ToString());

            // 普通代码块必须原样放行，包括围栏本身——Markdown 渲染要靠它。
            var plain = new TextToolGate();
            var plainText = plain.Push("公式如下：\n```text\n=SUM(A1:A9)\n```\n照此填入。");
            plainText += plain.Flush();

            report("普通代码块原样放行", plainText.Contains("=SUM(A1:A9)"), plainText);
            report("普通代码块保留围栏", plainText.Contains("```"), plainText);
            report("普通代码块不产生调用", plain.Calls.Count == 0, plain.Calls.Count.ToString());

            // 未闭合的块：模型被长度上限截在块中间。攥着不放会让这段文字
            // 彻底消失，用户看到一句没头没尾的话。
            var truncated = new TextToolGate();
            var truncatedText = truncated.Push("正在写入。\n```" + TextToolProtocol.BlockTag + "\n{\"tool\": \"write_val");
            truncatedText += truncated.Flush();

            report("未闭合块的块外正文仍交付", truncatedText.Contains("正在写入"), truncatedText);
            report(
                "未闭合的指令块被收束而非丢失",
                truncated.Calls.Count == 1 || truncatedText.Contains("write_val"),
                $"调用 {truncated.Calls.Count} 个，正文 {truncatedText}");

            // 未闭合的普通代码块要把文字交出去。
            var openPlain = new TextToolGate();
            var openText = openPlain.Push("例子：\n```text\n=SUM(A1:A9)");
            openText += openPlain.Flush();
            report("未闭合普通块的内容仍交付", openText.Contains("=SUM(A1:A9)"), openText);

            // 行中的反引号不是围栏，不该触发攥住。
            var inline = new TextToolGate();
            var inlineText = inline.Push("请在 `A1` 里填 1。");
            inlineText += inline.Flush();
            report("行内反引号原样放行", inlineText == "请在 `A1` 里填 1。", inlineText);

            // 多个块按顺序执行。
            var multi = new TextToolGate();
            var multiText = multi.Push(
                "两步走。\n```" + TextToolProtocol.BlockTag + "\n{\"tool\": \"get_selection\", \"args\": {}}\n```\n" +
                "```" + TextToolProtocol.BlockTag + "\n{\"tool\": \"read_range\", \"args\": {\"range\": \"A1\"}}\n```\n完成。");
            multiText += multi.Flush();

            report("一次回复里的两个块都解析出来", multi.Calls.Count == 2, multi.Calls.Count.ToString());
            report(
                "两个块按出现顺序",
                multi.Calls.Count == 2 && multi.Calls[0].Name == "get_selection" && multi.Calls[1].Name == "read_range",
                multi.Calls.Count == 2 ? multi.Calls[0].Name + "," + multi.Calls[1].Name : "");
            report("多块之间的正文仍交付", multiText.Contains("两步走") && multiText.Contains("完成"), multiText);
        }

        private static void TestCapabilityStore(Action<string, bool, string> report)
        {
            ModelCapabilities.Reset();

            var a = ModelCapabilities.For("CustomApi|openai|https://a/v1", "m1");
            a.ToolMode = ToolProtocolMode.Text;

            // 同一个模型名经不同网关转发，能力可以完全不同。
            var b = ModelCapabilities.For("CustomApi|openai|https://b/v1", "m1");
            report("换连接后档案独立", b.ToolMode == ToolProtocolMode.Native, b.ToolMode.ToString());

            var again = ModelCapabilities.For("CustomApi|openai|https://a/v1", "m1");
            report("同一连接同一模型取到同一档案", again.ToolMode == ToolProtocolMode.Text, again.ToolMode.ToString());

            // 手动指定要盖过探测结果，否则设置页那个选项等于没生效。
            report(
                "手动指定原生时忽略探测结果",
                ModelCapabilities.ResolveMode(ToolProtocolPreference.Native, a) == ToolProtocolMode.Native,
                "");

            report(
                "自动探测时采用档案",
                ModelCapabilities.ResolveMode(ToolProtocolPreference.Auto, a) == ToolProtocolMode.Text,
                "");

            ModelCapabilities.Reset();
            var fresh = ModelCapabilities.For("CustomApi|openai|https://a/v1", "m1");
            report("重置后档案回到默认", fresh.ToolMode == ToolProtocolMode.Native, fresh.ToolMode.ToString());

            report(
                "DeepSeek V4 Flash 自动采用文本协议",
                ModelCapabilities.ResolveMode(
                    ToolProtocolPreference.Auto,
                    fresh,
                    "deepseek-v4-flash") == ToolProtocolMode.Text,
                "");

            report(
                "带 provider 前缀的 DeepSeek Flash 自动采用文本协议",
                ModelCapabilities.ResolveMode(
                    ToolProtocolPreference.Auto,
                    fresh,
                    "deepseek/deepseek-v4-flash-vision-preview") == ToolProtocolMode.Text,
                "");

            report(
                "DeepSeek Flash 型号匹配不区分大小写",
                ModelCapabilities.IsDeepSeekFlashModel("DeepSeek-V4-FLASH") &&
                    ModelCapabilities.IsDeepSeekFlashModel("deepseek-flash"),
                "");

            report(
                "普通 Flash 型号不强制文本协议",
                !ModelCapabilities.IsDeepSeekFlashModel("some-flash-model") &&
                    ModelCapabilities.ResolveMode(
                        ToolProtocolPreference.Auto,
                        fresh,
                        "some-flash-model") == ToolProtocolMode.Native,
                "");

            report(
                "DeepSeek Flashlight 型号不强制文本协议",
                !ModelCapabilities.IsDeepSeekFlashModel("deepseek-flashlight"),
                "");

            report(
                "手动原生覆盖 DeepSeek Flash 兼容提示",
                ModelCapabilities.ResolveMode(
                    ToolProtocolPreference.Native,
                    fresh,
                    "deepseek-v4-flash") == ToolProtocolMode.Native,
                "");

            report(
                "手动指定时不再探测",
                !ModelCapabilities.DetectionEnabled(ToolProtocolPreference.Text) &&
                    ModelCapabilities.DetectionEnabled(ToolProtocolPreference.Auto),
                "");

            TestTextProtocolTally(report);
        }

        /// <summary>
        /// 未命中计数的归零。
        ///
        /// 「有进展就归零」是这套计数唯一容易写错的地方：只增不清的话，
        /// 文本协议下每一句寒暄都记一笔，攒够两笔就把模型降级成顾问——
        /// 而降级是不可见地把它的动手能力拿掉，之后它再也改不了表格。
        /// </summary>
        private static void TestTextProtocolTally(Action<string, bool, string> report)
        {
            // 一步没有指令块还不能判死：可能它正在正常作答。
            var first = ModelCapabilities.TallyTextProtocolStep(0, madeProgress: false);
            report("首次未命中不降级", first.Misses == 1 && !first.ShouldDegrade, first.Misses.ToString());

            // 连续第二步仍无，才判定它不会用。
            var second = ModelCapabilities.TallyTextProtocolStep(first.Misses, madeProgress: false);
            report("连续两次未命中才降级", second.Misses == 2 && second.ShouldDegrade, second.Misses.ToString());

            // 关键：中间有一步成功，计数必须归零，不能接着往上数。
            var afterProgress = ModelCapabilities.TallyTextProtocolStep(1, madeProgress: true);
            report(
                "有进展即归零",
                afterProgress.Misses == 0 && !afterProgress.ShouldDegrade,
                afterProgress.Misses.ToString());

            // 归零之后再来一次未命中，只能算第 1 次。
            var afterReset = ModelCapabilities.TallyTextProtocolStep(afterProgress.Misses, madeProgress: false);
            report(
                "归零后重新从 1 数起",
                afterReset.Misses == 1 && !afterReset.ShouldDegrade,
                afterReset.Misses.ToString());

            // 「未命中、命中、未命中、未命中」这条序列最能说明问题：
            // 前两次未命中之间隔着一次命中，因此只有最后两次连续的才触发降级。
            var tally = ModelCapabilities.TallyTextProtocolStep(0, false);   // 1
            var hit = ModelCapabilities.TallyTextProtocolStep(tally.Misses, true);  // 0
            var again = ModelCapabilities.TallyTextProtocolStep(hit.Misses, false); // 1
            report("命中打断连续性", !again.ShouldDegrade && again.Misses == 1, again.Misses.ToString());

            var final = ModelCapabilities.TallyTextProtocolStep(again.Misses, false); // 2
            report("此后再连续一次才降级", final.ShouldDegrade, final.Misses.ToString());

            report("阈值为 2", ModelCapabilities.MissesBeforeAdvisor == 2,
                ModelCapabilities.MissesBeforeAdvisor.ToString());

            // 计数不该留在档案上：档案跨轮存活，过程计数留在那里就会跨轮累加。
            var capability = ModelCapabilities.For("k", "m");
            var fields = typeof(ModelCapability).GetProperties();
            var hasMissCounter = false;
            foreach (var field in fields)
            {
                if (field.Name.IndexOf("Miss", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    hasMissCounter = true;
                }
            }

            report("档案上不存过程计数", !hasMissCounter && capability != null,
                hasMissCounter ? "档案仍带未命中计数，会跨轮累加" : "");
        }
    }
}
