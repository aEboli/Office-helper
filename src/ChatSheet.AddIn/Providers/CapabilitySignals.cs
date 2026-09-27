using System;

namespace ChatSheet.AddIn.Providers
{
    /// <summary>
    /// 从服务端错误与模型正文里认出「这个模型缺哪项能力」。
    ///
    /// 全是启发式，因为没有任何协议提供「该模型支持什么」的查询接口——
    /// GET /models 只给名字。判断错的代价是多跑一步（换个形态重试），
    /// 因此宁可宽一点也不要漏：漏了就是整轮失败，用户只看到一条原始 400。
    ///
    /// IsClientError 与 Mentions 是 internal，供 ModelAvailability 复用——
    /// 可用性判定与这里同源同风格，各写一份迟早会分叉。
    /// </summary>
    internal static class CapabilitySignals
    {
        /// <summary>
        /// 只对客户端错误做能力判定。
        ///
        /// 5xx 是服务端故障，重试同一个请求就可能成功，交给 RetryPolicy；
        /// 把它当成「不支持工具」会让一次网关抖动永久降级掉这个模型。
        /// </summary>
        internal static bool IsClientError(ProviderException ex)
        {
            return ex != null &&
                ex.Code != null &&
                ex.Code.StartsWith("HTTP_4", StringComparison.Ordinal);
        }

        internal static bool Mentions(string text, params string[] needles)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            foreach (var needle in needles)
            {
                if (text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 错误是否在说「我不支持工具/函数调用」。
        ///
        /// 认字段名（tools、tool_choice、functionDeclarations）比认自然语言可靠：
        /// 各家的措辞五花八门，但字段名来自协议，是固定的。
        /// </summary>
        internal static bool LooksLikeToolUnsupported(ProviderException ex)
        {
            if (!IsClientError(ex))
            {
                return false;
            }

            // 错误在说「问题出在模型本身」时，它不是任何能力信号。
            if (BlamesModelItself(ex))
            {
                return false;
            }

            var message = ex.Message ?? string.Empty;

            // 图片相关的错误绝不能落到这里：两条回退链各自记档，
            // 混淆会让「看不了图」把工具也一并降级掉。
            if (LooksLikeVisionUnsupported(ex))
            {
                return false;
            }

            return Mentions(
                message,
                "tools",
                "tool_choice",
                "tool_calls",
                "function call",
                "function_call",
                "functioncalling",
                "function calling",
                "functiondeclarations",
                "工具调用",
                "不支持函数");
        }

        /// <summary>
        /// 错误是否在说「我不支持图片输入」。
        ///
        /// 本判据认裸子串 image，而模型名本身就可能含 image
        /// （gpt-image-1、*-image-preview），那样一条 404 会被读成「不支持图片」。
        /// 排除靠 BlamesModelItself，详见那里。
        /// </summary>
        internal static bool LooksLikeVisionUnsupported(ProviderException ex)
        {
            if (!IsClientError(ex))
            {
                return false;
            }

            if (BlamesModelItself(ex))
            {
                return false;
            }

            var message = ex.Message ?? string.Empty;

            return Mentions(
                message,
                "image_url",
                "input_image",
                "inlinedata",
                "image",
                "vision",
                "multimodal",
                "multi-modal",
                "media_type",
                "图片",
                "图像",
                "视觉",
                "多模态");
        }

        /// <summary>
        /// 错误是否在说「问题出在这个模型本身」，而不是在说缺哪项能力。
        ///
        /// 这是两条能力判据共同的前置排除。没有它就有一个真实缺陷：
        /// LooksLikeVisionUnsupported 认裸子串 "image"，于是选 gpt-image-1、附一张图、
        /// 发送时，那句 `model 'gpt-image-1' does not exist` 是 4xx 且含 image，
        /// 会被记成「不支持图片输入」——白花一次视觉中转请求去描述图片、剥掉所有图、
        /// 再用同一个不存在的模型重试一遍，最后告诉用户「当前模型没有视觉能力」。
        /// 一条错误产生两个记录，其中一个是假的，而假的那个才是用户看到的。
        ///
        /// 判定委托给 ModelAvailability：那边只读 Detail（服务端原文），
        /// 不读拼过 hint 的 Message，避免我们自己的「请检查……模型名……」变成证据。
        /// </summary>
        private static bool BlamesModelItself(ProviderException ex)
        {
            return ModelAvailability.BlamesModel(ex);
        }

        /// <summary>
        /// 错误是否在说「输出上限那个字段名用错了」。
        ///
        /// OpenAI 的推理模型只接受 max_completion_tokens，对 max_tokens 回 400，
        /// 错误里会点名这个字段。认字段名而不是认模型名：模型名与行为没有可靠对应，
        /// 而字段名来自协议，是固定的。
        ///
        /// 判据要求同时出现「这个字段」与「不支持/请改用」这类措辞——只出现字段名
        /// 不够，网关把请求体回显在错误里时每条错误都会含 max_tokens。
        /// </summary>
        internal static bool LooksLikeOutputLimitFieldWrong(ProviderException ex)
        {
            if (!IsClientError(ex))
            {
                return false;
            }

            // 这条判据只对 Chat Completions 有意义，但不在这里判协议：
            // 调用方比这里更清楚自己发的是什么。
            var text = (ex.Detail ?? string.Empty) + "\n" + (ex.Message ?? string.Empty);

            if (!Mentions(text, "max_tokens", "max_completion_tokens"))
            {
                return false;
            }

            return Mentions(
                text,
                "unsupported",
                "not supported",
                "use \"max_completion_tokens\"",
                "use 'max_completion_tokens'",
                "use max_completion_tokens",
                "instead",
                "unrecognized",
                "unknown parameter",
                "invalid parameter",
                "不支持",
                "请改用");
        }

        /// <summary>
        /// 正文是否明确表示不具备工具能力，或推辞「我碰不到你的表格」。
        ///
        /// 服务端可能收下工具声明、不报任何错，模型却一个调用都不发。
        /// 直接说「我没有工具调用能力」也是能力信号，不必再要求它提到表格。
        ///
        /// 普通的「我不能」仍不足以触发降级：模型拒绝越权请求也会这么说，
        /// 那是正确行为，不该因此切换协议。
        /// </summary>
        internal static bool LooksLikeToolRefusal(string assistantText)
        {
            if (string.IsNullOrWhiteSpace(assistantText))
            {
                return false;
            }

            if (Mentions(
                assistantText,
                "不支持工具调用",
                "不支持函数调用",
                "不支持调用工具",
                "不支持调用函数",
                "不具备工具调用能力",
                "不具备工具调用的能力",
                "不具备调用工具的能力",
                "没有工具调用能力",
                "没有工具调用的能力",
                "没有调用工具的能力",
                "无法进行工具调用",
                "不能进行工具调用",
                "无法调用工具",
                "不能调用工具",
                "工具调用不受支持",
                "函数调用不受支持",
                "tool calling is not supported",
                "function calling is not supported",
                "tool calls are not supported",
                "does not support tool calling",
                "doesn't support tool calling",
                "do not support tool calling",
                "don't support tool calling",
                "does not support function calling",
                "doesn't support function calling",
                "cannot call tools",
                "can't call tools",
                "unable to call tools",
                "cannot invoke tools",
                "can't invoke tools",
                "unable to invoke tools",
                "do not have the ability to call tools",
                "don't have the ability to call tools",
                "do not have tool calling capability",
                "don't have tool calling capability",
                "do not have tool-calling capability",
                "don't have tool-calling capability",
                "tool calling is unavailable",
                "function calling is unavailable",
                "tool calls are unavailable"))
            {
                return true;
            }

            // 兼容旧有表述：先要出现「做不到」，再要求它谈到工作簿。
            var deniesAbility = Mentions(
                assistantText,
                "无法访问",
                "无法直接访问",
                "不能访问",
                "没有权限",
                "无法读取",
                "无法获取",
                "无法操作",
                "无法直接操作",
                "无法修改",
                "不能修改",
                "看不到",
                "无法查看",
                "没有办法访问",
                "不具备访问",
                "无法连接",
                "cannot access",
                "can't access",
                "unable to access",
                "do not have access",
                "don't have access",
                "no access to",
                "cannot read",
                "unable to read",
                "cannot modify");

            if (!deniesAbility)
            {
                return false;
            }

            return Mentions(
                assistantText,
                "表格",
                "工作簿",
                "工作表",
                "单元格",
                "excel",
                "spreadsheet",
                "workbook",
                "worksheet",
                "cell");
        }
    }
}
