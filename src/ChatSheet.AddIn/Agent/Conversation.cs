using System;
using System.Collections.Generic;
using System.Linq;
using ChatSheet.AddIn.Providers;

namespace ChatSheet.AddIn.Agent
{
    /// <summary>
    /// 会话与上下文管理。
    ///
    /// 负责在 token 预算内保留最有价值的上下文：
    /// 系统提示与最近若干轮必须保留，较早的轮次在超限时压缩成摘要。
    /// 工具结果通常体积最大，是首要压缩对象。
    /// </summary>
    internal sealed class Conversation
    {
        private readonly List<ChatMessage> _messages = new List<ChatMessage>();

        internal IReadOnlyList<ChatMessage> Messages => _messages;

        internal int LastPromptTokens { get; private set; }

        internal int LastCompletionTokens { get; private set; }

        internal int TotalPromptTokens { get; private set; }

        internal int TotalCompletionTokens { get; private set; }

        internal void RecordUsage(int promptTokens, int completionTokens)
        {
            if (promptTokens > 0)
            {
                LastPromptTokens = promptTokens;
                TotalPromptTokens += promptTokens;
            }

            if (completionTokens > 0)
            {
                LastCompletionTokens = completionTokens;
                TotalCompletionTokens += completionTokens;
            }
        }

        internal void Add(ChatMessage message)
        {
            _messages.Add(message);
        }

        internal void Clear()
        {
            _messages.Clear();
            LastPromptTokens = 0;
            LastCompletionTokens = 0;
        }

        /// <summary>
        /// 用修复后的消息替换全部历史。
        ///
        /// 只用于摘掉配不上的工具调用/结果：留着它们，服务端会整轮拒绝，
        /// 而历史只增不删，一次故障会把之后每一轮都拖死。
        /// </summary>
        internal void ReplaceAll(IEnumerable<ChatMessage> messages)
        {
            _messages.Clear();
            if (messages != null)
            {
                _messages.AddRange(messages);
            }
        }

        /// <summary>
        /// 估算 token 数。
        ///
        /// 刻意用启发式而非精确分词：真实分词需要引入模型专属分词器，
        /// 体积大且各模型不同；此处只需判断「是否接近预算」，
        /// 按中文约 1.5 字/token、英文约 4 字符/token 估算已足够。
        /// 估算偏保守（偏高），以免真正超限。
        /// </summary>
        internal static int EstimateTokens(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var cjk = 0;
            var other = 0;
            foreach (var c in text)
            {
                // CJK 基本区与扩展区
                if ((c >= 0x4E00 && c <= 0x9FFF) || (c >= 0x3400 && c <= 0x4DBF) || (c >= 0x3000 && c <= 0x303F))
                {
                    cjk++;
                }
                else
                {
                    other++;
                }
            }

            return (int)Math.Ceiling(cjk / 1.5) + (int)Math.Ceiling(other / 4.0);
        }

        internal int EstimateTotalTokens()
        {
            var total = 0;
            foreach (var message in _messages)
            {
                total += EstimateTokens(message.Content);
                foreach (var call in message.ToolCalls)
                {
                    total += EstimateTokens(call.Name) + EstimateTokens(call.ArgumentsJson);
                }

                // 每条消息的角色与结构开销。
                total += 8;
            }

            return total;
        }

        /// <summary>
        /// 触发压缩的占比阈值。
        ///
        /// 取 90% 而非 100%：等到真正超限时，本轮请求已经带着超限的上下文
        /// 发出去了，服务端会自行截断（截掉哪部分不受我们控制）或直接报错。
        /// 提前在 90% 动手，压缩后的这一轮仍能完整发出。
        /// </summary>
        internal const double CompressionThreshold = 0.9;

        /// <summary>
        /// 按需裁剪上下文。
        ///
        /// 达到预算的 90% 即开始压缩，目标是降到 70% 以下，留出后续轮次的余量；
        /// 若压到 90% 以下就停手，下一轮很快又会触发，反复压缩既费时也会
        /// 让模型不断丢失刚建立的上下文。
        ///
        /// 策略：系统提示始终保留；从最早的非系统消息开始压缩工具结果，
        /// 仍不够则整段丢弃并留一条说明。最近两轮永不压缩，
        /// 否则模型会失去刚刚发生的操作记忆而重复劳动。
        /// </summary>
        internal ContextTrimResult TrimToBudget(int budgetTokens)
        {
            var result = new ContextTrimResult { BudgetTokens = budgetTokens };
            result.TokensBefore = EstimateTotalTokens();
            result.TriggerTokens = (int)(budgetTokens * CompressionThreshold);
            // 压缩目标定在 70%，避免刚压完又立刻触发。
            var target = (int)(budgetTokens * 0.7);

            if (result.TokensBefore < result.TriggerTokens)
            {
                result.TokensAfter = result.TokensBefore;
                return result;
            }

            budgetTokens = target;

            // 保护最近的消息：从末尾往前数，保留最后 6 条（约两轮往返）。
            const int protectedTail = 6;
            var firstProtectedIndex = Math.Max(0, _messages.Count - protectedTail);

            // 第一步：压缩较早的工具结果，它们通常最占体积。
            for (var i = 0; i < firstProtectedIndex; i++)
            {
                if (EstimateTotalTokens() <= budgetTokens)
                {
                    break;
                }

                var message = _messages[i];
                // 文本协议下工具结果以 user 消息回灌，角色不再是 Tool，
                // 只看角色会把它们漏掉，于是压缩转去丢真正的对话历史。
                var isToolResult = message.Role == ChatRole.Tool || message.IsTextProtocolToolResult;
                if (!isToolResult || string.IsNullOrEmpty(message.Content))
                {
                    continue;
                }

                if (message.Content.Length <= 120)
                {
                    continue;
                }

                message.Content = Summarize(message.Content);
                result.CompressedToolResults++;
            }

            // 第二步：仍超限则丢弃最早的非系统消息。
            while (EstimateTotalTokens() > budgetTokens)
            {
                var index = _messages.FindIndex(m => m.Role != ChatRole.System);
                if (index < 0 || index >= Math.Max(0, _messages.Count - protectedTail))
                {
                    // 已无可丢弃的消息，只能带着超限继续，交由服务端裁剪。
                    break;
                }

                _messages.RemoveAt(index);
                result.DroppedMessages++;
            }

            if (result.DroppedMessages > 0)
            {
                // 明确告知模型有历史被移除，避免它以为上下文完整。
                var notice = ChatMessage.FromUser(
                    $"（系统提示：为控制上下文长度，本会话较早的 {result.DroppedMessages} 条记录已被移除。" +
                    "若需要早前的数据，请重新读取相关范围。）");

                var insertAt = _messages.FindIndex(m => m.Role != ChatRole.System);
                if (insertAt < 0) { insertAt = _messages.Count; }
                _messages.Insert(insertAt, notice);
            }

            result.TokensAfter = EstimateTotalTokens();
            return result;
        }

        private static string Summarize(string content)
        {
            // 保留头部信息：工具结果的关键字段（范围、影响数量）通常在前部。
            var head = content.Length <= 120 ? content : content.Substring(0, 120);
            return head + $"…（原始结果共 {content.Length} 字符，已压缩）";
        }

        /// <summary>替换或插入系统提示。系统提示每轮都要刷新以反映最新的工作簿状态。</summary>
        internal void SetSystemPrompt(string prompt)
        {
            var existing = _messages.FirstOrDefault(m => m.Role == ChatRole.System);
            if (existing != null)
            {
                existing.Content = prompt;
                return;
            }

            _messages.Insert(0, ChatMessage.FromSystem(prompt));
        }
    }

    internal sealed class ContextTrimResult
    {
        internal int BudgetTokens { get; set; }

        /// <summary>触发压缩的 token 数，即预算的 90%。</summary>
        internal int TriggerTokens { get; set; }

        internal int TokensBefore { get; set; }

        internal int TokensAfter { get; set; }

        internal int CompressedToolResults { get; set; }

        internal int DroppedMessages { get; set; }

        internal bool Trimmed => CompressedToolResults > 0 || DroppedMessages > 0;
    }
}
