using System;
using System.Collections.Generic;

namespace ChatSheet.AddIn.Providers
{
    /// <summary>
    /// 请求发出前的结构修复：保证「助手的工具调用」与「工具结果」成对。
    ///
    /// 这条配平是服务端的硬要求，措辞各家不同，但都整轮拒绝：
    ///   OpenAI 系：An assistant message with 'tool_calls' must be followed by
    ///               tool messages responding to each 'tool_call_id'.
    ///   网关：      No tool output found for function call fc_txt1-739998.
    ///
    /// 会话只要跑过几步，就可能留下配不上的半边：中途取消、执行抛异常、
    /// 上下文裁剪只丢了结果。坏消息留在历史里会被每一轮重发，症状是
    /// 「模型刚才还好好的，之后不管换哪个模型都整轮失败」，而错误里提到的
    /// 调用标识早已无从追溯。因此在唯一的出口处摘掉不成对的部分：
    /// 调用缺结果就只留正文，结果缺调用就整条丢弃。
    /// </summary>
    internal static class MessageRepair
    {
        /// <summary>助手消息的调用全部缺结果、正文又为空时，用来顶替正文的一句说明。</summary>
        internal const string OrphanedCallPlaceholder = "（此前的工具调用没有返回结果，已跳过。）";

        internal sealed class Result
        {
            internal List<ChatMessage> Messages { get; } = new List<ChatMessage>();

            /// <summary>摘掉的、没有结果的工具调用数。</summary>
            internal int DroppedCalls { get; set; }

            /// <summary>摘掉的、没有调用的工具结果数。</summary>
            internal int DroppedResults { get; set; }

            internal bool Changed => DroppedCalls > 0 || DroppedResults > 0;
        }

        internal static Result Repair(IReadOnlyList<ChatMessage> messages)
        {
            var result = new Result();
            if (messages == null)
            {
                return result;
            }

            var index = 0;
            while (index < messages.Count)
            {
                var message = messages[index];
                if (message == null)
                {
                    index++;
                    continue;
                }

                // 前面没有调用的结果：整条丢弃。留着会被服务端当成
                // 「结果找不到对应的调用」而拒绝。
                if (message.Role == ChatRole.Tool)
                {
                    result.DroppedResults++;
                    index++;
                    continue;
                }

                if (message.Role != ChatRole.Assistant || message.ToolCalls.Count == 0)
                {
                    result.Messages.Add(message);
                    index++;
                    continue;
                }

                // 结果必须紧跟调用，中间夹任何消息都不成立，因此只看后面
                // 连续的一段 tool 消息。
                var following = index + 1;
                var results = new List<ChatMessage>();
                while (following < messages.Count && messages[following] != null
                    && messages[following].Role == ChatRole.Tool)
                {
                    results.Add(messages[following]);
                    following++;
                }

                var matched = new bool[results.Count];
                var kept = new List<ToolCall>();
                foreach (var call in message.ToolCalls)
                {
                    var hit = FindResult(results, matched, call);
                    if (hit < 0)
                    {
                        result.DroppedCalls++;
                        continue;
                    }

                    matched[hit] = true;
                    kept.Add(call);
                }

                if (kept.Count == message.ToolCalls.Count)
                {
                    result.Messages.Add(message);

                    // 逐条放行，不能整段 AddRange：紧跟其后的那段结果里可能混着
                    // 不属于任何调用的孤儿（重放的历史、被裁掉调用的残留、
                    // 并发分支互相覆盖后的剩饭）。
                    //
                    // OpenAI 要求每条 role:"tool" 都回应紧邻那条 assistant 的某个
                    // tool_call_id，多出来的那条会让整轮 400——这正是本类存在的理由。
                    // 更糟的是「全部配上」是绝大多数情况，于是最常走的分支恰好是唯一
                    // 没设防的那个：故障只在历史里已经混进孤儿结果时出现，看起来像偶发。
                    for (var i = 0; i < results.Count; i++)
                    {
                        if (matched[i])
                        {
                            result.Messages.Add(results[i]);
                        }
                        else
                        {
                            result.DroppedResults++;
                        }
                    }
                }
                else if (kept.Count == 0)
                {
                    // 整条调用都不成立：退化成一条普通助手消息。
                    // 正文为空时补一句说明——Anthropic 不接受空的正文块。
                    var text = string.IsNullOrWhiteSpace(message.Content)
                        ? OrphanedCallPlaceholder
                        : message.Content;
                    result.Messages.Add(message.CloneWithToolCalls(null, text));
                    result.DroppedResults += results.Count;
                }
                else
                {
                    result.Messages.Add(message.CloneWithToolCalls(kept));
                    for (var i = 0; i < results.Count; i++)
                    {
                        if (matched[i])
                        {
                            result.Messages.Add(results[i]);
                        }
                        else
                        {
                            result.DroppedResults++;
                        }
                    }
                }

                index = following;
            }

            return result;
        }

        /// <summary>
        /// 找这条调用对应的结果。有标识就按标识配对；Gemini 与部分自建代理
        /// 不带标识，只能退回按工具名配对。
        /// </summary>
        private static int FindResult(List<ChatMessage> results, bool[] matched, ToolCall call)
        {
            for (var i = 0; i < results.Count; i++)
            {
                if (matched[i])
                {
                    continue;
                }

                var id = results[i].ToolCallId;
                if (!string.IsNullOrEmpty(call?.Id))
                {
                    if (string.Equals(id, call.Id, StringComparison.Ordinal))
                    {
                        return i;
                    }

                    continue;
                }

                if (string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(call?.Name)
                    && string.Equals(results[i].ToolName, call.Name, StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return -1;
        }
    }
}
