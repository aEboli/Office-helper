using System;
using System.Collections.Generic;
using System.Text;

namespace ChatSheet.AddIn.Providers
{
    /// <summary>
    /// 文本协议下的流式闸门：把模型的正文分成「给用户看的」与「工具调用」两股。
    ///
    /// 必须在流式过程中就拦住指令块，而不是等整段收完再清理：正文是逐字推给面板的，
    /// 一旦 JSON 的头几个字已经进了气泡，再撤回就要改已经渲染的内容。
    /// 因此策略是「疑似即攥住」——看到围栏开头就停止外放，等到能判定时再决定
    /// 整块吞掉还是原样放行。普通代码块只会因此晚一点显示，不会显示错。
    /// </summary>
    internal sealed class TextToolGate
    {
        private const string Fence = "```";

        private readonly StringBuilder _buffer = new StringBuilder();
        private readonly List<TextToolCall> _calls = new List<TextToolCall>();

        /// <summary>已确认是围栏块的开头，正在等闭合。</summary>
        private bool _inBlock;
        private bool _bufferAtLineStart = true;

        internal IReadOnlyList<TextToolCall> Calls => _calls;

        /// <summary>本次交付里是否出现过指令块。</summary>
        internal bool SawToolBlock => _calls.Count > 0;

        /// <summary>吃进一段增量，返回其中可以立即显示的部分。</summary>
        internal string Push(string delta)
        {
            if (!string.IsNullOrEmpty(delta))
            {
                _buffer.Append(delta);
            }

            return Drain(final: false);
        }

        /// <summary>
        /// 流结束时收束。
        ///
        /// 未闭合的块必须在这里定性：模型可能被长度上限截在块中间，
        /// 一直攥着会让这段文字彻底消失，用户看到的是一句没头没尾的话。
        /// </summary>
        internal string Flush()
        {
            return Drain(final: true);
        }

        private string Drain(bool final)
        {
            var visible = new StringBuilder();

            while (true)
            {
                var text = _buffer.ToString();

                if (!_inBlock)
                {
                    var fenceIndex = FindFenceAtLineStart(text, 0);
                    if (fenceIndex < 0)
                    {
                        // 没有围栏开头。末尾若正好是「半个围栏」，留到下一段再判断，
                        // 否则会先把 ``` 的前两个反引号显示出去。
                        var hold = final ? 0 : PartialFenceTail(text);
                        visible.Append(text, 0, text.Length - hold);
                        ConsumePrefix(text.Length - hold);
                        break;
                    }

                    visible.Append(text, 0, fenceIndex);
                    ConsumePrefix(fenceIndex);
                    _inBlock = true;
                    continue;
                }

                // 攥住状态：text 以围栏开头。
                var infoEnd = text.IndexOf('\n');
                if (infoEnd < 0)
                {
                    // 信息串这一行还没收完。
                    if (!final)
                    {
                        break;
                    }

                    // 流断在信息串上，块里什么都没有，原样交出去。
                    visible.Append(text);
                    ConsumePrefix(text.Length);
                    _inBlock = false;
                    break;
                }

                var fenceLength = Fence.Length;
                while (fenceLength < infoEnd && text[fenceLength] == '`') { fenceLength++; }
                var infoString = text.Substring(fenceLength, infoEnd - fenceLength).Trim();
                var closeIndex = FindClosingFence(text, infoEnd + 1, fenceLength, final);

                if (closeIndex < 0 && !final)
                {
                    break;
                }

                var bodyEnd = closeIndex < 0 ? text.Length : closeIndex;
                var body = text.Substring(infoEnd + 1, bodyEnd - infoEnd - 1);

                if (TextToolProtocol.TryParseBlockBody(infoString, body, out var call))
                {
                    _calls.Add(call);
                }
                else
                {
                    // 不是指令块，整块原样放行——包括围栏本身，
                    // Markdown 渲染要靠它才知道这是代码。
                    visible.Append(text, 0, closeIndex < 0 ? text.Length : EndOfFenceLine(text, closeIndex));
                }

                var consumed = closeIndex < 0 ? text.Length : EndOfFenceLine(text, closeIndex);
                ConsumePrefix(consumed);
                _inBlock = false;

                if (closeIndex < 0)
                {
                    break;
                }
            }

            return visible.ToString();
        }

        private void ConsumePrefix(int count)
        {
            if (count == 0) { return; }
            // 增量边界不等于行首；放行前半行后，下一帧仍可能是同一行的示例。
            _bufferAtLineStart = _buffer[count - 1] == '\n';
            _buffer.Remove(0, count);
        }

        private int FindClosingFence(string text, int from, int fenceLength, bool final)
        {
            while (true)
            {
                var index = FindFenceAtLineStart(text, from);
                if (index < 0) { return -1; }
                var lineEnd = text.IndexOf('\n', index);
                // 等整行收齐再判断，否则下一帧的正文可能证明它不是闭合围栏。
                if (lineEnd < 0 && !final) { return -1; }
                if (lineEnd < 0) { lineEnd = text.Length; }
                var end = index;
                while (end < lineEnd && text[end] == '`') { end++; }
                var enoughTicks = end - index >= fenceLength;
                while (end < lineEnd && (text[end] == ' ' || text[end] == '\t' || text[end] == '\r')) { end++; }
                if (enoughTicks && end == lineEnd) { return index; }
                if (lineEnd == text.Length) { return -1; }
                from = lineEnd + 1;
            }
        }

        /// <summary>围栏闭合行的结束位置（含换行符）。</summary>
        private static int EndOfFenceLine(string text, int fenceIndex)
        {
            var newline = text.IndexOf('\n', fenceIndex);
            return newline < 0 ? text.Length : newline + 1;
        }

        /// <summary>找出行首的围栏位置。围栏必须顶行，否则缩进的代码里到处都是。</summary>
        private int FindFenceAtLineStart(string text, int from)
        {
            for (var i = from; i + Fence.Length <= text.Length; i++)
            {
                if (text[i] != '`')
                {
                    continue;
                }

                if (i == 0 ? !_bufferAtLineStart : text[i - 1] != '\n')
                {
                    continue;
                }

                if (string.CompareOrdinal(text, i, Fence, 0, Fence.Length) == 0)
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// 末尾可能正在形成围栏的字符数。
        ///
        /// 只需看最后几个字符：围栏是行首的三个反引号，因此「换行 + 一两个反引号」
        /// 或「开头的一两个反引号」都要留着。多攥几个字符只是晚一帧显示。
        /// </summary>
        private int PartialFenceTail(string text)
        {
            if (text.Length == 0)
            {
                return 0;
            }

            var backticks = 0;
            var index = text.Length - 1;
            while (index >= 0 && text[index] == '`' && backticks < Fence.Length)
            {
                backticks++;
                index--;
            }

            if (backticks == 0)
            {
                // 结尾是换行时也要留着：下一段的头一个字符可能就是反引号。
                return text[text.Length - 1] == '\n' ? 1 : 0;
            }

            // 反引号要顶行才算围栏，因此前一个字符必须是换行或就是开头。
            var atLineStart = index < 0 ? _bufferAtLineStart : text[index] == '\n';
            if (!atLineStart)
            {
                return 0;
            }

            return index < 0 ? backticks : backticks + 1;
        }
    }
}
