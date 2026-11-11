using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ChatSheet.AddIn.Providers
{
    /// <summary>一条已分帧的 SSE 事件。</summary>
    internal sealed class SseFrame
    {
        internal string EventName { get; set; }

        internal string Data { get; set; }
    }

    /// <summary>
    /// SSE 分帧读取器。
    ///
    /// 四种协议的事件内容差异很大，但分帧规则相同：以空行分隔事件，
    /// 每行形如 "field: value"，同一事件内多行 data 需以换行拼接。
    /// 把分帧独立出来，各协议解析器只面对已完整的一帧。
    ///
    /// 不使用 ReadLineAsync 逐行读：它无法区分「流结束」与「服务端仍在思考」，
    /// 且在只发心跳注释的连接上会造成误判。这里按字节推进，显式处理 \r\n 与 \n。
    /// </summary>
    internal static class SseReader
    {
        internal static async Task ReadAsync(
            Stream stream,
            Func<SseFrame, Task<bool>> onFrame,
            CancellationToken cancellationToken)
        {
            if (stream == null) { throw new ArgumentNullException(nameof(stream)); }
            if (onFrame == null) { throw new ArgumentNullException(nameof(onFrame)); }

            var buffer = new byte[8192];
            var pending = new StringBuilder();
            var decoder = new UTF8Encoding(false).GetDecoder();
            // 解码器可能带着上一批未完成的四字节字符，本批会额外产出一个代理项。
            var chars = new char[Encoding.UTF8.GetMaxCharCount(buffer.Length)];

            var eventName = (string)null;
            var dataLines = new List<string>();
            var firstCharacter = true;

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var read = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);

                // 用增量解码器：多字节字符可能跨越两次读取的边界，
                // 直接 GetString 会把中文等字符切坏。
                var charCount = decoder.GetChars(buffer, 0, read, chars, 0, flush: read == 0);
                var offset = firstCharacter && charCount > 0 && chars[0] == '\uFEFF' ? 1 : 0;
                if (charCount > 0) { firstCharacter = false; }
                pending.Append(chars, offset, charCount - offset);
                // EOF 也结束最后一行；否则没有末尾换行的最后一条 data 会丢失。
                if (read == 0 && pending.Length > 0) { pending.Append('\n'); }

                while (true)
                {
                    var lineEnd = FindLineEnd(pending, out var skip);
                    if (lineEnd < 0)
                    {
                        break;
                    }

                    var line = pending.ToString(0, lineEnd);
                    pending.Remove(0, lineEnd + skip);

                    if (line.Length == 0)
                    {
                        // 空行表示一帧结束。
                        if (dataLines.Count > 0)
                        {
                            var frame = new SseFrame
                            {
                                EventName = eventName,
                                Data = string.Join("\n", dataLines),
                            };

                            dataLines.Clear();
                            eventName = null;

                            if (!await onFrame(frame).ConfigureAwait(false))
                            {
                                return;
                            }
                        }
                        else
                        {
                            eventName = null;
                        }

                        continue;
                    }

                    // 以冒号开头的是注释/心跳，忽略。
                    if (line[0] == ':')
                    {
                        continue;
                    }

                    var colon = line.IndexOf(':');
                    string field, value;
                    if (colon < 0)
                    {
                        field = line;
                        value = string.Empty;
                    }
                    else
                    {
                        field = line.Substring(0, colon);
                        value = line.Substring(colon + 1);
                        // 规范要求去掉冒号后的一个前导空格。
                        if (value.Length > 0 && value[0] == ' ')
                        {
                            value = value.Substring(1);
                        }
                    }

                    switch (field)
                    {
                        case "event":
                            eventName = value;
                            break;
                        case "data":
                            dataLines.Add(value);
                            break;
                        default:
                            // id、retry 等字段本项目不需要。
                            break;
                    }
                }
                if (read == 0) { break; }
            }

            cancellationToken.ThrowIfCancellationRequested();
            // 流结束时若仍有未闭合的一帧，也要交付，避免丢掉最后一条消息。
            if (dataLines.Count > 0)
            {
                await onFrame(new SseFrame
                {
                    EventName = eventName,
                    Data = string.Join("\n", dataLines),
                }).ConfigureAwait(false);
            }
        }

        /// <summary>找到行尾位置，并返回需要跳过的分隔符长度（\r\n 为 2）。</summary>
        private static int FindLineEnd(StringBuilder builder, out int separatorLength)
        {
            for (var i = 0; i < builder.Length; i++)
            {
                var c = builder[i];
                if (c == '\n')
                {
                    separatorLength = 1;
                    return i;
                }

                if (c == '\r')
                {
                    // \r 出现在缓冲末尾时无法判断是否为 \r\n，等下一批数据。
                    if (i + 1 >= builder.Length)
                    {
                        break;
                    }

                    separatorLength = builder[i + 1] == '\n' ? 2 : 1;
                    return i;
                }
            }

            separatorLength = 0;
            return -1;
        }
    }
}
