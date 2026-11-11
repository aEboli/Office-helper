using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn.Providers;

namespace ChatSheet.ToolTests
{
    /// <summary>
    /// 重试计数的归零验证。
    ///
    /// policy 层的单测（RetryTests）只验「第几次该等多久」，验不到「这个第几次
    /// 是从哪里数起的」。而后者才是会伤到用户的地方：计数若跨请求累加，
    /// 一轮里前几步偶发抖动就会把重试预算提前耗光，之后随便一次网关抖动
    /// 都会被当成「已经试到第 5 次」而直接失败。
    ///
    /// 因此这里起一个真的 TCP 服务，按脚本先拒后放，驱动真正的 ChatClient
    /// 重试循环，看它报出来的次数。不用 HttpListener：那个要预先注册 URL ACL，
    /// 没有管理员权限就跑不起来，而验证不该依赖这种前提。
    /// </summary>
    internal static class RetryResetTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            TestUnrelatedRequestDoesNotConsumeReply(report);
            TestCountResetsAfterSuccess(report);
            TestCountResetsBetweenRequests(report);
            TestBudgetNotConsumedAcrossRequests(report);
        }

        private static void TestUnrelatedRequestDoesNotConsumeReply(Action<string, bool, string> report)
        {
            using (var server = new ScriptedServer())
            {
                server.Enqueue(Reply.Stream("仍可使用"));
                using (var probe = new TcpClient("127.0.0.1", server.Port))
                using (var stream = probe.GetStream())
                {
                    stream.ReadTimeout = 3000;
                    var bytes = Encoding.ASCII.GetBytes("GET / HTTP/1.1\r\nHost: localhost\r\nConnection: close\r\n\r\n");
                    stream.Write(bytes, 0, bytes.Length);
                    var response = new byte[1024];
                    while (stream.Read(response, 0, response.Length) > 0) { }
                }
                report("本机端口探测不消耗重试响应或计数", server.RequestCount == 0, server.RequestCount.ToString());
            }
        }

        /// <summary>
        /// 一次请求内：先失败两次再成功。
        /// 报出的次数必须是 1、2，且成功后不再有重试通知。
        /// </summary>
        private static void TestCountResetsAfterSuccess(Action<string, bool, string> report)
        {
            using (var server = new ScriptedServer())
            {
                server.Enqueue(Reply.Unavailable());
                server.Enqueue(Reply.Unavailable());
                server.Enqueue(Reply.Stream("好了。"));

                var attempts = new List<int>();
                var text = Drive(server, attempts, out var error);

                report("先拒两次再放行能拿到正文", text == "好了。", error ?? text);
                report(
                    "重试次数从 1 数起且逐次递增",
                    attempts.Count == 2 && attempts[0] == 1 && attempts[1] == 2,
                    "实际 " + string.Join(",", attempts));
                report("成功后不再有重试通知", attempts.Count == 2, "实际 " + attempts.Count + " 次");
                report("服务端共收到 3 次请求", server.RequestCount == 3, server.RequestCount.ToString());
            }
        }

        /// <summary>
        /// 两次独立请求，共用同一个 ChatClient。
        ///
        /// 这是本文件的核心：第一次请求用掉两次重试并成功，第二次请求再失败时
        /// 必须重新从「第 1 次」数起。若报成「第 3 次」，说明计数跟着客户端跑，
        /// 而不是跟着这一次请求跑。
        /// </summary>
        private static void TestCountResetsBetweenRequests(Action<string, bool, string> report)
        {
            using (var server = new ScriptedServer())
            using (var client = new ChatClient())
            {
                // 第一次请求：拒两次后成功。
                server.Enqueue(Reply.Unavailable());
                server.Enqueue(Reply.Unavailable());
                server.Enqueue(Reply.Stream("第一次。"));

                var first = new List<int>();
                var firstText = DriveWith(client, server, first, out var firstError);

                report("第一次请求最终成功", firstText == "第一次。", firstError ?? firstText);
                report(
                    "第一次请求报出 1、2",
                    first.Count == 2 && first[0] == 1 && first[1] == 2,
                    "实际 " + string.Join(",", first));

                // 第二次请求：同一个客户端，再拒一次后成功。
                server.Enqueue(Reply.Unavailable());
                server.Enqueue(Reply.Stream("第二次。"));

                var second = new List<int>();
                var secondText = DriveWith(client, server, second, out var secondError);

                report("第二次请求最终成功", secondText == "第二次。", secondError ?? secondText);
                report(
                    "第二次请求重新从第 1 次数起",
                    second.Count == 1 && second[0] == 1,
                    "实际 " + string.Join(",", second));
            }
        }

        /// <summary>
        /// 上一次请求用掉的重试次数不占用下一次的预算。
        ///
        /// 第一次请求耗掉 4 次重试后成功；第二次请求若继承这个计数，
        /// 就只剩 1 次可用，连续拒 3 次必然失败。实际应当照样能成功。
        /// </summary>
        private static void TestBudgetNotConsumedAcrossRequests(Action<string, bool, string> report)
        {
            using (var server = new ScriptedServer())
            using (var client = new ChatClient())
            {
                for (var i = 0; i < 4; i++) { server.Enqueue(Reply.Unavailable()); }
                server.Enqueue(Reply.Stream("甲"));

                var first = new List<int>();
                DriveWith(client, server, first, out _);
                report("第一次请求用掉 4 次重试", first.Count == 4, "实际 " + first.Count);

                // 第二次请求连续被拒 3 次。预算若被继承（只剩 1 次）这里必失败。
                for (var i = 0; i < 3; i++) { server.Enqueue(Reply.Unavailable()); }
                server.Enqueue(Reply.Stream("乙"));

                var second = new List<int>();
                var text = DriveWith(client, server, second, out var error);

                report(
                    "上一次用掉的次数不占用这一次的预算",
                    text == "乙",
                    error ?? ("正文=" + text + "，重试 " + second.Count + " 次"));
                report(
                    "这一次自己数到 3",
                    second.Count == 3 && second[2] == 3,
                    "实际 " + string.Join(",", second));
            }
        }

        private static string Drive(ScriptedServer server, List<int> attempts, out string error)
        {
            using (var client = new ChatClient())
            {
                return DriveWith(client, server, attempts, out error);
            }
        }

        /// <summary>发一次请求，收集正文与重试通知里报出的次数。</summary>
        private static string DriveWith(
            ChatClient client,
            ScriptedServer server,
            List<int> attempts,
            out string error)
        {
            var request = new ChatRequest
            {
                Protocol = ProtocolKind.OpenAiChatCompletions,
                BaseUrl = server.BaseUrl,
                Token = "t",
                Model = "m",
                IncludeTools = false,
                MaxOutputTokens = 256,
            };

            request.Messages.Add(ChatMessage.FromUser("在吗"));

            var text = new StringBuilder();
            error = null;

            try
            {
                // 退避真实存在（首次 1 秒），因此这几条用例会慢几秒，
                // 但这是唯一能验到真实循环的办法。
                client.StreamAsync(
                    request,
                    chatEvent =>
                    {
                        if (chatEvent.Kind == ChatEventKind.TextDelta) { text.Append(chatEvent.Text); }
                        return Task.CompletedTask;
                    },
                    CancellationToken.None,
                    (attempt, delay, reason) =>
                    {
                        attempts.Add(attempt);
                        return Task.CompletedTask;
                    }).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                error = ex.GetType().Name + "：" + ex.Message;
            }

            return text.ToString();
        }

        /// <summary>一条预定的响应。</summary>
        private sealed class Reply
        {
            private Reply(int status, string body, bool sse)
            {
                Status = status;
                Body = body;
                Sse = sse;
            }

            internal int Status { get; }

            internal string Body { get; }

            internal bool Sse { get; }

            /// <summary>503：可重试故障。不带 Retry-After，走本地退避。</summary>
            internal static Reply Unavailable()
            {
                return new Reply(503, "{\"error\":{\"message\":\"upstream unavailable\"}}", false);
            }

            /// <summary>一段最小的 OpenAI 兼容 SSE 流。</summary>
            internal static Reply Stream(string content)
            {
                var builder = new StringBuilder();
                builder.Append("data: {\"choices\":[{\"delta\":{\"content\":")
                    .Append(Quote(content))
                    .Append("},\"index\":0}]}\n\n");
                builder.Append("data: {\"choices\":[{\"delta\":{},\"index\":0,\"finish_reason\":\"stop\"}]}\n\n");
                builder.Append("data: [DONE]\n\n");
                return new Reply(200, builder.ToString(), true);
            }

            private static string Quote(string value)
            {
                return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
            }
        }

        /// <summary>
        /// 按脚本逐条应答的最小 HTTP 服务。
        ///
        /// 自己解析请求行与头部，只为读完请求体好让连接干净关闭——
        /// 不读完的话客户端那侧可能先看到连接重置，把 503 换成 NETWORK_ERROR，
        /// 两者都可重试，但错误文本会变，断言就不稳。
        /// </summary>
        private sealed class ScriptedServer : IDisposable
        {
            private readonly TcpListener _listener;
            private readonly Queue<Reply> _replies = new Queue<Reply>();
            private readonly object _gate = new object();
            private readonly CancellationTokenSource _stopping = new CancellationTokenSource();
            private int _requestCount;

            internal ScriptedServer()
            {
                _listener = new TcpListener(IPAddress.Loopback, 0);
                _listener.Start();
                Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
                Task.Run(AcceptLoop);
            }

            internal int Port { get; }

            internal string BaseUrl => $"http://127.0.0.1:{Port}/v1";

            internal int RequestCount => Volatile.Read(ref _requestCount);

            internal void Enqueue(Reply reply)
            {
                lock (_gate) { _replies.Enqueue(reply); }
            }

            private async Task AcceptLoop()
            {
                while (!_stopping.IsCancellationRequested)
                {
                    TcpClient connection;
                    try
                    {
                        connection = await _listener.AcceptTcpClientAsync().ConfigureAwait(false);
                    }
                    catch
                    {
                        return;
                    }

                    // 逐条串行处理即可：被验的客户端一次只发一个请求。
                    try
                    {
                        Serve(connection);
                    }
                    catch
                    {
                        // 客户端提前断开属正常，不该让监听循环停掉。
                    }
                    finally
                    {
                        try { connection.Close(); } catch { }
                    }
                }
            }

            private void Serve(TcpClient connection)
            {
                using (var stream = connection.GetStream())
                {
                    var headers = ReadHeaders(stream);
                    if (headers == null) { return; }

                    if (!headers.StartsWith("POST /v1/chat/completions HTTP/", StringComparison.Ordinal))
                    {
                        Write(stream, Reply.Unavailable());
                        return;
                    }

                    DrainBody(stream, headers);
                    Interlocked.Increment(ref _requestCount);

                    Reply reply;
                    lock (_gate)
                    {
                        // 脚本用完就一律 503：多出来的请求说明重试没有按预期停下，
                        // 让它继续失败比静默成功更容易定位。
                        reply = _replies.Count > 0 ? _replies.Dequeue() : Reply.Unavailable();
                    }

                    Write(stream, reply);
                }
            }

            private static string ReadHeaders(Stream stream)
            {
                var buffer = new List<byte>();
                var one = new byte[1];

                while (true)
                {
                    var read = stream.Read(one, 0, 1);
                    if (read <= 0) { return null; }

                    buffer.Add(one[0]);
                    var count = buffer.Count;
                    if (count >= 4 &&
                        buffer[count - 4] == '\r' && buffer[count - 3] == '\n' &&
                        buffer[count - 2] == '\r' && buffer[count - 1] == '\n')
                    {
                        return Encoding.ASCII.GetString(buffer.ToArray());
                    }
                }
            }

            private static void DrainBody(Stream stream, string headers)
            {
                var length = 0;
                foreach (var line in headers.Split('\n'))
                {
                    var trimmed = line.Trim();
                    if (trimmed.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                    {
                        int.TryParse(trimmed.Substring("Content-Length:".Length).Trim(), out length);
                        break;
                    }
                }

                var remaining = length;
                var chunk = new byte[4096];
                while (remaining > 0)
                {
                    var read = stream.Read(chunk, 0, Math.Min(chunk.Length, remaining));
                    if (read <= 0) { break; }
                    remaining -= read;
                }
            }

            private static void Write(Stream stream, Reply reply)
            {
                var body = Encoding.UTF8.GetBytes(reply.Body);
                var contentType = reply.Sse ? "text/event-stream; charset=utf-8" : "application/json";

                var head = new StringBuilder();
                head.Append("HTTP/1.1 ").Append(reply.Status).Append(' ')
                    .Append(reply.Status == 200 ? "OK" : "Service Unavailable").Append("\r\n");
                head.Append("Content-Type: ").Append(contentType).Append("\r\n");
                head.Append("Content-Length: ").Append(body.Length).Append("\r\n");
                head.Append("Connection: close\r\n\r\n");

                var headBytes = Encoding.ASCII.GetBytes(head.ToString());
                stream.Write(headBytes, 0, headBytes.Length);
                stream.Write(body, 0, body.Length);
                stream.Flush();
            }

            public void Dispose()
            {
                _stopping.Cancel();
                try { _listener.Stop(); } catch { }
                _stopping.Dispose();
            }
        }
    }
}
