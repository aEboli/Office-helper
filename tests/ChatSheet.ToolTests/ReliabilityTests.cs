using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn;
using ChatSheet.AddIn.Bridge;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;
using Newtonsoft.Json.Linq;

namespace ChatSheet.ToolTests
{
    internal static class ReliabilityTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            Storage(report);
            foreach (var uri in new[] { "https://chatsheet.local/index.html", "https://chatsheet.local:443/index.html#settings" })
            { report("面板允许本地来源 " + uri, TaskPaneControl.IsTrustedPanelUri(uri), null); }
            foreach (var uri in new[] { "https://chatsheet.local.evil.test/", "https://chatsheet.local@evil.test/", "https://evil.test@chatsheet.local/", "http://chatsheet.local/", "https://chatsheet.local:444/", "file:///C:/test.html", "data:text/html,test", "about:blank" })
            { report("面板拒绝外部或伪装来源 " + uri, !TaskPaneControl.IsTrustedPanelUri(uri), null); }
            Framing(report);
            Task.Run(() => NetworkCancellationAsync(report)).GetAwaiter().GetResult();
            Task.Run(() => ChannelLifetimeAsync(report)).GetAwaiter().GetResult();
            Task.Run(() => ObjectRootRepairAsync(report)).GetAwaiter().GetResult();
        }

        private static void Storage(Action<string, bool, string> report)
        {
            var directory = Path.Combine(Path.GetTempPath(), "ChatSheet-storage-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, "settings.json");
            var secretKey = "audit-" + Guid.NewGuid().ToString("N");
            try
            {
                AtomicFile.WriteAllText(path, "{\"value\":\"原设置\"}");
                try { AtomicFile.Write(path, stream => { stream.WriteByte(123); throw new IOException("fixture interruption"); }); }
                catch (IOException) { }
                report("保存中断后原设置仍完整可读", JObject.Parse(File.ReadAllText(path)).Value<string>("value") == "原设置", null);
                var lockedFailure = false;
                using (var handle = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    try { AtomicFile.WriteAllText(path, "{\"value\":\"replacement\"}"); }
                    catch (IOException) { lockedFailure = true; }
                }
                report("替换被锁拒绝时不删除旧设置", lockedFailure && File.ReadAllText(path).Contains("原设置"), null);
                AtomicFile.WriteAllText(path, "{\"value\":\"新设置\"}");
                report("失败后可以重新保存且无临时残留", JObject.Parse(File.ReadAllText(path)).Value<string>("value") == "新设置" && Directory.GetFiles(directory, "*.tmp").Length == 0, null);
                Parallel.For(0, 8, i => AtomicFile.WriteAllText(path, new JObject { ["value"] = i }.ToString()));
                report("并发原子写入不会生成半截 JSON", JObject.Parse(File.ReadAllText(path))["value"]?.Type == JTokenType.Integer && Directory.GetFiles(directory, "*.tmp").Length == 0, null);
                Parallel.For(0, 12, i => FavoriteModels.Add("fixture", "model-" + i, directory));
                report("并发收藏不丢失模型", FavoriteModels.Load("fixture", directory).Count == 12, "实际 " + FavoriteModels.Load("fixture", directory).Count);
                SecretStore.Save(secretKey, "fixture-测试-secret");
                report("DPAPI 保存与读取保持兼容", SecretStore.Load(secretKey) == "fixture-测试-secret", null);
                SecretStore.Save(secretKey, "fixture-updated-secret");
                report("DPAPI 原子更新后读取新密钥", SecretStore.Load(secretKey) == "fixture-updated-secret", null);
            }
            finally { SecretStore.Delete(secretKey); Directory.Delete(directory, true); }
        }

        private static void Framing(Action<string, bool, string> report)
        {
            foreach (var ending in new[] { "", "\r", "\n", "\r\n" })
            {
                var frames = new List<SseFrame>();
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("\uFEFFevent: result\r\ndata: 首行\r\ndata: 末行" + ending)))
                {
                    SseReader.ReadAsync(stream, frame => { frames.Add(frame); return Task.FromResult(true); }, CancellationToken.None).GetAwaiter().GetResult();
                }
                report("SSE BOM 和末尾分隔符长度 " + ending.Length, frames.Count == 1 && frames[0].EventName == "result" && frames[0].Data == "首行\n末行", null);
            }
            using (var cts = new CancellationTokenSource())
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("data: abandoned\n")))
            {
                cts.Cancel();
                var delivered = false;
                try { SseReader.ReadAsync(stream, _ => { delivered = true; return Task.FromResult(true); }, cts.Token).GetAwaiter().GetResult(); }
                catch (OperationCanceledException) { report("已取消的 SSE 不交付尾帧", !delivered, null); return; }
                report("已取消的 SSE 应返回取消", false, null);
            }
        }

        private static async Task NetworkCancellationAsync(Action<string, bool, string> report)
        {
            var listener = new TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            using (var client = new ChatClient())
            using (var cancel = new CancellationTokenSource(TimeSpan.FromSeconds(5)))
            {
                TcpClient peer = null;
                try
                {
                    var endpoint = (IPEndPoint)listener.LocalEndpoint;
                    var pending = client.StreamAsync(new ChatRequest
                    {
                        Protocol = ProtocolKind.OpenAiChatCompletions,
                        BaseUrl = "http://127.0.0.1:" + endpoint.Port + "/v1",
                        Model = "fixture", Token = "fixture",
                    }, _ => Task.CompletedTask, cancel.Token, maxRetries: 0);
                    peer = await listener.AcceptTcpClientAsync();
                    var stream = peer.GetStream();
                    using (var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true))
                    {
                        while (!string.IsNullOrEmpty(await reader.ReadLineAsync())) { }
                    }
                    var headers = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Type: text/event-stream\r\nContent-Length: 99999\r\n\r\n");
                    await stream.WriteAsync(headers, 0, headers.Length);
                    await stream.FlushAsync();
                    await Task.Delay(150);
                    cancel.Cancel();
                    var finished = await Task.WhenAny(pending, Task.Delay(2000));
                    var cancelled = false;
                    if (finished == pending)
                    {
                        try { await pending; } catch (OperationCanceledException) { cancelled = true; }
                    }
                    report("响应头后服务器挂起仍能在两秒内停止", cancelled, null);
                    if (finished != pending) { peer.Close(); try { await pending; } catch { } }
                }
                finally { peer?.Close(); listener.Stop(); }
            }
        }

        private static async Task ChannelLifetimeAsync(Action<string, bool, string> report)
        {
            using (var channels = new AgentChannels(() => null, _ => Task.CompletedTask, _ => Task.CompletedTask,
                work => Task.FromResult(work()), () => new Settings { Mode = ConnectionMode.AuthorizedInternational }))
            using (var probeCancellation = new CancellationTokenSource())
            using (var bulk = new CancellationTokenSource())
            {
                var type = typeof(AgentChannels);
                var stateType = type.GetNestedType("WorkBuddyProbeState", BindingFlags.NonPublic);
                var state = Activator.CreateInstance(stateType, true);
                var pendingProbe = new TaskCompletionSource<WorkBuddyModelsResult>();
                stateType.GetField("Task", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(state, pendingProbe.Task);
                stateType.GetField("Cancellation", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(state, probeCancellation);
                var probes = (IDictionary)type.GetField("_workBuddyProbes", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(channels);
                probes[ConnectionMode.AuthorizedInternational] = state;
                var handlers = new Dictionary<string, Func<JObject, Task<object>>>();
                channels.Register(handlers);
                var first = handlers["chat.send"](new JObject { ["text"] = "fixture" });
                var busy = false;
                try { await handlers["chat.send"](new JObject { ["text"] = "duplicate" }); }
                catch (ProviderException ex) { busy = ex.Code == "BUSY"; }
                report("授权探测期间阻止重复对话", busy && !first.IsCompleted, null);
                await handlers["chat.stop"](new JObject());
                var stopped = await Task.WhenAny(first, Task.Delay(2000));
                report("停止立即结束授权准备阶段", stopped == first && JObject.FromObject(await first).Value<bool>("stopped"), null);
                var next = handlers["chat.send"](new JObject { ["text"] = "fixture-again" });
                type.GetField("_currentBulkProbe", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(channels, bulk);
                channels.Dispose();
                var disposed = await Task.WhenAny(next, Task.Delay(2000));
                report("关闭面板取消准备阶段与批量任务", disposed == next && bulk.IsCancellationRequested, null);
                if (disposed == next) { await next; }
            }
        }

        private static async Task ObjectRootRepairAsync(Action<string, bool, string> report)
        {
            var pushed = new List<object>();
            using (var channels = new AgentChannels(
                () => null,
                _ => Task.CompletedTask,
                message => { pushed.Add(message); return Task.CompletedTask; },
                work => Task.FromResult(work()),
                () => new Settings { Mode = ConnectionMode.AuthorizedInternational },
                createAgent: false))
            using (var cancellation = new CancellationTokenSource())
            {
                var handlers = new Dictionary<string, Func<JObject, Task<object>>>();
                channels.Register(handlers);

                var choose = channels.RequestObjectRootAsync("call-root", cancellation.Token);
                var request = pushed.Count == 1 ? JObject.FromObject(pushed[0]) : new JObject();
                report("缺失 root 时桥接请求与原工具调用 ID 关联",
                    request.Value<string>("kind") == "tool-root-request"
                        && request.Value<string>("id") == "call-root",
                    null);

                var invalid = JObject.FromObject(await handlers["tool.root.respond"](
                    new JObject { ["id"] = "call-root", ["root"] = "application" }));
                report("桥接拒绝不在白名单中的对象入口",
                    invalid.Value<bool>("accepted") == false && !choose.IsCompleted,
                    invalid.ToString(Newtonsoft.Json.Formatting.None));

                var valid = JObject.FromObject(await handlers["tool.root.respond"](
                    new JObject { ["id"] = "call-root", ["root"] = "worksheet" }));
                var chosenRoot = await choose;
                report("桥接接受合法对象入口并返回原等待调用",
                    valid.Value<bool>("accepted") && chosenRoot == "worksheet",
                    chosenRoot);

                var cancelChoice = channels.RequestObjectRootAsync("call-root-cancel", cancellation.Token);
                var cancelled = JObject.FromObject(await handlers["tool.root.respond"](
                    new JObject { ["id"] = "call-root-cancel", ["cancelled"] = true }));
                report("用户取消会收束原对象入口请求", cancelled.Value<bool>("accepted")
                    && await cancelChoice == null, null);

                using (var stopped = new CancellationTokenSource())
                {
                    var waiting = channels.RequestObjectRootAsync("call-root-stop", stopped.Token);
                    stopped.Cancel();
                    var cancellationObserved = false;
                    try { await waiting; }
                    catch (OperationCanceledException) { cancellationObserved = true; }
                    report("任务停止能收束待处理的对象入口选择", cancellationObserved, null);
                }

                var late = JObject.FromObject(await handlers["tool.root.respond"](
                    new JObject { ["id"] = "call-root-stop", ["root"] = "window" }));
                report("取消后的迟到选择被拒绝", late.Value<bool>("accepted") == false, null);
            }
        }
    }
}
