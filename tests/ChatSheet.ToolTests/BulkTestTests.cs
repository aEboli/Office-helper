using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ChatSheet.AddIn.Providers;
using ChatSheet.AddIn.Storage;

namespace ChatSheet.ToolTests
{
    /// <summary>
    /// 批量测试整份目录的并发行为。
    ///
    /// 这里盯的不是请求形态（那在 ProbeTests），而是「几十个模型一起跑」这件事本身
    /// 会出的错：
    ///
    ///   · 并发数发成 1 不会报错，只是慢十几倍——必须证明真的有多个在飞。
    ///   · 一个模型失败不该中断整批：它本身就是一条判定。
    ///   · 取消要能在半路生效，且已得结果保留——那些请求已经付过钱了。
    ///   · 整批只占一次单飞闸门：闸门若被放宽到 N，用户零散点的「试一下」也会并发，
    ///     那时同时有几个在飞谁也说不清。
    ///
    /// 打的是真实的 ModelProbe.ProbeManyAsync，不是它的复刻。曾经这里放的是一份
    /// 抄来的调度形状，那只能证明我抄的那份对——改了真实实现照样全绿。现在起一个
    /// 进程内的并发 HTTP 服务，按模型名分流应答，并记录同时在飞的峰值。
    ///
    /// 服务必须并发处理连接。RetryResetTests 里那个 ScriptedServer 是刻意串行的
    /// （它被验的客户端一次只发一个请求），直接拿来用会把并发度压成 1，
    /// 那时无论实现对不对，峰值都是 1，断言永远是绿的。
    /// </summary>
    internal static class BulkTestTests
    {
        internal static void Run(Action<string, bool, string> report)
        {
            TestUnrelatedRequestDoesNotAffectCounts(report);
            TestConcurrencyIsHonoured(report);
            TestSerialWhenAskedFor(report);
            TestOneFailureDoesNotStopBatch(report);
            TestCancellationKeepsResults(report);
            TestConcurrencyClamped(report);
            TestBatchHoldsTheSingleFlightGate(report);
            TestStartCallbackTracksInFlight(report);
            TestStopsAfterEnoughAvailable(report);
            TestTargetCountsOnlyAvailable(report);
            TestUnreachableTargetRunsEverything(report);
            TestTargetClampedAndGateReleased(report);
        }

        private static void TestUnrelatedRequestDoesNotAffectCounts(Action<string, bool, string> report)
        {
            using (var server = new ConcurrentServer(null, delayMs: 0))
            using (var probe = new TcpClient("127.0.0.1", server.Port))
            using (var stream = probe.GetStream())
            {
                stream.ReadTimeout = 3000;
                var bytes = Encoding.ASCII.GetBytes("GET / HTTP/1.1\r\nHost: localhost\r\nConnection: close\r\n\r\n");
                stream.Write(bytes, 0, bytes.Length);
                var response = new byte[1024];
                while (stream.Read(response, 0, response.Length) > 0) { }
                report("本机端口探测不计入模型并发或请求数", server.Served == 0 && server.Peak == 0,
                    "请求 " + server.Served + "，峰值 " + server.Peak);
            }
        }

        /// <summary>
        /// 探出目标个数的可用模型之后，不再发新请求。
        ///
        /// 判据是**服务端服务了几个连接**，不是回调次数：回调次数对不对说明不了有没有
        /// 省钱，而省钱是这个功能的全部理由。
        ///
        /// 20 个模型、并发 5、目标 1，实发条数必须是 5——不是 1。前五条在达标之前就
        /// 已经派发出去了，这是并发的固有代价，也是按钮上必须写出下界的原因。
        /// </summary>
        private static void TestStopsAfterEnoughAvailable(Action<string, bool, string> report)
        {
            var models = Enumerable.Range(0, 20).Select(i => "e" + i).ToList();

            // 第一个答得快、其余慢：达标那一刻真的有四条还在飞。
            // 全部同时返回的话，达标时那几条早就拿到判定了，于是「用取消做提前停会丢掉
            // 已付费的判定」这件事在结果里看不出来——那条断言会对着一个不受影响的
            // 场景通过。
            var run = RunSweep(
                models,
                concurrency: 5,
                stopAfterAvailable: 1,
                delayMs: 20,
                fastModel: models[0],
                slowDelayMs: 400);

            // 上界必须写死 5（= 并发数），不能写 6。
            //
            // 每个模型都答话，所以第一条结果必然置位 enough，而置位排在 slots.Release()
            // 之前；派发循环从 WaitAsync 醒来时读到的一定是新值，于是恰好停在 5。
            // 写成「5-6」会把「检查放在 WaitAsync 之前」那个变异容纳进去——实测那个
            // 变异正好给出 6，断言照样绿，等于这一条什么都没验。
            report(
                $"达标后不再派发（实发 {run.Served}/{models.Count} 条）",
                run.Served == 5,
                $"实发 {run.Served} 条，期望恰好 5（= 并发数）。等于 20 说明根本没停；" +
                    "等于 6 说明达标检查放在了 slots.WaitAsync 之前——那样会多派发一条，" +
                    "因为唤醒循环的那次 Release 来自已经把计数加上去的任务，" +
                    "而那一轮的检查早就过去了");

            report(
                "结局报「达标」",
                run.Outcome.TargetMet,
                "TargetMet 为假，面板就会把这次运行说成「整份目录都测完了」");

            report(
                $"已经发出去的都拿到了判定（{run.Results.Count} 条判定 / 服务端 {run.Served} 条请求）",
                run.Results.Count == run.Served,
                $"判定 {run.Results.Count} 条，服务端服务了 {run.Served} 条——差额就是" +
                    "花了钱没拿到答案的那几条。把达标改成 cts.Cancel() 会让这条变红：" +
                    "在飞那几条会抛 OCE 而不走 onResult。拿 outcome.Dispatched 做这条" +
                    "断言则会假绿，因为取消路径上它停在 0，0 == 0 照样通过");

            report(
                $"闸门在整批结束后放开（剩余许可 {run.GatePermitsAfter}）",
                run.GatePermitsAfter == 1,
                "达标直接 return 而不等 WhenAll 的话，闸门会在请求还在飞的时候放开，" +
                    "零散「试一下」当场与它们并发");

            report(
                "达标不算取消（判定一条没丢）",
                run.Results.Values.All(v => v == "Available"),
                string.Join(",", run.Results.Values.Distinct()));
        }

        /// <summary>
        /// 目标数只数 Available，不数 Unknown。
        ///
        /// 限流判「未确认」——请求付了钱、答案没拿到。把它算进目标等于在最坏结果上
        /// 宣布成功，而且会让运行在一个「谁也不知道能不能用」的模型上停下来。
        /// </summary>
        private static void TestTargetCountsOnlyAvailable(Action<string, bool, string> report)
        {
            // 只有最后一个会答话，其余全部限流。目标 1 因此必须一直跑到最后。
            var models = Enumerable.Range(0, 8).Select(i => "r" + i).ToList();
            var lucky = models[models.Count - 1];
            var run = RunSweep(
                models,
                concurrency: 2,
                stopAfterAvailable: 1,
                okOnly: lucky,
                othersRateLimited: true);

            report(
                $"限流不算达标，运行继续到底（实发 {run.Served}/{models.Count}）",
                run.Served == models.Count,
                $"实发 {run.Served} 条。少于 {models.Count} 说明把「未确认」当成了" +
                    "「找到一个能用的」——那是花了钱没拿到答案却宣布成功");

            report(
                "限流的都判「未确认」",
                models.Where(m => m != lucky).All(m =>
                    run.Results.TryGetValue(m, out var v) && v == "Unknown"),
                string.Join(",", run.Results.Select(kv => kv.Key + "=" + kv.Value)));

            report(
                "最后那个真能用的被找到了",
                run.Results.TryGetValue(lucky, out var luckyVerdict) &&
                    luckyVerdict == "Available" && run.Outcome.AvailableFound == 1,
                $"{lucky}={(run.Results.ContainsKey(lucky) ? run.Results[lucky] : "无判定")}，" +
                    $"AvailableFound={run.Outcome.AvailableFound}");
        }

        /// <summary>
        /// 目标达不到时把整份目录跑满，并且不谎报达标。
        ///
        /// 这是「用户选了省钱的选项却付了全额」那个场景在后端侧的锁。面板必须据此
        /// 明说，否则界面上看不出这次花的是全款。
        /// </summary>
        private static void TestUnreachableTargetRunsEverything(Action<string, bool, string> report)
        {
            var models = Enumerable.Range(0, 12).Select(i => "u" + i).ToList();
            var only = models[3];
            var run = RunSweep(models, concurrency: 4, stopAfterAvailable: 2, okOnly: only);

            report(
                $"目标达不到就跑满全程（实发 {run.Served}/{models.Count}）",
                run.Served == models.Count,
                $"实发 {run.Served} 条，期望 {models.Count}——目标数不该改变运行的范围");

            report(
                "不谎报达标",
                !run.Outcome.TargetMet && run.Outcome.AvailableFound == 1,
                $"TargetMet={run.Outcome.TargetMet}，AvailableFound=" +
                    $"{run.Outcome.AvailableFound}（目标 2，只有 1 个能用）");

            report(
                "每个候选都有判定",
                run.Results.Count == models.Count,
                $"{run.Results.Count}/{models.Count}");
        }

        /// <summary>目标数为 0 或负数时全部测完；闸门每条路都要放开。</summary>
        private static void TestTargetClampedAndGateReleased(Action<string, bool, string> report)
        {
            var models = Enumerable.Range(0, 6).Select(i => "z" + i).ToList();

            var zero = RunSweep(models, concurrency: 3, stopAfterAvailable: 0);
            report(
                $"目标 0 表示全部测完（实发 {zero.Served}/{models.Count}）",
                zero.Served == models.Count && !zero.Outcome.TargetMet,
                $"实发 {zero.Served}，TargetMet={zero.Outcome.TargetMet}");

            report(
                "目标 0 时 AvailableFound 仍如实计数（面板要用它，不能恒为 0）",
                zero.Outcome.AvailableFound == models.Count,
                $"AvailableFound={zero.Outcome.AvailableFound}，期望 {models.Count}。" +
                    "恒为 0 会让收尾那条推送把中途已经涨上去的数字打回去");

            var negative = RunSweep(models, concurrency: 3, stopAfterAvailable: -2);
            report(
                "目标给负数时全部测完，不提前停",
                negative.Served == models.Count && !negative.Outcome.TargetMet,
                $"实发 {negative.Served}，TargetMet={negative.Outcome.TargetMet}");

            report(
                $"闸门放开（0：{zero.GatePermitsAfter}，负数：{negative.GatePermitsAfter}）",
                zero.GatePermitsAfter == 1 && negative.GatePermitsAfter == 1,
                "闸门漏放会让之后所有探测永久排队");
        }

        /// <summary>
        /// onStart 在「拿到并发槽位之后」回调，而不是在入队时。
        ///
        /// 这条守的是面板上一个看得见的行为：批量测试时正在探的那几行有扫光。
        /// 面板按 onStart/onResult 两端维护「在飞集合」，所以同时在飞的个数
        /// 必须等于并发数——而不是模型总数。
        ///
        /// 若 onStart 提前到入队时回调，12 个模型会在一瞬间全部标成「正在探」，
        /// 列表里每一行都在扫，标记于是什么也没说。那种写法不会报错、不会变慢，
        /// 只是把标记变成噪音，因此只有这条断言拦得住它。
        ///
        /// 判据是「开始但未结束」的峰值，不是回调次数：次数对不对说明不了时机。
        /// </summary>
        private static void TestStartCallbackTracksInFlight(Action<string, bool, string> report)
        {
            var models = Enumerable.Range(0, 12).Select(i => "f" + i).ToList();
            var started = new List<string>();
            var finished = new HashSet<string>();
            var sync = new object();
            var inFlight = 0;
            var peakInFlight = 0;
            var startedBeforeFinish = true;

            using (var server = new ConcurrentServer(null, 45))
            {
                ModelProbe.ProbeManyAsync(
                    Connection(server.BaseUrl),
                    models,
                    concurrency: 5,
                    _ => (OutputLimitField?)null,
                    (model, verdict, done) =>
                    {
                        lock (sync)
                        {
                            // 每个探完的模型必须先被 onStart 报过。反了就说明面板
                            // 会先收到 settled 再收到 starting，那一行会一直挂着扫光。
                            if (!started.Contains(model)) { startedBeforeFinish = false; }
                            finished.Add(model);
                            inFlight--;
                        }

                        return Task.CompletedTask;
                    },
                    CancellationToken.None,
                    model =>
                    {
                        lock (sync)
                        {
                            started.Add(model);
                            inFlight++;
                            if (inFlight > peakInFlight) { peakInFlight = inFlight; }
                        }

                        return Task.CompletedTask;
                    }).GetAwaiter().GetResult();
            }

            report(
                $"每个模型开始探时都回调了 onStart（{started.Count}/{models.Count}）",
                started.Count == models.Count &&
                    models.All(m => started.Contains(m)),
                $"报了 {started.Count} 个：{string.Join(",", started)}");

            report(
                $"同时在飞的峰值不超过并发数（实测 {peakInFlight}）",
                peakInFlight >= 4 && peakInFlight <= 5,
                $"峰值 {peakInFlight}，期望 4-5。超过 5 说明 onStart 在模型真的开始探" +
                    "之前就回调了（例如提到等槽位之前），那时面板会把还在排队的行" +
                    "也标成正在探——标记于是说了假话");

            report(
                "onStart 早于同一个模型的 onResult",
                startedBeforeFinish,
                "先收到判定再收到「开始探」的话，那一行会一直挂着扫光");

            report(
                "每个开始探的最终都有判定（没有漏在飞里的）",
                finished.Count == models.Count,
                $"开始 {started.Count} 个，落定 {finished.Count} 个——差额就是面板上" +
                    "永远扫下去的那几行");

            report(
                "不传 onStart 也能跑（这个参数是可选的）",
                RunReal(new List<string> { "n1", "n2" }, 2, out _).Count == 2,
                "onStart 为 null 时抛异常，会让没升级的调用方直接崩");
        }

        /// <summary>
        /// 并发度确实达到给定值。
        ///
        /// 判据是「同时在飞的峰值」，不是耗时：耗时随机器负载漂，而峰值是确定的。
        /// 每个应答停一小会儿，好让并发真的叠起来。
        /// </summary>
        private static void TestConcurrencyIsHonoured(Action<string, bool, string> report)
        {
            var models = Enumerable.Range(0, 12).Select(i => "m" + i).ToList();
            int peak;
            var result = RunReal(models, concurrency: 5, peak: out peak);

            report(
                $"并发达到 5（实测峰值 {peak}）",
                peak >= 4 && peak <= 5,
                $"峰值 {peak}，期望 4-5（12 个模型、并发 5）");

            report(
                "并发不超过给定值",
                peak <= 5,
                $"峰值 {peak} 超过 5，闸门没生效");

            report(
                "每个模型都测到了",
                result.Count == models.Count,
                $"测了 {result.Count}/{models.Count}");

            report(
                "都判成可用（服务正常应答）",
                result.Values.All(v => v == "Available"),
                string.Join(",", result.Values.Distinct()));
        }

        /// <summary>
        /// 并发给 1 时必须真的串行。
        ///
        /// 这条是上一条的反例：没有它，「峰值 ≤ 5」在实现根本没并发时也成立。
        /// </summary>
        private static void TestSerialWhenAskedFor(Action<string, bool, string> report)
        {
            var models = Enumerable.Range(0, 6).Select(i => "s" + i).ToList();
            int peak;
            var result = RunReal(models, concurrency: 1, peak: out peak);

            report(
                $"并发给 1 时真的串行（峰值 {peak}）",
                peak == 1,
                $"峰值 {peak}，期望 1——说明并发数没被当回事");

            report(
                "串行也把每个都测到",
                result.Count == models.Count,
                $"测了 {result.Count}/{models.Count}");
        }

        /// <summary>单个失败不中断整批：它本身就是一条判定。</summary>
        private static void TestOneFailureDoesNotStopBatch(Action<string, bool, string> report)
        {
            var models = new List<string> { "ok1", "mock-absent", "ok2", "ok3" };
            int peak;
            var result = RunReal(models, concurrency: 2, peak: out peak, failOn: "mock-absent");

            report(
                "一个失败不中断整批",
                result.Count == models.Count,
                $"只测了 {result.Count}/{models.Count}");

            report(
                "失败的那个落成「不可用」（服务端点名了模型）",
                result.TryGetValue("mock-absent", out var verdict) && verdict == "Unavailable",
                $"实际 {(result.ContainsKey("mock-absent") ? result["mock-absent"] : "无判定")}");

            report(
                "其余仍判可用",
                models.Where(m => m != "mock-absent").All(m =>
                    result.TryGetValue(m, out var v) && v == "Available"),
                string.Join(",", result.Select(kv => kv.Key + "=" + kv.Value)));
        }

        /// <summary>取消在半路生效，已得结果保留。</summary>
        private static void TestCancellationKeepsResults(Action<string, bool, string> report)
        {
            var models = Enumerable.Range(0, 24).Select(i => "c" + i).ToList();
            int peak;
            var result = RunReal(
                models,
                concurrency: 3,
                peak: out peak,
                cancelAfter: done => done >= 4);

            report(
                $"取消在半路生效（测了 {result.Count}/{models.Count}）",
                result.Count < models.Count,
                $"取消后仍把 {result.Count} 个跑完了");

            report(
                "取消前已得的判定保留（那些请求已经付过钱）",
                result.Count >= 4,
                $"只剩 {result.Count} 条，取消把已付过钱的结果一起丢了");
        }

        /// <summary>并发数越界时收拢到合理区间，不接受 0 或负数。</summary>
        private static void TestConcurrencyClamped(Action<string, bool, string> report)
        {
            var models = new List<string> { "a", "b", "c" };

            int zeroPeak;
            var zero = RunReal(models, concurrency: 0, peak: out zeroPeak);
            report(
                $"并发给 0 时仍能跑完（收拢为至少 1，峰值 {zeroPeak}）",
                zero.Count == models.Count && zeroPeak >= 1,
                $"测了 {zero.Count}/{models.Count}，峰值 {zeroPeak}");

            int negPeak;
            var negative = RunReal(models, concurrency: -3, peak: out negPeak);
            report(
                "并发给负数时仍能跑完",
                negative.Count == models.Count,
                $"测了 {negative.Count}/{models.Count}");
        }

        /// <summary>
        /// 整批占着单飞闸门期间，零散的「试一下」必须排队等着，不能加入并发。
        ///
        /// 这一条是分层设计的全部理由：闸门若被放宽到 N，用户随手点的确认也会并发，
        /// 那时同时有几个在飞就说不清了。判据是峰值——若单个探测挤进了批内，
        /// 峰值会变成 4。
        /// </summary>
        private static void TestBatchHoldsTheSingleFlightGate(Action<string, bool, string> report)
        {
            var models = Enumerable.Range(0, 9).Select(i => "g" + i).ToList();

            using (var server = new ConcurrentServer(null, delayMs: 45))
            {
                var connection = Connection(server.BaseUrl);
                var single = new List<AvailabilityVerdict>();

                // 批跑起来之后再插一个零散探测。
                var batch = ModelProbe.ProbeManyAsync(
                    connection,
                    models,
                    3,
                    _ => (OutputLimitField?)null,
                    (m, v, done) => Task.CompletedTask,
                    CancellationToken.None);

                Thread.Sleep(60);

                // 先确认批真的占着闸门：不占着的话下面那条「不重叠」会因为
                // 根本没有竞争而白通过。
                var permitsDuringBatch = ModelProbe.GatePermits;

                var probe = ModelProbe.ProbeAsync(
                    connection, "lonely", null, CancellationToken.None)
                    .ContinueWith(t =>
                    {
                        if (t.Status == TaskStatus.RanToCompletion) { single.Add(t.Result); }
                    });

                Task.WaitAll(batch, probe);

                // 判据是「区间不重叠」，不是全局峰值。峰值会被一个竞态骗到：
                // 服务端在写完响应之后才减在飞计数，而客户端此刻已认为请求结束，
                // 于是排在后面的探测会与「正在收尾」的批内连接重叠，峰值虚高一格。
                // 区间比较不受这个影响——它问的是「这个探测有没有等批跑完」。
                report(
                    $"批跑期间闸门被占着（剩余许可 {permitsDuringBatch}）",
                    permitsDuringBatch == 0,
                    "闸门没被占着，下面那条「不重叠」会因为没有竞争而白通过");

                var spans = server.Spans;
                long[] lonely;
                var haveLonely = spans.TryGetValue("lonely", out lonely);
                var batchEnd = spans
                    .Where(kv => kv.Key != "lonely")
                    .Select(kv => kv.Value[1])
                    .DefaultIfEmpty(0)
                    .Max();
                var overlapping = spans
                    .Where(kv => kv.Key != "lonely" && haveLonely)
                    .Count(kv => kv.Value[0] < lonely[1] && lonely[0] < kv.Value[1]);

                report(
                    "批跑期间零散探测排队等着，不与批内请求重叠" +
                        (haveLonely ? $"（重叠 {overlapping} 个）" : "（没拿到它的区间）"),
                    haveLonely && overlapping == 0,
                    haveLonely
                        ? $"零散探测 {lonely[0]}-{lonely[1]}ms，批内最晚结束于 {batchEnd}ms，" +
                            $"重叠 {overlapping} 个——说明它挤进了批内"
                        : "服务端没记到 lonely 的区间");

                report(
                    "它在批跑结束之后才开始（是排队，不是被丢掉）",
                    haveLonely && lonely[0] >= batchEnd - 5,
                    haveLonely
                        ? $"它 {lonely[0]}ms 开始，而批内最晚结束于 {batchEnd}ms"
                        : "没拿到区间");

                report(
                    "零散探测最终也跑完了",
                    single.Count == 1,
                    $"拿到 {single.Count} 条判定");

                report(
                    $"批内并发仍达到上限（峰值 {server.Peak}）",
                    server.Peak >= 3,
                    $"峰值 {server.Peak}，期望至少 3——否则这条断言没测到并发");
            }
        }

        private static ResolvedConnection Connection(string baseUrl)
        {
            return new ResolvedConnection
            {
                Protocol = ProtocolKind.OpenAiChatCompletions,
                BaseUrl = baseUrl,
                Token = "t",
            };
        }

        /// <summary>一次带目标数的运行的全部观测值。</summary>
        private sealed class SweepRun
        {
            internal Dictionary<string, string> Results { get; set; }

            internal ProbeSweepOutcome Outcome { get; set; }

            internal int Peak { get; set; }

            /// <summary>服务端真的服务了几个请求。这就是钱。</summary>
            internal int Served { get; set; }

            internal int GatePermitsAfter { get; set; }
        }

        /// <summary>
        /// 带目标数跑一批，并把结局与服务端计数一起带回来。
        ///
        /// 刻意不改 RunReal 的签名：out 参数在 C# 里不能有默认值，加进去就是把现有
        /// 七处调用点一起打断；而把返回值换成元组同样打断，那些调用点都在结果上
        /// 直接取 .Count / .Values。
        /// </summary>
        private static SweepRun RunSweep(
            IReadOnlyList<string> models,
            int concurrency,
            int stopAfterAvailable,
            string okOnly = null,
            bool othersRateLimited = false,
            int delayMs = 45,
            Func<int, bool> cancelAfter = null,
            string fastModel = null,
            int slowDelayMs = 0)
        {
            var results = new Dictionary<string, string>();
            var sync = new object();
            var run = new SweepRun();

            using (var server = new ConcurrentServer(
                null, delayMs, okOnly, othersRateLimited, fastModel, slowDelayMs))
            using (var cts = new CancellationTokenSource())
            {
                var completed = 0;

                try
                {
                    run.Outcome = ModelProbe.ProbeManyAsync(
                        Connection(server.BaseUrl),
                        models,
                        concurrency,
                        _ => (OutputLimitField?)null,
                        (model, verdict, done) =>
                        {
                            lock (sync) { results[model] = verdict.ToString(); }
                            var n = Interlocked.Increment(ref completed);
                            if (cancelAfter != null && cancelAfter(n)) { cts.Cancel(); }
                            return Task.CompletedTask;
                        },
                        cts.Token,
                        onStart: null,
                        stopAfterAvailable: stopAfterAvailable).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException)
                {
                    // 取消是预期路径之一。
                }
                catch (AggregateException ex)
                    when (ex.InnerExceptions.Any(e => e is OperationCanceledException))
                {
                    // 同上。
                }

                run.Peak = server.Peak;
                run.Served = server.Served;
                run.GatePermitsAfter = ModelProbe.GatePermits;
            }

            run.Results = results;
            return run;
        }

        /// <summary>起一个进程内的并发 HTTP 服务，跑真实的 ProbeManyAsync。</summary>
        private static Dictionary<string, string> RunReal(
            IReadOnlyList<string> models,
            int concurrency,
            out int peak,
            string failOn = null,
            int delayMs = 45,
            Func<int, bool> cancelAfter = null)
        {
            var results = new Dictionary<string, string>();
            var sync = new object();

            using (var server = new ConcurrentServer(failOn, delayMs))
            using (var cts = new CancellationTokenSource())
            {
                var completed = 0;

                try
                {
                    ModelProbe.ProbeManyAsync(
                        Connection(server.BaseUrl),
                        models,
                        concurrency,
                        _ => (OutputLimitField?)null,
                        (model, verdict, done) =>
                        {
                            lock (sync) { results[model] = verdict.ToString(); }
                            var n = Interlocked.Increment(ref completed);
                            if (cancelAfter != null && cancelAfter(n)) { cts.Cancel(); }
                            return Task.CompletedTask;
                        },
                        cts.Token).GetAwaiter().GetResult();
                }
                catch (OperationCanceledException)
                {
                    // 取消是预期路径之一。
                }
                catch (AggregateException ex)
                    when (ex.InnerExceptions.Any(e => e is OperationCanceledException))
                {
                    // 同上。
                }

                peak = server.Peak;
            }

            return results;
        }

        /// <summary>
        /// 并发应答的最小 HTTP 服务，记录同时在飞的峰值。
        ///
        /// 与 RetryResetTests 里那个刻意串行的服务不同：这里必须并发，
        /// 否则并发度断言测的是服务而不是被测代码。
        /// </summary>
        private sealed class ConcurrentServer : IDisposable
        {
            private readonly TcpListener _listener;
            private readonly CancellationTokenSource _stopping = new CancellationTokenSource();
            private readonly string _failOn;
            private readonly int _delayMs;
            private readonly string _okOnly;
            private readonly bool _othersRateLimited;
            private readonly string _fastModel;
            private readonly int _slowDelayMs;
            private readonly object _gate = new object();
            private readonly Dictionary<string, long[]> _spans =
                new Dictionary<string, long[]>();
            private readonly System.Diagnostics.Stopwatch _clock =
                System.Diagnostics.Stopwatch.StartNew();
            private int _inFlight;
            private int _peak;
            private int _served;

            internal ConcurrentServer(
                string failOn,
                int delayMs,
                string okOnly = null,
                bool othersRateLimited = false,
                string fastModel = null,
                int slowDelayMs = 0)
            {
                _failOn = failOn;
                _delayMs = delayMs;
                _okOnly = okOnly;
                _othersRateLimited = othersRateLimited;
                _fastModel = fastModel;
                _slowDelayMs = slowDelayMs;
                _listener = new TcpListener(IPAddress.Loopback, 0);
                _listener.Start();
                Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
                Task.Run(AcceptLoop);
            }

            internal int Port { get; }

            internal string BaseUrl { get { return "http://127.0.0.1:" + Port + "/v1"; } }

            internal int Peak { get { lock (_gate) { return _peak; } } }

            /// <summary>
            /// 一共服务了几个请求。
            ///
            /// 这就是钱：判「真的少发了请求」只能看它，不能拿 Spans.Count 代替——
            /// span 是在 Thread.Sleep 之后才写进去的，达标停那一刻还在飞的几条
            /// 可能一条都没写，于是读到的数偏小，而偏小的方向恰好让断言更容易假绿。
            /// 这个计数在模型请求被确认后就加，早于任何延时与应答。
            /// </summary>
            internal int Served { get { lock (_gate) { return _served; } } }

            /// <summary>每个模型被服务的起止时刻（毫秒）。用来判断有没有真的排队。</summary>
            internal Dictionary<string, long[]> Spans
            {
                get { lock (_gate) { return new Dictionary<string, long[]>(_spans); } }
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

                    // 每个连接一个任务：并发是这组断言的前提。
                    var captured = connection;
                    var ignored = Task.Run(() =>
                    {
                        try { Serve(captured); }
                        catch { }
                        finally { try { captured.Close(); } catch { } }
                    });
                }
            }

            private void Serve(TcpClient connection)
            {
                var counted = false;
                try
                {
                    using (var stream = connection.GetStream())
                    {
                        var body = ReadRequest(stream);
                        if (body == null) { Write(stream, 404, "{}", false); return; }
                        var model = ExtractModel(body);
                        lock (_gate)
                        {
                            _inFlight++;
                            _served++;
                            counted = true;
                            if (_inFlight > _peak) { _peak = _inFlight; }
                        }
                        var began = _clock.ElapsedMilliseconds;

                        // 停一会儿，好让并发叠起来。不停的话每个请求瞬间结束，
                        // 峰值永远是 1，并发度断言就测不到东西。
                        //
                        // fastModel 让其中一个先答完而其余仍在飞：达标那一刻真的有几条
                        // 没结束，「用取消做提前停会丢掉已付费的判定」才观察得到。
                        // 一起返回时取消落得太晚，那几条已经拿到判定了。
                        Thread.Sleep(
                            _fastModel != null && model != _fastModel && _slowDelayMs > 0
                                ? _slowDelayMs
                                : _delayMs);

                        lock (_gate)
                        {
                            _spans[model] = new[] { began, _clock.ElapsedMilliseconds };
                        }

                        var named404 = (_failOn != null && model == _failOn) ||
                            (_okOnly != null && !_othersRateLimited && model != _okOnly);

                        if (named404)
                        {
                            // 点名模型的 404：这是唯一该判「不可用」的形状。
                            Write(
                                stream,
                                404,
                                "{\"error\":{\"message\":\"The model `" + model +
                                    "` does not exist\",\"code\":\"model_not_found\"}}",
                                false);
                            return;
                        }

                        // 限流：说的是账号而不是模型，判「未确认」。用来验目标数只数
                        // Available——把未确认算进去等于在「花了钱没拿到答案」上宣布成功。
                        if (_othersRateLimited && _okOnly != null && model != _okOnly)
                        {
                            Write(
                                stream,
                                429,
                                "{\"error\":{\"message\":\"Rate limit reached for this account\"," +
                                    "\"code\":\"rate_limit_exceeded\"}}",
                                false);
                            return;
                        }

                        var sse = new StringBuilder();
                        sse.Append("data: {\"choices\":[{\"delta\":{\"content\":\"ok\"},\"index\":0}]}\n\n");
                        sse.Append("data: {\"choices\":[{\"delta\":{},\"index\":0,\"finish_reason\":\"stop\"}]}\n\n");
                        sse.Append("data: [DONE]\n\n");
                        Write(stream, 200, sse.ToString(), true);
                    }
                }
                finally
                {
                    if (counted) { lock (_gate) { _inFlight--; } }
                }
            }

            /// <summary>读完请求头与请求体。不读完会让客户端那侧先看到连接重置。</summary>
            private static string ReadRequest(NetworkStream stream)
            {
                var head = new StringBuilder();
                var one = new byte[1];
                var length = 0;

                while (true)
                {
                    if (stream.Read(one, 0, 1) <= 0) { break; }
                    head.Append((char)one[0]);
                    if (head.Length >= 4 &&
                        head[head.Length - 1] == '\n' && head[head.Length - 2] == '\r' &&
                        head[head.Length - 3] == '\n' && head[head.Length - 4] == '\r')
                    {
                        break;
                    }
                }

                if (!head.ToString().StartsWith("POST /v1/chat/completions HTTP/", StringComparison.Ordinal)) { return null; }
                foreach (var line in head.ToString().Split('\n'))
                {
                    var trimmed = line.Trim();
                    if (trimmed.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                    {
                        int.TryParse(trimmed.Substring("Content-Length:".Length).Trim(), out length);
                    }
                }

                if (length <= 0) { return string.Empty; }

                var buffer = new byte[length];
                var read = 0;
                while (read < length)
                {
                    var got = stream.Read(buffer, read, length - read);
                    if (got <= 0) { break; }
                    read += got;
                }

                return Encoding.UTF8.GetString(buffer, 0, read);
            }

            private static string ExtractModel(string body)
            {
                var match = Regex.Match(body ?? string.Empty, "\"model\"\\s*:\\s*\"([^\"]*)\"");
                return match.Success ? match.Groups[1].Value : string.Empty;
            }

            private static void Write(NetworkStream stream, int status, string body, bool sse)
            {
                var bytes = Encoding.UTF8.GetBytes(body);
                var header = new StringBuilder();
                header.Append("HTTP/1.1 ").Append(status).Append(" X\r\n");
                header.Append("Content-Type: ")
                    .Append(sse ? "text/event-stream" : "application/json")
                    .Append("\r\n");
                header.Append("Content-Length: ").Append(bytes.Length).Append("\r\n");
                header.Append("Connection: close\r\n\r\n");

                var headerBytes = Encoding.UTF8.GetBytes(header.ToString());
                stream.Write(headerBytes, 0, headerBytes.Length);
                stream.Write(bytes, 0, bytes.Length);
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
