using System;
using System.Collections.Generic;
using System.IO;
using ChatSheet.AddIn.Providers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace ChatSheet.AddIn.Storage
{
    /// <summary>接入模式。</summary>
    internal enum ConnectionMode
    {
        /// <summary>读取本机 CLI 配置中的接口地址与令牌。</summary>
        LocalCli = 0,

        /// <summary>自定义接口地址、密钥与模型。</summary>
        CustomApi = 1,

        /// <summary>使用 WorkBuddy 当前授权登录。</summary>
        Authorized = 2,

        /// <summary>使用 WorkBuddy 国际版授权登录。</summary>
        AuthorizedInternational = 3,
    }

    internal static class ConnectionModeExtensions
    {
        internal static bool IsWorkBuddy(this ConnectionMode mode)
        {
            return mode == ConnectionMode.Authorized || mode == ConnectionMode.AuthorizedInternational;
        }

        internal static bool IsDomesticWorkBuddy(this ConnectionMode mode)
        {
            return mode == ConnectionMode.Authorized;
        }
    }

    /// <summary>审批策略。</summary>
    internal enum ApprovalPolicy
    {
        /// <summary>写操作逐项审批，读操作自动执行。</summary>
        PerWrite = 0,

        /// <summary>本轮每张工作表/文档的同类操作首次确认；授权只在本轮有效。</summary>
        PerTurn = 1,

        /// <summary>全自动执行，依赖撤销兜底。</summary>
        Automatic = 2,
    }

    /// <summary>
    /// 用户设置。密钥不存在这里，只存于 DPAPI 加密的 SecretStore，
    /// 本文件是明文 JSON，任何敏感值都不得写入。
    /// </summary>
    internal sealed class Settings
    {
        internal const string CustomApiSecretKey = "custom-api-token";

        internal const string ProxySecretKey = "proxy-password";

        private const string LegacyProxyProfileId = "legacy";

        internal ConnectionMode Mode { get; set; } = ConnectionMode.LocalCli;

        internal CliKind CliSource { get; set; } = CliKind.Auto;

        internal ProtocolKind CustomProtocol { get; set; } = Protocols.Default;

        internal string CustomBaseUrl { get; set; } = string.Empty;

        internal ProxyKind ProxyType { get; set; } = ProxyKind.Direct;

        internal string ProxyHost { get; set; } = string.Empty;

        internal int ProxyPort { get; set; }

        internal string ProxyUsername { get; set; } = string.Empty;

        /// <summary>可切换的代理配置；密码按配置单独保存在 DPAPI 密钥槽。</summary>
        internal List<ProxyProfile> ProxyProfiles { get; set; } = new List<ProxyProfile>();

        internal string ActiveProxyProfileId { get; set; } = string.Empty;

        /// <summary>当前代理失败时是否按列表顺序尝试其他配置。</summary>
        internal bool AutoSwitchProxy { get; set; } = true;

        internal string Model { get; set; } = string.Empty;

        /// <summary>
        /// <see cref="Model"/> 是为哪个接入连接选的（见 <see cref="ConnectionKey"/>）。
        ///
        /// 必须持久化：模型只对选它的那个连接有意义。没有这个标记时，
        /// 从「自定义接口」切回「本机 CLI 配置」后，旧模型会继续留在 Model 里，
        /// 而 <see cref="ResolveConnection"/> 又优先用 Model，于是 CLI 配置自带的模型
        /// 被一个它根本没有的模型名顶掉，界面上还显示为「已就绪」。
        ///
        /// 空串表示来历不明（旧版设置文件），由 <see cref="AdoptOrDropUnstampedModel"/> 决定去留。
        /// </summary>
        internal string ModelConnection { get; set; } = string.Empty;

        /// <summary>
        /// 默认 High：多数模型自身的默认档也是 high，
        /// 且表格任务常涉及多步推理，档位过低会让模型跳过必要的确认。
        /// </summary>
        internal ThinkingLevel Thinking { get; set; } = ThinkingLevel.High;

        // 默认直接执行表格修改；用户仍可在设置页主动切回逐项或每轮审批。
        internal ApprovalPolicy Approval { get; set; } = ApprovalPolicy.Automatic;

        internal double? Temperature { get; set; }

        internal int MaxOutputTokens { get; set; } = 8192;

        /// <summary>
        /// 上下文 token 预算。超出后压缩较早的轮次。
        ///
        /// 取 200000 而非窗口全额 272000：估算器不计入图片（只累加文本与工具调用），
        /// 输出 token 在多数服务商与输入共享同一窗口，且启发式估算在纯数字上可能偏低。
        /// 留出的余量正是给这三项。
        /// </summary>
        internal int ContextBudgetTokens { get; set; } = 200_000;

        /// <summary>单轮工具步数；0 表示持续执行，仍可随时停止。</summary>
        internal int MaxSteps { get; set; } = 0;

        internal bool AutoIncludeSelection { get; set; } = true;

        /// <summary>
        /// 工具形态。默认自动探测：先按原生函数声明发，被拒或模型推辞后降级。
        ///
        /// 之所以留出手动档，是因为探测只能对「服务端报错」和「模型明说做不到」
        /// 起作用。有的网关既不报错也不调用工具，静默把声明丢掉——那种情况下
        /// 用户比探测更早知道真相，需要一个直接指定的地方。
        /// </summary>
        internal ToolProtocolPreference ToolProtocol { get; set; } = ToolProtocolPreference.Auto;

        /// <summary>
        /// 视觉中转模型。主模型没有视觉能力时，先用它把图片转成文字。
        ///
        /// 空串表示不启用，此时看不了图的模型会去掉图片继续这一轮。
        /// 沿用当前连接与密钥，只替换模型名。
        /// </summary>
        internal string VisionRelayModel { get; set; } = string.Empty;

        /// <summary>
        /// 当前接入连接的稳定标识。
        ///
        /// 只包含会改变「有哪些模型可用」的字段：自定义接口看协议、地址和代理，
        /// 本机 CLI 看用的是哪个 CLI 与代理。密钥绝不进入该键，它会随设置一起明文落盘。
        /// </summary>
        internal string ConnectionKey()
        {
            if (Mode != ConnectionMode.CustomApi)
            {
                return Mode + "|" + CliSource + "|" + ProxyKey();
            }

            // 地址先规范化，避免尾斜杠、缺少 /v1 这类等价写法被判成换了连接。
            var address = (CustomBaseUrl ?? string.Empty).Trim();
            try
            {
                address = Protocols.NormalizeBaseUrl(address, CustomProtocol);
            }
            catch (ProviderException)
            {
                // 地址还没填完或填错时用原样文本，此时本就不该复用别处的模型。
            }

            return Mode + "|" + Protocols.Get(CustomProtocol).Id + "|" + address + "|" + ProxyKey();
        }

        private string ProxyKey()
        {
            var profile = FindActiveProxyProfile();
            var kind = profile?.Kind ?? ProxyType;
            var host = profile?.Host ?? ProxyHost;
            var port = profile?.Port ?? ProxyPort;
            var username = profile?.Username ?? ProxyUsername;
            if (kind == ProxyKind.Direct) { return "direct"; }
            if (kind == ProxyKind.System) { return "system"; }
            return $"{kind}|{(host ?? string.Empty).Trim().ToLowerInvariant()}|{port}|{(username ?? string.Empty).Trim()}";
        }

        /// <summary>
        /// 常用模型名单的归属键。
        ///
        /// 与 ConnectionKey 只差一处：本机 CLI 按**解析后**的 CliKind 归组，
        /// 而不是配置里的 CliSource。必须如此——用户刚弄清自己在用哪个 CLI
        /// 就去把下拉从「自动」钉成「Claude」时，LocalCliConfig.Resolve 返回的凭据
        /// 一模一样（Auto 的第一候选就是 Claude），而 ConnectionKey 会从
        /// LocalCli|Auto 变成 LocalCli|Claude，于是攒起来的名单原地失联，
        /// 旧分组还留在盘上够不着。
        ///
        /// 解析不出来时退回 ConnectionKey：此时这个连接根本发不出请求
        /// （ResolveConnection 会抛），没有哪个网关可以让一份名单张冠李戴。
        ///
        /// 刻意不记地址：读不到地址与「就是官方地址」分不开
        /// （TryReadCodexBaseUrl 吞掉异常返回 null，调用方补上官方端点），
        /// 而且 token 缺失时连地址都拿不到。名单失效交给面板侧的展示期阀门——
        /// 名单里没有一个模型出现在当前目录时就显示完整目录，那条一次管住全部原因。
        /// </summary>
        internal string FavoritesKey()
        {
            if (Mode == ConnectionMode.CustomApi || Mode.IsWorkBuddy())
            {
                return ConnectionKey();
            }

            try
            {
                return Mode + "|" + LocalCliConfig.Resolve(CliSource).Source;
            }
            catch (ProviderException)
            {
                return ConnectionKey();
            }
        }

        /// <summary>把当前模型登记为「为当前连接所选」。用户主动选定模型后调用。</summary>
        internal void StampModelConnection()
        {
            ModelConnection = string.IsNullOrWhiteSpace(Model) ? string.Empty : ConnectionKey();
        }

        /// <summary>
        /// 丢弃不属于当前连接的模型。
        ///
        /// 这是「切回本机 CLI 配置后仍在用自定义接口的模型」的根治点：
        /// 无论这个值是怎么留下来的（保存竞态、面板状态过期、旧版设置文件），
        /// 只要它的登记连接和当前连接不一致就不再生效。
        /// </summary>
        internal bool DropModelFromOtherConnection()
        {
            if (string.IsNullOrWhiteSpace(Model) ||
                string.IsNullOrEmpty(ModelConnection) ||
                ModelConnection == ConnectionKey())
            {
                return false;
            }

            Model = string.Empty;
            ModelConnection = string.Empty;
            return true;
        }

        /// <summary>
        /// 保存时决定是否留用传入的模型。
        ///
        /// <paramref name="chosenForConnection"/> 由面板给出，表示这个模型是用户在
        /// 当前这套接入配置下选的。没有这个确认而连接又变了，模型只能是上一套配置的
        /// 残留——面板的表单状态可能比磁盘旧，所以不能只看模式有没有变。
        /// </summary>
        internal void KeepModelOnlyIfChosenForConnection(string previousConnectionKey, bool chosenForConnection)
        {
            if (chosenForConnection || ConnectionKey() == previousConnectionKey)
            {
                StampModelConnection();
                return;
            }

            Model = string.Empty;
            ModelConnection = string.Empty;
        }

        /// <summary>
        /// 处理旧版设置文件里没有登记连接的模型。
        ///
        /// 自定义接口的模型只可能是为它自己选的，直接认领；本机 CLI 则无从判断，
        /// 而这正是缺陷的高发处，因此丢弃并回落到 CLI 配置自带的模型——
        /// 代价是升级后可能需要重选一次，换来的是不会继续用一个错的模型名发请求。
        /// </summary>
        internal void AdoptOrDropUnstampedModel()
        {
            if (string.IsNullOrWhiteSpace(Model))
            {
                Model = string.Empty;
                ModelConnection = string.Empty;
                return;
            }

            if (Mode == ConnectionMode.CustomApi)
            {
                ModelConnection = ConnectionKey();
                return;
            }

            Model = string.Empty;
            ModelConnection = string.Empty;
        }

        /// <summary>
        /// 侧边栏宽度，单位是宿主的窗格单位（随显示缩放变化，并非 CSS 像素）。
        /// 0 表示尚未记录，此时由面板自行校准一次。
        ///
        /// 必须持久化：不记住的话每次打开都要重新按当前视口反推宽度，
        /// 而反推依赖一次瞬时测量，结果每次都不同，表现为面板宽度自己跳动。
        /// </summary>
        internal int PaneWidth { get; set; }

        /// <summary>
        /// 面板主题："light"、"dark"，空串表示还没收到面板的报告。
        ///
        /// 权威值在面板侧的 localStorage：主题必须在首屏绘制之前定下来，
        /// 而那时读不到本文件（要走一次异步消息桥往返）。这里存的是给宿主
        /// 自己上色用的副本——面板外面那圈 WinForms 控件和 WebView2 的默认底色
        /// 都不受页面 CSS 管辖，写死白色的话深色主题下每次开面板都先闪一块白。
        /// </summary>
        internal string Theme { get; set; } = string.Empty;

        /// <summary>
        /// 选择器是否只显示常用名单里的模型。默认关——老用户升级后选择器逐字不变。
        ///
        /// 只是「要不要筛」这一个意愿，筛不筛得动由面板按当前目录决定：
        /// 名单里没有一个模型出现在目录里时一律显示完整目录，否则开关会把人锁在外面。
        /// </summary>
        internal bool OnlyFavoriteModels { get; set; }

        internal void SyncActiveProxyProfileFromLegacy()
        {
            EnsureProxyProfiles();
            var active = FindActiveProxyProfile();
            if (active == null) { return; }

            active.Kind = ProxyType;
            active.Host = ProxyHost;
            active.Port = ProxyPort;
            active.Username = ProxyUsername;
        }

        private void EnsureProxyProfiles()
        {
            if (ProxyProfiles == null) { ProxyProfiles = new List<ProxyProfile>(); }
            if (ProxyProfiles.Count == 0)
            {
                ProxyProfiles.Add(new ProxyProfile
                {
                    Id = ProxyType == ProxyKind.Direct ? "direct" : LegacyProxyProfileId,
                    Name = ProxyType == ProxyKind.Direct ? "直连" : "代理 1",
                    Kind = ProxyType,
                    Host = ProxyHost,
                    Port = ProxyPort,
                    Username = ProxyUsername,
                });
            }

            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < ProxyProfiles.Count; index++)
            {
                var profile = ProxyProfiles[index] ?? new ProxyProfile();
                profile.Id = NormalizeProxyProfileId(profile.Id, index + 1, used);
                profile.Name = string.IsNullOrWhiteSpace(profile.Name) ? $"代理 {index + 1}" : profile.Name.Trim();
                profile.Host = (profile.Host ?? string.Empty).Trim();
                profile.Username = (profile.Username ?? string.Empty).Trim();
                if (!Enum.IsDefined(typeof(ProxyKind), profile.Kind)) { profile.Kind = ProxyKind.Direct; }
                if (profile.Kind == ProxyKind.Direct || profile.Kind == ProxyKind.System)
                {
                    profile.Host = string.Empty;
                    profile.Port = 0;
                    profile.Username = string.Empty;
                }
                else if (profile.Port < 1 || profile.Port > 65535)
                {
                    profile.Port = 0;
                }

                ProxyProfiles[index] = profile;
            }

            if (FindActiveProxyProfile() == null)
            {
                ActiveProxyProfileId = ProxyProfiles[0].Id;
            }

            SyncLegacyFieldsFromActive();
        }

        private ProxyProfile FindActiveProxyProfile()
        {
            if (ProxyProfiles == null || ProxyProfiles.Count == 0) { return null; }
            foreach (var profile in ProxyProfiles)
            {
                if (profile != null && string.Equals(profile.Id, ActiveProxyProfileId, StringComparison.OrdinalIgnoreCase))
                {
                    return profile;
                }
            }

            return null;
        }

        internal void SyncLegacyFieldsFromActive()
        {
            var active = FindActiveProxyProfile();
            if (active == null) { return; }
            ProxyType = active.Kind;
            ProxyHost = active.Host ?? string.Empty;
            ProxyPort = active.Port;
            ProxyUsername = active.Username ?? string.Empty;
        }

        private static string NormalizeProxyProfileId(string value, int fallbackNumber, HashSet<string> used)
        {
            var source = (value ?? string.Empty).Trim();
            var builder = new System.Text.StringBuilder();
            foreach (var ch in source)
            {
                if (char.IsLetterOrDigit(ch) || ch == '-' || ch == '_') { builder.Append(ch); }
            }

            var id = builder.Length == 0 ? "proxy-" + fallbackNumber : builder.ToString();
            var baseId = id;
            var suffix = 2;
            while (!used.Add(id)) { id = baseId + "-" + suffix++; }
            return id;
        }

        private static List<ProxyProfile> ParseProxyProfiles(JArray array)
        {
            var profiles = new List<ProxyProfile>();
            foreach (var token in array ?? new JArray())
            {
                var item = token as JObject;
                if (item == null) { continue; }
                var profile = new ProxyProfile
                {
                    Id = item.Value<string>("id") ?? string.Empty,
                    Name = item.Value<string>("name") ?? string.Empty,
                    Host = item.Value<string>("proxyHost") ?? string.Empty,
                    Port = item.Value<int?>("proxyPort") ?? 0,
                    Username = item.Value<string>("proxyUsername") ?? string.Empty,
                };
                if (Enum.TryParse(item.Value<string>("proxyType"), out ProxyKind kind)) { profile.Kind = kind; }
                profiles.Add(profile);
            }
            return profiles;
        }

        private List<ProxyProfile> CloneProxyProfiles()
        {
            var clone = new List<ProxyProfile>();
            foreach (var profile in ProxyProfiles ?? new List<ProxyProfile>())
            {
                if (profile == null) { continue; }
                clone.Add(new ProxyProfile
                {
                    Id = profile.Id,
                    Name = profile.Name,
                    Kind = profile.Kind,
                    Host = profile.Host,
                    Port = profile.Port,
                    Username = profile.Username,
                });
            }
            return clone;
        }

        internal static string ProxySecretKeyFor(string profileId)
        {
            if (string.Equals(profileId, LegacyProxyProfileId, StringComparison.OrdinalIgnoreCase))
            {
                return ProxySecretKey;
            }

            var builder = new System.Text.StringBuilder();
            foreach (var ch in profileId ?? string.Empty)
            {
                if (char.IsLetterOrDigit(ch) || ch == '-' || ch == '_') { builder.Append(ch); }
            }
            var safeId = builder.Length == 0 ? "unknown" : builder.ToString();
            return ProxySecretKey + "-" + safeId;
        }

        private static string LoadProxyPassword(ProxyProfile profile)
        {
            if (profile == null) { return string.Empty; }
            return SecretStore.Load(ProxySecretKeyFor(profile.Id)) ?? string.Empty;
        }

        /// <summary>
        /// 复制一份仅用于组装异步响应的设置快照。
        ///
        /// ACP 探测可能跨越多个面板消息；响应回来时 _settings 可能已经被保存或
        /// 快捷设置替换。快照让返回的授权目录、模式和模型始终属于同一次请求。
        /// </summary>
        internal Settings Clone()
        {
            return new Settings
            {
                Mode = Mode,
                CliSource = CliSource,
                CustomProtocol = CustomProtocol,
                CustomBaseUrl = CustomBaseUrl,
                ProxyType = ProxyType,
                ProxyHost = ProxyHost,
                ProxyPort = ProxyPort,
                ProxyUsername = ProxyUsername,
                ProxyProfiles = CloneProxyProfiles(),
                ActiveProxyProfileId = ActiveProxyProfileId,
                AutoSwitchProxy = AutoSwitchProxy,
                Model = Model,
                ModelConnection = ModelConnection,
                Thinking = Thinking,
                Approval = Approval,
                Temperature = Temperature,
                MaxOutputTokens = MaxOutputTokens,
                ContextBudgetTokens = ContextBudgetTokens,
                MaxSteps = MaxSteps,
                AutoIncludeSelection = AutoIncludeSelection,
                ToolProtocol = ToolProtocol,
                VisionRelayModel = VisionRelayModel,
                PaneWidth = PaneWidth,
                Theme = Theme,
                OnlyFavoriteModels = OnlyFavoriteModels,
            };
        }

        private static string FilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ChatSheet",
            "settings.json");

        internal static Settings Load()
        {
            try
            {
                var path = FilePath;
                if (!File.Exists(path))
                {
                    return new Settings();
                }

                var root = JObject.Parse(File.ReadAllText(path));
                var settings = new Settings
                {
                    CustomBaseUrl = root.Value<string>("customBaseUrl") ?? string.Empty,
                    ProxyHost = root.Value<string>("proxyHost") ?? string.Empty,
                    ProxyPort = root.Value<int?>("proxyPort") ?? 0,
                    ProxyUsername = root.Value<string>("proxyUsername") ?? string.Empty,
                    ActiveProxyProfileId = root.Value<string>("activeProxyProfileId") ?? string.Empty,
                    AutoSwitchProxy = root.Value<bool?>("autoSwitchProxy") ?? true,
                    Model = root.Value<string>("model") ?? string.Empty,
                    ModelConnection = root.Value<string>("modelConnection") ?? string.Empty,
                    Temperature = root.Value<double?>("temperature"),
                    MaxOutputTokens = root.Value<int?>("maxOutputTokens") ?? 8192,
                    ContextBudgetTokens = root.Value<int?>("contextBudgetTokens") ?? 200_000,
                    MaxSteps = root.Value<int?>("maxSteps") ?? 0,
                    AutoIncludeSelection = root.Value<bool?>("autoIncludeSelection") ?? true,
                    VisionRelayModel = root.Value<string>("visionRelayModel") ?? string.Empty,
                    PaneWidth = root.Value<int?>("paneWidth") ?? 0,
                    Theme = root.Value<string>("theme") ?? string.Empty,
                    OnlyFavoriteModels = root.Value<bool?>("onlyFavoriteModels") ?? false,
                };

                var profiles = root["proxyProfiles"] as JArray;
                if (profiles != null)
                {
                    settings.ProxyProfiles = ParseProxyProfiles(profiles);
                }

                if (Enum.TryParse(root.Value<string>("mode"), out ConnectionMode mode)) { settings.Mode = mode; }
                if (Enum.TryParse(root.Value<string>("cliSource"), out CliKind cli)) { settings.CliSource = cli; }
                if (Protocols.TryParse(root.Value<string>("customProtocol"), out var protocol)) { settings.CustomProtocol = protocol; }
                if (Enum.TryParse(root.Value<string>("proxyType"), out ProxyKind proxyType)) { settings.ProxyType = proxyType; }
                // 必须用完整限定名：本类的 Thinking 属性会遮蔽 Providers.Thinking 静态类。
                if (Providers.Thinking.TryParse(root.Value<string>("thinking"), out var thinking))
                {
                    settings.Thinking = thinking;
                }
                if (Enum.TryParse(root.Value<string>("approval"), out ApprovalPolicy approval)) { settings.Approval = approval; }
                if (Enum.TryParse(root.Value<string>("toolProtocol"), out ToolProtocolPreference toolProtocol))
                {
                    settings.ToolProtocol = toolProtocol;
                }

                // 模式等字段都读完才能判断模型归属，因此收敛放在最后。
                settings.Normalize();
                return settings;
            }
            catch (Exception ex)
            {
                // 设置损坏不应阻塞使用，退回默认值并保留原文件供排查。
                Log.Error("读取设置失败，已回退默认值", ex);
                return new Settings();
            }
        }

        internal void Save()
        {
            try
            {
                Normalize();

                var root = new JObject
                {
                    ["mode"] = Mode.ToString(),
                    ["cliSource"] = CliSource.ToString(),
                    ["customProtocol"] = Protocols.Get(CustomProtocol).Id,
                    ["customBaseUrl"] = CustomBaseUrl ?? string.Empty,
                    ["proxyType"] = ProxyType.ToString(),
                    ["proxyHost"] = ProxyHost ?? string.Empty,
                    ["proxyPort"] = ProxyPort,
                    ["proxyUsername"] = ProxyUsername ?? string.Empty,
                    ["activeProxyProfileId"] = ActiveProxyProfileId ?? string.Empty,
                    ["autoSwitchProxy"] = AutoSwitchProxy,
                    ["model"] = Model ?? string.Empty,
                    ["modelConnection"] = ModelConnection ?? string.Empty,
                    ["thinking"] = Thinking.ToString(),
                    ["approval"] = Approval.ToString(),
                    ["maxOutputTokens"] = MaxOutputTokens,
                    ["contextBudgetTokens"] = ContextBudgetTokens,
                    ["maxSteps"] = MaxSteps,
                    ["autoIncludeSelection"] = AutoIncludeSelection,
                    ["toolProtocol"] = ToolProtocol.ToString(),
                    ["visionRelayModel"] = VisionRelayModel ?? string.Empty,
                    ["paneWidth"] = PaneWidth,
                    ["theme"] = Theme ?? string.Empty,
                    // 必须与 Load 成对：本文件按白名单整体重建，只读不写的键会被
                    // 下一次任意写入方（面板宽度、主题都算）抹掉。
                    ["onlyFavoriteModels"] = OnlyFavoriteModels,
                };

                var profiles = new JArray();
                foreach (var profile in ProxyProfiles ?? new List<ProxyProfile>())
                {
                    profiles.Add(new JObject
                    {
                        ["id"] = profile.Id ?? string.Empty,
                        ["name"] = profile.Name ?? string.Empty,
                        ["proxyType"] = profile.Kind.ToString(),
                        ["proxyHost"] = profile.Host ?? string.Empty,
                        ["proxyPort"] = profile.Port,
                        ["proxyUsername"] = profile.Username ?? string.Empty,
                    });
                }
                root["proxyProfiles"] = profiles;

                if (Temperature.HasValue)
                {
                    root["temperature"] = Temperature.Value;
                }

                var path = FilePath;
                AtomicFile.WriteAllText(path, root.ToString(Formatting.Indented));
            }
            catch (Exception ex)
            {
                Log.Error("保存设置失败", ex);
                throw new ProviderException("SETTINGS_SAVE_FAILED", "保存设置失败：" + ex.Message, ex);
            }
        }

        /// <summary>把越界值收敛到合理区间，避免用户或损坏文件导致异常行为。</summary>
        private void Normalize()
        {
            EnsureProxyProfiles();

            // 兼容旧版没有代理字段的模型归属键：默认直连与旧键等价，
            // 不应让升级后的第一次读取丢掉用户已选模型。
            if (ProxyType == ProxyKind.Direct && !string.IsNullOrEmpty(ModelConnection))
            {
                var currentKey = ConnectionKey();
                const string directSuffix = "|direct";
                if (currentKey.EndsWith(directSuffix, StringComparison.Ordinal) &&
                    ModelConnection == currentKey.Substring(0, currentKey.Length - directSuffix.Length))
                {
                    ModelConnection = currentKey;
                }
            }

            // 模型归属先收敛：读盘与保存都会经过这里，是唯一能保证
            // 「内存里的 Model 一定属于当前连接」的地方。
            if (string.IsNullOrEmpty(ModelConnection))
            {
                AdoptOrDropUnstampedModel();
            }
            else
            {
                DropModelFromOtherConnection();
            }

            if (MaxOutputTokens < 256) { MaxOutputTokens = 256; }
            if (MaxOutputTokens > 200_000) { MaxOutputTokens = 200_000; }

            if (ContextBudgetTokens < 8_000) { ContextBudgetTokens = 8_000; }
            if (ContextBudgetTokens > 2_000_000) { ContextBudgetTokens = 2_000_000; }

            if (MaxSteps < 0) { MaxSteps = 0; }
            if (MaxSteps > 200) { MaxSteps = 200; }

            // 0 保留「未记录」语义；其余值收敛到可用区间，
            // 损坏的极端值会让面板窄到无法操作或宽到盖住工作表。
            if (PaneWidth != 0)
            {
                if (PaneWidth < 200) { PaneWidth = 200; }
                if (PaneWidth > 4000) { PaneWidth = 4000; }
            }

            // 只认这两个值。别的一律收回空串，让宿主退回浅色——
            // 拿一个不认识的主题名去查颜色表只会得到默认色，不如显式表达「不知道」。
            if (Theme != "light" && Theme != "dark") { Theme = string.Empty; }

            VisionRelayModel = (VisionRelayModel ?? string.Empty).Trim();

            if (Temperature.HasValue)
            {
                if (Temperature.Value < 0) { Temperature = 0; }
                if (Temperature.Value > 2) { Temperature = 2; }
            }
        }

        /// <summary>
        /// 解析出可用于发起请求的接入信息。
        /// 令牌只在返回值中短暂存在，不写日志、不回传面板。
        /// </summary>
        internal ResolvedConnection ResolveConnection()
        {
            switch (Mode)
            {
                case ConnectionMode.CustomApi:
                {
                    if (string.IsNullOrWhiteSpace(CustomBaseUrl))
                    {
                        throw new ProviderException("BASE_URL_REQUIRED", "尚未填写接口地址。");
                    }

                    var token = SecretStore.Load(CustomApiSecretKey);
                    if (string.IsNullOrWhiteSpace(token))
                    {
                        throw new ProviderException("TOKEN_REQUIRED", "尚未填写接口密钥。");
                    }

                    return new ResolvedConnection
                    {
                        Protocol = CustomProtocol,
                        BaseUrl = Protocols.NormalizeBaseUrl(CustomBaseUrl, CustomProtocol),
                        Token = token,
                        Model = Model,
                        ProxyCandidates = ResolveProxyCandidates(),
                        Proxy = ResolveProxy(),
                        SourceLabel = "自定义接口",
                    };
                }

                case ConnectionMode.Authorized:
                case ConnectionMode.AuthorizedInternational:
                    // WorkBuddy 的令牌与对话协议由 WorkBuddy 自己管理。
                    // 这里返回的是路由标记，不是可以交给普通 HTTP ChatClient 的 API 连接。
                    return new ResolvedConnection
                    {
                        Protocol = Protocols.Default,
                        Model = Model,
                        Proxy = new ProxyOptions(),
                        ProxyCandidates = new List<ProxyOptions> { new ProxyOptions() },
                        SourceLabel = Mode == ConnectionMode.AuthorizedInternational
                            ? "WorkBuddy 国际版授权"
                            : "WorkBuddy 授权",
                        IsWorkBuddy = true,
                        WorkBuddyMode = Mode,
                    };

                default:
                {
                    var credentials = LocalCliConfig.Resolve(CliSource);
                    return new ResolvedConnection
                    {
                        Protocol = credentials.Protocol,
                        BaseUrl = credentials.BaseUrl,
                        Token = credentials.Token,
                        // 用户未指定模型时沿用 CLI 配置中的模型。
                        Model = string.IsNullOrWhiteSpace(Model) ? credentials.Model : Model,
                        ProxyCandidates = ResolveProxyCandidates(),
                        Proxy = ResolveProxy(),
                        SourceLabel = credentials.DisplayName,
                    };
                }
            }
        }

        /// <summary>读取当前代理配置；代理密码只从 DPAPI 密钥槽加载。</summary>
        internal ProxyOptions ResolveProxy(string passwordOverride = null)
        {
            NormalizeProxyForUse();
            return BuildProxyOptions(FindActiveProxyProfile(), passwordOverride);
        }

        internal IReadOnlyList<ProxyOptions> ResolveProxyCandidates(string passwordOverride = null)
        {
            NormalizeProxyForUse();
            var candidates = new List<ProxyOptions>();
            var active = FindActiveProxyProfile();
            if (active == null) { active = ProxyProfiles[0]; }
            candidates.Add(BuildProxyOptions(active, passwordOverride));

            if (AutoSwitchProxy)
            {
                foreach (var profile in ProxyProfiles)
                {
                    if (profile != null && !string.Equals(profile.Id, active.Id, StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            candidates.Add(BuildProxyOptions(profile, null));
                        }
                        catch (ProviderException)
                        {
                            // 无效的备用配置不应阻塞当前有效代理；用户切换到它时仍会得到校验错误。
                        }
                    }
                }
            }

            return candidates;
        }

        private ProxyOptions BuildProxyOptions(ProxyProfile profile, string passwordOverride)
        {
            profile = profile ?? new ProxyProfile
            {
                Id = "direct",
                Kind = ProxyType,
                Host = ProxyHost,
                Port = ProxyPort,
                Username = ProxyUsername,
            };

            if (profile.Kind != ProxyKind.Direct && profile.Kind != ProxyKind.System &&
                (string.IsNullOrWhiteSpace(profile.Host) || profile.Port < 1 || profile.Port > 65535))
            {
                throw new ProviderException("PROXY_INVALID", "代理地址或端口无效。");
            }

            return new ProxyOptions
            {
                ProfileId = profile.Id,
                Kind = profile.Kind,
                Host = profile.Host ?? string.Empty,
                Port = profile.Port,
                Username = profile.Username ?? string.Empty,
                // null means the caller did not override the password; an empty
                // string is an explicit unsaved clear from the settings page.
                Password = passwordOverride != null ? passwordOverride : LoadProxyPassword(profile),
            };
        }

        private void NormalizeProxyForUse()
        {
            EnsureProxyProfiles();
        }

        /// <summary>
        /// 解析视觉中转的目标。未配置中转模型时返回 null。
        ///
        /// 沿用主连接的协议、地址与密钥，只换模型名：同一个服务商下通常本来就有
        /// 带视觉的型号，要求用户再配一整套接入信息会把「贴张截图」变成配置作业。
        /// </summary>
        internal ResolvedRelayTarget ResolveVisionRelay(ResolvedConnection connection)
        {
            var model = (VisionRelayModel ?? string.Empty).Trim();
            if (model.Length == 0 || connection == null)
            {
                return null;
            }

            // 中转模型与主模型同名时不必中转：主模型看不了图，同名的它也看不了。
            if (string.Equals(model, connection.Model, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return new ResolvedRelayTarget
            {
                Protocol = connection.Protocol,
                BaseUrl = connection.BaseUrl,
                Token = connection.Token,
                Proxy = connection.Proxy,
                ProxyCandidates = connection.ProxyCandidates,
                Model = model,
            };
        }
    }

    internal sealed class ResolvedConnection
    {
        internal ProtocolKind Protocol { get; set; }

        internal string BaseUrl { get; set; }

        internal string Token { get; set; }

        internal ProxyOptions Proxy { get; set; } = new ProxyOptions();

        internal IReadOnlyList<ProxyOptions> ProxyCandidates { get; set; } = new List<ProxyOptions>();

        internal string Model { get; set; }

        internal string SourceLabel { get; set; }

        /// <summary>该连接由 WorkBuddy ACP 处理，不使用 BaseUrl 或 Token。</summary>
        internal bool IsWorkBuddy { get; set; }

        /// <summary>WorkBuddy 连接所属的授权空间。</summary>
        internal ConnectionMode WorkBuddyMode { get; set; } = ConnectionMode.Authorized;
    }
}
