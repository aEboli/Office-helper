# 安装 Office-helper Windows 发行包

Office-helper 同时安装 Excel/WPS 表格和 Word/WPS Writer 两个 COM 入口。旧 Excel 入口仍使用 ChatSheet 的 ProgID、CLSID 和本地目录，仅作为兼容标识，不代表产品仍只支持 Excel。

本文适用于 GitHub Release 中的 `ChatSheet-v0.10.3.16-win.zip`，不适用于从 Git 克隆的源码目录。

> [!IMPORTANT]
> 这是带 PowerShell 安装入口的预构建 ZIP，不是 MSI 或 EXE 安装器。发行包未进行代码签名；请先核对 SHA-256，再运行安装脚本。

## 发行包内容

解压 `ChatSheet-v0.10.3.16-win.zip` 后，目录结构如下：

```text
ChatSheet-v0.10.3.16-win/
├── install.bat                  # 双击打开安装菜单
├── app/                         # 预构建 COM 加载项、WebView2 面板和依赖
├── scripts/
│   ├── install.ps1              # 安装、卸载和只读诊断入口
│   ├── menu.ps1                 # install.bat 背后的交互菜单
│   └── ChatSheet.Registration.psm1
├── INSTALL.md                   # 本文副本，方便离线查阅
├── RELEASE-NOTES.md             # v0.10.3.16 发行说明
└── SHA256SUMS.txt               # 包内文件校验清单
```

GitHub Release 页面同时提供 `ChatSheet-v0.10.3.16-win.zip.sha256`，用于核对整个 ZIP 文件。

## 系统要求与支持范围

| 项目 | 要求或边界 |
| --- | --- |
| 操作系统 | Windows 10/11 |
| 主要宿主 | Microsoft Excel、Microsoft Word 桌面版，已验证 x86 与 x64 |
| 额外宿主 | WPS 表格 ET、WPS Writer，已验证 x86/x64 的 12.0 Build 19823；其他版本请先运行诊断 |
| 必需运行时 | .NET Framework 4.8、Microsoft Edge WebView2 Runtime |
| 安装权限 | 安装和卸载时需要接受 UAC 管理员授权，用于在 HKLM 注册托管 COM 类；宿主加载项登记写入当前用户 |
| 不需要 | .NET SDK、Node.js、Office.js 开发证书、常驻本地 HTTP 服务 |
| 不支持 | Excel for Mac、Excel 网页版 |

安装包不会自动更新，也不承诺静默安装或跨用户安装。Windows 或组织策略可能对网络下载的 ZIP 或 PowerShell 脚本显示安全提示；请只从 [Office-helper GitHub Release](https://github.com/aEboli/Office-helper/releases/tag/v0.10.3.16) 下载。

## 1. 下载并验证 SHA-256

从 [v0.10.3.16 Release](https://github.com/aEboli/Office-helper/releases/tag/v0.10.3.16) 下载下面两个文件到同一目录：

- `ChatSheet-v0.10.3.16-win.zip`
- `ChatSheet-v0.10.3.16-win.zip.sha256`

在 PowerShell 中进入下载目录：

```powershell
Get-FileHash -Algorithm SHA256 .\ChatSheet-v0.10.3.16-win.zip
Get-Content .\ChatSheet-v0.10.3.16-win.zip.sha256
```

以 Release 页面提供的 `.sha256` 文件为准；两条命令中的 64 位十六进制值必须完全一致。若不一致，请不要解压或运行脚本，删除文件后重新下载。

解压后还可以核对包内文件：

```powershell
$expected = Get-Content .\SHA256SUMS.txt
foreach ($line in $expected) {
    $hash, $relative = $line -split '\s{2}', 2
    $actual = (Get-FileHash -Algorithm SHA256 $relative).Hash.ToLowerInvariant()
    if ($actual -ne $hash) { throw "校验失败：$relative" }
}
'包内文件 SHA-256 校验通过'
```

## 2. 解压并安装

1. 右键 ZIP，选择“全部解压缩”，保留完整的 `ChatSheet-v0.10.3.16-win` 目录结构，不要只复制 DLL。
2. 保存并完全关闭 Microsoft Excel、Microsoft Word 和 WPS。覆盖升级时旧宿主可能占用 DLL，安装器会拒绝半新半旧的复制。
3. 双击解压根目录下的 `install.bat`，在菜单中输入 `1`：

   ```text
     [1] 安装或更新    覆盖安装并注册，Excel/WPS 需重启才生效
     [2] 卸载          反注册并删除安装目录，保留日志与设置
     [3] 诊断          检查运行时、COM 注册、宿主登记与日志
     [4] 退出
   ```

   也可以在解压根目录打开 PowerShell，执行：

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Action install
   ```

4. 接受 UAC。脚本会把 `app` 复制到 `%LOCALAPPDATA%\ChatSheet\app`，向 HKLM 的 32 位和 64 位 Classes 视图注册两个托管 COM 类，并登记 Excel、Word、WPS 表格和 WPS Writer；WPS Writer 的 `OfficeHelper.Word.AddIn` 同时写入当前用户及安装级 x86/x64 `AddinsWL` 白名单。
5. 完全重新打开宿主：Excel/WPS 表格与 Word/WPS Writer 均使用 **Office-helper** 入口。首次使用请在“设置”页选择接入模式和模型。

> [!NOTE]
> 发行包使用 `app\` 中的预构建产物，不会调用 `dotnet build`，因此不需要 .NET SDK。源码安装和 ZIP 安装是两条独立路径。

安装时还会清理 WPS Writer 对 `OfficeHelper.Word.AddIn` 留下的用户级和安装级 `AddinsCL`/`AddinsBL` 禁用项。它们会让 WPS 在双击打开 `.docx` 时跳过 COM 加载，即使白名单和 `LoadBehavior` 都正确。

## 3. 诊断、升级与卸载

安装器菜单可以执行诊断、安装/升级和卸载；命令行等价操作如下：

| 命令 | 用途 |
| --- | --- |
| `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Action diagnose` | 检查 .NET Framework、WebView2、COM 注册、宿主登记、`LoadBehavior` 与日志 |
| `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Action install` | 安装或覆盖升级；旧 DLL 被占用时会明确提示先关闭宿主 |
| `powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Action uninstall` | 反注册并删除 `%LOCALAPPDATA%\ChatSheet\app`；需要先关闭宿主并接受 UAC |

卸载不会自动删除 `%LOCALAPPDATA%\ChatSheet` 下的设置、DPAPI 加密密钥、WebView2 用户数据或日志。如需彻底清理，请先备份需要保留的信息，再手动删除该目录。

## 4. 重要边界

- 哈希一致只能证明下载的 ZIP 与发布时的字节一致；它不等同于你的 Excel、组织策略、模型服务或工作簿一定可用。
- 发布包的验证覆盖 Release 构建、工具和可靠性回归、网页测试、真实 WebView2 面板安全测试、OpenSpec 校验、解包布局、包内哈希和宿主加载；首次在新机器上使用仍应先做小范围手工验收。
- ChatSheet 默认逐项审批写入、格式、排序和结构变化。重要工作簿仍应先备份并人工复核模型生成的修改。
- 使用模型时，完成请求所需的提示词、工作簿结构、选区/读取范围结果以及附加的图片和文本文件内容可能会发送给你配置的服务商。请只配置可信端点，并遵守其隐私、计费和数据政策。

更多功能和开发说明请见仓库根目录的 [README](../README.md) 与 [架构说明](architecture.md)。
