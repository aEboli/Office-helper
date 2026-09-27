<div align="center">

# Office-helper

**把 AI 放进 Excel、Word 和 WPS 侧边栏，用自然语言处理表格与文档。**

Office-helper 是产品显示名；Excel 入口继续保留 `ChatSheet.AddIn` 的旧 ProgID/CLSID 和安装目录，确保已有安装可以无感升级。

Windows COM 加载项 · WebView2 面板 · 流式对话 · 可审阅的 Excel/Word 工具

[![Latest release](https://img.shields.io/github/v/release/aEboli/Office-helper?display_name=tag&sort=semver)](https://github.com/aEboli/Office-helper/releases)
[![Windows](https://img.shields.io/badge/Windows-10%2F11-0078D4?logo=windows&logoColor=white)](#安装)
[![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?logo=dotnet&logoColor=white)](#从源码构建)

[下载 v0.10.3.17](https://github.com/aEboli/Office-helper/releases/tag/v0.10.3.17) · [发行说明](docs/releases/v0.10.3.17.md) · [提交问题](https://github.com/aEboli/Office-helper/issues)

</div>

> [!TIP]
> 普通用户下载 Windows ZIP，解压后运行 `install.bat`。不需要 .NET SDK、Node.js、Office.js 旁加载或本地 HTTP 服务。

## 先看这里

Office-helper 在 Excel/WPS 表格中读取当前工作簿和选区，在 Word/WPS Writer 中读取当前文档和明确 Story 目标，再通过对应工具完成分析与修改。当前源码版本 0.10.3.17 默认全自动执行，操作卡片报告实际结果；仍可在设置中主动选择逐项或每轮审批。Excel 与 Word 入口统一显示为 Office-helper，旧 ChatSheet ProgID/CLSID 仅用于兼容已有安装。

~~~text
选区 / 工作簿 → 模型理解上下文 → 表格工具执行 → 操作卡片与结果核验
~~~

适合解释公式、清理数据、批量改格式、合并标题、创建表格或图表，也适合把自然语言要求变成能复核、能撤销的表格操作。

## 能做什么

| 场景 | 能力 | 默认保护 |
| --- | --- | --- |
| 读取与分析 | 工作簿结构、选区、值、公式、格式和异常 | 大范围分页读取 |
| 写入与公式 | 写值、写公式、清除内容或格式 | 默认自动执行，大范围分块写入 |
| 格式与结构 | 字体、边框、筛选、规则、行列、图表、透视表及视图 | 专用工具与通用对象入口，宿主支持以读回为准 |
| 长任务 | 流式输出、输入排队、持续执行和停止 | 步数 0 表示不限，连续相同失败终止 |
| 文档编辑 | Word/WPS Writer 文档、段落、表格与明确 Story 目标 | 写入后读回；仅在可创建快照时提供撤销 |
| 图片与附件 | PNG、JPEG、WebP 图片和文本附件 | 限制数量与大小，拒绝二进制文件 |
| 模型接入 | OpenAI、Anthropic、Gemini、兼容网关、本机 CLI、WorkBuddy | API Key 用 Windows DPAPI 保护 |

### Word 文档编辑

起草内容写入当前光标；改写时替换明确选区。若非空选区、Story 或段落范围不清楚，会先询问。发送请求会等待整轮文档任务结束，完成或停止后再更新状态；写入结果会读回核对，只有能创建可靠快照时才提供撤销。

### Word 草稿预览与版式设计

起草、续写和生成式改写会先生成侧栏草稿，不会直接修改文档。预览保留标题、段落、列表、表格、引用和重点文字等结构，可选择沿用文档、简洁、商务或报告风格，也可以继续用自然语言调整。确认插入或替换后，系统会重新核对文档与选区，再按当前审批和保护规则写入；丢弃草稿不会改变文档。结构化草稿只有在快照能够完整恢复时才提供撤销。

### 模型能力适配

自动模式会按当前连接和模型记录工具能力。服务商拒绝原生工具字段，或模型明确表示不支持工具/函数调用时，会在同一步切换到文本指令协议，继续使用同一套审批、参数校验和撤销链路；普通安全拒绝不会触发切换。明确匹配 DeepSeek Flash 变体时，自动模式首轮直接使用文本指令协议，手动选择原生调用仍优先。

### Word 信函格式

格式规范只在用户明确提到文种、版式或标准时应用；只使用中文或英文，不会触发格式推断。文档已有结构会保留，格式变体不明确时先询问。

| 用户点名的格式 | 常见排版约定 | 称呼与开场 |
| --- | --- | --- |
| 中文一般书信 | 称呼顶格；正文另起段并首行缩进两字；祝颂语、署名和日期置于文末，落款通常靠右。一般书信通常没有居中标题；这属于常见惯例，不是统一强制标准。 | 称呼可按关系使用“尊敬的”“亲爱的”等表达；称呼行通常以中文冒号收尾，问候或正文另起段。 |
| 英文商务信函 full block | 各部分左对齐，正文不缩进，以空行分段。 | 常用 Dear 加合适称谓或姓名；首段自然说明写信目的。 |
| 英文商务信函 modified block | 正文仍左对齐且不缩进；日期及结尾敬语、署名移至页面中部附近。 | 称呼和开场遵循英文商务语境，不逐词翻译“您好”“冒昧打扰”等中文表达。 |

Word 段落工具以字符为单位设置首行缩进，并读回检查实际值；不会用空格模拟缩进。标题居中、落款右对齐等细节只在被点名的格式要求下应用，且只处理本轮新增或明确指定的段落。

## 安装

### 预构建 Windows ZIP（推荐）

1. 下载 [`ChatSheet-v0.10.3.17-win.zip`](https://github.com/aEboli/Office-helper/releases/download/v0.10.3.17/ChatSheet-v0.10.3.17-win.zip) 和同名 `.sha256` 文件。
2. 在 PowerShell 中核对哈希：

   ~~~powershell
   Get-FileHash -Algorithm SHA256 .\ChatSheet-v0.10.3.17-win.zip
   Get-Content .\ChatSheet-v0.10.3.17-win.zip.sha256
   ~~~

3. 完整解压 ZIP，双击根目录 `install.bat`，选择安装或更新。
4. 接受 UAC 后重启 Excel、Word 或 WPS：Excel/WPS 表格与 Word/WPS Writer 均显示 **Office-helper** 入口。
5. 在“设置”中选择 API 接入、WorkBuddy 国内版或 WorkBuddy 国际版；代理在独立的“代理”页配置。

升级前请保存并完全退出 Excel、Word 和 WPS，避免旧 DLL 被宿主占用。安装器会部署到 `%LOCALAPPDATA%\ChatSheet\app`（旧路径保留以兼容设置），同时注册表格和文档 COM 加载项并运行自检。

完整校验、诊断、升级和卸载说明见 [Windows 发行包安装指南](docs/windows-release-install.md)。

### 从源码构建

需要 Windows、.NET Framework 4.8 Developer Pack、.NET SDK 和 WebView2 Runtime：

~~~powershell
dotnet build ChatSheet.sln --configuration Release
.\scripts\install.ps1 -Action install
~~~

## WorkBuddy 模型目录

国际版 WorkBuddy 会继承官方桌面端的产品上下文，因此 ChatSheet 与桌面端使用同一份完整模型目录，包括官方返回的 `0.00x` 模型；独立 CodeBuddy CLI 则使用自己的目录。国内版与国际版的账号、配置和模型来源彼此隔离，ChatSheet 不复制 WorkBuddy token，也不主动注销官方客户端。

## 安全边界

- 模型只能调用当前宿主公开的表格或文档工具，不能直接执行 PowerShell、读取文件系统或随意联网。
- 审批、执行、撤销和恢复都绑定原工作簿；切换工作簿、停止任务或关闭面板后，迟到回调会被丢弃。
- WebView2 顶层导航仅允许 `https://chatsheet.local/...`，不监听本地端口，也不需要开发证书。
- 设置、密钥和常用模型名单使用原子写入；异常中断不会清空旧数据。

## 兼容性

| 项目 | 支持范围 |
| --- | --- |
| 操作系统 | Windows 10 / 11 |
| 主要宿主 | Microsoft Excel、Microsoft Word 桌面版，已验证 x86 / x64 |
| 额外宿主 | WPS 表格 ET、WPS Writer，已验证 x86 / x64；其他版本请先诊断 |
| 运行时 | .NET Framework 4.8、Microsoft Edge WebView2 Runtime |
| 不支持 | Excel for Mac、Excel 网页版 |

发布 ZIP 未进行代码签名，也不是 MSI/EXE。请只从 GitHub Release 下载并先核对 SHA-256。

## 开发与测试

~~~powershell
dotnet build ChatSheet.sln --configuration Release
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --reliability-tests
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --project-audit-tests
Get-ChildItem tests\web\*.test.mjs | ForEach-Object { node $_.FullName }
~~~

涉及 COM 注册、Excel/WPS 宿主调用或 WebView2 初始化时，请先阅读 [架构说明](docs/architecture.md)。规范变更按仓库内 OpenSpec 流程维护。

## 文档

- [v0.10.3.17 发行说明](docs/releases/v0.10.3.17.md)
- [Windows 发行包安装、校验与卸载](docs/windows-release-install.md)
- [架构说明与常见宿主陷阱](docs/architecture.md)
- [全部 GitHub Releases](https://github.com/aEboli/Office-helper/releases)
- [问题反馈](https://github.com/aEboli/Office-helper/issues)

## 许可证

GitHub 仓库目前为公开状态，但尚未附带许可证。公开可见不等于授予复制、修改或分发权限；如需复用，请先与维护者确认许可证安排。


## 一键安装与 macOS 边界

Windows 用户可以运行：

    irm https://raw.githubusercontent.com/aEboli/Office-helper/main/scripts/install-online.ps1 | iex

安装器会从 GitHub Release 下载 ZIP 和 SHA-256 sidecar，校验一致后再调用本地安装脚本。也可以下载 ZIP 后双击根目录的 install.bat，菜单提供安装、卸载和诊断。

当前版本是 Windows COM 加载项，依赖 Excel for Windows、.NET Framework 4.8 和 WebView2，不能安装到 Excel for Mac。仓库提供 install.command 和 scripts/install-macos.sh 作为一键兼容性检查；它不会修改 macOS 或伪装成已安装。macOS 原生支持需要 Office.js 加载项和跨平台本地服务，尚未包含在 v0.10.3.17。

- [macOS 安装边界](docs/macos-install.md)
