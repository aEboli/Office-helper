# 验证

- `dotnet build ChatSheet.sln --configuration Release --nologo`：通过，0 个警告，0 个错误。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --agent-tool-tests`：通过，177/177。新增 22 条覆盖孤儿结果配平、参数校验宽容度、四种协议的工具痕迹降级与历史对象未被改写。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --reliability-tests`：通过，32/32，与上一版基线一致，无回归。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --project-audit-tests`：通过，13/13（真实 Excel 宿主，含参数纠错后实际写入并读回）。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --codex-config-tests`：通过，17/17。
- `tests/web/*.test.mjs`：通过，37 个测试文件、1129 条用例全部通过。
- `scripts/verify-all.ps1`（完整链路，含 Excel 宿主测试）：通过 42，失败 0，跳过 0，退出码 0；`-SkipHost` 为通过 41、失败 0、跳过 1。
- 独立探针复核：全配对混入孤儿结果时 `DroppedResults=1`、配对结果保留、历史 assistant 的调用数仍为 1；未声明工具的请求体中不含 `role:"tool"`/`tool_calls`/`function_call`/`function_call_use`/`tool_result`，而历史对象仍是工具角色。
- 安装新版本：通过。安装目录中的 `ChatSheet.AddIn.dll` 与 Release 构建逐字节一致（两者 SHA-256 相同），安装后文件版本为 `0.10.3.20`；x86/x64 COM 注册与 Excel、Word、WPS 表格与 WPS Writer 宿主登记正常；用户 `settings.json` 未被改动。
- `scripts/package-release.ps1` 实跑：通过。ZIP 哈希与外部 sidecar 一致，包内 `SHA256SUMS.txt` 与 `app/ChatSheet.AddIn.dll` 一致，包内发行说明与仓库中的发行说明逐字节相同。
- 发布包来源可追溯：包内 `ChatSheet.AddIn.dll` 与 `ChatWord.AddIn.dll` 的 `ProductVersion` 内嵌修订号 SHALL 等于被标记的发布提交，从而按 tag 取到的源码与包内二进制同源。发布时逐项核对：
  - 先确认本地提交已推到远端，再打包；
  - 解出包内两个程序集的 `ProductVersion`，与 `git rev-parse <tag>^{commit}` 比对；
  - 不一致即判定为「包与 tag 不同源」，重打而不是照发。
  - 打包必须在发布提交完成之后运行。在提交前打包会内嵌上一个提交号，包内二进制与 tag 指向的源码不同源；本次首轮打包即命中该问题（内嵌 `d63ca40`，那个提交里的源码版本还是 0.10.3.18），已重打并核对。
- `openspec validate repair-request-shape-and-argument-tolerance --strict --no-interactive`：通过。
- `git diff --check`：通过。
- GitHub `main` 与 `v0.10.3.20` Release：待本次发布完成后回填。