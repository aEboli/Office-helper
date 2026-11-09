# 验证

- `dotnet build ChatSheet.sln --configuration Release --nologo`：通过，0 个警告，0 个错误。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --agent-tool-tests`：通过，177/177。新增 22 条覆盖孤儿结果配平、参数校验宽容度、四种协议的工具痕迹降级与历史对象未被改写。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --reliability-tests`：通过，32/32，与上一版基线一致，无回归。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --project-audit-tests`：通过，13/13（真实 Excel 宿主，含参数纠错后实际写入并读回）。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --codex-config-tests`：通过，17/17。
- `tests/web/*.test.mjs`：通过，37 个测试文件、1129 条用例全部通过。
- `scripts/verify-all.ps1`（完整链路，含 Excel 宿主测试）：通过 42，失败 0，跳过 0，退出码 0；`-SkipHost` 为通过 41、失败 0、跳过 1。
- 独立探针复核：全配对混入孤儿结果时 `DroppedResults=1`、配对结果保留、历史 assistant 的调用数仍为 1；未声明工具的请求体中不含 `role:"tool"`/`tool_calls`/`function_call`/`function_call_use`/`tool_result`，而历史对象仍是工具角色。
- 安装新版本：通过。安装目录中的 `ChatSheet.AddIn.dll` 与 Release 构建 SHA-256 一致（`E597FB7E3C26B412BEE2E1168ACE4EFBF2CA6EC94F5BF6A20C41497A5061D0D4`），安装后文件版本为 `0.10.3.20`；x86/x64 COM 注册与 Excel、Word、WPS 表格与 WPS Writer 宿主登记正常；用户 `settings.json` 未被改动。
- `scripts/package-release.ps1` 实跑：通过。ZIP 哈希与 sidecar 一致（`255aec18c4b4857fcc73f734585417a9686f7b71d4a45e4e6feb880a54d0dd7f`），包内 `SHA256SUMS.txt` 与 `app/ChatSheet.AddIn.dll` 一致，包内发行说明与仓库中的发行说明逐字节相同。
- `openspec validate repair-request-shape-and-argument-tolerance --strict --no-interactive`：通过。
- `git diff --check`：通过。
- GitHub `main` 与 `v0.10.3.20` Release：待本次发布完成后回填。