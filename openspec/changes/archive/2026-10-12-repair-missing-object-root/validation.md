# 验证

- `dotnet build ChatSheet.sln --configuration Release --nologo`：通过，0 个警告，0 个错误。
- `tests/ChatSheet.ToolTests/bin/Release/ChatSheet.ToolTests.exe --reliability-tests`：通过，32/32；包含对象入口请求关联、白名单校验、显式取消、停止取消和迟到响应拒绝。
- `tests/web/*.test.mjs`：通过，37/37 个测试文件。
- `node --check src/web/scripts/chat.js` 与 `node --check tests/web/approval-review.test.mjs`：通过。
- `openspec validate repair-missing-object-root --strict --no-interactive`：通过。
- `git diff --check`：通过。
- 安装 `0.10.3.18`：通过。安装目录中的 `ChatSheet.AddIn.dll` 与 Release 构建 SHA-256 一致；安装前后设置文件哈希一致；`install.ps1 -Action diagnose` 报告 x86/x64 COM 注册、Excel/WPS 宿主登记、WebView2 和 .NET Framework 正常。
- WPS 表格实机启动：通过，日志确认 `OnConnection`、`GetCustomUI`、功能区加载；测试工作簿哈希未变化，启动的 WPS 窗口已正常关闭。
- `ChatSheet.PaneHarness.exe --auto 12`：通过，WebView2 初始化成功、页面消息桥连通、面板布局日志正常。
- GitHub `main`：提交 `cd60202` 已推送，远端分支更新成功；未创建公开 Release。
