# 验证记录

- `dotnet build src/ChatWord.AddIn/ChatWord.AddIn.csproj --configuration Release`：通过，0 个警告，0 个错误。
- `Get-Content -Raw src/web/scripts/chat.js | node --check --input-type=module -`：通过。
- `openspec validate word-draft-design-preview --strict --no-interactive`：通过。
- `git diff --check`（本次修改文件）：通过。
- 13 个相关源文件与 OpenSpec 文件完成 UTF-8 替换字符检查：未发现乱码标记。
- 未新增或运行自动化测试；尚未在 Word/WPS 宿主中验证实时 COM 操作。
