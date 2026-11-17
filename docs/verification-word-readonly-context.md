# Word 对话误撤销正文修复验证

验证日期：2026-11-17。

## 问题与复现

Word 对话开始、文档摘要刷新以及结构读取都会调用 `WordDocumentContext.GetSummary()` 或 `CapabilityPayload()`。旧实现使用 `WordCom.TryGet(document, "Undo", ...)` 探测撤销能力，相当于尝试执行该成员；在本机 WPS Writer 中，按属性读取 `Undo` 仍会触发原生撤销。

修复前，在专门创建的 WPS 测试文档中插入 `readonly probe body`，读取一次摘要：正文从 20 个字符变为仅剩 1 个段落结束符。没有调用模型，也没有执行替换或删除工具。刚输入或整篇粘贴的正文因此可能被整次撤销。旧摘要每次会探测两次 `Undo`，结构读取还会重复生成能力摘要。

本机 Microsoft Word 的同一检查未触发正文变化，不能把 WPS 的复现结果泛化为所有 Word 版本。

## 修复范围

- `WordCom.HasMember`：COM 对象通过 `IDispatch.GetIDsOfNames` 查询成员名；托管测试对象使用反射元数据查询，不读取属性、不调用方法。
- `CapabilityPayload`：撤销能力改用上述查询，每次生成能力摘要只查询一次，保留现有 `undo` 和 `unsupportedMembers` 数据结构。
- 快照撤销、工具审批、文档写入规则及对话界面没有改动。
- 回归测试覆盖读取摘要、读取结构、重复读取不改变正文；真实宿主检查还验证原生撤销和重做记录仍可使用。
- 宿主冒烟测试只关闭自己创建的测试文档；存在原有文档时恢复原活动文档，不退出该宿主。

这是只读路径的缺陷修复，没有新增能力或改变公开接口，按现有 OpenSpec 规范直接修复。

## 验证结果

| 检查 | 结果 |
| --- | --- |
| Release 解决方案构建 | 通过，0 警告、0 错误 |
| Word 工具回归 | 127 项通过，0 失败 |
| Agent 工具回归 | 279 项通过，0 失败 |
| Microsoft Word 真实宿主冒烟 | 19 项通过，包含正文保留、原生撤销/重做、写入读回及快照撤销 |
| WPS Writer 真实宿主冒烟 | 19 项通过，包含与 Microsoft Word 相同的检查 |
| Git 差异空白检查 | 通过 |

修复前新增的两项正文保留回归均失败，修复后通过。宿主测试使用专门创建的临时文档，没有调用真实模型或发送付费 API 请求；本次没有验证实际侧栏的完整模型对话。

本地测试日志位于 `artifacts/validation/word-readonly-context/`，不纳入版本控制。

## 复跑

```powershell
dotnet build ChatSheet.sln --configuration Release --nologo
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --word-tests
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --agent-tool-tests
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --word-smoke
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --wps-word-smoke
git diff --check
```

## 本机生效范围

本修复纳入 `v0.10.3.21` Windows 发行包。发布不会自动替换 `%LOCALAPPDATA%\ChatSheet\app` 内已安装的加载项。保存文档并完全退出 Excel、Word 和 WPS 后，运行命令安装器或发行包中的 `install.bat` 更新安装，再重新打开宿主，才能使用修复后的版本。

```powershell
irm https://raw.githubusercontent.com/aEboli/Office-helper/main/scripts/install-online.ps1 | iex
```
