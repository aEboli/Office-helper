# 工具调用 BUG 排查与验证

验证日期：2026-11-11。范围：当前 Office-helper 工作区的 Excel、Word/WPS 工具执行、模型协议解析和取消处理。

使用三名 agent 分别检查 Excel 工具、Word 工具和协议层，主 agent 处理调用入口、停止/审批、集成与真实宿主验证。协议成员因服务限流退出后，其已写入的修改由主 agent 复核并验证。

## 修复内容

| 范围 | 已复现的问题 | 修复后的行为 |
| --- | --- | --- |
| Word 调用入口 | 非法 JSON 被当成空参数；缺失工具名或草稿内部类型错误使整轮失败 | 返回可纠正的参数错误，区分截断；参数校验先于审批和宿主访问 |
| Word 停止与关闭 | 审批等待没有随停止结束；已排队回调在关闭后仍执行；取消被包装为宿主错误 | 审批等待绑定取消令牌，实际宿主回调重新检查取消与面板状态 |
| Word 文本写入 | 缺少 `replace`/`text` 时可能清空内容；读回异常被当成空文本成功 | 缺少值与显式空字符串分别处理；读回失败如实返回；折叠光标合法的 null 文本仍可用 |
| Word 表格与目标 | 嵌套表格定位被顶层空值覆盖；替换已有单元格被当成插入；书签/内容控件实际 Story 未核对 | 保留有效定位，拒绝冲突，按实际 Story 执行与判断撤销能力 |
| Word 单元格边界 | 真实 Word 的 CR+BEL 可能只占一个字符位置，固定减 2 会漏改最后一个字 | 按宿主返回的实际末尾位置剔除标记，替换和清空不残留末字 |
| Word 分页符 | 非空选区可能被分页符替换；没有确认插入结果 | 要求折叠位置，写入后读回分页符 |
| Excel 对象工具 | 可选 `args:null` 被拒、路径 `kind` 校验和执行不一致、大整数溢出 | 接受合法空值并统一解析；保留 Int32 枚举和 Int64 数值；路径执行前校验参数 |
| Excel 矩阵与格式 | 不等长行被补空造成意外清空；部分格式已写入后才发现另一个参数非法 | 写入前拒绝不等长行和非标量单元格；格式参数完整验证后再修改 |
| Excel 分页与分块 | 宿主事件切换工作簿后再次按活动工作簿解析同名工作表 | 分页、分块和读回沿用首次解析的工作表对象 |
| 流式协议 | UTF-8 四字节字符跨缓冲区可能溢出；重复终态重复交付；调用顺序不稳定 | 扩大解码缓冲区，终态仅交付一次，按工具索引合并 |
| 多协议工具往返 | Anthropic 初始参数丢失；并行结果被拆散；Gemini 调用标识丢失 | 保留初始/增量参数、合并同轮结果并保留调用 ID |
| 文本工具块 | 流分段把行内示例误当行首指令；围栏闭合处理与分段有关 | 维护真实行首状态和围栏长度，普通示例保持正文，工具块仅执行一次 |
| 测试服务 | 本机 PortDock 的 `GET /` 探测消耗脚本响应并污染并发统计 | 两个本机测试服务仅统计目标模型 POST 请求，并加入探测隔离回归 |

工作区在任务开始时已存在 Schema 缓存、Word 文本工具历史处理及 `WordRequestShapeTests.cs` 等未提交修改，均保留并纳入验证；不将这些既有修改计作本次新增成果。

## 复现证据

- Word 调用入口首批新增用例在修复前为 9 失败、1 通过；另行复现了面板关闭回调、审批取消竞态和草稿内部类型错误。
- 使用独立的 HEAD 源码副本构建原始程序集，将新增 Excel/协议回归运行于原始程序集：31 失败、10 通过。修复后的同组 41 项全部通过。
- Word 工具成员记录的原始模拟回归为 27 失败、5 通过；集成后 32 项全部通过。
- 真实 Word 测试发现并纠正了模拟宿主未覆盖的折叠光标 null 文本，以及表格替换残留末字的问题。
- 首次完整测试 928 通过、11 失败，失败集中在本机网络重试与并发统计。独立原始 HEAD 在同环境也出现失败；记录请求头确认额外请求来自 PortDock。修正测试服务后，相关 51 项全部通过，最终完整测试全部通过。

## 最终验证

| 检查 | 结果 |
| --- | --- |
| Release 解决方案构建 | 0 警告、0 错误 |
| 默认完整 .NET 测试，包含真实 Excel | 947 通过、0 失败 |
| Agent 工具调用专项 | 275 通过、0 失败；相较任务开始的 183 项新增 92 项 |
| Word 工具、调用入口与请求形状 | 123 通过、0 失败 |
| Excel 项目审查专项 | 13 通过、0 失败 |
| WPS 表格项目审查专项 | 13 通过、0 失败 |
| Microsoft Word 真实宿主 | 13 通过、0 失败 |
| WPS Writer 真实宿主 | 13 通过、0 失败 |
| WorkBuddy ACP 本地 fixture | 6 通过、0 失败 |
| Codex 配置 fixture | 17 通过、0 失败 |
| Web 测试 | 37 个文件全部通过 |
| OpenSpec `validate --all --strict` | 22 项通过 |
| 差异与编码 | `git diff --check`、严格 UTF-8 解码及替换字符检查通过 |

不同入口含重复用例，表中数量不能相加。真实宿主测试只使用新建的临时工作簿/文档，结束后关闭。

主要日志位于 `artifacts/bug-audit-20261111/`：`full-tools-final.log`、`agent-tools.log`、`word-tools.log`、`word-smoke.log`、`wps-word-smoke.log`、`excel-audit.log`、`wps-excel-audit.log`、`network-tests-fixed.log`、`original-provider-excel-regressions.log`。基线源码副本和诊断探针也保存在该忽略目录内。

## 复跑

```powershell
dotnet build ChatSheet.sln --configuration Release
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --agent-tool-tests
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --word-tests
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --word-smoke
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --wps-word-smoke
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --wps --project-audit-tests
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --workbuddy-tests
.\tests\ChatSheet.ToolTests\bin\Release\ChatSheet.ToolTests.exe --codex-config-tests
Get-ChildItem tests\web\*.test.mjs | ForEach-Object { node $_.FullName; if ($LASTEXITCODE -ne 0) { throw $_.Name } }
openspec validate --all --strict
```

## 边界

协议层使用录制形状的本地事件与 HTTP/ACP fixture 验证，没有调用真实付费模型。源码和本地 Release 构建已更新，本次未发布发行包、未重新注册或安装 Office 插件。修复恢复既有规范约束，没有新增能力或改变外部接口，因此未创建或归档其他在途 OpenSpec 提案。

协议复核参考：OpenAI 官方 Function calling 文档、Claude 官方 Handle tool calls 文档，以及 Google 官方 `python-genai` 的 `FunctionCall`/`FunctionResponse` 类型定义；抓取副本保存在同一 artifacts 目录。
