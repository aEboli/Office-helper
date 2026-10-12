# 修复 Excel 对象工具缺少 root 时无法继续

## Why

部分模型会发出缺少必需 `root` 参数的 `excel_object` 调用。当前执行器将其作为普通参数错误返回，整次对象操作无法继续；用户无法补上缺失的对象入口。

## What Changes

- `excel_object` 缺少 `root` 时，在对应工具卡片上暂停并要求用户选择工作簿、工作表或窗口。
- 只接受这三个允许值；选择后把值补到原调用参数，并继续相同调用的既有校验、审批和执行流程。
- 用户取消时不执行该调用并返回明确失败结果；停止或关闭面板时取消等待并显示未执行。
- 其他缺失参数继续使用既有校验和错误处理。

## Capabilities

### Modified Capabilities

- `spreadsheet-operation-access`：补充缺少对象根入口时的显式选择与恢复行为。

## Impact

涉及 `AgentRunner`、面板消息桥、工具操作卡和 `spreadsheet-operation-access` 规范。此变更不推断对象入口，不改变其他工具缺参行为，也不改变审批策略。
