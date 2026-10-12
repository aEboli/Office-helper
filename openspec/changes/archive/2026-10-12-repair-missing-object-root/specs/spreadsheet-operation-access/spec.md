## ADDED Requirements

### Requirement: 缺少对象根入口时由用户补全

当 `excel_object` 调用缺少必需的 `root` 时，系统 SHALL 在对应工具卡片上暂停并要求用户选择 `workbook`、`worksheet` 或 `window`。系统 SHALL 将选择写回原调用参数，并以相同调用继续既有校验、审批和执行流程；不得推断该值或执行第二次调用。

#### Scenario: 用户补全对象根入口

- **WHEN** 模型发起缺少 `root` 的 `excel_object` 调用
- **THEN** 对应工具卡显示三个允许的对象入口和取消选项
- **WHEN** 用户选择一个允许的对象入口
- **THEN** 系统在原工具卡中更新参数并继续原调用
- **AND THEN** 原有参数校验、审批策略和执行边界保持生效

#### Scenario: 用户取消对象根入口选择

- **WHEN** 用户显式取消
- **THEN** 该 `excel_object` 调用不执行，并返回明确失败结果

#### Scenario: 停止或关闭面板时取消等待

- **WHEN** 用户在选择期间停止任务或关闭面板
- **THEN** 待处理调用不执行，等待随任务取消或面板关闭而收束

#### Scenario: 其他缺失参数

- **WHEN** `excel_object` 或其他表格工具缺少 `root` 以外的必需参数
- **THEN** 系统继续使用既有参数错误处理，不弹出对象入口选择
