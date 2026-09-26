## ADDED Requirements

### Requirement: 文档起草内容写入当前文档

当用户要求撰写、续写或补充面向当前文档的内容，且存在明确的折叠光标目标时，Word 助手 SHALL 使用 Word 写入工具将生成内容写入该目标，并 SHALL 根据工具读回结果报告状态。该写入 SHALL 遵守当前审批设置、文档保护检查和现有撤销能力。

#### Scenario: 在光标处起草内容

- **WHEN** 用户在 Word 面板要求撰写一段面向当前文档的内容，且当前选区是折叠光标
- **THEN** 助手 SHALL 将生成内容插入该光标位置
- **AND THEN** 只有工具读回成功后才可报告文档已更新

#### Scenario: 明确改写选中文字

- **WHEN** 用户明确要求改写当前非空选区
- **THEN** 助手 SHALL 将选区作为唯一写入目标并替换其内容
- **AND THEN** 选区外文档内容 SHALL 保持不变

#### Scenario: 起草请求遇到非空选区

- **WHEN** 用户请求新写内容但当前选区非空，且没有明确要求替换该选区
- **THEN** 助手 SHALL 询问用户是替换选区还是指定其他插入位置
- **AND THEN** 用户明确目标前文档 SHALL 保持不变

#### Scenario: 不面向文档的普通问答

- **WHEN** 用户提出知识问答、建议或其他不要求写入当前文档的请求
- **THEN** 助手 SHALL 只在聊天区回答，不得自动修改文档

#### Scenario: 写入目标或写入结果不可确认

- **WHEN** 当前文档目标不明确、文档受保护或写入读回失败
- **THEN** 助手 SHALL 询问目标或报告具体失败状态
- **AND THEN** 助手 SHALL NOT 声称文档已成功更新

### Requirement: Word 文档操作提供三种审批模式

Word/WPS Writer 写操作 SHALL 支持与 Excel 相同的 `PerWrite`、`PerTurn` 和 `Automatic` 三种审批模式。读操作 SHALL 自动执行。每轮授权 SHALL 限定在当前文档和单一操作类别，并且只在本轮有效；结构操作 SHALL 与格式、普通写入和删除分开授权。面板 SHALL 显示现存授权并允许收回，运行中更改模式 SHALL 对下一次写操作生效。

#### Scenario: 逐项审批

- **WHEN** 当前模式为 `PerWrite` 且没有适用的本轮授权
- **THEN** 每个写操作 SHALL 显示审批卡，读操作 SHALL 自动执行

#### Scenario: 每轮同文档同类确认

- **WHEN** 当前模式为 `PerTurn` 且用户批准当前文档的一项写操作
- **THEN** 本轮同一文档、同一类别的后续操作 SHALL 不再询问
- **AND THEN** 其他类别和其他文档的操作 SHALL 仍按需询问，结构操作 SHALL 单独授权

#### Scenario: 全自动

- **WHEN** 当前模式为 `Automatic`
- **THEN** 写操作 SHALL 直接执行，不显示审批卡
- **AND THEN** 只有工具实际返回快照时才显示撤销入口

#### Scenario: 收回授权及新一轮隔离

- **WHEN** 用户收回本轮授权或开始新一轮任务
- **THEN** 相应授权 SHALL 被清除
- **AND THEN** 后续未授权写操作 SHALL 按当前模式重新处理

#### Scenario: 运行中切换模式

- **WHEN** 用户在任务运行期间切换审批模式
- **THEN** 下一次写操作 SHALL 使用新模式
