## ADDED Requirements

### Requirement: Word 格式规范由用户明确点名触发

Word 助手 SHALL 仅在用户明确提到文种、格式或标准时应用对应的写作与排版规范。助手 SHALL NOT 仅凭正文语言、文档内容或笼统的美化要求推断书信格式或公文标准。用户点名的格式存在会改变布局的未指定变体时，助手 SHALL 先澄清。格式操作 SHALL 仅针对本轮新增内容或用户明确指定的段落。

#### Scenario: 未点名具体格式

- **WHEN** 用户要求撰写或整理内容，但没有提到文种、格式或标准
- **THEN** 助手 SHALL 沿用文档现有结构，不得仅按中文或英文自动套用书信或公文格式

#### Scenario: 明确点名中文书信格式

- **WHEN** 用户明确要求按中文一般书信格式撰写或排版
- **THEN** 助手 SHALL 按中文书信惯例处理称呼、开头、正文、祝颂语、署名和日期，包括称呼另起一行顶格、正文段首常空两字、文末落款
- **AND THEN** 助手 SHALL 将该惯例描述为通行写法，不得声称其为适用于所有私人书信的统一强制标准

#### Scenario: 明确点名英文商务信函版式

- **WHEN** 用户明确要求英文商务信函的 full block 或 modified block 格式
- **THEN** 助手 SHALL 只应用被点名的版式及对应对齐、段落缩进规则
- **AND THEN** full block SHALL 全部左对齐且正文不缩进；modified block SHALL 将日期及结尾签名移至中间位置，并保持正文不缩进、地址及正文左对齐
- **AND THEN** 助手 SHALL 使用符合收件人关系和正式程度的英文称呼，并在首段自然说明目的，不得逐词翻译中文问候和客套语

#### Scenario: 未指定英文商务信函变体

- **WHEN** 用户要求英文商务信函格式，但未说明 block、modified block 等会改变布局的变体
- **THEN** 助手 SHALL 先询问所需变体，不得静默选择并重排段落

#### Scenario: 格式目标涉及已有内容

- **WHEN** 格式化目标无法与相邻既有段落分开
- **THEN** 助手 SHALL 在修改前询问，不得将格式扩展到整篇文档

### Requirement: Word 段落工具支持首行字符缩进

Word `format_paragraph` 工具 SHALL 接受字符单位的 `first_line_indent_chars` 数值参数，并 SHALL 将其写入 Word `ParagraphFormat.CharacterUnitFirstLineIndent`。工具 SHALL 在成功返回前读回并核对该值；宿主不支持该属性或读回不匹配时 SHALL 返回失败状态。

#### Scenario: 设置首行缩进并读回

- **WHEN** 工具收到明确目标和 `first_line_indent_chars` 数值
- **THEN** 工具 SHALL 按字符单位设置段落首行缩进
- **AND THEN** 工具 SHALL 返回与请求相符的实际读回值

#### Scenario: 宿主无法支持或确认缩进

- **WHEN** Word/WPS 不支持首行缩进属性或实际读回值与请求不符
- **THEN** 工具 SHALL 返回明确失败，不得报告格式已成功应用
