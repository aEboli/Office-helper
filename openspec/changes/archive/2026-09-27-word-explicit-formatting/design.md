# 设计

## 触发边界

- 只有用户明确提到文种、格式或标准时才套用相应规范；仅有中文或英文内容、或笼统要求“排版好看”不构成触发条件。
- 对 block 与 modified block 等会改变布局的变体，用户未指定时先询问。
- 中文一般书信按被点名的中文惯例处理：称呼另起一行顶格，正文段首常空两字，祝颂语、署名和日期放在文末，落款常靠右；一般书信通常没有居中标题。该惯例描述常见写法，不宣称是适用于所有私人信件的强制国家标准；党政公文格式不用于普通私人书信。
- 英文商务信函按用户点名的版式处理。full block 各部分左对齐、正文段落不缩进并用空行分隔；modified block 的正文和地址左对齐，日期及结尾签名移至中间位置，正文不缩进。中文称谓可用“尊敬的/亲爱的”等关系和礼貌标记；英文商务称呼通常用 Dear 加正确称谓或姓名，首段自然说明目的，不逐词翻译中文问候和客套语。

## Word 格式执行

- 用户要求起草并格式化时，先写入目标，再读取段落范围，使用 `format_paragraph` 对新增内容涉及的段落逐项设置对齐、间距和首行缩进。
- `first_line_indent_chars` 映射到 Word `ParagraphFormat.CharacterUnitFirstLineIndent`，单位为字符；负值可表示悬挂缩进。写入后读回实际值，不一致时报错。
- 每次格式操作只针对本轮新增或用户明确指定的段落；无法将目标段落与相邻既有内容分开时先询问。

## 在线参考

- Purdue OWL 的基础商务信函指南区分 block、modified block 等格式，适用范围是商务信函，不泛化到所有英文信件：<https://owl.purdue.edu/owl/subject_specific_writing/professional_technical_writing/basic_business_letters/index.html>
- Microsoft Learn 说明 Word 的 `CharacterUnitFirstLineIndent` 以字符为单位设置首行或悬挂缩进：<https://learn.microsoft.com/en-us/office/vba/api/word.paragraphformat.characterunitfirstlineindent>
- 中文一般私人书信缺少适用于所有场景的统一强制版式依据；因此仅在用户点名时采用明确说明为惯例的写法。
