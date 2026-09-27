# 任务

## 实现

- [x] 增加 DeepSeek Flash/V4 Flash 兼容匹配和自动模式解析。
- [x] 保持手动 Native/Text/None 选择优先。
- [x] 补充单元测试、mock-provider 场景和端到端判定。

## 验证

- [x] 运行工具测试、Release 构建和 `openspec validate force-deepseek-flash-text-protocol --strict`。
- [x] 在真实 Excel + WebView2 中确认首个请求不带原生 tools 且文本指令真实执行。
