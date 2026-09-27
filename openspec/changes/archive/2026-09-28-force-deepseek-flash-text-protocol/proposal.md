# 预先兼容 DeepSeek Flash 的文本工具协议

## Why

部分网关提供的 DeepSeek V4 Flash/Flash 变体会接受请求，却不会产生原生工具调用，直接回复没有工具调用能力。当前自动探测要先发送一次带 `tools` 的请求，再依靠回复触发回退；这会浪费一次请求，也会让用户先看到模型拒绝。

## What Changes

- 在工具协议保持“自动探测”且模型 ID 明确匹配 DeepSeek Flash/V4 Flash 变体时，首轮直接使用已有文本指令协议。
- 文本协议继续走现有工具解析、审批、参数校验、执行和撤销链路。
- 用户手动选择“原生函数调用”时保持原有行为，不被型号启发式覆盖；用户也可以手动选择“文本指令”或“不用工具”。
- 其他模型继续按现有自动探测逻辑处理，避免把同名模型的不同网关行为混为一谈。

## Capabilities

### Modified Capabilities

- `model-capability-fallback`：增加 DeepSeek Flash 自动模式的首轮文本协议兼容规则。

## Impact

涉及模型能力模式解析、对应单测、mock/Excel 端到端验证和增量规范。不会新增工具权限或旁路执行能力。
