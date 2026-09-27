## ADDED Requirements

### Requirement: An explicit tool capability refusal switches to the text protocol

When automatic detection is enabled and a model receives native tool declarations but
returns no tool calls while explicitly stating that it cannot use or does not support
tool/function calling, the add-in SHALL switch to the text instruction protocol and retry
the same step. The refusal SHALL NOT be retained in the conversation history.

This detection SHALL only treat a reply as a tool capability refusal when it explicitly
connects inability or lack of support to tool/function calling. Generic safety refusals
and unrelated statements SHALL NOT trigger the switch. The existing per-connection and
model one-time probe limit SHALL continue to apply.

#### Scenario: Chinese reply says tool calling is unsupported

- **WHEN** a model emits no native tool call and replies that it does not support or cannot perform tool calling
- **THEN** the add-in switches to the text instruction protocol and retries the same step
- **AND THEN** the refusal is omitted from conversation history

#### Scenario: English reply says it cannot call tools

- **WHEN** a model emits no native tool call and explicitly says it cannot call or invoke tools
- **THEN** the add-in switches to the text instruction protocol and retries the same step

#### Scenario: A generic safety refusal remains a normal answer

- **WHEN** a model emits no native tool call and refuses an unsafe request without claiming that tool/function calling is unavailable
- **THEN** the add-in does not switch protocols because of that refusal

#### Scenario: Manual protocol selection remains authoritative

- **WHEN** the user selected a tool protocol other than automatic detection
- **THEN** a refusal does not rewrite the saved capability mode or trigger automatic fallback
