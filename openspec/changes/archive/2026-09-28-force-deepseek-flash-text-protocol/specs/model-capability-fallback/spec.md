## ADDED Requirements

### Requirement: Automatic mode preselects text protocol for DeepSeek Flash variants

When the tool protocol preference is automatic and the selected model identifier clearly
denotes a DeepSeek Flash or V4 Flash variant, the add-in SHALL start the turn in the text
instruction protocol without sending native tool declarations first.

The match SHALL be case-insensitive and SHALL tolerate provider prefixes and suffixes such
as `deepseek/deepseek-v4-flash-vision-preview`, while not matching unrelated models that
merely contain the word `flash`. The rule SHALL be based on the model identifier only as a
conservative compatibility hint; an explicit user protocol choice SHALL take precedence.

#### Scenario: DeepSeek V4 Flash uses text protocol immediately

- **WHEN** automatic mode is selected and the model is `deepseek-v4-flash`
- **THEN** the first request omits native tool declarations and uses the text instruction prompt

#### Scenario: Provider-prefixed Flash variant uses text protocol

- **WHEN** automatic mode is selected and the model is `deepseek/deepseek-v4-flash-vision-preview`
- **THEN** the first request omits native tool declarations and uses the text instruction prompt

#### Scenario: Manual native selection remains authoritative

- **WHEN** the user explicitly selects native function calling for a DeepSeek Flash model
- **THEN** the add-in starts with native tool declarations and does not apply the automatic hint

#### Scenario: Unrelated Flash model is not forced to text

- **WHEN** automatic mode is selected and the model identifier is `some-flash-model`
- **THEN** the add-in retains the normal native-first automatic behavior
