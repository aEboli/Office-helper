# model-capability-fallback Specification

## Purpose
Match each request to what the selected model can actually do. The add-in previously
assumed every model could call functions and accept images, so a model lacking either
capability produced a raw provider error, or — worse — silently produced an answer about
a worksheet it had never read. This capability treats "cannot call tools" and "cannot see
images" as conditions to be detected and worked around rather than as user error, and it
holds the fallbacks to one rule: a request that degrades must lose capability, never
safeguards.

## Requirements

### Requirement: Tool capability is a per-model property with three modes

The add-in SHALL track, for each combination of connection and model name, which tool
protocol to use: native function declarations, a text instruction protocol, or no tools at
all. The tracked value SHALL default to native declarations, so a model that works today
keeps behaving exactly as before.

Tracking SHALL be keyed by connection as well as model name, because the same model name
served through a different gateway may support different features. The key SHALL NOT
include the API key.

Detection results SHALL live only for the current panel session and SHALL NOT be written
to the settings file, so that a gateway gaining tool support is picked up without the user
clearing anything.

#### Scenario: A model that supports tools is unaffected

- **WHEN** a model accepts native function declarations
- **THEN** requests carry the tool declarations as before
- **AND THEN** no fallback notice is shown

#### Scenario: The same model name on two connections

- **WHEN** one connection is found not to support tools
- **THEN** the other connection's entry for that same model name is unaffected

### Requirement: A rejected tool declaration switches to the text protocol

When a request carrying tool declarations is rejected with a client error whose message
identifies a tool-related field, the add-in SHALL record that the model needs the text
instruction protocol, SHALL retry the same step under that protocol, and SHALL report the
switch to the user. The turn SHALL NOT fail for this reason alone.

The retry SHALL NOT consume an additional step from the per-turn step budget, since no
progress was made.

#### Scenario: Provider rejects the tools field

- **WHEN** the provider rejects a request because it does not support tools
- **THEN** the step is retried without tool declarations, using the text instruction protocol
- **AND THEN** the user is told the model does not support native tool calls
- **AND THEN** the turn continues rather than failing

#### Scenario: An unrelated client error still fails

- **WHEN** a request is rejected for a reason unrelated to tools, such as an invalid key
- **THEN** the error is reported as before and no protocol switch is recorded

### Requirement: A model that ignores tools but claims it cannot act switches too

When a model under native declarations produces no tool calls on the first step of a turn
and its reply states that it cannot access, see, or modify the workbook, or explicitly
states that it cannot use or does not support tool/function calling, the add-in SHALL
treat that as absent tool capability, SHALL record the text instruction protocol, and
SHALL retry that step. The reply that made this claim SHALL NOT be kept in the
conversation, so the model does not read its own refusal as established fact.

The explicit tool capability path SHALL only match a reply that connects inability or lack
of support to tool/function calling. Generic safety refusals and unrelated statements
SHALL NOT trigger the switch.

This heuristic SHALL be attempted at most once per connection and model, because a genuine
refusal must not cause every later turn to be re-run.

#### Scenario: Model says it cannot access the workbook

- **WHEN** a model given tool declarations replies that it has no access to the spreadsheet, without calling any tool
- **THEN** the step is retried under the text instruction protocol
- **AND THEN** the refusal is not present in the conversation history

#### Scenario: Model explicitly says tool calling is unavailable

- **WHEN** a model given tool declarations replies that it cannot use or does not support tool/function calling, without calling any tool
- **THEN** the step is retried under the text instruction protocol
- **AND THEN** the refusal is not present in the conversation history

#### Scenario: Generic safety refusal does not trigger the heuristic

- **WHEN** a model refuses an unsafe request without claiming that tool/function calling is unavailable
- **THEN** the reply is delivered as the answer and the protocol is not switched

#### Scenario: A refusal for other reasons is not retried repeatedly

- **WHEN** the heuristic has already been tried for this connection and model
- **THEN** a later reply without tool calls is delivered to the user as the answer

### Requirement: The text instruction protocol executes through the normal tool path

Under the text instruction protocol, requests SHALL NOT carry tool declarations. The
system prompt SHALL instead describe the available tools with their names, parameters and
purpose, and SHALL specify a fenced block carrying the tool name and arguments as JSON.

A block parsed out of the model's reply SHALL be executed through the same path as a
native tool call: the same approval policy, the same argument validation, the same limits,
the same undo registration, and the same panel operation card. The text protocol SHALL NOT
provide access to any tool, or bypass any check, that a native tool call does not.

A block naming an unknown tool, or carrying arguments that are not valid JSON, SHALL be
reported back to the model as a failed tool call stating the problem, in the same form as
any other tool failure.

Tool results SHALL be returned to the model in a form the protocol accepts, and SHALL
remain recognisable as tool results so that context compression continues to treat them as
its first target.

#### Scenario: Model emits a well-formed instruction block

- **WHEN** the reply contains a fenced block naming a write tool with valid arguments
- **THEN** the call is subject to the active approval policy before it executes
- **AND THEN** an operation card appears as it would for a native tool call
- **AND THEN** an undo entry is registered as it would be for a native tool call

#### Scenario: Model emits a block naming a tool that does not exist

- **WHEN** the block names an unknown tool
- **THEN** the model receives a failed result identifying the unknown tool
- **AND THEN** the turn continues

#### Scenario: Model emits a block whose arguments are not valid JSON

- **WHEN** the block's arguments cannot be parsed
- **THEN** the model receives a failed result stating that the arguments were not valid JSON
- **AND THEN** truncated arguments are reported as truncation rather than as a syntax mistake

#### Scenario: Compression still targets tool results first

- **WHEN** context reaches the compression threshold in a session using the text protocol
- **THEN** tool results are compressed before other messages are dropped

### Requirement: Instruction blocks do not appear in the conversation transcript

While streaming under the text instruction protocol, the add-in SHALL withhold text from
the transcript once it may be the start of an instruction block, and SHALL discard the
block if it proves to be one. Text that proves not to be an instruction block SHALL be
delivered in full and in its original order.

A fenced block that is left unterminated when the stream ends SHALL be resolved rather
than withheld indefinitely: if it parses as an instruction it SHALL be executed, and
otherwise its text SHALL be delivered.

#### Scenario: A tool block is not shown as message text

- **WHEN** the model emits an instruction block
- **THEN** the transcript shows the surrounding prose and an operation card, not the block's JSON

#### Scenario: An ordinary code block is shown normally

- **WHEN** the model emits a fenced code block that is not an instruction block
- **THEN** the block appears in the transcript unchanged

### Requirement: A model that cannot follow the text protocol degrades to advisory mode

When the model produces neither instruction blocks nor tool calls under the text protocol
across consecutive steps, the add-in SHALL record that no tool protocol works for this
model and SHALL tell the user that the model can only advise.

In advisory mode the system prompt SHALL state that the model cannot read or modify the
workbook and SHALL direct it to answer with formulas, steps and explanations. The prompt
SHALL NOT tell the model that it is connected to the workbook, because a model that cannot
act on that claim tends to report operations it never performed.

#### Scenario: Model never emits a usable instruction block

- **WHEN** the text protocol yields no usable instruction block
- **THEN** the user is told that this model cannot operate the workbook and can only advise
- **AND THEN** later turns for that model omit tool declarations and use the advisory prompt

#### Scenario: Advisory mode does not claim workbook access

- **WHEN** a turn runs in advisory mode
- **THEN** the system prompt does not assert that the model is connected to the workbook

### Requirement: A rejected image marks the model as lacking vision

When a request carrying images is rejected with a client error whose message identifies an
image or multimodal field, the add-in SHALL record that the model cannot accept images and
SHALL handle the turn through a vision fallback rather than failing.

An image rejection SHALL NOT be recorded as a tool problem, and a tool rejection SHALL NOT
be recorded as a vision problem, so one missing capability does not disable the other.

#### Scenario: Provider rejects the image content

- **WHEN** the provider rejects a request because the model does not accept images
- **THEN** the model is recorded as lacking vision
- **AND THEN** the turn continues through the vision fallback

#### Scenario: A tool rejection leaves vision untouched

- **WHEN** a request is rejected over a tool field
- **THEN** the model's recorded vision capability is unchanged

### Requirement: A configured relay model describes images for a model without vision

The user SHALL be able to nominate a vision relay model. When set, and the selected model
cannot accept images, the add-in SHALL send each image to the relay model over the same
connection and SHALL replace the image in the conversation with the returned description,
attributed as a description rather than presented as the user's own words.

The relay request SHALL ask for what matters when reading a spreadsheet screenshot:
visible structure, headers, values and any error text.

Descriptions SHALL be reused for the remainder of the turn rather than requested again on
each step, so a multi-step turn does not pay for the same image repeatedly.

If the relay itself fails, the add-in SHALL fall back to continuing without the images and
SHALL say that the relay failed, rather than failing the turn.

#### Scenario: Relay describes a screenshot

- **WHEN** a model without vision receives a turn carrying an image and a relay model is configured
- **THEN** the image is described by the relay model
- **AND THEN** the main model receives that description as text, marked as a description of an attached image

#### Scenario: Multi-step turn reuses the description

- **WHEN** a turn that used a relayed description takes several steps
- **THEN** the relay is not asked to describe the same image again

#### Scenario: Relay model fails

- **WHEN** the relay request fails
- **THEN** the turn continues without the images and the failure is reported

### Requirement: Without a relay, the turn continues without the images

When the selected model cannot accept images and no relay model is configured, the add-in
SHALL retry the turn with the images removed, SHALL tell the model that the user attached
images it cannot see, and SHALL tell the user that this model cannot read images together
with what they can do about it.

Images SHALL NOT be dropped silently, because a user who believes the model examined their
screenshot will trust an answer that was never based on it.

#### Scenario: No relay configured

- **WHEN** a model without vision receives an image and no relay is configured
- **THEN** the turn proceeds with the images removed
- **AND THEN** the model is told that images were attached but are not visible to it
- **AND THEN** the user is told the model cannot read images and how to proceed

### Requirement: Both capabilities can be set by hand

The settings page SHALL let the user choose the tool protocol — automatic detection,
native, text instructions, or none — and SHALL let the user name a vision relay model.
An explicit tool protocol choice SHALL suppress detection for that model.

Manual selection is required because a provider that silently ignores tools or images
gives detection nothing to react to, while the user may already know the model's limits.

#### Scenario: User forces the text protocol

- **WHEN** the user selects the text instruction protocol
- **THEN** requests omit tool declarations from the first step, without waiting for a failure

#### Scenario: User leaves detection automatic

- **WHEN** the tool protocol is left on automatic
- **THEN** native declarations are used until a failure or refusal is detected

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
