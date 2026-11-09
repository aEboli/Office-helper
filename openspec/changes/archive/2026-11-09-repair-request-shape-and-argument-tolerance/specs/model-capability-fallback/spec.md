## ADDED Requirements

### Requirement: Requests never carry unpaired tool traces

The add-in SHALL keep every tool call and its result paired in the outgoing request, and SHALL remove anything that cannot be paired rather than sending it. A tool result that does not answer a call in the immediately preceding assistant message SHALL be dropped, and an assistant call with no matching result SHALL be reduced to its text, for every protocol the add-in supports.

The repair SHALL treat the case where all calls are paired exactly like the case where only some are paired, because the all-paired case is the common path and must not be the one left unguarded.

The repair SHALL NOT mutate the conversation history: operation cards and undo entries are registered against the original objects.

#### Scenario: All calls paired but an extra result present

- **WHEN** an assistant message's calls are all answered and one further tool message belongs to no call
- **THEN** the extra tool message is dropped from the request
- **AND THEN** the paired results are still sent and the history objects are unchanged

#### Scenario: Some calls unpaired

- **WHEN** only some of an assistant message's calls are answered
- **THEN** only the answered calls and their results are sent

### Requirement: Requests without tool declarations carry no native tool traces

When a request carries no tool declarations, the add-in SHALL convert any native tool traces remaining in the history into plain text before sending, so that the request contains no tool-role message, no assistant tool-call list, and no protocol-specific tool block.

The conversion SHALL preserve what the model needs in order to avoid repeating work — the tool name, its arguments and the result — and SHALL keep results recognisable to context compression.

The conversion SHALL NOT claim a tool trace where none exists and SHALL NOT mutate the conversation history.

#### Scenario: Switching from native to the text protocol mid-conversation

- **WHEN** a conversation began under native declarations and later sends requests without them
- **THEN** the request body contains no native tool trace for any supported protocol
- **AND THEN** the model can still read which tools ran and what they returned

#### Scenario: Requests that still declare tools are untouched

- **WHEN** a request carries tool declarations
- **THEN** its native tool calls and results are sent unchanged
