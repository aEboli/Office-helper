## ADDED Requirements

### Requirement: Argument validation matches the execution layer

The add-in SHALL validate tool arguments before approval and execution using the same tolerance the tools themselves apply, and SHALL reject only what cannot be read without loss.

A non-required argument whose value is `null` SHALL be treated as not supplied. Numeric and boolean arguments SHALL accept strings that convert losslessly, and SHALL reject strings that do not. Array and object arguments SHALL remain strictly typed. A required argument that is absent, `null`, or blank SHALL still be reported as missing.

A rejected argument SHALL name the argument and state what was received, so that the model can correct it rather than resend an equivalent form.

An argument the tool schema declares as nullable SHALL be declared nullable consistently in the declaration sent to the model, in the text-protocol catalogue, and in validation.

#### Scenario: Optional argument written as null

- **WHEN** an optional boolean argument is sent as `null`
- **THEN** the call proceeds as if the argument were omitted

#### Scenario: Scalar written as a lossless string

- **WHEN** an integer argument is sent as the string `"0"`
- **THEN** the call proceeds using that value

#### Scenario: Scalar written as a lossy string

- **WHEN** an integer argument is sent as `12.5` or as a non-numeric string
- **THEN** the call is rejected as an invalid argument

#### Scenario: Required argument still enforced

- **WHEN** a required argument is absent or `null`
- **THEN** the call is rejected as missing that argument

### Requirement: Object-entry completion does not precede argument validation

For a call whose only missing piece is the object entry, the add-in SHALL validate every non-interactive argument before asking the user to complete the object entry, and SHALL validate the complete arguments again afterwards.

Asking the user first would spend a click on a call that was already invalid.

#### Scenario: Object call missing both the entry and another required argument

- **WHEN** a call omits the object entry and one further required argument
- **THEN** the missing argument is reported without asking the user to choose an object entry

#### Scenario: Object call missing only the entry

- **WHEN** a call omits only the object entry
- **THEN** the user is asked once, and the completed arguments are validated before approval

#### Scenario: Validation failure does not consume the user's click

- **WHEN** a call is rejected before the object-entry prompt
- **THEN** the operation card reports the argument error and no object-entry choice is shown
