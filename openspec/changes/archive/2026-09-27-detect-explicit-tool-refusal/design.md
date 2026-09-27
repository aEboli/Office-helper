## Context

`CapabilitySignals.LooksLikeToolRefusal` currently detects inability phrases only when the
same reply also mentions a spreadsheet or workbook. This catches statements such as
"I cannot access your spreadsheet" but misses a direct statement such as "I cannot call
tools". `AgentRunner.ShouldProbeToolRefusal` already limits detection to automatic mode,
native mode, zero tool calls, and one probe per connection/model.

## Decision

Extend the refusal signal with a second, narrow path: an explicit inability/support denial
paired with a direct tool/function calling phrase. Keep the existing workbook-access path.
Do not infer capabilities from model names, because the same model identifier can behave
differently behind different connections or gateways.

The switch continues to use the existing text protocol and same-step retry. The initial
refusal is discarded, and parsed text-protocol calls continue through the normal execution
and approval path. If the model cannot use that protocol, the existing advisor fallback
remains responsible for stopping unsupported operations and explaining the limitation.

## Validation

Add positive tests for direct Chinese and English tool-capability denials and negative tests
for generic safety refusals or ordinary replies. Add a targeted agent/mock-provider
regression proving that the refusal is not retained and that a valid text instruction is
processed by the existing tool path. Run strict validation for this change and the focused
capability test suite before implementation is considered complete.
