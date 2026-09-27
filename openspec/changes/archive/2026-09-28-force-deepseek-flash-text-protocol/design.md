## Decision

Add a small model-name compatibility hint beside the existing capability cache. The hint is
consulted only by `ModelCapabilities.ResolveMode` when the preference is `Auto`; `Native`,
`Text`, and `None` remain explicit user choices. Matching is case-insensitive and accepts a
provider prefix before a DeepSeek Flash token, but requires both `deepseek` and `flash` plus
the V4 marker when present, so an unrelated provider's generic flash model is unaffected.

The selected text mode uses the existing `SystemPrompt`, `TextToolGate`, `ToolCall` execution,
approval policy, and undo registration. No special DeepSeek request format is added. If a
particular gateway's text protocol also fails, the existing consecutive-miss logic switches
to advisory mode.

## Validation

Add pure mode-resolution tests for exact, prefixed, case-insensitive, unrelated, and manual
preference cases. Extend the mock provider and end-to-end script with a DeepSeek Flash model
scenario that fails if the first request contains native tools and verifies a text instruction
call writes to Excel. Run the focused tool tests, Release build, strict OpenSpec validation,
and the real Excel/WebView2 scenario.
