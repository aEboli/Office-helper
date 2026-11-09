## ADDED Requirements

### Requirement: One command verifies the whole project

The project SHALL provide a single command that builds the solution, runs every .NET test entry point and runs every web test, and that fails as a whole when any step fails.

The command SHALL report a summary stating how many steps passed, failed and were skipped, SHALL list each failure with the location of its log, and SHALL count web test files from the directory rather than from a hard-coded number.

Steps that require a real Office host SHALL be skippable through an explicit switch, and a skipped step SHALL be reported as skipped rather than as passed.

#### Scenario: A step fails

- **WHEN** any build step, test entry point or web test fails
- **THEN** the command exits non-zero
- **AND THEN** the summary lists that step and points to its log

#### Scenario: Host tests skipped

- **WHEN** the command is run with the skip-host switch
- **THEN** host-dependent steps are reported as skipped
- **AND THEN** the remaining steps still run and the exit code reflects only them

### Requirement: Repository-root runtime residue stays out of version control

Runtime residue created at the repository root by child processes — directories named after unexpanded environment-variable literals, lock files written through relative paths, and files created by output redirection to a device name — SHALL be excluded from version control by explicit ignore rules, each accompanied by the reason it appears.

#### Scenario: Residue is created again

- **WHEN** a child process recreates such residue at the repository root
- **THEN** version control does not report it
- **AND THEN** a file that looks like source is still reported normally
