## ADDED Requirements

### Requirement: Version strings stay consistent across metadata, panel and tests

The version SHALL appear in the add-in project metadata, in the panel's declared fallback version and its text, and in the test that asserts that fallback. A change to any one of them SHALL be accompanied by the others in the same change, so that no test or document can disagree with the shipped build.

#### Scenario: Version is bumped

- **WHEN** the project version is increased
- **THEN** the panel's declared and displayed fallback version matches it
- **AND THEN** the version test asserts the new value and passes

### Requirement: Panel assertions read the panel's own claims

A panel test that checks the displayed version SHALL compare it against the version the page itself declares, and SHALL NOT hard-code a version constant, so that the check verifies page consistency rather than which build happened to be installed.

#### Scenario: A newer version is installed

- **WHEN** a build with a different version is installed and the panel test runs
- **THEN** the version assertion still passes provided the label matches the page's declared version

### Requirement: Release notes state measured verification

Release notes SHALL state verification results as measured, naming the entry point and the pass and total counts, and SHALL NOT leave unmeasured placeholders or carry forward a previous release's numbers.

#### Scenario: Release prepared

- **WHEN** a release is prepared
- **THEN** its notes report the counts actually observed for each verification entry point
- **AND THEN** the packaged notes match the notes in the repository
