# windows-release-distribution Specification

## Purpose
Provide Windows Excel users with a versioned ChatSheet distribution that can be installed without building the source, while retaining an auditable link between the published archive, its checksum, and the released source revision.

## Requirements

### Requirement: Versioned prebuilt Windows archive

The project SHALL provide a reproducible command that builds the release configuration and produces `ChatSheet-v<version>-win.zip`, where `<version>` is the add-in project version. The archive SHALL contain one top-level `ChatSheet-v<version>-win` directory with the complete prebuilt add-in payload under `app`, the PowerShell installation scripts under `scripts`, an `INSTALL.md`, a `RELEASE-NOTES.md`, and a manifest of SHA-256 checksums for the packaged files.

#### Scenario: Package a release from a clean source checkout

- **WHEN** a maintainer runs the documented packaging command from a checkout containing the declared add-in version
- **THEN** it creates the versioned ZIP and a same-named external SHA-256 sidecar in the ignored release-artifact directory
- **AND THEN** the ZIP contains the required root directory and a loadable `app/ChatSheet.AddIn.dll`

### Requirement: Install a prebuilt archive without an SDK

The installation command SHALL recognize the archive layout and deploy its `app` payload without invoking `dotnet build` or requiring a .NET SDK. It SHALL preserve the existing source-checkout behavior: a source installation builds by default and validates that copied web assets are current. Installation and uninstallation SHALL state that administrator authorization is required for machine-wide managed COM registration, while Excel add-in registration applies only to the current user.

#### Scenario: Install from an extracted release archive

- **WHEN** a user runs the documented install command from an extracted, complete Windows archive
- **THEN** the script uses the prebuilt `app` payload instead of attempting a source build
- **AND THEN** it continues to register the COM classes and the current user's Excel add-in entry through the existing installation flow

#### Scenario: Install from a source checkout

- **WHEN** a developer runs the install command from a source checkout without `-SkipBuild`
- **THEN** the script builds the release configuration before deploying it
- **AND THEN** a stale web payload is rejected when the developer explicitly requests `-SkipBuild`

### Requirement: A double-clickable installer with numbered choices

The archive and the source checkout SHALL each provide a launcher at their root that a user can start without typing a command, and that opens a menu offering install-or-update, uninstall, diagnose, and exit, each selected by entering its number. The launcher SHALL sit at the same location relative to the scripts directory in both layouts, so one launcher serves both.

The menu SHALL show, before the choices, the currently installed version if any, which layout it is running from, and where the add-in installs to. It SHALL request administrator authorization once when it starts, so that the actions it offers do not each prompt again and the menu survives them; it SHALL remain open after an action completes, and SHALL report an invalid entry and re-offer the choices rather than exiting.

Because a console reads a batch launcher using a non-UTF-8 codepage, the launcher itself SHALL contain no non-ASCII text; all text shown to the user SHALL come from the script it starts.

#### Scenario: User double-clicks the launcher

- **WHEN** a user double-clicks the launcher in an extracted archive
- **THEN** administrator authorization is requested once
- **AND THEN** a menu appears showing the installed version, the layout, the install location, and four numbered choices

#### Scenario: User selects an action by number

- **WHEN** a user enters the number for an action
- **THEN** that action runs and its output appears in the same window
- **AND THEN** the menu is offered again once the action finishes, without a further authorization prompt

#### Scenario: User enters something that is not a choice

- **WHEN** a user enters a value that is not one of the offered numbers
- **THEN** the menu reports it as invalid and offers the choices again

#### Scenario: Launcher runs from either layout

- **WHEN** the launcher runs from a prebuilt archive, and separately from a source checkout
- **THEN** each reports the layout it is running from
- **AND THEN** the action it performs matches that layout's documented install behavior

### Requirement: Release provenance and integrity information

The public release SHALL use the `v<version>` tag that points to the release-preparation commit and SHALL provide the prebuilt ZIP together with its SHA-256 sidecar. Release documentation SHALL identify the exact archive name, target environment, prerequisite runtimes, UAC requirement, absence of code signing, and the difference between archive validation and a full clean-machine installation acceptance test.

#### Scenario: Verify a downloaded public release

- **WHEN** a user downloads the ZIP and SHA-256 sidecar from the public release
- **THEN** the user can calculate the ZIP SHA-256 and compare it to the published sidecar
- **AND THEN** the user can identify the corresponding `v<version>` source tag and installation prerequisites from the release documentation

### Requirement: Accurate public installation guidance

The README and packaged installation document SHALL distinguish a prebuilt archive from a standalone MSI or EXE installer. They SHALL state support for Windows desktop Microsoft Excel only, require .NET Framework 4.8 and Microsoft Edge WebView2 Runtime, identify WPS/Excel Online/macOS as unsupported, and state that the package is not code-signed.

#### Scenario: A first-time user follows the published instructions

- **WHEN** a user reads the Release download instructions
- **THEN** the user is directed to extract the ZIP and run the PowerShell installation command
- **AND THEN** the user is not told that an SDK, Node.js, WPS support, a code-signed installer, or a clean-machine acceptance test is included when it is not

### Requirement: Consistent version and feature documentation

A public release SHALL use one version consistently across project metadata, the source tag, README download links, installation guidance, release notes, archive directory, archive filename, and checksum sidecar. Release notes SHALL summarize the user-visible changes since the previous public release and state the operational, privacy, support, signing, automated-verification, and manual-acceptance boundaries that materially affect those changes.

#### Scenario: Prepare a feature release

- **WHEN** a maintainer prepares a public release containing user-visible capabilities added since the previous version
- **THEN** the project version, tag, documentation links, release asset names, and packaged documents all identify the same version
- **AND THEN** the release notes describe those capabilities and their relevant limits without presenting local tests as clean-machine, configured-provider, or user-workbook acceptance

#### Scenario: Preserve historical release records

- **WHEN** a newer version becomes the current public release
- **THEN** current download guidance points to the newer version
- **AND THEN** historical release notes and previously published tags remain available without being rewritten as current facts

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
