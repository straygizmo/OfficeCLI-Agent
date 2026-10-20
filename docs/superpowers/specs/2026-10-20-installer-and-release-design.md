# Installer and GitHub Release — design

Date: 2026-10-20

## Intent

People who want to use the add-in should be able to download one file from
GitHub Releases, run it, open Word / Excel / PowerPoint, and start chatting.
Today the only way to run the add-in is to clone the repository, build the
Node packages, open three Visual Studio solutions and press F5.

The maintainer should be able to cut a release with one command on their own
Windows PC. CI builds are a later step, not part of this design.

## Goals

- One installer `OfficeCliAgent-Setup-<version>.exe` that installs the Word,
  Excel and PowerPoint add-ins together, without administrator rights.
- The installer carries everything the add-in itself needs: the three VSTO
  add-ins, the built host (`packages/host/dist` plus its production
  `node_modules`, which includes the GitHub Copilot runtime), and the built
  chat UI (`packages/ui/dist`).
- Uninstall through "Apps & features" removes the add-ins and files cleanly.
- One script builds everything, produces the installer, tags the commit and
  publishes a GitHub Release with the installer attached.
- README (English and Japanese) tells users how to install, what they must
  install themselves, and how to sign in to Copilot or set up BYOK.

## Non-goals

- Code signing. There is no certificate. SmartScreen will warn on first run;
  the README documents how to proceed ("More info" → "Run anyway").
- Bundling Node.js, OfficeCLI or the WebView2 runtime. Users install those
  themselves (Node.js 20.19+ is required; OfficeCLI is optional but needed
  for the advanced document tools; WebView2 is normally present).
- Signing in to GitHub Copilot from inside the task pane. Users who use a
  Copilot subscription still sign in once with the standalone Copilot CLI
  (`copilot login`). BYOK users do not need the standalone CLI at all.
  The README recommends installing the standalone CLI with
  `npm install -g @github/copilot`, since Node.js is already required.
  (My call.) The bundled runtime and the standalone CLI share the
  `~/.copilot` auth state; the builder confirms this on the maintainer's PC
  by checking that the host's `getAuthStatus` reports `authType: "user"`
  after a standalone `copilot login`, before the README presents this as the
  primary path.
- GitHub Actions / CI builds. Follow-up once the local flow works.
- Per-machine (all users) installation, MSI, ClickOnce publishing.
- Auto-update.

## Anti-goals (what would make this a failure)

- The add-in works for the maintainer but not on a clean PC because it still
  looks for the repository checkout.
- The installer silently breaks the maintainer's own F5 development setup
  (for example by overwriting `addin.json` with installed paths).
- Uninstall leaves add-in registry entries behind so Office shows a broken
  add-in on every start.
- A release gets published from a dirty tree or without the version being
  bumped, so two releases carry the same version number.

## Constraints

- Windows only. VSTO add-ins target .NET Framework 4.8.
- Building the VSTO projects requires Visual Studio with the "Office/SharePoint
  development" workload (present on the maintainer's PC, found via vswhere).
  `MSBuild` is not on PATH; the script locates it through vswhere.
- VSTO manifests must be signed at build time (`SignManifests=true`). The
  committed `.csproj` files reference the maintainer's test certificate by
  thumbprint and `*_TemporaryKey.pfx` is git-ignored. Release builds therefore
  happen on the maintainer's PC, where that certificate exists. No change to
  the signing setup.
- Repository: `github.com/straygizmo/OfficeCLI-Agent`. `gh` is installed and
  logged in. `pnpm` 9 and Node 22 are installed. Inno Setup is not installed
  yet; the build script must tell the maintainer how to install it
  (`winget install JRSoftware.InnoSetup`) if `ISCC.exe` is missing.
- The host's Copilot runtime lives in the optional dependency
  `@github/copilot-sdk-win32-x64` (about 140 MB on disk). The installer will
  be large (expect roughly 50–80 MB after compression); that is accepted.
- **x64 only.** The SDK picks its runtime package from the architecture of
  the Node that runs the host. The installer ships only `win32-x64`, so the
  README states that x64 Node.js is required. Windows on ARM is out of scope
  for the first release (ARM64 Node would fail to find its runtime package;
  x64 Node under emulation works). The installer's Node check does not
  inspect the architecture. (My call.)

## Decisions

### Installer technology and layout

- **Inno Setup 6**, script at `installer/OfficeCliAgent.iss`.
- Per-user install (`PrivilegesRequired=lowest`), default directory
  `{localappdata}\Programs\OfficeCliAgent`. No UAC prompt.
- Installed layout:

  ```
  OfficeCliAgent\
    addin\Word\        OfficeCliAgent.Word.dll, .dll.manifest, .vsto, Core, WebView2, Newtonsoft, runtimes\
    addin\Excel\       same for Excel
    addin\PowerPoint\  same for PowerPoint
    host\              package.json, dist\, node_modules\   (self-contained, production only)
    ui\                taskpane.html, assets\
    LICENSE, THIRD-PARTY-NOTICES (if easy; otherwise LICENSE only)
  ```

  Each app keeps its own folder because the three builds each ship their own
  copy of Core and the NuGet DLLs and their `.dll.manifest` lists those files.

- Add-in registration is done by the installer through the registry, exactly
  as Visual Studio does for F5 (the `|vstolocal` form). For each app
  (`Word`, `Excel`, `PowerPoint`):

  ```
  HKCU\Software\Microsoft\Office\<App>\Addins\OfficeCliAgent.<App>
    Description   = "OfficeCLI Agent"
    FriendlyName  = "OfficeCLI Agent"
    LoadBehavior  = 3 (DWORD)
    Manifest      = "file:///<install dir with forward slashes>/addin/<App>/OfficeCliAgent.<App>.vsto|vstolocal"
  ```

  The `|vstolocal` form loads the add-in from the local folder without a
  ClickOnce install and without a ClickOnce trust prompt, so no certificate
  needs to be trusted on the user's PC. The builder verifies this on the
  maintainer's PC before relying on it; if a trust prompt does appear, stop
  and report back rather than working around it.

- This is the **same registry key Visual Studio writes for F5**, so on the
  maintainer's PC the installed add-in and the development add-in are
  mutually exclusive by design: installing replaces the F5 registration, the
  next F5 replaces it back, and the two never load side by side (they would
  fight over port 4320 and the shared WebView2 profile). To honour the
  anti-goal about not breaking the F5 setup, the **uninstaller only deletes a
  key whose `Manifest` value points inside the install directory**; if the key
  points elsewhere (an F5 registration) it is left alone. The installer may
  overwrite the key unconditionally. (My call.)

- Uninstall removes the three registry keys (subject to the rule above) and
  the install directory. It keeps the user's settings and logs
  (`%LOCALAPPDATA%\OfficeCliAgent` and `%USERPROFILE%\.officecli-agent`).

- Before installing or uninstalling, the installer checks whether Word, Excel
  or PowerPoint is running (window classes `OpusApp`, `XLMAIN`,
  `PPTFrameClass`) and asks the user to close them. Files in use would make
  the install fail half-way otherwise.

- Prerequisite checks are **warnings, not blockers**:
  - Node.js: look for `node.exe` on PATH and in the Node.js install registry
    key. If missing, show a page/message with the download link and let the
    user continue (they can install Node afterwards; the add-in reports the
    problem when it starts the host).
  - WebView2 runtime: check the EdgeUpdate client registry key; if missing,
    show the download link, let the user continue.
  - OfficeCLI: not checked by the installer. The host already logs a warning
    and the task pane shows that the advanced tools are disabled.

### Add-in changes (OfficeCliAgent.Core)

- `AgentSettings` resolution currently walks up from the add-in DLL looking
  for a repository (`packages/host/dist/cli.js`). Extend it to also recognise
  the installed layout: walking up from the DLL, a directory that contains
  `host\dist\cli.js` and `ui\taskpane.html` yields `HostDir=<dir>\host` and
  `UiDir=<dir>\ui`. Repository layout is checked first at each level, then
  installed layout. The installer writes nothing to `addin.json`.
- Today the defaults file written on first run contains the *resolved*
  absolute `HostDir` / `UiDir`. Change this so the file written on first run
  contains only the fields the user may set, with `HostDir` / `UiDir` left
  null, and resolution happens on every load. Reason: a file pinned to the
  repository path (as on the maintainer's PC) would otherwise make an
  installed copy keep running the repository build, and a file pinned to the
  install path would break after uninstall. Explicit values in `addin.json`
  or the `OFFICECLI_AGENT_*` environment variables still win. (My call.)
- `OfficeCLI Agent` is the user-facing name in the registry `FriendlyName` /
  `Description`, matching the ribbon group label.

### Host packaging

- The release script produces the self-contained `host\` folder with
  `pnpm --filter @officecli-agent/host deploy --prod <staging>/host`, which
  copies `dist`, `package.json` and a flattened production `node_modules`
  including the workspace package `@officecli-agent/protocol`. The builder
  verifies that `node host\dist\cli.js --help` (or an equivalent smoke check)
  runs from the staged folder outside the repository, and that the Copilot
  runtime starts (`@github/copilot-sdk-win32-x64` present under
  `node_modules`).
- No changes to how the host finds its own config (`~/.officecli-agent`).

### Release script

- `scripts/release.ps1` (PowerShell 7), run from the repository root:

  1. Refuse to run if the working tree is dirty or the current branch is not
     `main` (override switch `-AllowDirty` for local testing of the installer
     only, which then skips the tag/release steps).
  2. Read the version from the root `package.json`. Tag is `v<version>`.
     Refuse if the tag already exists locally or on `origin`.
  3. `pnpm install --frozen-lockfile` and `pnpm build`.
  4. Stage the host with `pnpm deploy` and copy `packages/ui/dist` into a
     staging directory under `dist/installer/` (git-ignored).
  5. Build the three VSTO solutions in `Release` configuration with MSBuild
     found through vswhere (`-restore` for NuGet). Copy each `bin\Release`
     into the staging `addin\<App>` folder.
  6. Run `ISCC.exe` on `installer/OfficeCliAgent.iss`, passing the version
     and staging directory as defines. Output
     `dist/installer/OfficeCliAgent-Setup-<version>.exe` and a
     `.sha256` file next to it.
  7. Unless `-NoRelease` is given: create and push the tag, then
     `gh release create v<version> --title "OfficeCLI Agent v<version>"
     --generate-notes` with the exe and the `.sha256` attached. The release is
     published immediately. `-Draft` creates a draft instead. (My call: the
     user asked for a one-command publish; `-Draft` is the escape hatch.)
  8. Print the release URL.

- `scripts/build-installer.ps1` holds steps 3–6 so the installer can be built
  and tested without touching git or GitHub; `release.ps1` calls it.
- Version bump is manual: edit `version` in the root `package.json`
  (`packages/host` and `packages/ui` versions are not used for the installer).
  The VSTO `AssemblyVersion` values stay as they are.

### README changes (English `README.md` and Japanese `README-ja.md`)

- New **Install** section near the top, before "How it works": download the
  latest `OfficeCliAgent-Setup-<version>.exe` from Releases, SmartScreen
  warning and how to proceed, what gets installed and where, how to
  uninstall.
- **Requirements** split into "to use" and "to build":
  - To use: Windows, Microsoft 365 / Office 2021+, WebView2 runtime, Node.js
    20.19+, OfficeCLI (optional but recommended, with the install command),
    and *either* a Copilot subscription + one-time `copilot login` with the
    standalone Copilot CLI *or* BYOK. State explicitly that the Copilot
    runtime itself is bundled, so the standalone CLI is only for signing in.
  - To build: Visual Studio 2022 with the Office/SharePoint workload, pnpm 9,
    Inno Setup 6 (for the installer).
- **BYOK** section: keep the existing content; add a short numbered
  "first run" walkthrough (open the task pane → Settings → BYOK providers →
  Add → pick a preset, paste the key, choose a model → select it as the
  provider) so a new user can find it without reading `config.json` examples.
- New **Releasing** section for maintainers: bump version, run
  `scripts/release.ps1`, what it does, prerequisites.
- `vsto/README.md`: replace the "distribution is ClickOnce, bundling not
  supported" note with a pointer to the installer.

### Verification before publishing (acceptance)

The first release is only published after the builder has:

1. Built the installer on the maintainer's PC.
2. Installed it under a state where the repository build is **not** reachable
   (for example by temporarily renaming `addin.json` and confirming the
   installed `HostDir` / `UiDir` are resolved from the install directory, as
   shown in `addin.log`).
3. Opened Word, Excel and PowerPoint, confirmed the ribbon toggle appears, the
   task pane loads, the host starts from the installed folder and a simple
   request (for example "insert hello") succeeds with the configured provider.
4. Uninstalled and confirmed the registry keys and files are gone and Office
   starts without add-in errors.
5. Re-registered the development (F5) add-ins if they were disturbed, or
   confirmed they still work.
6. Confirmed the Copilot sign-in path as described under Non-goals (standalone
   `copilot login` → bundled runtime authenticated). BYOK is verified with at
   least one provider.

Acceptance on a second, clean PC (no repository, no prior `~/.copilot`,
ideally a non-ASCII Windows user name) would be preferable but is not
available; see below.

## Decided: acceptance environment

No second PC or VM is available. The first release is verified on the
maintainer's PC only, with the repository build made unreachable as described
in step 2, and the release notes say that it has not yet been tested on a
clean machine.

## Left to the builder

- Exact Inno Setup script contents (pages, messages in English with Japanese
  where the installer language is Japanese, icons, compression settings).
- How the Node.js / WebView2 checks are implemented in `[Code]`.
- How `pnpm deploy` is invoked and whether `node_modules` needs pruning of
  non-Windows platform packages (only `win32-x64` is needed; drop others if
  they got installed).
- Whether to include `THIRD-PARTY-NOTICES`.
- MSBuild invocation details (solution vs. project, `/p:Platform`), and how
  `bin\Release` output is collected.
- Structure of the PowerShell scripts, error handling and output.
- Wording of the README sections.
