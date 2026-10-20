# Installer and GitHub Release Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** One `OfficeCliAgent-Setup-<version>.exe` that installs the Word / Excel / PowerPoint add-ins with the bundled host and UI, and one script that builds it and publishes a GitHub Release.

**Architecture:** The add-in (`OfficeCliAgent.Core`) learns to find host and UI next to its own DLL in an installed layout, so the installer never touches `addin.json`. PowerShell scripts stage the host with `pnpm deploy`, build the three VSTO solutions with MSBuild in Release, and feed the staged tree to an Inno Setup script that registers the add-ins through the registry in the same `|vstolocal` form Visual Studio uses for F5. `release.ps1` wraps that and calls `gh release create`.

**Tech Stack:** C# / .NET Framework 4.8 (VSTO), xunit via `dotnet test`, pnpm 9, PowerShell 7, MSBuild (VS 2022 Community via vswhere), Inno Setup 6, GitHub CLI.

**Spec:** `docs/superpowers/specs/2026-10-20-installer-and-release-design.md`

## Global Constraints

- Windows x64 only; VSTO targets .NET Framework 4.8; the host ships only `@github/copilot-sdk-win32-x64`.
- Per-user install, no administrator rights: `PrivilegesRequired=lowest`, default dir `{localappdata}\Programs\OfficeCliAgent`.
- Installed layout: `addin\Word`, `addin\Excel`, `addin\PowerPoint`, `host\` (package.json, dist, node_modules), `ui\` (taskpane.html, assets), `LICENSE`.
- Registry key per app: `HKCU\Software\Microsoft\Office\<App>\Addins\OfficeCliAgent.<App>` with `Description` = `FriendlyName` = `OfficeCLI Agent`, `LoadBehavior` = 3, `Manifest` = `file:///<install dir, forward slashes>/addin/<App>/OfficeCliAgent.<App>.vsto|vstolocal`.
- Uninstall deletes a key only when its `Manifest` value points inside the install directory. User data in `%LOCALAPPDATA%\OfficeCliAgent` and `%USERPROFILE%\.officecli-agent` is kept.
- Installer never writes `addin.json`. Explicit `addin.json` values and `OFFICECLI_AGENT_*` environment variables still override layout detection.
- Prerequisite checks (Node.js, WebView2) warn and let the user continue. Office running (window classes `OpusApp`, `XLMAIN`, `PPTFrameClass`) blocks until closed.
- Version source: `version` in the root `package.json`. Tag `v<version>`. Release title `OfficeCLI Agent v<version>`. Output `dist/installer/OfficeCliAgent-Setup-<version>.exe` plus `.sha256`.
- No code signing. SmartScreen warning is documented, not worked around.
- Commit messages end with `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`.

## Review Focus

1. **Install path with a space or non-ASCII characters** (Japanese user names). The `Manifest` URL must still load. Pinned in Task 5 by installing to a directory containing a space and a Japanese character and opening Word.
2. **Upgrade or uninstall while the host process is still running** (Office closed less than 15 minutes ago). Files under `host\` are locked unless the installer stops that node process first. Pinned in Task 5.
3. **An `addin.json` from an earlier build that pins `HostDir` / `UiDir` to the repository** keeps an installed copy running the repository build. Explicit values win by spec; the acceptance in Task 7 renames the file, and Task 1 stops writing resolved paths so new users never get such a file.
4. **Node.js missing when the task pane opens**: the current error says "Build the repository (pnpm build)", which means nothing to an installed user. Pinned in Task 1 with a message naming Node.js and the download page.
5. **Uninstalling the released build on the maintainer's PC deletes the F5 registration.** Pinned in Task 5: with the key pointing at `bin\Debug`, uninstall leaves it alone.

---

### Task 1: Core resolves host and UI from the installed layout

**Files:**
- Create: `vsto/OfficeCliAgent.Core/InstallLayout.cs`
- Create: `vsto/OfficeCliAgent.Core.Tests/OfficeCliAgent.Core.Tests.csproj`
- Create: `vsto/OfficeCliAgent.Core.Tests/InstallLayoutTests.cs`
- Create: `vsto/OfficeCliAgent.Core.Tests/AgentSettingsTests.cs`
- Modify: `vsto/OfficeCliAgent.Core/AgentSettings.cs` (Load, FindRepoRoot, HostCliPath/UiEntryPath)
- Modify: `vsto/OfficeCliAgent.Core/HostProcessManager.cs:32-35` (messages)
- Modify: `vsto/OfficeCliAgent.Core/AgentTaskPane.cs:36-42` (message)

**Interfaces:**
- Produces: `public static class InstallLayout` with
  `public static InstallLayout.Result Find(string startDir, Func<string, bool> fileExists, int maxLevels = 8)` where
  `public sealed class Result { public string HostDir; public string UiDir; public bool IsRepository; }`, returning `null` when nothing is found.
  `public static string SerializeDefaults(AgentSettings s)` on `AgentSettings` returns the JSON written on first run.

- [ ] **Step 1: Create the test project**

`vsto/OfficeCliAgent.Core.Tests/OfficeCliAgent.Core.Tests.csproj`: SDK-style, `<TargetFramework>net48</TargetFramework>`, `<UseWindowsForms>true</UseWindowsForms>`, `<IsPackable>false</IsPackable>`, PackageReferences `Microsoft.NET.Test.Sdk` 17.12.0, `xunit` 2.9.2, `xunit.runner.visualstudio` 2.8.2, ProjectReference `..\OfficeCliAgent.Core\OfficeCliAgent.Core.csproj`. The project folder name does not match `vsto/**/packages/` or `vsto/**/bin/` so only its `bin`/`obj` are ignored.

Run: `dotnet test vsto/OfficeCliAgent.Core.Tests`
Expected: builds, 0 tests, exit 0.

- [ ] **Step 2: Write failing tests for `InstallLayout.Find`**

`InstallLayoutTests.cs`, using a `HashSet<string>` of existing files (case-insensitive, backslash paths) as `fileExists`:

```csharp
[Fact] public void FindsRepositoryLayoutWalkingUp()
// files: C:\repo\packages\host\dist\cli.js, C:\repo\packages\ui\dist\taskpane.html
// start: C:\repo\vsto\OfficeCliAgent.Word\OfficeCliAgent.Word\bin\Debug
// assert HostDir == C:\repo\packages\host, UiDir == C:\repo\packages\ui\dist, IsRepository == true

[Fact] public void FindsInstalledLayoutWalkingUp()
// files: C:\inst\host\dist\cli.js, C:\inst\ui\taskpane.html
// start: C:\inst\addin\Word
// assert HostDir == C:\inst\host, UiDir == C:\inst\ui, IsRepository == false

[Fact] public void PrefersRepositoryOverInstalledAtSameLevel()
// both layouts present under C:\x; start C:\x\addin\Word; assert IsRepository == true

[Fact] public void ReturnsNullWhenNothingFound()
// no files; start C:\nothing\a\b; assert Find(...) == null

[Fact] public void InstalledLayoutRequiresBothHostAndUi()
// only C:\inst\host\dist\cli.js; start C:\inst\addin\Word; assert null

[Fact] public void StopsAfterMaxLevels()
// files for installed layout at C:\inst; start 10 levels below; maxLevels 8; assert null
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `dotnet test vsto/OfficeCliAgent.Core.Tests`
Expected: compile error, `InstallLayout` not defined.

- [ ] **Step 4: Implement `InstallLayout.Find` in `vsto/OfficeCliAgent.Core/InstallLayout.cs`**

Walk up from `startDir` with `Path.GetDirectoryName`, at most `maxLevels` directories including the start. At each level check repository first (`packages\host\dist\cli.js` present → `HostDir = packages\host`, `UiDir = packages\ui\dist`), then installed (`host\dist\cli.js` **and** `ui\taskpane.html` present).

- [ ] **Step 5: Run tests to verify they pass**

Run: `dotnet test vsto/OfficeCliAgent.Core.Tests`
Expected: 6 passed.

- [ ] **Step 6: Write failing tests for the defaults file**

`AgentSettingsTests.cs`:

```csharp
[Fact] public void DefaultsJsonLeavesDirectoriesUnset()
// var s = new AgentSettings { HostDir = "C:\\resolved\\host", UiDir = "C:\\resolved\\ui" };
// var json = AgentSettings.SerializeDefaults(s);
// var o = JObject.Parse(json);
// assert o["HostDir"].Type == JTokenType.Null, o["UiDir"].Type == JTokenType.Null
// assert o["Port"] == 4320, o["NodePath"] == "node", o["AutoStartHost"] == true, o["HostIdleExitSeconds"] == 900

[Fact] public void DefaultsJsonHasNoDerivedPaths()
// assert json does not contain "HostCliPath" and does not contain "UiEntryPath"
```

- [ ] **Step 7: Run tests to verify they fail**

Run: `dotnet test vsto/OfficeCliAgent.Core.Tests`
Expected: compile error, `SerializeDefaults` not defined.

- [ ] **Step 8: Change `AgentSettings`**

- Add `[JsonIgnore]` to `HostCliPath` and `UiEntryPath`.
- Add `public static string SerializeDefaults(AgentSettings s)`: serialize a copy with `HostDir` and `UiDir` set to null, `Formatting.Indented`.
- In `Load()`: replace `FindRepoRoot()` with `InstallLayout.Find(assemblyDir, File.Exists)`; apply `HostDir` / `UiDir` only when unset; write the first-run file with `SerializeDefaults` **before** or independently of resolution so the resolved paths never land in the file. Log one line `settings: layout <repository|installed|none> host=<..> ui=<..>`.
- Delete `FindRepoRoot`. Update the class summary comment to mention both layouts.

- [ ] **Step 9: Run tests to verify they pass**

Run: `dotnet test vsto/OfficeCliAgent.Core.Tests`
Expected: 8 passed.

- [ ] **Step 10: Replace the developer-only error messages**

- `HostProcessManager.EnsureRunningAsync`: before `Process.Start`, when `NodePath` is `"node"` and no `node.exe` is found on `PATH` (search each `PATH` entry for `node.exe`), throw `FileNotFoundException` with the message
  `Node.js was not found. Install Node.js 20.19 or later (x64) from https://nodejs.org/ and restart Office, or set NodePath in <ConfigPath>.`
  Keep the existing "Host not found" exception but reword to
  `Host not found. Reinstall OfficeCLI Agent, or (for development) run pnpm build and set HostDir in <ConfigPath>.`
- `AgentTaskPane`: reword the UI-not-found text in both languages to say reinstall OfficeCLI Agent, or for development run `pnpm build` / set `UiDir`.

- [ ] **Step 11: Build the three VSTO projects still compile**

Run (PowerShell):
```powershell
$msbuild = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -requires Microsoft.VisualStudio.Workload.Office -find 'MSBuild\**\Bin\MSBuild.exe'
& $msbuild vsto\OfficeCliAgent.Word\OfficeCliAgent.Word.sln -restore -p:Configuration=Debug -p:RestorePackagesConfig=true -v:m
```
Expected: `Build succeeded.` with 0 errors. Then press F5 in Visual Studio for Word (or start Word with the existing Debug registration) and confirm `addin.log` shows `settings: layout repository`.

- [ ] **Step 12: Commit**

```bash
git add vsto/OfficeCliAgent.Core vsto/OfficeCliAgent.Core.Tests
git commit -m "feat(core): resolve host and UI from the installed layout"
```

---

### Task 2: Stage the host with `pnpm deploy`

**Files:**
- Create: `scripts/stage-host.ps1`
- Create: `scripts/lib/common.ps1`

**Interfaces:**
- Produces: `scripts/stage-host.ps1 -StagingDir <path>` creates `<StagingDir>\host` (package.json, dist, node_modules) and `<StagingDir>\ui` (copy of `packages/ui/dist`). Exits non-zero on failure.
- Produces: `scripts/lib/common.ps1` with `Get-RepoRoot`, `Get-RootVersion` (reads `version` from root `package.json`), `Invoke-Checked { scriptblock } "description"` (throws when `$LASTEXITCODE` is non-zero), `Write-Step "text"`.

- [ ] **Step 1: Write `scripts/lib/common.ps1`** with the four functions above. `Get-RepoRoot` is the parent of `scripts/`.

- [ ] **Step 2: Write `scripts/stage-host.ps1`**

Parameters: `-StagingDir` (mandatory). Steps: `Set-StrictMode -Version Latest`, `$ErrorActionPreference = 'Stop'`; remove and recreate `<StagingDir>\host` and `<StagingDir>\ui`; run `pnpm --filter @officecli-agent/host deploy --prod "<StagingDir>\host"` from the repo root; fail unless `host\dist\cli.js`, `host\node_modules\@github\copilot-sdk\package.json` and `host\node_modules\@github\copilot-sdk-win32-x64\prebuilds\win32-x64\copilot-runtime.exe` exist; copy `packages/ui/dist/*` to `<StagingDir>\ui`; fail unless `ui\taskpane.html` exists. (`pnpm deploy` in pnpm 9 copies workspace dependencies, so `@officecli-agent/protocol` arrives without extra work; if it does not, add `--legacy` and note it in the script.)

- [ ] **Step 3: Run it against a built tree**

Run:
```powershell
pnpm install --frozen-lockfile; pnpm build
pwsh scripts/stage-host.ps1 -StagingDir dist/installer/staging
Get-ChildItem dist/installer/staging/host/node_modules/@github
```
Expected: exit 0; `copilot-sdk` and `copilot-sdk-win32-x64` listed, no other platform packages.

- [ ] **Step 4: Smoke-test the staged host outside the repository**

Run:
```powershell
Copy-Item dist/installer/staging/host "$env:TEMP\ocagent-host" -Recurse -Force
$p = Start-Process node -ArgumentList '"' + "$env:TEMP\ocagent-host\dist\cli.js" + '" --port 4399 --no-tls --idle-exit 5' -WorkingDirectory "$env:TEMP\ocagent-host" -RedirectStandardOutput "$env:TEMP\ocagent-host.log" -PassThru
Start-Sleep 20; Stop-Process $p -ErrorAction SilentlyContinue; Get-Content "$env:TEMP\ocagent-host.log"
```
Expected: the log contains `[host] Copilot CLI runtime started` and a line showing the port. Record whether the process exited on its own after the idle timeout; it does not matter for the installer.

- [ ] **Step 5: Commit**

```bash
git add scripts/lib/common.ps1 scripts/stage-host.ps1
git commit -m "build: stage the host and UI for the installer"
```

---

### Task 3: Build the VSTO add-ins in Release and collect them

**Files:**
- Create: `scripts/build-vsto.ps1`

**Interfaces:**
- Produces: `scripts/build-vsto.ps1 -StagingDir <path> [-Configuration Release]` builds the three solutions and fills `<StagingDir>\addin\Word`, `addin\Excel`, `addin\PowerPoint`.

- [ ] **Step 1: Write `scripts/build-vsto.ps1`**

Locate MSBuild with `vswhere.exe -latest -requires Microsoft.VisualStudio.Workload.Office -find 'MSBuild\**\Bin\MSBuild.exe'`; fail with a message naming the Office/SharePoint workload if nothing is found. For each app in `Word, Excel, PowerPoint`: run `MSBuild vsto\OfficeCliAgent.<App>\OfficeCliAgent.<App>.sln -restore -p:Configuration=<Configuration> -p:Platform="Any CPU" -p:RestorePackagesConfig=true -v:m -nologo` via `Invoke-Checked`; remove and recreate `<StagingDir>\addin\<App>`; copy `vsto\OfficeCliAgent.<App>\OfficeCliAgent.<App>\bin\<Configuration>\*` recursively excluding `*.pdb` and `*.xml`; fail unless `OfficeCliAgent.<App>.dll`, `OfficeCliAgent.<App>.dll.manifest`, `OfficeCliAgent.<App>.vsto`, `OfficeCliAgent.Core.dll`, `Microsoft.Web.WebView2.Core.dll`, `Newtonsoft.Json.dll` and `runtimes\` exist in the target.

- [ ] **Step 2: Run it**

Run: `pwsh scripts/build-vsto.ps1 -StagingDir dist/installer/staging`
Expected: three `Build succeeded.` lines, exit 0, and `Get-ChildItem dist/installer/staging/addin/*` lists the three folders with the files above. Confirm with `Select-String '<deployment install="false"' dist/installer/staging/addin/Word/OfficeCliAgent.Word.vsto` that the manifest is the local-load form.

- [ ] **Step 3: Commit**

```bash
git add scripts/build-vsto.ps1
git commit -m "build: release build of the VSTO add-ins into the staging tree"
```

---

### Task 4: Inno Setup script and `build-installer.ps1`

**Files:**
- Create: `installer/OfficeCliAgent.iss`
- Create: `scripts/build-installer.ps1`

**Interfaces:**
- Consumes: `stage-host.ps1`, `build-vsto.ps1` from Tasks 2 and 3.
- Produces: `scripts/build-installer.ps1 [-Version <x.y.z>] [-SkipBuild]` → `dist/installer/OfficeCliAgent-Setup-<version>.exe` and `.sha256`. Returns the exe path. `-SkipBuild` reuses an existing staging tree.
- Produces: `installer/OfficeCliAgent.iss` compiled with `ISCC.exe /DAppVersion=<v> /DStagingDir=<abs path> /DOutputDir=<abs path> installer\OfficeCliAgent.iss`.

- [ ] **Step 1: Install Inno Setup if missing**

Run: `winget install --id JRSoftware.InnoSetup -e --accept-source-agreements --accept-package-agreements` when `${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe` is absent.
Expected: `Test-Path "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"` is `True` and `Languages\Japanese.isl` exists there.

- [ ] **Step 2: Write `installer/OfficeCliAgent.iss`**

`[Setup]`: `AppId={{2F6C1C0E-7B2A-4D6B-9E0C-OFFICECLIAGENT}` (replace with a real GUID generated once with `[guid]::NewGuid()` and keep it fixed forever), `AppName=OfficeCLI Agent`, `AppVersion={#AppVersion}`, `AppPublisher=straygizmo`, `AppPublisherURL=https://github.com/straygizmo/OfficeCLI-Agent`, `DefaultDirName={localappdata}\Programs\OfficeCliAgent`, `DisableProgramGroupPage=yes`, `PrivilegesRequired=lowest`, `ArchitecturesAllowed=x64compatible`, `ArchitecturesInstallIn64BitMode=x64compatible`, `OutputDir={#OutputDir}`, `OutputBaseFilename=OfficeCliAgent-Setup-{#AppVersion}`, `Compression=lzma2/ultra64`, `SolidCompression=yes`, `LZMAUseSeparateProcess=yes`, `UninstallDisplayName=OfficeCLI Agent`, `CloseApplications=no` (Office is checked in code instead), `LicenseFile={#StagingDir}\..\..\..\LICENSE` (or pass `/DRepoRoot`).

`[Languages]`: `english` (`compiler:Default.isl`) and `japanese` (`compiler:Languages\Japanese.isl`).

`[Files]`: `{#StagingDir}\addin\*` → `{app}\addin` recursesubdirs; `{#StagingDir}\host\*` → `{app}\host` recursesubdirs; `{#StagingDir}\ui\*` → `{app}\ui` recursesubdirs; `LICENSE` → `{app}`.

`[Registry]`: for each of `Word`, `Excel`, `PowerPoint`, root `HKCU`, subkey `Software\Microsoft\Office\<App>\Addins\OfficeCliAgent.<App>`: `Description` string `OfficeCLI Agent`, `FriendlyName` string `OfficeCLI Agent`, `LoadBehavior` dword `3`, `Manifest` string `{code:ManifestUrl|<App>}`. No `uninsdeletekey` flag.

`[Code]`:
- `function ManifestUrl(App: String): String;` → `'file:///' + install dir with `\` replaced by `/` + '/addin/' + App + '/OfficeCliAgent.' + App + '.vsto|vstolocal'`.
- `function OfficeRunning: Boolean;` → `FindWindowByClassName('OpusApp') <> 0` or `XLMAIN` or `PPTFrameClass`.
- `function InitializeSetup: Boolean;` loops while `OfficeRunning`, showing a retry/cancel MsgBox `Please close Word, Excel and PowerPoint before continuing.` (Japanese via `CustomMessages`); returns False on cancel. Same loop in `InitializeUninstall`.
- `procedure StopHost;` runs `powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process -Filter \"Name='node.exe'\" | Where-Object { $_.CommandLine -like '*<app dir>\host\dist\cli.js*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"` hidden, waits. Called from `CurStepChanged(ssInstall)` and `CurUninstallStepChanged(usUninstall)` before files are touched.
- `function NodeFound: Boolean;` → `FileSearch('node.exe', GetEnv('PATH')) <> ''` or `RegQueryStringValue(HKCU, 'Software\Node.js', 'InstallPath', ...)` or same under `HKLM`.
- `function WebView2Found: Boolean;` → `RegValueExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv')` or the same key under `HKCU\Software\Microsoft\EdgeUpdate\Clients\...`.
- `procedure CurPageChanged(CurPageID: Integer);` on `wpReady`: if not `NodeFound`, MsgBox warning with `https://nodejs.org/` (x64, 20.19+); if not `WebView2Found`, MsgBox warning with `https://developer.microsoft.com/microsoft-edge/webview2/`. Both continue.
- `procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);` on `usUninstall`: `StopHost`; for each app read `Manifest`; if it starts with `'file:///' + install dir with forward slashes` (case-insensitive), `RegDeleteKeyIncludingSubkeys`; else leave and log.
- `[CustomMessages]` English and Japanese for the close-Office prompt and the two warnings.

- [ ] **Step 3: Write `scripts/build-installer.ps1`**

Parameters `-Version` (default `Get-RootVersion`), `-SkipBuild`. Unless `-SkipBuild`: `pnpm install --frozen-lockfile`, `pnpm build`, `stage-host.ps1`, `build-vsto.ps1` into `dist/installer/staging`. Locate `ISCC.exe` at `${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe`; if missing, fail with `Inno Setup 6 is required: winget install --id JRSoftware.InnoSetup -e`. Run `ISCC.exe /Qp /DAppVersion=<v> /DStagingDir=<abs> /DOutputDir=<abs dist/installer> /DRepoRoot=<abs> installer\OfficeCliAgent.iss`. Write `<exe>.sha256` containing `<lowercase hex>  OfficeCliAgent-Setup-<v>.exe`. Print and return the exe path.

- [ ] **Step 4: Build the installer**

Run: `pwsh scripts/build-installer.ps1 -SkipBuild` (staging from Tasks 2–3 still present)
Expected: `dist/installer/OfficeCliAgent-Setup-0.1.0.exe` and `.sha256` exist; exe size roughly 50–90 MB; `Get-FileHash` matches the `.sha256` content.

- [ ] **Step 5: Commit**

```bash
git add installer/OfficeCliAgent.iss scripts/build-installer.ps1
git commit -m "build: Inno Setup installer for the Word/Excel/PowerPoint add-ins"
```

---

### Task 5: Install, verify, uninstall on the maintainer's PC

**Files:** none (manual verification; fix anything found in the files from Tasks 1–4 and commit the fixes).

- [ ] **Step 1: Hide the development configuration**

Run: `Rename-Item "$env:LOCALAPPDATA\OfficeCliAgent\addin.json" addin.dev.json`; note the current Word/Excel/PowerPoint `Manifest` registry values (`reg query ...\Addins\OfficeCliAgent.<App>`) so they can be restored.

- [ ] **Step 2: Install with the default directory**

Run the exe. Expected: no UAC prompt; SmartScreen may warn ("More info" → "Run anyway"); the Node.js / WebView2 warnings do **not** appear on this PC; install completes; `reg query HKCU\Software\Microsoft\Office\Word\Addins\OfficeCliAgent.Word` shows `Manifest = file:///C:/Users/mtmar/AppData/Local/Programs/OfficeCliAgent/addin/Word/OfficeCliAgent.Word.vsto|vstolocal`.

- [ ] **Step 3: Verify in all three apps**

Open Word, then Excel, then PowerPoint. Expected for each: no ClickOnce / publisher trust prompt; the ribbon group with the Agent toggle appears; the task pane loads; `addin.log` contains `settings: layout installed host=C:\Users\mtmar\AppData\Local\Programs\OfficeCliAgent\host`; `host.log` shows `Copilot CLI runtime started`; a request such as "insert the text hello at the cursor" (Word), "put 1 in A1" (Excel), "add a slide titled hello" (PowerPoint) succeeds with the configured provider. If a trust prompt appears, stop and report; do not work around it.

- [ ] **Step 4: Verify Copilot sign-in path**

With the standalone CLI signed in (`copilot login` done earlier on this PC), confirm the task pane lists Copilot models (the host calls `listModels` at start; it fails without auth). Also confirm one BYOK provider works from the task pane settings. Record both results for the release notes.

- [ ] **Step 5: Upgrade while the host is running (Review Focus 2)**

Close all Office apps (host keeps running for up to 15 minutes; confirm with `Get-Process node`). Run the same installer again. Expected: install completes without a file-in-use error, and the host process from before is gone.

- [ ] **Step 6: Install path with a space and non-ASCII (Review Focus 1)**

Uninstall from "Apps & features", then run the installer choosing `C:\Users\mtmar\AppData\Local\Programs\OfficeCLI エージェント test`. Open Word. Expected: add-in loads, task pane works. Uninstall again.

- [ ] **Step 7: Uninstall leaves F5 registration alone (Review Focus 5)**

Install to the default directory again. Overwrite the Word key's `Manifest` with the Debug value noted in Step 1. Uninstall. Expected: Excel and PowerPoint keys are gone; the Word key still exists with the Debug manifest; `C:\Users\mtmar\AppData\Local\Programs\OfficeCliAgent` is gone; `%LOCALAPPDATA%\OfficeCliAgent` still exists.

- [ ] **Step 8: Restore the development setup**

Rename `addin.dev.json` back to `addin.json`; restore the three Debug `Manifest` values (or press F5 once per project). Open Word and confirm the F5 add-in loads.

- [ ] **Step 9: Commit fixes**

If any step required changes, commit them with a `fix:` message naming what failed.

---

### Task 6: `release.ps1`

**Files:**
- Create: `scripts/release.ps1`
- Modify: `.gitignore` only if `dist/installer` is not already covered (it is: `dist/`).

**Interfaces:**
- Consumes: `scripts/build-installer.ps1` (Task 4), `Get-RootVersion` (Task 2).
- Produces: `scripts/release.ps1 [-Draft] [-NoRelease] [-AllowDirty]`.

- [ ] **Step 1: Write `scripts/release.ps1`**

Order: (1) unless `-AllowDirty`: fail if `git status --porcelain` is non-empty or `git branch --show-current` is not `main`; (2) `$v = Get-RootVersion`, `$tag = "v$v"`; unless `-NoRelease` or `-AllowDirty`: fail if `git tag -l $tag` is non-empty or `git ls-remote --tags origin $tag` returns a line; fail if `gh release view $tag` succeeds; (3) `$exe = build-installer.ps1 -Version $v`; (4) if `-NoRelease` or `-AllowDirty`: print the exe path and stop; (5) `git tag -a $tag -m "OfficeCLI Agent $tag"`, `git push origin $tag`; (6) `gh release create $tag "$exe" "$exe.sha256" --title "OfficeCLI Agent $tag" --generate-notes` plus `--draft` when `-Draft`; (7) print the URL from `gh release view $tag --json url -q .url`. All external calls through `Invoke-Checked`.

- [ ] **Step 2: Dry run**

Run: `pwsh scripts/release.ps1 -NoRelease`
Expected: refuses while the tree is dirty (uncommitted plan work), then after committing, builds and prints the exe path without creating a tag (`git tag -l v0.1.0` empty).

- [ ] **Step 3: Commit**

```bash
git add scripts/release.ps1
git commit -m "build: one-command GitHub release script"
```

---

### Task 7: Documentation

**Files:**
- Modify: `README.md` (Install section before "How it works"; Requirements; BYOK first-run; Releasing)
- Modify: `README-ja.md` (same sections in Japanese)
- Modify: `vsto/README.md` (last bullet of 既知の制約・今後; 設定 section mentions installed layout)

- [ ] **Step 1: Write the English sections**

- **Install** (new, after the intro bullets): download `OfficeCliAgent-Setup-<version>.exe` from the Releases page; SmartScreen: "More info" → "Run anyway" because the installer is not code-signed; installs per user to `%LOCALAPPDATA%\Programs\OfficeCliAgent` and registers the Word / Excel / PowerPoint add-ins; uninstall from Settings → Apps; settings and logs in `%LOCALAPPDATA%\OfficeCliAgent` and `%USERPROFILE%\.officecli-agent` are kept.
- **Requirements** split into **To use** (Windows 10/11 x64, Microsoft 365 / Office 2021+, WebView2 runtime, Node.js 20.19+ x64, OfficeCLI with its install command marked optional-but-recommended, and either a Copilot subscription with one-time `npm install -g @github/copilot` + `copilot login`, or BYOK) and **To build** (existing list plus Inno Setup 6). State that the Copilot runtime is bundled and the standalone CLI is only used to sign in.
- **BYOK**: add a numbered first-run walkthrough before the JSON example (open the task pane → Settings → BYOK providers → Add → preset → key → model → select as provider).
- **Releasing** (new, before Credits): bump `version` in root `package.json`, commit on `main`, run `pwsh scripts/release.ps1` (`-Draft`, `-NoRelease`), what it needs (VS 2022 with Office workload, Inno Setup 6, `gh auth login`), and that `x64` only / not tested on a clean machine for 0.1.0.

- [ ] **Step 2: Write the same in `README-ja.md`** (sections インストール, 前提 split into 使うには / ビルドするには, BYOK の初回手順, リリース手順).

- [ ] **Step 3: Update `vsto/README.md`**: replace the ClickOnce bullet with "配布は `installer/OfficeCliAgent.iss`（Inno Setup）。`scripts/release.ps1` 参照。" and in 設定 add that an installed copy resolves `host\` and `ui\` next to `addin\` automatically.

- [ ] **Step 4: Check links and commands**

Run: `Select-String -Path README.md,README-ja.md -Pattern 'release.ps1|OfficeCliAgent-Setup|copilot login|@github/copilot'`
Expected: each pattern appears in both files.

- [ ] **Step 5: Commit**

```bash
git add README.md README-ja.md vsto/README.md
git commit -m "docs: install, requirements, BYOK first run and releasing"
```

---

### Task 8: Publish v0.1.0

**Files:** none.

- [ ] **Step 1: Confirm the tree is clean on `main` and pushed**

Run: `git status --porcelain; git branch --show-current; git push origin main`
Expected: empty status, `main`, push succeeds.

- [ ] **Step 2: Release**

Run: `pwsh scripts/release.ps1`
Expected: builds, tags `v0.1.0`, publishes; prints `https://github.com/straygizmo/OfficeCLI-Agent/releases/tag/v0.1.0`.

- [ ] **Step 3: Verify the release**

Run: `gh release view v0.1.0 --json assets,isDraft,url -q '{assets: [.assets[].name], draft: .isDraft, url: .url}'`
Expected: assets `OfficeCliAgent-Setup-0.1.0.exe` and `OfficeCliAgent-Setup-0.1.0.exe.sha256`, `draft: false`. Download the exe with `gh release download v0.1.0 -p '*.exe' -D "$env:TEMP\rel"` and confirm its hash matches the `.sha256`.

- [ ] **Step 4: Add the clean-machine note to the release**

Run: `gh release edit v0.1.0 --notes-file -` with the generated notes plus a final paragraph: not yet tested on a machine without the repository; x64 only; Copilot sign-in requires the standalone CLI (`npm install -g @github/copilot`, `copilot login`).
Expected: `gh release view v0.1.0` shows the paragraph.
