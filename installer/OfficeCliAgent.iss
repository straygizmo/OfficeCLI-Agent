; OfficeCLI Agent installer (Inno Setup 6).
; Compiled by scripts\build-installer.ps1 with:
;   ISCC.exe /DAppVersion=<v> /DStagingDir=<abs> /DOutputDir=<abs> /DRepoRoot=<abs> installer\OfficeCliAgent.iss
; Installs per user (no administrator rights) and registers the Word / Excel / PowerPoint
; VSTO add-ins through the registry in the same |vstolocal form Visual Studio uses for F5.

#ifndef AppVersion
  #error AppVersion is required (/DAppVersion=x.y.z)
#endif
#ifndef StagingDir
  #error StagingDir is required (/DStagingDir=<path>)
#endif
#ifndef OutputDir
  #define OutputDir "."
#endif
#ifndef RepoRoot
  #define RepoRoot ".."
#endif

#define AppName "OfficeCLI Agent"
#define AppDirName "OfficeCliAgent"

[Setup]
AppId={{7D3B5E2A-4C1F-4A8E-9B6D-2E5F0C7A9D11}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=straygizmo
AppPublisherURL=https://github.com/straygizmo/OfficeCLI-Agent
AppSupportURL=https://github.com/straygizmo/OfficeCLI-Agent/issues
AppUpdatesURL=https://github.com/straygizmo/OfficeCLI-Agent/releases
VersionInfoVersion={#AppVersion}
DefaultDirName={localappdata}\Programs\{#AppDirName}
DisableProgramGroupPage=yes
DisableDirPage=no
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
LicenseFile={#RepoRoot}\LICENSE
OutputDir={#OutputDir}
OutputBaseFilename=OfficeCliAgent-Setup-{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes
UninstallDisplayName={#AppName}
CloseApplications=no
SetupLogging=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[CustomMessages]
english.CloseOffice=Please close Word, Excel and PowerPoint before continuing.
japanese.CloseOffice=続行する前に Word、Excel、PowerPoint を終了してください。
english.NodeMissing=Node.js was not found on this PC.%n%nOfficeCLI Agent needs Node.js 20.19 or later (x64). Install it from https://nodejs.org/ after this setup finishes, then restart Office.
japanese.NodeMissing=この PC に Node.js が見つかりませんでした。%n%nOfficeCLI Agent には Node.js 20.19 以降 (x64) が必要です。セットアップ完了後に https://nodejs.org/ からインストールし、Office を再起動してください。
english.WebView2Missing=The Microsoft Edge WebView2 Runtime was not found.%n%nThe task pane needs it. Install it from https://developer.microsoft.com/microsoft-edge/webview2/ after this setup finishes.
japanese.WebView2Missing=Microsoft Edge WebView2 ランタイムが見つかりませんでした。%n%nタスクペインの表示に必要です。セットアップ完了後に https://developer.microsoft.com/microsoft-edge/webview2/ からインストールしてください。

[Files]
Source: "{#StagingDir}\addin\*"; DestDir: "{app}\addin"; Flags: recursesubdirs ignoreversion
Source: "{#StagingDir}\host\*"; DestDir: "{app}\host"; Flags: recursesubdirs ignoreversion
Source: "{#StagingDir}\ui\*"; DestDir: "{app}\ui"; Flags: recursesubdirs ignoreversion
Source: "{#RepoRoot}\LICENSE"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; No uninsdeletekey: the uninstaller only removes keys that point inside {app} (see [Code]),
; so a developer's F5 registration of the same add-in is left alone.
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeCliAgent.Word"; ValueType: string; ValueName: "Description"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeCliAgent.Word"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeCliAgent.Word"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\OfficeCliAgent.Word"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:ManifestUrl|Word}"
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\OfficeCliAgent.Excel"; ValueType: string; ValueName: "Description"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\OfficeCliAgent.Excel"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\OfficeCliAgent.Excel"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\OfficeCliAgent.Excel"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:ManifestUrl|Excel}"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\OfficeCliAgent.PowerPoint"; ValueType: string; ValueName: "Description"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\OfficeCliAgent.PowerPoint"; ValueType: string; ValueName: "FriendlyName"; ValueData: "{#AppName}"
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\OfficeCliAgent.PowerPoint"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\OfficeCliAgent.PowerPoint"; ValueType: string; ValueName: "Manifest"; ValueData: "{code:ManifestUrl|PowerPoint}"

[UninstallDelete]
; Only our own subfolders: a user may have pointed the installer at an existing directory.
Type: filesandordirs; Name: "{app}\addin"
Type: filesandordirs; Name: "{app}\host"
Type: filesandordirs; Name: "{app}\ui"

[Code]
const
  AddinKeyFmt = 'Software\Microsoft\Office\%s\Addins\OfficeCliAgent.%s';
  { VSTO inclusion list: trusts a manifest URL + signing public key for the current user, which is
    what Visual Studio does for F5. Without it the runtime shows "publisher cannot be verified". }
  InclusionKeyFmt = 'Software\Microsoft\VSTO\Security\Inclusion\%s';
  InclusionGuidWord = '{3A1C9F4E-6B2D-4E7A-8C5F-1D0B9E7A4C21}';
  InclusionGuidExcel = '{5C8E2B7A-9D14-4F36-B0A2-7E6C3D1F8B42}';
  InclusionGuidPowerPoint = '{8F2D6A1C-3E5B-47C9-A4D8-0B7E9C2F6A63}';
  WebView2ClientId = '{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}';

function ForwardSlashes(const Path: String): String;
begin
  Result := Path;
  StringChangeEx(Result, '\', '/', True);
end;

{ file:///C:/Users/x/AppData/Local/Programs/OfficeCliAgent/addin/Word/OfficeCliAgent.Word.vsto }
function ManifestFileUrl(App: String): String;
begin
  Result := 'file:///' + ForwardSlashes(ExpandConstant('{app}')) + '/addin/' + App + '/OfficeCliAgent.' + App + '.vsto';
end;

{ The registry Manifest value: the file URL plus |vstolocal (load from the local folder). }
function ManifestUrl(App: String): String;
begin
  Result := ManifestFileUrl(App) + '|vstolocal';
end;

function InclusionGuid(App: String): String;
begin
  if App = 'Word' then Result := InclusionGuidWord
  else if App = 'Excel' then Result := InclusionGuidExcel
  else Result := InclusionGuidPowerPoint;
end;

procedure WriteInclusionEntry(App: String);
var
  Key, KeyFile: String;
  PublicKey: AnsiString;
begin
  Key := Format(InclusionKeyFmt, [InclusionGuid(App)]);
  KeyFile := ExpandConstant('{app}') + '\addin\' + App + '\OfficeCliAgent.' + App + '.publickey.txt';
  if not LoadStringFromFile(KeyFile, PublicKey) then
  begin
    Log('Inclusion list: cannot read ' + KeyFile + '; Office will ask to trust the add-in once.');
    Exit;
  end;
  RegWriteStringValue(HKCU, Key, 'Url', ManifestFileUrl(App));
  RegWriteStringValue(HKCU, Key, 'PublicKey', String(PublicKey));
  Log('Inclusion list: ' + Key + ' -> ' + ManifestFileUrl(App));
end;

function OfficeRunning: Boolean;
begin
  Result := (FindWindowByClassName('OpusApp') <> 0)        { Word }
         or (FindWindowByClassName('XLMAIN') <> 0)         { Excel }
         or (FindWindowByClassName('PPTFrameClass') <> 0); { PowerPoint }
end;

{ Retry/Cancel loop until Office is closed; False when the user cancels. }
function WaitForOfficeToClose: Boolean;
begin
  Result := True;
  while OfficeRunning do
  begin
    if MsgBox(CustomMessage('CloseOffice'), mbError, MB_RETRYCANCEL) <> IDRETRY then
    begin
      Result := False;
      Exit;
    end;
  end;
end;

{ A PowerShell single-quoted literal: only the quote itself needs doubling. }
function PsQuoted(const S: String): String;
begin
  Result := S;
  StringChangeEx(Result, '''', '''''', True);
  Result := '''' + Result + '''';
end;

{ Stop a host (node host\dist\cli.js) started from this install directory so its files are not locked.
  The path is matched with IndexOf (no wildcard syntax) so spaces, quotes and brackets are literal. }
procedure StopHost;
var
  Cmd: String;
  ResultCode: Integer;
begin
  Cmd := '-NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -eq ''node.exe'' -and $_.CommandLine -and $_.CommandLine.IndexOf('
       + PsQuoted(ExpandConstant('{app}') + '\host\dist\cli.js') + ', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"';
  Log('StopHost: powershell.exe ' + Cmd);
  Exec('powershell.exe', Cmd, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

function NodeFound: Boolean;
var
  Dummy: String;
begin
  Result := (FileSearch('node.exe', GetEnv('PATH')) <> '')
         or RegQueryStringValue(HKCU, 'Software\Node.js', 'InstallPath', Dummy)
         or RegQueryStringValue(HKLM, 'SOFTWARE\Node.js', 'InstallPath', Dummy);
end;

function WebView2Found: Boolean;
begin
  Result := RegValueExists(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\' + WebView2ClientId, 'pv')
         or RegValueExists(HKLM, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\' + WebView2ClientId, 'pv')
         or RegValueExists(HKCU, 'Software\Microsoft\EdgeUpdate\Clients\' + WebView2ClientId, 'pv');
end;

function InitializeSetup: Boolean;
begin
  Result := WaitForOfficeToClose;
end;

function InitializeUninstall: Boolean;
begin
  Result := WaitForOfficeToClose;
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = wpReady then
  begin
    if not NodeFound then
      MsgBox(CustomMessage('NodeMissing'), mbInformation, MB_OK);
    if not WebView2Found then
      MsgBox(CustomMessage('WebView2Missing'), mbInformation, MB_OK);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
    StopHost
  else if CurStep = ssPostInstall then
  begin
    WriteInclusionEntry('Word');
    WriteInclusionEntry('Excel');
    WriteInclusionEntry('PowerPoint');
  end;
end;

{ Delete Key only when its string value ValueName starts with Prefix (points inside this install). }
procedure DeleteKeyIfOurs(const Key, ValueName, Prefix: String);
var
  Value: String;
begin
  if not RegQueryStringValue(HKCU, Key, ValueName, Value) then
    Exit;
  if CompareText(Copy(Value, 1, Length(Prefix)), Prefix) = 0 then
  begin
    Log('Removing ' + Key);
    RegDeleteKeyIncludingSubkeys(HKCU, Key);
  end
  else
    Log('Keeping ' + Key + ' (points elsewhere: ' + Value + ')');
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Apps: TArrayOfString;
  I: Integer;
  Prefix: String;
begin
  if CurUninstallStep <> usUninstall then
    Exit;
  StopHost;
  Prefix := 'file:///' + ForwardSlashes(ExpandConstant('{app}')) + '/';
  SetArrayLength(Apps, 3);
  Apps[0] := 'Word';
  Apps[1] := 'Excel';
  Apps[2] := 'PowerPoint';
  for I := 0 to 2 do
  begin
    DeleteKeyIfOurs(Format(AddinKeyFmt, [Apps[I], Apps[I]]), 'Manifest', Prefix);
    DeleteKeyIfOurs(Format(InclusionKeyFmt, [InclusionGuid(Apps[I])]), 'Url', Prefix);
  end;
end;
