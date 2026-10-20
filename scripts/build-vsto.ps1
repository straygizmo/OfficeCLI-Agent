<#
.SYNOPSIS
Builds the Word / Excel / PowerPoint VSTO add-ins and collects them for the installer.

Requires Visual Studio 2022 with the "Office/SharePoint development" workload (found
through vswhere). Output: <StagingDir>\addin\Word, addin\Excel, addin\PowerPoint.
#>
param(
    [Parameter(Mandatory)] [string] $StagingDir,
    [string] $Configuration = 'Release'
)
. (Join-Path $PSScriptRoot 'lib\common.ps1')

$repo = Get-RepoRoot
$StagingDir = [System.IO.Path]::GetFullPath($StagingDir)

$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
if (-not (Test-Path $vswhere)) { throw "vswhere.exe not found at $vswhere. Install Visual Studio 2022." }
$msbuild = & $vswhere -latest -requires Microsoft.VisualStudio.Workload.Office -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
if (-not $msbuild) {
    throw 'MSBuild with the "Office/SharePoint development" workload was not found. Install it with the Visual Studio Installer.'
}
Write-Step "Using $msbuild"

$apps = 'Word', 'Excel', 'PowerPoint'
$required = @(
    'OfficeCliAgent.{0}.dll', 'OfficeCliAgent.{0}.dll.manifest', 'OfficeCliAgent.{0}.vsto',
    'OfficeCliAgent.Core.dll', 'Microsoft.Web.WebView2.Core.dll', 'Newtonsoft.Json.dll', 'runtimes'
)

foreach ($app in $apps) {
    $sln = Join-Path $repo "vsto\OfficeCliAgent.$app\OfficeCliAgent.$app.sln"
    Write-Step "Building $app ($Configuration)"
    Invoke-Checked {
        & $msbuild $sln -restore "-p:Configuration=$Configuration" '-p:Platform=Any CPU' -p:RestorePackagesConfig=true -v:m -nologo
    } "MSBuild $app"

    $bin = Join-Path $repo "vsto\OfficeCliAgent.$app\OfficeCliAgent.$app\bin\$Configuration"
    $target = Join-Path $StagingDir "addin\$app"
    if (Test-Path $target) { Remove-Item $target -Recurse -Force }
    New-Item -ItemType Directory -Force $target | Out-Null
    # robocopy: mirror the build output without debug symbols and XML docs.
    robocopy $bin $target /E /XF *.pdb *.xml /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed for $app (exit $LASTEXITCODE)" }

    foreach ($pattern in $required) {
        $name = $pattern -f $app
        if (-not (Test-Path (Join-Path $target $name))) { throw "addin\$app is missing $name" }
    }

    # The installer trusts the add-in for the current user through the VSTO inclusion list
    # (HKCU\Software\Microsoft\VSTO\Security\Inclusion), which needs the signing key's public
    # part as <RSAKeyValue> XML. Extract it from the signed deployment manifest.
    $key = Get-ManifestPublicKey (Join-Path $target "OfficeCliAgent.$app.vsto")
    Set-Content (Join-Path $target "OfficeCliAgent.$app.publickey.txt") $key -NoNewline -Encoding ascii
}

Write-Step "Add-ins staged under $StagingDir\addin"
