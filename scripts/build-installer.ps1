<#
.SYNOPSIS
Builds OfficeCliAgent-Setup-<version>.exe (plus a .sha256 file) under dist\installer.

Steps: pnpm install + build, stage the host and UI (stage-host.ps1), build the VSTO add-ins in
Release (build-vsto.ps1), then compile installer\OfficeCliAgent.iss with Inno Setup 6.
-SkipBuild reuses an existing dist\installer\staging tree.
#>
param(
    [string] $Version,
    [switch] $SkipBuild
)
. (Join-Path $PSScriptRoot 'lib\common.ps1')

$repo = Get-RepoRoot
if (-not $Version) { $Version = Get-RootVersion }
$outDir = Join-Path $repo 'dist\installer'
$staging = Join-Path $outDir 'staging'

$isccCandidates = @(
    (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'),
    (Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe'),
    (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe')
)
$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) {
    throw 'Inno Setup 6 is required: winget install --id JRSoftware.InnoSetup -e'
}

if (-not $SkipBuild) {
    Push-Location $repo
    try {
        Write-Step 'pnpm install'
        Invoke-Checked { pnpm install --frozen-lockfile } 'pnpm install'
        Write-Step 'pnpm build'
        Invoke-Checked { pnpm build } 'pnpm build'
    } finally {
        Pop-Location
    }
    & (Join-Path $PSScriptRoot 'stage-host.ps1') -StagingDir $staging
    & (Join-Path $PSScriptRoot 'build-vsto.ps1') -StagingDir $staging
} elseif (-not (Test-Path (Join-Path $staging 'addin\Word\OfficeCliAgent.Word.vsto'))) {
    throw "-SkipBuild was given but $staging has no staged add-ins"
}

New-Item -ItemType Directory -Force $outDir | Out-Null
$exe = Join-Path $outDir "OfficeCliAgent-Setup-$Version.exe"
if (Test-Path $exe) { Remove-Item $exe -Force }

Write-Step "Compiling installer $Version with $iscc"
Invoke-Checked {
    & $iscc /Qp "/DAppVersion=$Version" "/DStagingDir=$staging" "/DOutputDir=$outDir" "/DRepoRoot=$repo" (Join-Path $repo 'installer\OfficeCliAgent.iss')
} 'ISCC'
if (-not (Test-Path $exe)) { throw "ISCC did not produce $exe" }

$hash = (Get-FileHash $exe -Algorithm SHA256).Hash.ToLowerInvariant()
$shaFile = "$exe.sha256"
"$hash  $(Split-Path $exe -Leaf)`n" | Set-Content $shaFile -NoNewline -Encoding ascii

Write-Step ("Built {0} ({1:N1} MB)" -f $exe, ((Get-Item $exe).Length / 1MB))
$exe
