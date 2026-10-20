<#
.SYNOPSIS
Stages the Node host and the chat UI for the installer.

Creates <StagingDir>\host (package.json, dist, production node_modules including the
GitHub Copilot runtime) with `pnpm deploy`, and <StagingDir>\ui (copy of packages/ui/dist).
Run `pnpm install` and `pnpm build` first.
#>
param(
    [Parameter(Mandatory)] [string] $StagingDir
)
. (Join-Path $PSScriptRoot 'lib\common.ps1')

$repo = Get-RepoRoot
$StagingDir = [System.IO.Path]::GetFullPath($StagingDir)
$hostDir = Join-Path $StagingDir 'host'
$uiDir = Join-Path $StagingDir 'ui'

foreach ($d in @($hostDir, $uiDir)) {
    if (Test-Path $d) { Remove-Item $d -Recurse -Force }
}
New-Item -ItemType Directory -Force $StagingDir | Out-Null

Write-Step "Deploying host to $hostDir"
Push-Location $repo
try {
    # hoisted: a flat node_modules with real files (no junctions), which the installer can copy as-is.
    Invoke-Checked { pnpm --filter @officecli-agent/host deploy --prod --config.node-linker=hoisted $hostDir } 'pnpm deploy'
} finally {
    Pop-Location
}

$required = @(
    'dist\cli.js',
    'node_modules\@github\copilot-sdk\package.json',
    'node_modules\@github\copilot-sdk-win32-x64\prebuilds\win32-x64\copilot-runtime.exe',
    'node_modules\@officecli-agent\protocol\package.json'
)
foreach ($rel in $required) {
    if (-not (Test-Path (Join-Path $hostDir $rel))) { throw "staged host is missing $rel" }
}

Write-Step "Copying UI to $uiDir"
New-Item -ItemType Directory -Force $uiDir | Out-Null
Copy-Item (Join-Path $repo 'packages\ui\dist\*') $uiDir -Recurse -Force
if (-not (Test-Path (Join-Path $uiDir 'taskpane.html'))) { throw 'staged UI is missing taskpane.html (run pnpm build)' }

Write-Step "Staged host and UI under $StagingDir"
