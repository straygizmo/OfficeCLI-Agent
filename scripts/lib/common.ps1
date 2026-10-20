# Shared helpers for the build and release scripts. Dot-source this file.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-RepoRoot {
    (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}

function Get-RootVersion {
    $pkg = Get-Content (Join-Path (Get-RepoRoot) 'package.json') -Raw | ConvertFrom-Json
    if (-not $pkg.version) { throw 'version is missing from the root package.json' }
    $pkg.version
}

function Write-Step([string] $Text) {
    Write-Host "==> $Text" -ForegroundColor Cyan
}

function Invoke-Checked([scriptblock] $Command, [string] $Description) {
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed (exit code $LASTEXITCODE)"
    }
}
