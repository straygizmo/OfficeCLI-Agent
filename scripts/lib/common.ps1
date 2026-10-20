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

# <RSAKeyValue><Modulus>…</Modulus><Exponent>…</Exponent></RSAKeyValue> from a signed
# ClickOnce/VSTO manifest, in the exact form the VSTO inclusion list expects (no xmlns).
function Get-ManifestPublicKey([string] $ManifestPath) {
    [xml] $doc = Get-Content $ManifestPath -Raw
    $ns = @{ ds = 'http://www.w3.org/2000/09/xmldsig#' }
    $node = Select-Xml -Xml $doc -Namespace $ns -XPath '//ds:KeyValue/ds:RSAKeyValue' | Select-Object -First 1
    if (-not $node) { throw "$ManifestPath has no RSAKeyValue signature; is the manifest signed?" }
    $modulus = ($node.Node.GetElementsByTagName('Modulus') | Select-Object -First 1).InnerText
    $exponent = ($node.Node.GetElementsByTagName('Exponent') | Select-Object -First 1).InnerText
    "<RSAKeyValue><Modulus>$modulus</Modulus><Exponent>$exponent</Exponent></RSAKeyValue>"
}

# Runs a native command, shows its output on the host (never in the caller's pipeline),
# and throws when it exits non-zero. -ArgumentList is available inside the block as $args.
function Invoke-Checked([scriptblock] $Command, [string] $Description, [object[]] $ArgumentList = @()) {
    & $Command @ArgumentList 2>&1 | Out-Host
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed (exit code $LASTEXITCODE)"
    }
}
