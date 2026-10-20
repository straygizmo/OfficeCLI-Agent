<#
.SYNOPSIS
Builds the installer and publishes a GitHub Release for the version in the root package.json.

  pwsh scripts/release.ps1              build, tag v<version>, push the tag, publish the release
  pwsh scripts/release.ps1 -Draft       same, but the release is created as a draft
  pwsh scripts/release.ps1 -NoRelease   build the installer only (no tag, no release)
  pwsh scripts/release.ps1 -AllowDirty  build from a dirty tree / another branch; implies -NoRelease

Requires: a clean tree on main, the tag not existing yet, Visual Studio 2022 with the Office
workload, Inno Setup 6, pnpm, and `gh auth login`.
#>
param(
    [switch] $Draft,
    [switch] $NoRelease,
    [switch] $AllowDirty
)
. (Join-Path $PSScriptRoot 'lib\common.ps1')

$repo = Get-RepoRoot
Push-Location $repo
try {
    if ($AllowDirty) { $NoRelease = $true }

    if (-not $AllowDirty) {
        $status = git status --porcelain
        if ($status) { throw "working tree is not clean:`n$($status -join "`n")" }
        $branch = git branch --show-current
        if ($branch -ne 'main') { throw "release from main only (current branch: $branch)" }
    }

    $version = Get-RootVersion
    $tag = "v$version"
    Write-Step "Version $version (tag $tag)"

    if (-not $NoRelease) {
        if (git tag -l $tag) { throw "tag $tag already exists locally; bump version in package.json" }
        $remoteTag = git ls-remote --tags origin $tag
        if ($remoteTag) { throw "tag $tag already exists on origin; bump version in package.json" }
        $null = gh release view $tag 2>$null
        if ($LASTEXITCODE -eq 0) { throw "release $tag already exists on GitHub" }
    }

    $exe = & (Join-Path $PSScriptRoot 'build-installer.ps1') -Version $version
    if (-not $exe -or -not (Test-Path $exe)) { throw 'build-installer.ps1 did not return an installer path' }
    $exe = [string] ($exe | Select-Object -Last 1)

    if ($NoRelease) {
        Write-Step "Installer ready (no release): $exe"
        return
    }

    Write-Step "Tagging $tag"
    Invoke-Checked { git tag -a $tag -m "OfficeCLI Agent $tag" } 'git tag'
    Invoke-Checked { git push origin $tag } 'git push tag'

    Write-Step "Creating GitHub release $tag"
    $args = @($tag, $exe, "$exe.sha256", '--title', "OfficeCLI Agent $tag", '--generate-notes')
    if ($Draft) { $args += '--draft' }
    Invoke-Checked { gh release create @args } 'gh release create'

    $url = gh release view $tag --json url -q .url
    Write-Step "Release: $url"
    $url
} finally {
    Pop-Location
}
