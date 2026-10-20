<#
.SYNOPSIS
Builds the installer and publishes a GitHub Release for the version in the root package.json.

  pwsh scripts/release.ps1              build, tag v<version>, push the tag, publish the release
  pwsh scripts/release.ps1 -Draft       same, but the release is created as a draft
  pwsh scripts/release.ps1 -NoRelease   build the installer only (no tag, no release)
  pwsh scripts/release.ps1 -AllowDirty  build from a dirty tree / another branch; implies -NoRelease

Requires: a clean tree on main that matches origin/main, the tag not existing yet, Visual Studio
2022 with the Office workload, Inno Setup 6, pnpm, and `gh auth login`.
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
        # Fail fast on everything that would otherwise surface after minutes of building.
        Invoke-Checked { gh auth status } 'gh auth status (run: gh auth login)'
        Invoke-Checked { git fetch origin main --quiet } 'git fetch origin main'
        $local = git rev-parse HEAD
        $remote = git rev-parse origin/main
        if ($local -ne $remote) { throw "local main ($local) does not match origin/main ($remote); push or pull first" }
        if (git tag -l $tag) { throw "tag $tag already exists locally; bump version in package.json" }
        if (git ls-remote --tags origin $tag) { throw "tag $tag already exists on origin; bump version in package.json" }
        gh release view $tag *> $null
        if ($LASTEXITCODE -eq 0) { throw "release $tag already exists on GitHub" }
    }

    $exe = [string] (& (Join-Path $PSScriptRoot 'build-installer.ps1') -Version $version | Select-Object -Last 1)
    if (-not $exe -or -not (Test-Path $exe)) { throw 'build-installer.ps1 did not return an installer path' }

    if ($NoRelease) {
        Write-Step "Installer ready (no release): $exe"
        return
    }

    Write-Step "Tagging $tag"
    Invoke-Checked { git tag -a $tag -m "OfficeCLI Agent $tag" } 'git tag'
    Invoke-Checked { git push origin $tag } 'git push tag'

    Write-Step "Creating GitHub release $tag"
    $ghArgs = @('release', 'create', $tag, $exe, "$exe.sha256", '--title', "OfficeCLI Agent $tag", '--generate-notes')
    if ($Draft) { $ghArgs += '--draft' }
    try {
        Invoke-Checked { gh @args } 'gh release create' -ArgumentList $ghArgs
    } catch {
        Write-Host "The tag $tag is already pushed. Either retry the release by hand:" -ForegroundColor Yellow
        Write-Host "  gh $($ghArgs -join ' ')" -ForegroundColor Yellow
        Write-Host "or remove the tag and run this script again:" -ForegroundColor Yellow
        Write-Host "  git push --delete origin $tag; git tag -d $tag" -ForegroundColor Yellow
        throw
    }

    $url = gh release view $tag --json url -q .url
    Write-Step "Release: $url"
    $url
} finally {
    Pop-Location
}
