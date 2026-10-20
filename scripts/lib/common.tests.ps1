# Minimal self-checks for scripts/lib/common.ps1. Run: pwsh scripts/lib/common.tests.ps1
. (Join-Path $PSScriptRoot 'common.ps1')
$failed = 0
function Check([string] $Name, [bool] $Ok) {
    if ($Ok) { Write-Host "ok   $Name" } else { Write-Host "FAIL $Name" -ForegroundColor Red; $script:failed++ }
}

# Invoke-Checked must pass -ArgumentList through to the script block as $args.
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("ic-" + [guid]::NewGuid().ToString('N') + '.txt')
$env:IC_TEST_OUT = $tmp
Invoke-Checked { cmd /c "echo ARGS:$($args -join ' ') > `"%IC_TEST_OUT%`"" } 'echo' -ArgumentList @('v9.9.9', 'file.exe')
$line = (Get-Content $tmp -Raw).Trim()
Remove-Item $tmp -Force
Check 'Invoke-Checked forwards -ArgumentList as $args' ($line -eq 'ARGS:v9.9.9 file.exe')

# Invoke-Checked must throw on a non-zero exit code.
$threw = $false
try { Invoke-Checked { cmd /c exit 3 } 'exit 3' } catch { $threw = $_.Exception.Message -like '*exit 3 failed (exit code 3)*' }
Check 'Invoke-Checked throws with the exit code' $threw

# Invoke-Checked must not leak native output into the caller's pipeline (it goes to the host).
$captured = Invoke-Checked { cmd /c echo noise } 'echo'
Check 'Invoke-Checked returns nothing to the pipeline' ($null -eq $captured)

if ($failed) { exit 1 }
