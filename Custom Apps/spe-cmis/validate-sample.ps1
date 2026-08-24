param(
    [switch]$SkipInstall,
    [switch]$SkipTests,
    [switch]$KeepProcesses,
    [int]$TimeoutSec = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot '../../Tools/powershell/SampleValidation.ps1')

$appRoot = $PSScriptRoot
$localSettingsPath = Join-Path $appRoot 'local.settings.json'
$nodeEnvironment = Get-ValidationNodeEnvironment
$runtimeHandle = $null
$runtimeSkipReasons = @()

try {
    Write-Step 'Preflight checks'
    Assert-CommandExists 'node'
    Assert-CommandExists 'npm'

    if (-not $SkipInstall) {
        Write-Step 'Installing dependencies'
        Invoke-ExternalCommand -FilePath 'npm' -Arguments @('install') -WorkingDirectory $appRoot -Environment $nodeEnvironment
    }

    Write-Step 'Building function app'
    Invoke-ExternalCommand -FilePath 'npm' -Arguments @('run', 'build') -WorkingDirectory $appRoot -Environment $nodeEnvironment

    Write-Host 'No automated test script is defined for this sample.' -ForegroundColor Yellow

    $hasFunctionConfig = Test-Path $localSettingsPath
    $hasFunctionHost = $null -ne (Get-Command 'func' -ErrorAction SilentlyContinue)
    if (-not $hasFunctionConfig) {
        $runtimeSkipReasons += 'local.settings.json is missing (copy local.settings.json.example and fill in your Entra app values to enable the runtime smoke check)'
    }
    if (-not $hasFunctionHost) {
        $runtimeSkipReasons += 'Azure Functions Core Tools (func) is not installed'
    }

    if ($hasFunctionConfig -and $hasFunctionHost) {
        Write-Step 'Starting function host'
        $logPath = New-ValidationLogPath -WorkingDirectory $appRoot -Name 'spe-cmis-func'
        $runtimeHandle = Start-LoggedProcess -FilePath 'npx' -Arguments @('func', 'start') -WorkingDirectory $appRoot -LogPath $logPath -Environment $nodeEnvironment

        # A placeholder containerTypeId is enough to prove the CMIS getRepositories route is
        # wired up end to end; without real tenant credentials the call is expected to fail
        # auth/lookup (401/403/500) rather than return CMIS repository data.
        $probeUrl = 'http://127.0.0.1:7071/api/storage/fileStorage/containerTypes/00000000-0000-0000-0000-000000000000/cmis/browser'
        [void](Wait-ForHttpEndpoint -Url $probeUrl -TimeoutSec $TimeoutSec -AllowedStatusCodes @(200, 401, 403, 500) -ProcessHandle $runtimeHandle)

        Write-Step 'Capturing HTTP validation artifact'
        $artifactPath = New-ValidationArtifactPath -WorkingDirectory $appRoot -Kind 'http' -Name 'get-repositories' -Extension 'http.txt'
        Save-HttpArtifact -ArtifactPath $artifactPath -Url $probeUrl -Method 'GET' -AllowedStatusCodes @(200, 401, 403, 500) | Out-Null

        Write-ValidationSummary -Status 'PASS' -Message 'Build succeeded and the function host responded to a CMIS getRepositories request.'
    }
    else {
        Write-ValidationSummary -Status 'SKIP_CONFIG' -Message "Build succeeded; runtime smoke skipped because $($runtimeSkipReasons -join '; ')."
    }

    Write-Host 'spe-cmis sample validation completed.' -ForegroundColor Green
}
catch {
    Write-ValidationSummary -Status 'FAIL' -Message $_.Exception.Message
    throw
}
finally {
    if (-not $KeepProcesses -and $null -ne $runtimeHandle) {
        Stop-LoggedProcess -Handle $runtimeHandle
    }
}
