param(
    [Parameter(Mandatory=$true)]
    [string]$ActorUpn,

    [Parameter(Mandatory=$true)]
    [string]$AdminUpn,

    [Parameter(Mandatory=$true)]
    [string]$HaloId,

    [Parameter(Mandatory=$true, HelpMessage="Format: yyyy-MM-dd")]
    [DateTime]$StartDate,

    [Parameter(Mandatory=$true, HelpMessage="Format: yyyy-MM-dd")]
    [DateTime]$EndDate
)

function Connect-ExchangeEnv {
    param([string]$Upn)
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -UserPrincipalName $Upn
}

function Initialize-OutputDirectory {
    param([string]$RootPath, [string]$Id)
    $baseSessionDir = Join-Path $RootPath "sessiondata"
    if (-not (Test-Path $baseSessionDir)) { New-Item -ItemType Directory -Path $baseSessionDir | Out-Null }

    $outputDir = Join-Path $baseSessionDir $Id
    if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
    return $outputDir
}

function Get-AuditLogs {
    param(
        [DateTime]$Start,
        [DateTime]$End,
        [string]$TargetUser,
        [string]$Destination
    )
    $accumulatedRecords = @()
    for ($d = $Start; $d -lt $End; $d = $d.AddDays(1)) {
        $sliceStart = $d
        $sliceEnd = $d.AddDays(1)
        if ($sliceEnd -gt $End) { $sliceEnd = $End }
        
        $sessionId = [guid]::NewGuid().Guid
        $records = Search-UnifiedAuditLog -StartDate $sliceStart -EndDate $sliceEnd -UserIds $TargetUser -ResultSize 5000 -SessionId $sessionId -HighCompleteness
        
        if ($records) { 
            $accumulatedRecords += $records 
            # Export each day's results to disk in JSON format
            $jsonPath = Join-Path $Destination "AuditLog-$($sliceStart.ToString('yyyyMMdd')).json"
            $records | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8
        }
    }
    return $accumulatedRecords
}

function Export-CombinedLogs {
    param($Records, $Destination)
    if ($Records) {
        $Records | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $Destination "AuditLog-All.json") -Encoding UTF8
    }
}

# Main Execution
Connect-ExchangeEnv -Upn $AdminUpn
$logDir = Initialize-OutputDirectory -RootPath $PSScriptRoot -Id $HaloId
$allLogs = Get-AuditLogs -Start $StartDate -End $EndDate -TargetUser $ActorUpn -Destination $logDir
Export-CombinedLogs -Records $allLogs -Destination $logDir