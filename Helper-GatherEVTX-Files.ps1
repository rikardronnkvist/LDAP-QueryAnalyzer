[CmdLetBinding()]
PARAM (
    [String] $EventLogPath = "C:\GIT\LDAP-QueryAnalyzer",
    [string] $ServerPath = "ADMIN$\system32\winevt\Logs\Directory Service.evtx",
    [String[]] $DomainControllers = @(
        "DC01",
        "DC02.domain.com",
        "DC98",
        "DC99.anotherdomain.com"
    )
)

Function Write-Info {
    PARAM (
        [string] $Info
    )

    Write-Host (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -NoNewline -ForegroundColor Gray
    Write-Host "   " -NoNewline
    Write-Host $Info -ForegroundColor White
}


# ---------------------------------------------------------------------------------------
#  Start!

Write-Info "Starting to gather EVTX-files"
$scriptStartTime = Get-Date


ForEach ($DomainController in $DomainControllers) {
    $remotePath = Join-Path "\\$($DomainController)" $ServerPath
    $localFile = Join-Path $EventLogPath "$($DomainController.Replace(".", "-")) - Directory Service.evtx"

    Write-Info "Copy from '$($remotePath)' to '$($localFile)'"
    Copy-Item -Path $remotePath -Destination $localFile -Force
    
}

$span = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
Write-Info "Elapsed time for script: $($span.Hours) h, $($span.Minutes) m, $($span.Seconds) s (PowerShell: $($PSVersionTable.PSVersion) $($PSVersionTable.PSEdition))"

# ---------------------------------------------------------------------------------------
#  Done!
