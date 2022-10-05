[CmdLetBinding()]
PARAM (
    [String] $EventLogPath = "C:\GIT\LDAP-QueryAnalyzer",
    [String] $CompleteCSVFileName = "_1644-Events.csv",
    [bool] $gatherExtraEventData = $false,  # Set to $true if you want to analyze LDAP query performance in Excel
    [String[]] $filterClientIP = @(         # Filters are disabled if you gather extra event data
        "SAM",
        "Internal",
        "KCC",
        "LSA",
        "NTDSAPI",
        "127.0.*",
        "192.168.*",
        "172.16.1.12",
        "172.16.1.13",
        "172.16.1.14",
        "172.16.2.101",
        "172.16.2.102"
    ),
    [String[]] $filterLdapQuery = @(
        "*objectCategory=CN=MS-SMS-Management-Point,CN=Schema*",
        "*objectCategory=CN=MS-SMS-Site,CN=Schema*"
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

Function Test-Filters {
    PARAM (
        $eventToCheck
    )

    foreach ($filter in $filterClientIP) {
        if ($filter -match "\*") {
            If ($eventToCheck.ClientIP -like $filter) {
                Write-Verbose "Filter match! $($eventToCheck.ClientIP) -match $($filter)"
                Return $true
                Break
            }
        } else {
            If ($eventToCheck.ClientIP -eq $filter) {
                Write-Verbose "Filter match! $($eventToCheck.ClientIP) -eq $($filter)"
                Return $true
                Break
            }
        }

    }
    Write-Verbose "No ClientIP match for $($eventToCheck.ClientIP)"

    foreach ($filter in $filterLdapQuery) {
        If ($eventToCheck.LdapFilter -like $filter) {
            Write-Verbose "Filter match! $($eventToCheck.LdapFilter) -match $($filter)"
            Return $true
            break
        }
    }
    Write-Verbose "No Ldap Query match for $($eventToCheck.LdapFilter)"

    Return $false
}

Function Get-ClientIPorPort {
    PARAM (
        $Client
    )

	$Return = 'Unknown'
	[regex]$regexIPV6 = '(?<IP>\[[A-Fa-f0-9:%]{1,}\])\:(?<Port>([0-9]+))'
	[regex]$regexIPV4 =	'((?<IP>(\d{1,3}\.){3}\d{1,3})\:(?<Port>[0-9]+))|(?<IP>(\d{1,3}\.){3}\d{1,3})'
	[regex]$KnownClient = '(?<IP>([G-Z])\w+)'

	Switch -RegEx ($Client[0]) {
	    $regexIPV6   { $Return = $matches.($client[1]) }
	    $regexIPV4   { $Return = $matches.($client[1]) }
		$KnownClient { $Return = $matches.($client[1]) }
	}

	Return $Return
}

# ---------------------------------------------------------------------------------------
#  Start!
Clear-Host

Write-Info "Starting to process directory $($EventLogPath)"
$scriptStartTime = Get-Date

# ---------------------------------------------------------------------------------------
#  Process EVTX-Files

$evtxFilesStartTime = Get-Date

Write-Info "Searching for EVTX-files"
$eventFiles = Get-ChildItem -Path $EventLogPath -Filter "*.evtx" | Sort-Object Name
$eventsMb = [math]::Round(($eventFiles | Measure-Object -Property Length -Sum).Sum / 1MB)
$completedMb = 0
$eventFileNo = 1

Write-Info "Found $($eventFiles.Count) evtx-files, total of $($eventsMb) Mb"

$eventFiles | ForEach-Object ($_) {
    $evtxStartTime = Get-Date
    Write-Progress -Activity "Reading EVTX" -Status "$($_.Name)   File $($eventFileNo) of $($eventFiles.Count)   Size: $( [Math]::Round($_.Length / 1MB) ) of $($eventsMb) Mb" -Id 1 -PercentComplete (($completedMb / $eventsMb) * 100)

	Write-Info "Reading file $($eventFileNo) of $($eventFiles.Count) ($($_.Name)), Size: $( [Math]::Round($_.Length / 1MB) ) Mb"
	$allEvents = Get-WinEvent -FilterHashtable @{Path=$_.FullName; LogName="Directory Service"; id="1644" } -ErrorAction SilentlyContinue


    $nofEvents = ($allEvents | Measure-Object).Count
    $csvExport = @()
    Write-Info "   Processing $($nofEvents) events"
    If ($gatherExtraEventData) { Write-Info "   Gathering extra EventData, will take extra time" }

    $row = 0
    $addedRows = 0
	ForEach ($Event in $allEvents) {
        Write-Progress -Activity "Processing events" -Status "Processing $($nofEvents) events" -CurrentOperation "Row $($row) (Keeping $($addedRows))" -PercentComplete (($row / $nofEvents) * 100) -ParentId 1 -Id 2
	    $newEvent = New-Object System.Object

		$newEvent | Add-Member -MemberType NoteProperty -Name LDAPServer        -force -Value $Event.MachineName
		$newEvent | Add-Member -MemberType NoteProperty -Name TimeGenerated     -force -Value $Event.TimeCreated
		$newEvent | Add-Member -MemberType NoteProperty -Name ClientIP          -force -Value (Get-ClientIPorPort($Event.Properties[4].Value,'IP'))
		$newEvent | Add-Member -MemberType NoteProperty -Name ClientPort        -force -Value (Get-ClientIPorPort($Event.Properties[4].Value,'Port'))
		$newEvent | Add-Member -MemberType NoteProperty -Name StartingNode      -force -Value $Event.Properties[0].Value
		$newEvent | Add-Member -MemberType NoteProperty -Name LdapFilter        -force -Value $Event.Properties[1].Value

        If ($gatherExtraEventData) {
		    $newEvent | Add-Member -MemberType NoteProperty -Name SearchScope -force -Value $Event.Properties[5].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name AttributeSelection -force -Value $Event.Properties[6].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name ServerControls -force -Value $Event.Properties[7].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name VisitedEntries -force -Value $Event.Properties[2].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name ReturnedEntries -force -Value $Event.Properties[3].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name UsedIndexes -force -Value $Event.Properties[8].Value # KB 2800945 or later has extra data fields.
		    $newEvent | Add-Member -MemberType NoteProperty -Name PagesReferenced -force -Value $Event.Properties[9].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name PagesReadFromDisk -force -Value $Event.Properties[10].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name PagesPreReadFromDisk -force -Value $Event.Properties[11].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name CleanPagesModified -force -Value $Event.Properties[12].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name DirtyPagesModified -force -Value $Event.Properties[13].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name SearchTimeMS -force -Value $Event.Properties[14].Value
		    $newEvent | Add-Member -MemberType NoteProperty -Name AttributesPreventingOptimization -force -Value $Event.Properties[15].Value

            # Do not filter rows when adding extra data
            $csvExport += $newEvent
            $addedRows ++
        } else {
            if (! (Test-Filters -eventToCheck $newEvent) ) {
                $csvExport += $newEvent
                $addedRows ++
            }
        }

        $row ++
	}
    Write-Progress -Activity "Processing events" -ParentId 1 -Id 2 -Completed

    If ($gatherExtraEventData) {
        $csvFileName = Join-Path $EventLogPath "$($_.BaseName) - 1644-Events.csv"
    } else {}
        $csvFileName = Join-Path $EventLogPath "$($_.BaseName) - 1644-Events-Filtered.csv"
    }

    Write-Info "   Export $( ($csvExport | Measure-Object).Count ) rows to CSV - $($csvFileName)"
    $csvExport | Export-Csv -Path $csvFileName -NoTypeInformation -Encoding UTF8 -Force

    $span = New-TimeSpan -Start $evtxStartTime -End (Get-Date)
    Write-Info "   Elapsed time for file: $($span.Hours) h, $($span.Minutes) m, $($span.Seconds) s"

    $completedMb += ( $_.Length / 1MB )
    $eventFileNo ++
    $allEvents = $null
}
Write-Progress -Activity "Reading EVTX" -id 1 -Completed

    
Write-Info "Done processing EVTX-files!"

$span = New-TimeSpan -Start $evtxFilesStartTime -End (Get-Date)
Write-Info "Elapsed time for all EVTX-files: $($span.Hours) h, $($span.Minutes) m, $($span.Seconds) s"
Write-Info ""


# ---------------------------------------------------------------------------------------
#  Process CSV-Files

$csvFilesStartTime = Get-Date

If ($gatherExtraEventData) {
    Write-Info "Processing CSV-files"
    $csvFiles = Get-ChildItem -Path $EventLogPath -Filter "*1644-Events.csv" | Sort-Object Name
} else {}
    Write-Info "Processing filtered CSV-files"
    $csvFiles = Get-ChildItem -Path $EventLogPath -Filter "*1644-Events-Filtered.csv" | Sort-Object Name
}


$csvMb = [math]::Round(($csvFiles | Measure-Object -Property Length -Sum).Sum / 1MB)
$csvFileNo = 1

$csvExportFileName = (Join-Path $EventLogPath $CompleteCSVFileName)
Remove-Item -Path $csvExportFileName -Force -ErrorAction SilentlyContinue | Out-Null
Write-Info "Found $($eventFiles.Count) CSV-files, total of $($csvMb) Mb"


$csvFiles | ForEach-Object {
	Write-Info "Reading $($_.Name) ( File $($csvFileNo) of $($csvFiles.Count) )"

    Write-Progress -Activity "Reading CSV" -Status "Reading $($_.Name) ( File $($csvFileNo) of $($csvFiles.Count) )" -PercentComplete (($csvFileNo / $csvFiles.Count) * 100)
        
    $csv = Import-Csv -Path $_.FullName -Encoding UTF8

    Write-Info "   Adding $( ($csv | Measure-Object).Count ) rows to $($CompleteCSVFileName)"
    $csv | Export-Csv -Path $csvExportFileName -Append -NoTypeInformation -Encoding UTF8

    $csvFileNo ++

}
Write-Progress -Activity "Reading CSV" -Completed

Write-Info "Done processing CSV-files!"

$span = New-TimeSpan -Start $csvFilesStartTime -End (Get-Date)
Write-Info "Elapsed time for all CSV-files: $($span.Hours) h, $($span.Minutes) m, $($span.Seconds) s"
Write-Info ""

# ---------------------------------------------------------------------------------------

$span = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
Write-Info "Elapsed time for script: $($span.Hours) h, $($span.Minutes) m, $($span.Seconds) s (PowerShell: $($PSVersionTable.PSVersion) $($PSVersionTable.PSEdition))"

# ---------------------------------------------------------------------------------------
#  Done!
