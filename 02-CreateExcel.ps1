[CmdLetBinding()]
PARAM (
    [String] $EventLogPath = "C:\GIT\LDAP-QueryAnalyzer",
    [String] $CompleteCSVFileName = "_Filtered-events.csv",
    [String] $ExcelFileName = "_1644Analysis.xlsx"
)

Function Write-Info {
    PARAM (
        [string] $Info
    )

    Write-Host (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -NoNewline -ForegroundColor Gray
    Write-Host "   " -NoNewline
    Write-Host $Info -ForegroundColor White
}


#-----Function supporting Excel Import.
Function mcSetPivotField($mcPivotFieldSetting)
{ #Set pivot field attributes per MSDN
    if ($null -ne $mcPivotFieldSetting[1]) { $mcPivotFieldSetting[0].Orientation  = $mcPivotFieldSetting[1]} # 1 Orientation { $xlRowField | $xlDataField }, in XlPivotFieldOrientation
	if ($null -ne $mcPivotFieldSetting[2]) { $mcPivotFieldSetting[0].NumberFormat = $mcPivotFieldSetting[2]} # 2 NumberFormat { $mcNumberF | $mcPercentF }
	if ($null -ne $mcPivotFieldSetting[3]) { $mcPivotFieldSetting[0].Function     = $mcPivotFieldSetting[3]} # 3 Function { $xlAverage | $xlSum | $xlCount }, in XlConsolidationFunction
	if ($null -ne $mcPivotFieldSetting[4]) { $mcPivotFieldSetting[0].Calculation  = $mcPivotFieldSetting[4]} # 4 Calculation { $xlPercentOfTotal | $xlPercentRunningTotal }, in XlPivotFieldCalculation
	if ($null -ne $mcPivotFieldSetting[5]) { $mcPivotFieldSetting[0].BaseField    = $mcPivotFieldSetting[5]} # 5 BaseField  <String>
    if ($null -ne $mcPivotFieldSetting[6]) { $mcPivotFieldSetting[0].Name         = $mcPivotFieldSetting[6]} # 6 Name <String>
    if ($null -ne $mcPivotFieldSetting[7]) { $mcPivotFieldSetting[0].Position     = $mcPivotFieldSetting[7]} # 7 Position
}

Function mcSetPivotTableFormat($mcPivotTable)
{ # Set pivotTable cosmetics and sheet name
    $mcPT=$mcPivotTable[0].PivotTables($mcPivotTable[1])
        $mcPT.HasAutoFormat = $False #2.turn of AutoColumnWidth
    for ($i=2; $i -lt 9; $i++)
    { #3. SetColumnWidth for Sheet($mcPivotTable[0]),PivotTable($mcPivotTable[1]),Column($mcPivotTable[2-8])
        if ($null -ne $mcPivotTable[$i]) { $mcPivotTable[0].columns.item(($i-1)).columnWidth = $mcPivotTable[$i]}
    }
    $mcPivotTable[0].Application.ActiveWindow.SplitRow = 3
    $mcPivotTable[0].Application.ActiveWindow.SplitColumn = 2
	$mcPivotTable[0].Application.ActiveWindow.FreezePanes = $true #1.Freeze R1C1
    $mcPivotTable[0].Cells.Item(1,1)="LDAPServer filter"
    $mcPivotTable[0].Cells.Item(3,1)=$mcPivotTable[9] #4 set TXT at R3C1 with PivotTableName$mcPivotTable[9]
    $mcPivotTable[0].Name=$mcPivotTable[10] #5 Set Sheet Name to $mcPivotTable[10]
        $mcRC = ($mcPivotTable[0].UsedRange.Cells).Rows.Count-1
    if ($null -ne $mcPivotTable[11])
    { # $mcPivotTable[11] Set ColorScale
        $mColorScaleRange='$'+$mcPivotTable[11]+'$4:$'+$mcPivotTable[11]+'$'+$mcRC
        [Void]$mcPivotTable[0].Range($mColorScaleRange).FormatConditions.AddColorScale(3) #$mcPivotTable[11]=ColorScale
        $mcPivotTable[0].Range($mColorScaleRange).FormatConditions.item(1).ColorScaleCriteria.item(1).type = 1 #xlConditionValueLowestValue
        $mcPivotTable[0].Range($mColorScaleRange).FormatConditions.item(1).ColorScaleCriteria.item(1).FormatColor.Color = 8109667
        $mcPivotTable[0].Range($mColorScaleRange).FormatConditions.item(1).ColorScaleCriteria.item(2).FormatColor.Color = 8711167
        $mcPivotTable[0].Range($mColorScaleRange).FormatConditions.item(1).ColorScaleCriteria.item(3).type = 2 #xlConditionValueHighestValue
        $mcPivotTable[0].Range($mColorScaleRange).FormatConditions.item(1).ColorScaleCriteria.item(3).FormatColor.Color = 7039480
    }
#    if ($mcPivotTable[12] -ne $null)
#    { # $mcPivotTable[12] Set DataBar
#        $mcDataBarRange='$'+$mcPivotTable[12]+'$4:$'+$mcPivotTable[12]+'$'+$mcRC
#		[void]$mcPivotTable[0].Range($mcDataBarRange).FormatConditions.Delete()
#        [void]$mcPivotTable[0].Range($mcDataBarRange).FormatConditions.AddDatabar()	#$mcPivotTable[12]:Set DataBar
#    }
}

Function mcSortPivotFields($mcPF)
{ #Sort on $mcPF and collapse later pivot fields
    for ($i=2; $i -lt 5; $i++) { #collapse later pivot fields
        if ($null -ne $mcPF[$i]) { $mcPF[$i].showDetail = $false }
    }
#    [void]($mcPF[0].Cells.Item(4,2)).sort(($mcPF[0].Cells.Item(4, 2)), 2)
#	$mcPF[1].showDetail = $false
    $mcPF[0].Cells(4,2).Select() | Out-Null
	$mcExcel.CommandBars.ExecuteMso("SortDescendingExcel")
}

Function mcSetPivotTableHeaderColor($mcSheet)
{ #Set PiviotTable Header Color for easier reading
	$mcSheet[0].Range("A4:"+[char]($mcSheet[0].UsedRange.Cells.Columns.count+64)+[string](($mcSheet[0].UsedRange.Cells).Rows.count-1)).interior.Color = 16056319 #Set Level0 color
    for ($i=1; $i -lt 5; $i++) { #Set header(s) color
		if ($null -ne $mcSheet[$i]) { $mcSheet[0].Range(($mcSheet[$i]+"3")).interior.Colorindex = 37 }
	}
}


#----Import csv to excel-----------------------------------------------------
Write-Info 'Import csv to excel.'
$scriptStartTime = Get-Date

$mcExcel = New-Object -ComObject excel.application
$mcWorkbooks = $mcExcel.Workbooks.Add()
	$Sheet1 = $mcWorkbooks.worksheets.Item(1)
$mcCurrentRow = 1

$mcConnector = $Sheet1.QueryTables.add(("TEXT;" + (Join-Path $EventLogPath $CompleteCSVFileName) ),$Sheet1.Range(('a'+($mcCurrentRow))))
$Sheet1.QueryTables.item($mcConnector.name).TextFileCommaDelimiter = $True
$Sheet1.QueryTables.item($mcConnector.name).TextFileParseType  = 1
[void]$Sheet1.QueryTables.item($mcConnector.name).Refresh()
if ($mcCurrentRow -ne 1) { [void]($Sheet1.Cells.Item($mcCurrentRow,1).entireRow).delete()} # Delete header on 2nd and later CSV.
$mcCurrentRow = $Sheet1.UsedRange.EntireRow.Count+1

#----Customize XLS-----------------------------------------------------------
Write-Info 'Customizing XLS.'
	$xlRowField = 1 #XlPivotFieldOrientation
	$xlPageField = 3 #XlPivotFieldOrientation
	$xlDataField = 4 #XlPivotFieldOrientation
	$xlAverage = -4106 #XlConsolidationFunction
	$xlSum = -4157 #XlConsolidationFunction
	$xlPercentOfTotal = 8 #XlPivotFieldCalculation
	$xlPercentRunningTotal = 13 #XlPivotFieldCalculation
	$mcNumberF = "###,###,###,###,###"
	$mcPercentF = "#0.00%"
	$mcDateGroupFlags=($false, $true, $true, $true, $false, $false, $false) #https://msdn.microsoft.com/en-us/library/office/ff839808.aspx

Write-Info "Sheet1 - RawData"
	$Sheet1.Range("A1").Autofilter() | Out-Null
	$Sheet1.Application.ActiveWindow.SplitRow = 1
	$Sheet1.Application.ActiveWindow.FreezePanes = $true
	$Sheet1.Columns.Item('J').numberformat = $Sheet1.Columns.Item('K').numberformat = $Sheet1.Columns.Item('M').numberformat = $Sheet1.Columns.Item('N').numberformat = $Sheet1.Columns.Item('O').numberformat = $Sheet1.Columns.Item('p').numberformat = $Sheet1.Columns.Item('Q').numberformat = $Sheet1.Columns.Item('R').numberformat = $mcNumberF

Write-Info "Sheet2 - PivotTable1"
	$Sheet2 = $mcWorkbooks.Worksheets.add()
	$PivotTable1 = $mcWorkbooks.PivotCaches().Create(1,"Sheet1!R1C1:R$($Sheet1.UsedRange.Rows.count)C$($Sheet1.UsedRange.Columns.count)",5) # xlDatabase=1 xlPivotTableVersion15=5 Excel2013
	$PivotTable1.CreatePivotTable("Sheet2!R1C1") | Out-Null
		$mcPF00 = $Sheet2.PivotTables("PivotTable1").PivotFields("LDAPServer")
			mcSetPivotField($mcPF00, $xlPageField, $null, $null, $null, $null, $null)
		$mcPF0 = $Sheet2.PivotTables("PivotTable1").PivotFields("StartingNode")
			mcSetPivotField($mcPF0, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF1 = $Sheet2.PivotTables("PivotTable1").PivotFields("LdapFilter")
			mcSetPivotField($mcPF1, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF2 = $Sheet2.PivotTables("PivotTable1").PivotFields("ClientIP")
			mcSetPivotField($mcPF2, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF = $Sheet2.PivotTables("PivotTable1").PivotFields("TimeGenerated")
			mcSetPivotField($mcPF, $xlRowField, $null, $null, $null, $null, $null)
			$mcCells=$mcPF.DataRange.Item(4)
			$mcCells.group($true,$true,1,$mcDateGroupFlags) | Out-Null
		$mcPF = $Sheet2.PivotTables("PivotTable1").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $null, $null, $null, "Search Count",1)
		$mcPF = $Sheet2.PivotTables("PivotTable1").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $xlAverage, $null, $null, "AvgSearchTime",2)
		$mcPF = $Sheet2.PivotTables("PivotTable1").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentOfTotal, $null, "%GrandTotal",3)
		$mcPF = $Sheet2.PivotTables("PivotTable1").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentRunningTotal, "StartingNode", "%RunningTotal",4)
    mcSetPivotTableFormat($Sheet2, "PivotTable1", 60, 12, 14, 12, 14, $null, $null,"StartingNode grouping", "2.TopIP-StartingNode", "D", "D")
    mcSortPivotFields($sheet2,$mcPF0,$mcPF1,$mcPF2)
	mcSetPivotTableHeaderColor($Sheet2, "B", "D", "E")

Write-Info "Sheet3 - PivotTable2"
	$Sheet3 = $mcWorkbooks.Worksheets.add()
	$PivotTable2 = $mcWorkbooks.PivotCaches().Create(1,"Sheet1!R1C1:R$($Sheet1.UsedRange.Rows.count)C$($Sheet1.UsedRange.Columns.count)",5) # xlDatabase=1 xlPivotTableVersion15=5 Excel2013
	$PivotTable2.CreatePivotTable("Sheet3!R1C1") | Out-Null
		$mcPF00 = $Sheet3.PivotTables("PivotTable2").PivotFields("LDAPServer")
			mcSetPivotField($mcPF00, $xlPageField, $null, $null, $null, $null, $null)
		$mcPF0 = $Sheet3.PivotTables("PivotTable2").PivotFields("ClientIP")
			mcSetPivotField($mcPF0, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF1 = $Sheet3.PivotTables("PivotTable2").PivotFields("LdapFilter")
			mcSetPivotField($mcPF1, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF2 = $Sheet3.PivotTables("PivotTable2").PivotFields("TimeGenerated")
			mcSetPivotField($mcPF2, $xlRowField, $null, $null, $null, $null, $null)
			$mcCells=$mcPF2.DataRange.Item(3)
			$mcCells.group($true,$true,1,$mcDateGroupFlags) | Out-Null
		$mcPF = $Sheet3.PivotTables("PivotTable2").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $null, $null, $null, "Search Count",1)
		$mcPF = $Sheet3.PivotTables("PivotTable2").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $xlAverage, $null, $null, "AvgSearchTime (MS)",2)
		$mcPF = $Sheet3.PivotTables("PivotTable2").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentOfTotal, $null, "%GrandTotal",3)
		$mcPF = $Sheet3.PivotTables("PivotTable2").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentRunningTotal, "ClientIP", "%RunningTotal",4)
	mcSetPivotTableFormat($Sheet3, "PivotTable2", 60, 12, 19, 12, 14, $null, $null,"IP grouping", "3.TopIP", "D", "D")
    mcSortPivotFields($sheet3,$mcPF0,$mcPF1)
	mcSetPivotTableHeaderColor($Sheet3, "B", "D", "E")

Write-Info "Sheet4 - PivotTable3"
	$Sheet4 = $mcWorkbooks.Worksheets.add()
	$PivotTable3 = $mcWorkbooks.PivotCaches().Create(1,"Sheet1!R1C1:R$($Sheet1.UsedRange.Rows.count)C$($Sheet1.UsedRange.Columns.count)",5) # xlDatabase=1 xlPivotTableVersion15=5 Excel2013
	$PivotTable3.CreatePivotTable("Sheet4!R1C1") | Out-Null
		$mcPF00 = $Sheet4.PivotTables("PivotTable3").PivotFields("LDAPServer")
			mcSetPivotField($mcPF00, $xlPageField, $null, $null, $null, $null, $null)
		$mcPF0 = $Sheet4.PivotTables("PivotTable3").PivotFields("LdapFilter")
			mcSetPivotField($mcPF0, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF1 = $Sheet4.PivotTables("PivotTable3").PivotFields("ClientIP")
			mcSetPivotField($mcPF1, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF = $Sheet4.PivotTables("PivotTable3").PivotFields("TimeGenerated")
			mcSetPivotField($mcPF, $xlRowField, $null, $null, $null, $null, $null)
			$mcCells=$mcPF.DataRange.Item(3)
			$mcCells.group($true,$true,1,$mcDateGroupFlags) | Out-Null
		$mcPF = $Sheet4.PivotTables("PivotTable3").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $null, $null, $null, "Search Count",1)
		$mcPF = $Sheet4.PivotTables("PivotTable3").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $xlAverage, $null, $null, "AvgSearchTime (MS)",2)
		$mcPF = $Sheet4.PivotTables("PivotTable3").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentOfTotal, $null, "%GrandTotal",3)
		$mcPF = $Sheet4.PivotTables("PivotTable3").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentRunningTotal, "LdapFilter", "%RunningTotal",4)
	mcSetPivotTableFormat($Sheet4, "PivotTable3", 60, 12, 19, 12, 14, $null, $null,"Filter grouping", "4.TopIP-Filters","D","D")
    mcSortPivotFields($sheet4,$mcPF0,$mcPF1)
	mcSetPivotTableHeaderColor($Sheet4, "B", "D", "E")

Write-Info "Sheet5 - PivotTable4"
	$Sheet5 = $mcWorkbooks.Worksheets.add()
	$PivotTable4 = $mcWorkbooks.PivotCaches().Create(1,"Sheet1!R1C1:R$($Sheet1.UsedRange.Rows.count)C$($Sheet1.UsedRange.Columns.count)",5) # xlDatabase=1 xlPivotTableVersion15=5 Excel2013
	$PivotTable4.CreatePivotTable("Sheet5!R1C1") | Out-Null
		$mcPF00 = $Sheet5.PivotTables("PivotTable4").PivotFields("LDAPServer")
			mcSetPivotField($mcPF00, $xlPageField, $null, $null, $null, $null, $null)
		$mcPF0 = $Sheet5.PivotTables("PivotTable4").PivotFields("ClientIP")
			mcSetPivotField($mcPF0, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF1 = $Sheet5.PivotTables("PivotTable4").PivotFields("LdapFilter")
			mcSetPivotField($mcPF1, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF = $Sheet5.PivotTables("PivotTable4").PivotFields("TimeGenerated")
			mcSetPivotField($mcPF, $xlRowField, $null, $null, $null, $null, $null)
			$mcCells=$mcPF.DataRange.Item(3)
			$mcCells.group($true,$true,1,$mcDateGroupFlags) | Out-Null
		$mcPF = $Sheet5.PivotTables("PivotTable4").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $xlSum, $null, $null, "Total SearchTime (MS)")
		$mcPF = $Sheet5.PivotTables("PivotTable4").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $null, $null, $null, "Search Count")
		$mcPF = $Sheet5.PivotTables("PivotTable4").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $xlAverage, $null, $null, "AvgSearchTime (MS)")
		$mcPF = $Sheet5.PivotTables("PivotTable4").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentOfTotal, $null, "%GrandTotal (MS)")
		$mcPF = $Sheet5.PivotTables("PivotTable4").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentRunningTotal, "ClientIP", "%RunningTotal (Ms)")
	mcSetPivotTableFormat($Sheet5, "PivotTable4", 50, 21, 12, 19, 17, 19, $null,"IP grouping", "5.TopTime-IP","E","E")
    mcSortPivotFields($sheet5,$mcPF0,$mcPF1)
	mcSetPivotTableHeaderColor($Sheet5, "B", "E", "F")

Write-Info "Sheet6 - PivotTable5"
	$Sheet6 = $mcWorkbooks.Worksheets.add()
	$PivotTable5 = $mcWorkbooks.PivotCaches().Create(1,"Sheet1!R1C1:R$($Sheet1.UsedRange.Rows.count)C$($Sheet1.UsedRange.Columns.count)",5) # xlDatabase=1 xlPivotTableVersion15=5 Excel2013
	$PivotTable5.CreatePivotTable("Sheet6!R1C1") | Out-Null
		$mcPF00 = $Sheet6.PivotTables("PivotTable5").PivotFields("LDAPServer")
			mcSetPivotField($mcPF00, $xlPageField, $null, $null, $null, $null, $null)
		$mcPF0 = $Sheet6.PivotTables("PivotTable5").PivotFields("LdapFilter")
			mcSetPivotField($mcPF0, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF1 = $Sheet6.PivotTables("PivotTable5").PivotFields("ClientIP")
			mcSetPivotField($mcPF1, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF = $Sheet6.PivotTables("PivotTable5").PivotFields("TimeGenerated")
			mcSetPivotField($mcPF, $xlRowField, $null, $null, $null, $null, $null)
			$mcCells=$mcPF.DataRange.Item(3)
			$mcCells.group($true,$true,1,$mcDateGroupFlags) | Out-Null
		$mcPF = $Sheet6.PivotTables("PivotTable5").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $xlSum, $null, $null, "Total SearchTime (MS)")
		$mcPF = $Sheet6.PivotTables("PivotTable5").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $null, $null, $null, "Search Count")
		$mcPF = $Sheet6.PivotTables("PivotTable5").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $xlAverage, $null, $null, "AvgSearchTime (MS)")
		$mcPF = $Sheet6.PivotTables("PivotTable5").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentOfTotal, $null, "%GrandTotal (MS)")
		$mcPF = $Sheet6.PivotTables("PivotTable5").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $null, $xlPercentRunningTotal, "LdapFilter", "%RunningTotal (MS)")
	mcSetPivotTableFormat($Sheet6, "PivotTable5", 50, 21, 12, 19, 17, 19, $null,"Filter grouping", "6.TopTime-Filters","E","E")
    mcSortPivotFields($sheet6,$mcPF0,$mcPF1)
	mcSetPivotTableHeaderColor($Sheet6, "B", "E", "F")

Write-Info "Sheet7 - PivotTable6"
	$Sheet7 = $mcWorkbooks.Worksheets.add()
	$PivotTable6 = $mcWorkbooks.PivotCaches().Create(1,"Sheet1!R1C1:R$($Sheet1.UsedRange.Rows.count)C$($Sheet1.UsedRange.Columns.count)",5) # xlDatabase=1 xlPivotTableVersion15=5 Excel2013
	$PivotTable6.CreatePivotTable("Sheet7!R1C1") | Out-Null
		$mcPF00 = $Sheet7.PivotTables("PivotTable6").PivotFields("LDAPServer")
			mcSetPivotField($mcPF00, $xlPageField, $null, $null, $null, $null, $null)
		$mcPF0 = $Sheet7.PivotTables("PivotTable6").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF0, $xlRowField, $null, $null, $null, $null, $null)
			$mcPF0.DataRange.Item(1).group(0,$true,50) | Out-Null
		$mcPF1 = $Sheet7.PivotTables("PivotTable6").PivotFields("LdapFilter")
			mcSetPivotField($mcPF1, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF2 = $Sheet7.PivotTables("PivotTable6").PivotFields("ClientIP")
			mcSetPivotField($mcPF2, $xlRowField, $null, $null, $null, $null, $null)
		$mcPF = $Sheet7.PivotTables("PivotTable6").PivotFields("TimeGenerated")
			mcSetPivotField($mcPF, $xlRowField, $null, $null, $null, $null, $null)
			$mcCells=$mcPF.DataRange.Item(4)
			$mcCells.group($true,$true,1,$mcDateGroupFlags) | Out-Null
		$mcPF = $Sheet7.PivotTables("PivotTable6").PivotFields("ClientIP")
			mcSetPivotField($mcPF, $xlDataField, $mcNumberF, $null, $null, $null, "Search Count")
		$mcPF = $Sheet7.PivotTables("PivotTable6").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $xlSum, $xlPercentOfTotal, $null, "%GrandTotal (MS)")
		$mcPF = $Sheet7.PivotTables("PivotTable6").PivotFields("SearchTimeMS")
			mcSetPivotField($mcPF, $xlDataField, $mcPercentF, $xlSum, $xlPercentRunningTotal, "SearchTimeMS", "%RunningTotal (MS)")
	mcSetPivotTableFormat($Sheet7, "PivotTable6", 60, 12, 17, 19,$null, $null, $null, "SearchTime (MS) grouping", "7.TimeRanks",$null,"C")
	$mcPF0.showDetail = $mcPF1.showDetail = $mcPF2.showDetail = $false
	mcSetPivotTableHeaderColor($Sheet7, "C", "D")

Write-Info "Sheet8 - PivotTable7"
	$Sheet8 = $mcWorkbooks.Worksheets.add()
	$PivotTable7 = $mcWorkbooks.PivotCaches().Create(1,"Sheet1!R1C1:R$($Sheet1.UsedRange.Rows.count)C$($Sheet1.UsedRange.Columns.count)",5) # xlDatabase=1 xlPivotTableVersion15=5 Excel2013
	$PivotTable7.CreatePivotTable("Sheet8!R1C1") | Out-Null
    $Sheet8.name = "8.SandBox"

Write-Info "Set Sheet1 name and sort sheet names in reverse"
	$Sheet1.Name = "1.RawData"
$Sheet2.Tab.ColorIndex = $Sheet3.Tab.ColorIndex = $Sheet4.Tab.ColorIndex = 35
$Sheet5.Tab.ColorIndex = $Sheet6.Tab.ColorIndex = $Sheet7.Tab.ColorIndex = 36
$sheet8.Tab.Color=8109667
$mcWorkSheetNames = New-Object System.Collections.ArrayList
	foreach ($mcWorkSheet in $mcWorkbooks.Worksheets) { $mcWorkSheetNames.add($mcWorkSheet.Name) | Out-null }
	$mctmp = $mcWorkSheetNames.Sort() | Out-Null
	For ($i=0; $i -lt $mcWorkSheetNames.Count-1; $i++){ #Sort name.
		$mcTmp = $mcWorkSheetNames[$i]
		$mcBefore = $mcWorkbooks.Worksheets.Item($mcTmp)
		$mcAfter = $mcWorkbooks.Worksheets.Item($i+1)
		$mcBefore.Move($mcAfter)
	}
$Sheet1.Activate()

$fileName = Join-Path $EventLogPath $ExcelFileName
Write-Info "Saving file to $fileName"
$mcWorkbooks.SaveAs($fileName)

$mcExcel.visible = $true

Write-Info 'Script completed.'

$span = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
Write-Info "Elapsed time for script: $($span.Hours) h, $($span.Minutes) m, $($span.Seconds) s (PowerShell: $($PSVersionTable.PSVersion) $($PSVersionTable.PSEdition))"
	
#################################################################################
