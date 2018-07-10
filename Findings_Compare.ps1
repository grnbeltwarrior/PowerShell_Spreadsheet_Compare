# Variables
# Already found vulns:
$vulnSpreadsheet = "C:\Existing.xlsx"
# New vuln scan spreadsheet:
$findingSpreadsheet = "C:\Findings.xlsx"
$vulnSheetName = 'Vuln Open'
$findingSheetName = 'Vuln List'
$newSheetArray = 'Finding #','State','Status','Finding Title','System Name','System IP','Port','Technical Risk','Business Risk','CVSS Base Score','CVS','Description','Recommendations','Notes','Detail'
# 3rd spreadsheet of new only vulns:
$outpath = "C:\New.xlsx"

function Closing-Excel {
	Write-Host "Closing out Excel."
	$vulnWorkBook.Close()
	$findingWorkBook.Close()
    $objExcel.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($objExcel)
}

function VulnBuilder {
    # App_Packet Locations
    $DeviceName = $vulnSheet.Range('B' + $args[0]).Text # Column B
    $IPAddress = $vulnSheet.Range('C' + $args[0]).Text # Column C
    $ServiceDetail = $vulnSheet.Range('D' + $args[0]).Text # Column D - need to split on (tcp/udp)/port
    $ServiceDetail = $ServiceDetail.split('/')[1]
    $VulnID = $vulnSheet.Range('E' + $args[0]).Text # Column E - need to split on column strings
    if ($VulnID -match 'string1') {
        $cutVulnID = ($VulnID -replace "[^0-9]", '')
    }
    Elseif ($VulnID -match 'string2') {
        $cutVulnID = ($VulnID -replace "[^0-9]", '')
    }
    Else {
        Write-Output "Something went wrong: " $vulnID
        }
    $VulnStatus = $vulnSheet.Range('G' + $args[0]).Text # Column G Active/Exception
    $LastFoundDate = $vulnSheet.Range('L' + $args[0]).Text # Column L
    $VulnName = $vulnSheet.Range('AN' + $args[0]).Text # Column AN
    $vulnString = "$cutVulnID,$VulnName,$DeviceName,$IPAddress,$ServiceDetail"
    return $vulnString
}

Function FindingBuilder {
    # Nessus Findings Locations
    $Finding = $findingSheet.Range('A' + $args[0]).Text # Column A - need to split on string
    $Finding = $Finding.split('-')[1]
    $FindingTitle = $findingSheet.Range('D' + $args[0]).Text # Column D
    $SystemName = $findingSheet.Range('E' + $args[0]).Text # Column E
    $SystemIP = $findingSheet.Range('F' + $args[0]).Text # Column F
    $Port = $findingSheet.Range('G' + $args[0]).Text # Column G - need to split on port/(tcp/udp)
    $Port = $Port.split('/')[0]
    $findingString = "$Finding,$FindingTitle,$SystemName,$SystemIP,$Port"
    return $findingString
}

function SaveAndCloseSpreadsheet {
    $excel.ActiveWorkbook.SaveAs($outpath)
	$newWorkbook.Close()
	$excel.Quit()
	[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}

function getFindingRow {
	# Go through and copy over data from the findings spreadsheet to the new spreadsheet.
	$lastRow = $newSpreadsheet.UsedRange.rows.count + 1
	$newCells.item($lastRow,'A') = $findingSheet.Range('A' + $args[0]).Text 
	$newCells.item($lastRow,'B') = $findingSheet.Range('B' + $args[0]).Text
	$newCells.item($lastRow,'C') = $findingSheet.Range('C' + $args[0]).Text
	$newCells.item($lastRow,'D') = $findingSheet.Range('D' + $args[0]).Text
	$newCells.item($lastRow,'E') = $findingSheet.Range('E' + $args[0]).Text
	$newCells.item($lastRow,'F') = $findingSheet.Range('F' + $args[0]).Text
	$newCells.item($lastRow,'G') = $findingSheet.Range('G' + $args[0]).Text
	$newCells.item($lastRow,'H') = $findingSheet.Range('H' + $args[0]).Text
	$newCells.item($lastRow,'I') = $findingSheet.Range('I' + $args[0]).Text
	$newCells.item($lastRow,'J') = $findingSheet.Range('J' + $args[0]).Text
	$newCells.item($lastRow,'K') = $findingSheet.Range('K' + $args[0]).Text
	$newCells.item($lastRow,'L') = $findingSheet.Range('L' + $args[0]).Text
	$newCells.item($lastRow,'M') = $findingSheet.Range('M' + $args[0]).Text
	$newCells.item($lastRow,'N') = $findingSheet.Range('N' + $args[0]).Text
	$newCells.item($lastRow,'O') = $findingSheet.Range('O' + $args[0]).Text
}

# New instance of excel
$objExcel = New-Object -ComObject Excel.Application
$objExcel.DisplayAlerts = $false
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false

# Resulting Spreadsheet
$newWorkbook = $excel.Workbooks.Add()

# Open spreadsheets
$vulnWorkBook = $objExcel.Workbooks.Open($vulnSpreadsheet)
Write-Host "Opening the Vuln Spreadsheet..."
$findingWorkBook = $objExcel.Workbooks.Open($findingSpreadsheet)
Write-Host "Opening the Finding Spreadsheet..."

# Look for vuln open sheet
$vulnSheet = $vulnWorkBook.sheets.item($vulnSheetName)
$findingSheet = $findingWorkBook.sheets.item($findingSheetName)

# Loop through Nessus Findings then through App Packet
$FindingRowMax = ($findingSheet.UsedRange.Rows).count
$VulnRowMax = ($vulnSheet.UsedRange.Rows).count

# Create new finding spread sheet
Write-Host "Creating new spreadsheet..."
$newSpreadsheet = $newWorkbook.Sheets.Item(1)
$newSpreadsheet.Activate()
$row = 1
$newCells =$newSpreadsheet.Cells
for ($i = 0;$i -le 14;$i++) {
	$newCells.Item($row,$i+1) = $newSheetArray[$i]
}

for($FindRow = 2; $FindRow -le $FindingRowMax; $FindRow++){
    $counter = 0
	$find = FindingBuilder $FindRow
    for($VulnRow = 2; $VulnRow -le $VulnRowMax; $VulnRow++){
        $vuln = VulnBuilder $VulnRow
        if ($find -eq $vuln) {
            $counter = $counter + 1
        }
   }
    if ($counter -ge 1) {
		getFindingRow $FindRow
    }
}

Write-Host "Entering information into the new spreadsheet..."
$newSpreadsheet.columns.item("A:D").EntireColumn.AutoFit() | out-null
$newSpreadsheet.UsedRange.rows.RowHeight = 15

SaveAndCloseSpreadsheet
Closing-Excel
Remove-Variable objExcel
