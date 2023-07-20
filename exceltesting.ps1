#the purpose of this scrpit is to act as a testing groud for working with excel
#Author James Collins


function MakeNewWorksheet($sheetName)
{

}



#ref excel object
$ExcelObject = New-Object -ComObject Excel.Application

#make the excel window visable
$ExcelObject.Visible=$true

#Open workbook
$ActiveWorkbook = $ExcelObject.workbooks.Open("D:\Projects\powershell\Powershell-excel\file1.xlsx")

#set active worksheet
$ActiveWorkSheet = $ActiveWorkbook.Sheets.Item("D1")
#$columnsToCopy = $sourceUsedRange.Columns.Item(1).Resize($sourceUsedRange.Rows.Count, 3)


$d1UsedRange = $ActiveWorkSheet.UsedRange


#Make a new worksheet
$newWorksheet = $ActiveWorkbook.Sheets.Add()
$newWorksheet.Name = "Compare D1 D2"

#get the used range from source
$CompareRange = $newWorksheet.Cells.Item(1,1)

$d1UsedRange.Copy($CompareRange)
$ActiveWorkSheet = $ActiveWorkbook.Sheets.Item("D1")
$d2UsedRange = $ActiveWorkSheet.UsedRange






echo "Summary of Excel Sheet"

foreach($row in $CompareRange.value2)
{
    foreach($cell in $row)
    {
        Write-Host $cell
    }

}

$ActiveWorkbook.Close($true)
$ExcelObject.Quit()