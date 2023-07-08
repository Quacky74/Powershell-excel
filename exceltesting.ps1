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
$ActiveWorkbook = $ExcelObject.workbooks.Open("D:\Projects\powershell\file1.xlsx")

#set active worksheet
$ActiveWorkSheet = $ActiveWorkbook.Sheets.Item("D1")

$d1UsedRange = $ActiveWorkSheet.UsedRange
$d1Data = $d1UsedRange.Value2



echo "Summary of Excel Sheet"

foreach($row in $d1Data)
{
    foreach($cell in $row)
    {
        Write-Host $cell
    }

}

$ActiveWorkbook.Close($true)
$ExcelObject.Quit()