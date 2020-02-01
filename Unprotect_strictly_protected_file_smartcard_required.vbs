'Unprotect Strictly Confidential Data (Smartcard Required)
'When running in February 2020, first need to download 'Report - Jan 2020 (1 Feb 2019 - 31 Jan 2020).xlsx' file (template) available in this repo, save it in a desired location >> Dir, and move it to a subfolder '202002'
'Note: This script was created for a specific type of report.

Set objExcel = CreateObject("Excel.Application")
Set objDelete = CreateObject("Scripting.FileSystemObject")
objExcel.Visible = False
objExcel.DisplayAlerts = False

'Define variables
'Directory path
Dir = "C:\Users\DOM\Downloads\"
'Dates and formats
t=Now()
PrMv=DatePart("m",DateAdd("m",-1,t)) 'Previous reporting month value
CMv=DatePart("m",DateAdd("m",0,t)) 'Current reporting month value
PrMLnm=MonthName(PrMv) 'Previous reporting month long name
PrMSnm=MonthName(PrMv,True) 'Previous reporting month short name
CMSnm=MonthName(CMv,True) 'Current reporting month short name
CYr=DatePart("yyyy",DateAdd("m",0,t)) 'Current year current month value
CYrPrM=DatePart("yyyy",DateAdd("m",0,t)) 'Current year previous month value
PrYr=DatePart("yyyy",DateAdd("m",-11,t)) 'Previous year
LastDayPrM=DatePart("d",DateSerial(PrYr,CMv,0)) 'Prior year current month last day
'Strings
FolderName=CYr&right(0&CMv,2)&"\" 'e.g. 202002\
FileName= "Report - "&PrMSnm&" "&CYrPrM&" (1 "&CMSnm&" "&PrYr&" - "&LastDayPrM&" "&PrMSnm&" "&CYrPrM&").xlsx"  'e.g. Report - Jan 2020 (1 Feb 2019 - 31 Jan 2020).xlsx
SheetName=PrMLnm&" "&CYr&" - Report" 'e.g. January 2020 - Report
pw="rlendzion"
Folder=Dir&FolderName
Path1=Folder&FileName
Path2=Folder&"vNoPass.xlsx" 'Temporary file
Path3=Folder&"Unprotected_"&FileName 'Final unprotected file

'Excel formula definition - Smartcard restrictions will be inherited by another file if we simply copy and paste an array
Loc="'"&Folder&"[vNoPass.xlsx]"
Frm="=IF("&Loc&SheetName&"'!A1"&cStr("="&chr(34)&chr(34)&","&chr(34)&chr(34)&",")&Loc&SheetName&"'!A1"&cStr(")")

Set objWbk = objExcel.Workbooks.Open(Path1,,,,cStr(pw))
objWbk.Password=""
objWbk.SaveAs Path2
Set NewWbk=objExcel.Workbooks.Add()
NewWbk.Sheets(1).Name = SheetName
lastrow=objWbk.Worksheets(SheetName).UsedRange.Rows.Count + 1 'Keep +1 in case the data starts from row 2
NewWbk.Sheets(SheetName).Range("A1").Formula = Frm
Const xlFillDefault = 0
NewWbk.Sheets(SheetName).Range("A1").AutoFill NewWbk.Sheets(SheetName).Range("A1:A"&lastrow), xlFillDefault
NewWbk.Sheets(SheetName).Range("A:A").AutoFill NewWbk.Sheets(SheetName).Range("A:Z"), xlFillDefault
NewWbk.RefreshAll 'Refresh all formulas; Alternatively can use WScript.Sleep
NewWbk.Sheets(SheetName).Range("A1:Z"&lastrow).Copy
NewWbk.Sheets(SheetName).Range("A1").PasteSpecial - 4163 'https://docs.microsoft.com/eu-us/office/vba/api/excel.xlpastetype
objWbk.Sheets(SheetName).Range("A1:Z"&lastrow).Copy
NewWbk.Sheets(SheetName).Range("A1").PasteSpecial - 4122 'xlPasteFormats

'Save output
NewWbk.Password=""
NewWbk.SaveAs Path3
objDelete.DeleteFile(Path2) 
objExcel.Visible = True
objExcel.DisplayAlerts = True
objExcel.Quit

'Summary
MsgBox "Process started at "&StartTime&vbCrLf&"Process completed successfully at "&Now()&vbCrLf&vbCrLf&lastrow&" rows were processed"