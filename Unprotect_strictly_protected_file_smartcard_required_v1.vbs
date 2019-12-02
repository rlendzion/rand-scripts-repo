'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Case description: 
'There is a new Excel file with highly sensitive data submitted regularly to a project team. The file and its sheets have
'their names changed by the data providers using a specific naming convention. The file's extra security restrictions prevent 
'it from simply dropping its password in Excel and saving as an unprotected file. Other tools and macros would still identify 
'the file as zipped/in use because of other constraints preventing it from further processing. 
'The only working solution is to open a new Excel instance and copy the restricted file's content with a formula. The script 
'can be run from any directory multiple times as it overwrites an unprotected file if it already exists.
'Author: Robert Lendzion
'Date:2019-11-29
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set objExcel = CreateObject("Excel.Application")
Set objDelete = CreateObject("Scripting.FileSystemObject")
objExcel.Visible = False
objExcel.DisplayAlerts = False

'Date/Time
StartTime=Now()
PrMv=DatePart("m",DateAdd("m",-1,"t")) 'Previous reporting month value
CMv=DatePart("m",DateAdd("m",0,"t")) 'Current reporting month value
PrMLnm=MonthName(PrMv) 'Previous reporting month long name
PrMSnm=MonthName(PrMv,True) 'Previous reporting month short name
CMSnm=MonthName(CMv,True) 'Current reporting month short name
CYr=DatePart("yyyy",DateAdd("m",0,t)) 'Current year current month value
CYrPrM=DatePart("yyyy",DateAdd("m",0,t)) 'Current year previous month value
PrYr=DatePart("yyyy",DateAdd("m",-11,t)) 'Previous year
LastDayPrM=DatePart("d",DateSerial(PrYr,CMv,0) 'Prior year current month last day

'Strings
FolderName=CYr&right(0&CMv,2) 'e.g. 201909
FileName= "File - "&PrMSnm&" "&CYrPrM&" (1 "&CMSnm&" "&PrYr&" - "&LastDayPrM&" "&PrMSnm&" "&CYrPrM&").xlsx"  'e.g. [File - Aug 2019 (1 Sep 2018 - 31 Aug 2019).xlsx]
SheetName=PrMLnm&" "&CYr&" - SheetName" 'e.g. [August 2019 - SheetName]
pw="password"
Folder="\\teams.cc.cnet.xxx.net@SSL\sites\folder\"&FolderName&"\"
Path1=Folder&FileName
Path2=Folder&"vNoPass.xlsx" 'Temporary file
Path3=Folder&"Unprotected_"&FileName 'Final unprotected file

'Excel formula definition
Loc="'https://teams.cc.cnet.xxx.net/sites/folder/"&FolderName&"/[vNoPass.xlsx]" 'When a file is located on a SharePoint site, the formula uses https reference
Frm="=IF("&Loc&SheetName&"'!A1"&cStr("="&chr(34)&chr(34)&","&chr(34)&chr(34)&",")&Loc&SheetName&"'!A1"&cStr(")")

Set objWbk = objExcel.Workbooks.Open(Path,,,,cStr(pw))
objWbk.Password=""
objWbk.SaveAs Path2
Set NewWbk=objExcel.Workbook.Add()
NewWbk.Sheets(1).Name = SheetName
lastrow=objWbk.Worksheets(SheetName).UsedRange.Rows.Count + 1
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