'Unprotect Strictly Confidential Data (Smartcard Required)
Set objExcel = CreateObject("Excel.Application")
Set objDelete = CreateObject("Scripting.FileSystemObject")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Dim sInput, RepPeriod, FileName, dInput, pw, wsInput, SheetCnt
Dim ListNames: ListNames = "List of Sheet Names: "
StartTime=Now()

	'Define input parameters
RepPeriod = Year(Now()) & Month(Now()) 'Returns YYYYMM date format
sInput = InputBox("Provide folder name: ",,"folder")
If IsEmpty(sInput) Then 'If cancelled
	objExcel.Visible=True
	objDelete.DisplayAlerts=True
	WScript.Quit
Else
	FileName = InputBox("Provide file name with extension (.xls, .xlsx):")
	pw = InputBox("Enter password:")
	dInput = InputBox("Insert reporting period:",,cStr(RepPeriod))

	Folder="\\teams.cc.cnet.xxx.net@SSL\sites\"&sInput&"\"&dInput&"\"
	Path1=Folder&FileName
	Path2=Folder&"vNoPass.xlsx"
	Path3=Folder&"Unprotected_"&FileName

	'Open file
	Set objWbk = objExcel.Workbooks.Open(Path1,,,,cStr(pw))
	objWbk.Password=""
	objWbk.SaveAs Path2

	'Create the list of all sheets in a file
	SheetCnt = objWbk.Sheets.Count
	for i = 0 to SheetCnt
		ListNames = ListNames & vbCrLf & "- " & objWbk.Sheets(i).Name & ";"
	Next

	SheetName = objWbk.Sheets(1).Name 'Default SheetName
	wsInput = InputBox("Select sheet for further processing." & vbCrLf & vbCrLf & ListNames,,cStr(SheetName)) 'Select first sheet by default. Use list to select another sheet if needed
	Set NewWbk = objExcel.Workbooks.Add()
	NewWbk.Sheets(1).Name = SheetName 'Rename first sheet of a new file

	'Excel formula definition
	Loc="'https://teams.cc.cnet.xxx.net/sites/"&sInput&"/"&dInput&"/[vNoPass.xlsx]"
	Frm="=IF("&Loc&SheetName&"'!A1"&cStr("="&chr(34)&chr(34)&","&chr(34)&chr(34)&",")&Loc&SheetName&"'!A1"&cStr(")")

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
End If