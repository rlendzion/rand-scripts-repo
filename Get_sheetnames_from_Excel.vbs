'Pull sheetnames from a file and put them into another file

Set ObjExcel = CreateObject("Excel.Application")
Set wshShell = WScript.CreateObject("WScript.Shell")
strUserName = wshShell.ExpandEnvironmentStrings("%USERNAME%")
objExcel.Visible = True
objExcel.DisplayAlerts = False
Path1="C:\Users\" & strUserName & "\Downloads\Infile.xlsx"
Path2="C:\Users\" & strUserName & "\Downloads\TempList.xlsx"
Set objWbk = objExcel.Workbooks.Open(Path1)
Set NewWbk = objExcel.Workbooks.Add
For i = 1 To objWbk.Sheeets.Count
	NewWbk.Sheets(1).Range("A" & i+1).Value = objWbk.Sheets(i).Name
Next
NewWbk.SaveAs Path2
objExcel.Quit