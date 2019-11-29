Set wshShell = WScript.CreateObject("WScript.Shell")
strUserName = wshShell.ExpandEnvironmentStrings("%USERNAME%")
Set objExcel = CreateObject("Excel.Application")
Set objDelete = CreateObject("Scripting.FileSystemObject")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Dim insURL
Dim Download
Dim objShell
regions = array("AMER","CAND","EMEA","JAPN","APAC","CHIN","GLOB")
Set NewWbk = objExcel.Workbooks.Add()
For each i in regions
	insURL = "https://my-app.com/console/region.htm?region=" & cStr(i)
	Downlaod = "https://my-app.com/console/download.htm?region=" & cStr(i) 'Before downloading the file, we need to be sure that the next page was loaded, so that the file for download is refreshed
	set ObjShell = CreateObject("Shell.Application")
	objShell.ShellExecute "chrome.exe", insURL, "", "", 1
	WScript.Sleep 5000 'Wait till the window appears so that the correct file gets downloaded
	objShell.ShellExecute "chrome.exe", Download, "", "", 1
	FileName = "download.xls" 'Note that the file will have the same name eveery time. That is why we need to remove it after we extracted data from it
	FullPath = "C:\Users\" & strUserName & "\Downloads\" & FileName
	WScript.Sleep 3000 'Let the file get saved
	set objWbk = objExcel.Workbooks.Open(cStr(FullPath))
	lastrow = objWbk.Worksheets("Latest 50 Run").UsedRange.Rows.Count
	'MsgBox "last row: " & lastrow 'for debugging only
	If i = "AMER" Then
		ObjWbk.Sheets("Latest 50 Run").Range("A1:L"&lastrow).Copy 'Copy with headers
		Else objWbk.Sheets("Latest 50 Run").Range("A2:L"&lastrow).Copy 'Copy without headers
	End If
	If i = "AMER" Then
		target = NewWbk.Worksheets("Sheet1").UsedRange.Rows.Count 'Start from 1st row
		Else target = NewWbk.Worksheets("Sheet1").UsedRange.Rows.Count +1 'Append one row after the last active row
	End If
	'MsgBox "target: " & target 'for debugging only
	NewWbk.Sheets("Sheet1").Range("A"&target).PasteSpecial - 4104
	objWbk.Close
	objWbk.DeleteFile(FullPath) 'Delete saved file before we move on to downloading another one
Next
NewWbk.SaveAs "\\PROD.SERVER.COMP.NET\UserData" & strUserName & "\Desktop\Mydata.xlsx"
MsgBox "Process finished successfully. The file has been saved on the desktop"
objExcel.Quit
objExcel.Visible = TRUE