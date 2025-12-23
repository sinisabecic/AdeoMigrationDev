Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then WScript.Quit
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set wbGlavni = objExcel.Workbooks.Open("C:\dev-vg\adeo-migracija\klijenti-uvoz.xlsm")
objExcel.Run "'" & wbGlavni.Name & "'!Module6.UvozKomitenataSpolja", CStr(objArgs(0))
If Err.Number <> 0 Then MsgBox "Greska: " & Err.Description Else MsgBox "Komitenti uspjesno izvezeni!", 64, "ADEO"