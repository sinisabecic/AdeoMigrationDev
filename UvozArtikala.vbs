Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then WScript.Quit
On Error Resume Next
izbor = MsgBox("Izaberite rezim izvoza:" & vbCrLf & vbCrLf & "[ YES ]  ---->  VELEPRODAJA" & vbCrLf & "[ NO  ]  ---->  MALOPRODAJA", vbYesNo + vbQuestion + vbSystemModal, "Izbor procedure")
If izbor = 6 Then rezimRada = "VELEPRODAJA" Else rezimRada = "MALOPRODAJA"
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set wbGlavni = objExcel.Workbooks.Open("C:\dev-vg\adeo-migracija\artikli-uvoz.xlsm")
objExcel.Run "'" & wbGlavni.Name & "'!Module14.UvozPodatakaSpolja", CStr(objArgs(0)), CStr(rezimRada)
If Err.Number <> 0 Then MsgBox "Greska: " & Err.Description Else MsgBox "Artikli uspjesno izvezeni!", 64, "ADEO"