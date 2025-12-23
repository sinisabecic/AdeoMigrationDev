Option Explicit

Dim fso, shell, currentDir, installDir, adeoDir
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' 1. DEFINISANJE PUTANJA
currentDir = fso.GetParentFolderName(WScript.ScriptFullName) & "\"
installDir = "C:\dev-vg\"
adeoDir = installDir & "adeo-migracija\"

On Error Resume Next

' 2. KREIRANJE FOLDERA
If Not fso.FolderExists(installDir) Then fso.CreateFolder(installDir)
If Not fso.FolderExists(adeoDir) Then fso.CreateFolder(adeoDir)

' 3. KOPIRANJE XLSM FAJLOVA
CopyFile currentDir & "artikli-uvoz.xlsm", adeoDir
CopyFile currentDir & "klijenti-uvoz.xlsm", adeoDir

' 4. KREIRANJE DINAMIČKOG: UvozArtikala.vbs
Dim vbsArtikli
vbsArtikli = "Set objArgs = WScript.Arguments" & vbCrLf & _
"If objArgs.Count = 0 Then WScript.Quit" & vbCrLf & _
"putanjaFajlaZaUvoz = objArgs(0)" & vbCrLf & _
"imeGlavnogFajla = ""artikli-uvoz.xlsm""" & vbCrLf & _
"putanjaDoGlavnog = """ & adeoDir & "artikli-uvoz.xlsm""" & vbCrLf & _
"Dim izbor, rezimRada" & vbCrLf & _
"izbor = MsgBox(""Izaberite rezim izvoza:"" & vbCrLf & vbCrLf & ""[ YES ]  ---->  VELEPRODAJA"" & vbCrLf & ""[ NO  ]  ---->  MALOPRODAJA"", vbYesNo + vbQuestion + vbSystemModal, ""Izbor procedure"")" & vbCrLf & _
"If izbor = 6 Then rezimRada = ""VELEPRODAJA"" Else rezimRada = ""MALOPRODAJA""" & vbCrLf & _
"On Error Resume Next" & vbCrLf & _
"Set objExcel = GetObject(, ""Excel.Application"")" & vbCrLf & _
"If Err.Number <> 0 Then Set objExcel = CreateObject(""Excel.Application"")" & vbCrLf & _
"Set wbGlavni = objExcel.Workbooks(imeGlavnogFajla)" & vbCrLf & _
"If wbGlavni Is Nothing Then Set wbGlavni = objExcel.Workbooks.Open(putanjaDoGlavnog)" & vbCrLf & _
"objExcel.Visible = True" & vbCrLf & _
"objExcel.Run ""'"" & wbGlavni.Name & ""'!Module14.UvozPodatakaSpolja"", CStr(putanjaFajlaZaUvoz), CStr(rezimRada)" & vbCrLf & _
"MsgBox ""Artikli su uspjesno izvezeni!"", 64, ""Status"""

CreateTextFile adeoDir & "UvozArtikala.vbs", vbsArtikli

' 5. KREIRANJE DINAMIČKOG: UvozKomitenata.vbs
Dim vbsKlijenti
vbsKlijenti = "Set objArgs = WScript.Arguments" & vbCrLf & _
"If objArgs.Count = 0 Then WScript.Quit" & vbCrLf & _
"putanjaFajlaZaUvoz = objArgs(0)" & vbCrLf & _
"imeGlavnogFajla = ""klijenti-uvoz.xlsm""" & vbCrLf & _
"putanjaDoGlavnog = """ & adeoDir & "klijenti-uvoz.xlsm""" & vbCrLf & _
"On Error Resume Next" & vbCrLf & _
"Set objExcel = GetObject(, ""Excel.Application"")" & vbCrLf & _
"If Err.Number <> 0 Then Set objExcel = CreateObject(""Excel.Application"")" & vbCrLf & _
"Set wbGlavni = objExcel.Workbooks(imeGlavnogFajla)" & vbCrLf & _
"If wbGlavni Is Nothing Then Set wbGlavni = objExcel.Workbooks.Open(putanjaDoGlavnog)" & vbCrLf & _
"objExcel.Visible = True" & vbCrLf & _
"objExcel.Run ""'"" & wbGlavni.Name & ""'!Module6.UvozKomitenataSpolja"", CStr(putanjaFajlaZaUvoz)" & vbCrLf & _
"MsgBox ""Komitenti su uspjesno izvezeni!"", 64, ""Gotovo"""

CreateTextFile adeoDir & "UvozKomitenata.vbs", vbsKlijenti

' 6. KREIRANJE I POKRETANJE REG FAJLOVA (Preko podfoldera adeo-migracija)
Dim regArtikli, regKomitenti, pathEscaped
pathEscaped = "C:\\dev-vg\\adeo-migracija\\" ' Duple kose crte za REG format

regArtikli = "Windows Registry Editor Version 5.00" & vbCrLf & vbCrLf & _
"[HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\IzveziArtikleVG]" & vbCrLf & _
"@=""Izvezi artikle (VG)""" & vbCrLf & _
"""Icon""=""excel.exe""" & vbCrLf & vbCrLf & _
"[HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\IzveziArtikleVG\command]" & vbCrLf & _
"@=""wscript.exe \""" & pathEscaped & "UvozArtikala.vbs\"" \""%1\"""""

regKomitenti = "Windows Registry Editor Version 5.00" & vbCrLf & vbCrLf & _
"[HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\IzveziKomitenteVG]" & vbCrLf & _
"@=""Izvezi komitente (VG)""" & vbCrLf & _
"""Icon""=""excel.exe""" & vbCrLf & vbCrLf & _
"[HKEY_CLASSES_ROOT\SystemFileAssociations\.xlsx\shell\IzveziKomitenteVG\command]" & vbCrLf & _
"@=""wscript.exe \""" & pathEscaped & "UvozKomitenata.vbs\"" \""%1\""""" & vbCrLf & vbCrLf & _
"[HKEY_CLASSES_ROOT\SystemFileAssociations\.xls\shell\IzveziKomitenteVG]" & vbCrLf & _
"@=""Izvezi komitente (VG)""" & vbCrLf & _
"""Icon""=""excel.exe""" & vbCrLf & vbCrLf & _
"[HKEY_CLASSES_ROOT\SystemFileAssociations\.xls\shell\IzveziKomitenteVG\command]" & vbCrLf & _
"@=""wscript.exe \""" & pathEscaped & "UvozKomitenata.vbs\"" \""%1\"""""

' Privremeno čuvanje i pokretanje .reg fajlova
CreateTextFile adeoDir & "context_menu_za_izvoz_artikala.reg", regArtikli
CreateTextFile adeoDir & "context_menu_za_izvoz_komitenata.reg", regKomitenti

' Tiho pokretanje regedita
shell.Run "regedit.exe /s """ & adeoDir & "context_menu_za_izvoz_artikala.reg""", 0, True
shell.Run "regedit.exe /s """ & adeoDir & "context_menu_za_izvoz_komitenata.reg""", 0, True

' 7. KREIRANJE PREČICA NA DESKTOPU
CreateShortcut "Artikli Uvoz", adeoDir & "artikli-uvoz.xlsm"
CreateShortcut "Klijenti Uvoz", adeoDir & "klijenti-uvoz.xlsm"

MsgBox "Instalacija uspesno zavrsena!" & vbCrLf & "Folder: C:\dev-vg\adeo-migracija\" & vbCrLf & "Context menu opcije su aktivne.", 64, "ADEO Migracija Setup"

' --- POMOĆNE FUNKCIJE ---
Sub CopyFile(source, dest)
    If fso.FileExists(source) Then fso.CopyFile source, dest, True
End Sub

Sub CreateTextFile(path, content)
    Dim f: Set f = fso.CreateTextFile(path, True)
    f.Write content
    f.Close
End Sub

Sub CreateShortcut(sName, target)
    Dim desktop, link
    desktop = shell.SpecialFolders("Desktop")
    Set link = shell.CreateShortcut(desktop & "\" & sName & ".lnk")
    link.TargetPath = target
    link.Save
End Sub