[Setup]
AppName=ADEO Migracija
AppVersion=1.0
; --- COPYRIGHT I LIČNI PODACI ---
AppCopyright=Copyright (C) 2025 Sinisa Becic
AppPublisher=Sinisa Becic
AppPublisherURL=https://github.com/sinisabecic
AppSupportURL=https://github.com/sinisabecic/AdeoMigrationDev
VersionInfoCompany=Sinisa Becic
VersionInfoDescription=Alat za migraciju artikala i komitenata iz ADEO sistema u VG sistem
VersionInfoVersion=1.0.0.1

DefaultDirName=C:\dev-vg\adeo-migracija
DefaultGroupName=ADEO Migracija
UninstallDisplayIcon={app}\artikli-uvoz.xlsm
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

; Slike za Wizard
WizardImageFile=logo_veliki.bmp
WizardSmallImageFile=logo_mali.bmp

[Languages]
; Koristimo podrazumijevani engleski, ali ćemo ga "pregaziti" našim prevodom ispod
Name: "default"; MessagesFile: "compiler:Default.isl"

[Messages]
; Ručni prevod ključnih djelova instalera
WelcomeLabel1=Dobrodošli u instalaciju ADEO Migracija alata
WelcomeLabel2=Ovaj program će instalirati alate za pripremu artikala i komitenata na VG sistem.
ClickNext=Kliknite na "Next" za nastavak ili "Cancel" za izlaz.
ButtonNext=&Dalje >
ButtonInstall=&Instaliraj
ButtonCancel=Otkaži
ButtonFinish=&Završi
DirBrowseText=Izaberite folder u koji želite instalirati program:
StatusInstalling=Instalacija u toku...
FinishedHeadingLabel=Instalacija je završena
FinishedLabelNoIcons=Program ADEO Migracija je uspješno instaliran.

[Files]
Source: "artikli-uvoz.xlsm"; DestDir: "{app}"; Flags: ignoreversion
Source: "klijenti-uvoz.xlsm"; DestDir: "{app}"; Flags: ignoreversion
Source: "artikli.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "komitenti.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userdesktop}\Artikli Uvoz"; Filename: "{app}\artikli-uvoz.xlsm"
Name: "{userdesktop}\Klijenti Uvoz"; Filename: "{app}\klijenti-uvoz.xlsm"

[Registry]
; --- ARTIKLI (Ikonice i Registry) ---
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\IzveziArtikleVG"; ValueType: string; ValueData: "Izvezi artikle (VG)"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\IzveziArtikleVG"; ValueName: "Icon"; ValueType: string; ValueData: "{app}\artikli.ico"
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\IzveziArtikleVG\command"; ValueType: string; ValueData: "wscript.exe ""{app}\UvozArtikala.vbs"" ""%1"""

Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\IzveziArtikleVG"; ValueType: string; ValueData: "Izvezi artikle (VG)"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\IzveziArtikleVG"; ValueName: "Icon"; ValueType: string; ValueData: "{app}\artikli.ico"
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\IzveziArtikleVG\command"; ValueType: string; ValueData: "wscript.exe ""{app}\UvozArtikala.vbs"" ""%1"""

; --- KOMITENTI (Ikonice i Registry) ---
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\IzveziKomitenteVG"; ValueType: string; ValueData: "Izvezi komitente (VG)"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\IzveziKomitenteVG"; ValueName: "Icon"; ValueType: string; ValueData: "{app}\komitenti.ico"
Root: HKCR; Subkey: "SystemFileAssociations\.xlsx\shell\IzveziKomitenteVG\command"; ValueType: string; ValueData: "wscript.exe ""{app}\UvozKomitenata.vbs"" ""%1"""

Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\IzveziKomitenteVG"; ValueType: string; ValueData: "Izvezi komitente (VG)"; Flags: uninsdeletekey
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\IzveziKomitenteVG"; ValueName: "Icon"; ValueType: string; ValueData: "{app}\komitenti.ico"
Root: HKCR; Subkey: "SystemFileAssociations\.xls\shell\IzveziKomitenteVG\command"; ValueType: string; ValueData: "wscript.exe ""{app}\UvozKomitenata.vbs"" ""%1"""

[Code]
// Procedura za dodavanje GitHub linka na samu stranicu instalera
procedure InitializeWizard;
var
  GitHubLabel: TNewStaticText;
begin
  GitHubLabel := TNewStaticText.Create(WizardForm);
  GitHubLabel.Parent := WizardForm;
  GitHubLabel.Left := ScaleX(10);
  GitHubLabel.Top := WizardForm.ClientHeight - ScaleY(30);
  GitHubLabel.Caption := 'Autor: Sinisa Becic | GitHub: github.com/sinisabecic';
  GitHubLabel.Font.Color := clGray;
  GitHubLabel.Cursor := crHand;
  // Opciono: Možeš dodati i klik funkciju da otvara browser, ali za sad je samo tekst
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  VBSContent: String;
begin
  if CurStep = ssPostInstall then
  begin
    // UvozArtikala.vbs
    VBSContent := 'Set objArgs = WScript.Arguments' + #13#10 +
                  'If objArgs.Count = 0 Then WScript.Quit' + #13#10 +
                  'On Error Resume Next' + #13#10 +
                  'Dim izbor, rezimRada' + #13#10 +
                  'izbor = MsgBox("Izaberite rezim izvoza:" & vbCrLf & vbCrLf & "[ YES ]  ---->  VELEPRODAJA" & vbCrLf & "[ NO  ]  ---->  MALOPRODAJA", vbYesNo + vbQuestion + vbSystemModal, "Izbor procedure")' + #13#10 +
                  'If izbor = 6 Then rezimRada = "VELEPRODAJA" Else rezimRada = "MALOPRODAJA"' + #13#10 +
                  'Set objExcel = CreateObject("Excel.Application")' + #13#10 +
                  'objExcel.Visible = False' + #13#10 +
                  'objExcel.DisplayAlerts = False' + #13#10 +
                  'Set wbGlavni = objExcel.Workbooks.Open("' + ExpandConstant('{app}') + '\artikli-uvoz.xlsm")' + #13#10 +
                  'objExcel.Run "''" & wbGlavni.Name & "''!Module14.UvozPodatakaSpolja", CStr(objArgs(0)), CStr(rezimRada)' + #13#10 +
                  'wbGlavni.Close False' + #13#10 +
                  'objExcel.Quit' + #13#10 +
                  'If Err.Number <> 0 Then MsgBox "Greska: " & Err.Description, 16, "ADEO" Else MsgBox "Artikli su uspjesno obradjeni!", 64, "ADEO"';
    SaveStringToFile(ExpandConstant('{app}\UvozArtikala.vbs'), VBSContent, False);

    // UvozKomitenata.vbs
    VBSContent := 'Set objArgs = WScript.Arguments' + #13#10 +
                  'If objArgs.Count = 0 Then WScript.Quit' + #13#10 +
                  'On Error Resume Next' + #13#10 +
                  'Set objExcel = CreateObject("Excel.Application")' + #13#10 +
                  'objExcel.Visible = False' + #13#10 +
                  'objExcel.DisplayAlerts = False' + #13#10 +
                  'Set wbGlavni = objExcel.Workbooks.Open("' + ExpandConstant('{app}') + '\klijenti-uvoz.xlsm")' + #13#10 +
                  'objExcel.Run "''" & wbGlavni.Name & "''!Module6.UvozKomitenataSpolja", CStr(objArgs(0))' + #13#10 +
                  'wbGlavni.Close False' + #13#10 +
                  'objExcel.Quit' + #13#10 +
                  'If Err.Number <> 0 Then MsgBox "Greska: " & Err.Description, 16, "ADEO" Else MsgBox "Komitenti su uspjesno obradjeni!", 64, "ADEO"';
    SaveStringToFile(ExpandConstant('{app}\UvozKomitenata.vbs'), VBSContent, False);
  end;
end;