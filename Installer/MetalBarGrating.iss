; ============================================================
; Metal Bar Grating — Inno Setup Installer Script
;
; Produces a single .exe installer that:
;   1. Copies the .bundle to Autodesk ApplicationPlugins
;   2. Shows license/readme if present
;   3. Registers uninstall entry in Windows
;
; Prerequisites:
;   - Inno Setup 6.x (free): https://jrsoftware.org/isinfo.php
;   - Build the solution in Release first so the Bundle folder
;     is populated at:
;     HandMadeGratingAddinVB\bin\Release\net48\Bundle\
;
; To compile:  Right-click this .iss → Compile (or iscc.exe)
; ============================================================

#define MyAppName        "Metal Bar Grating"
#define MyAppVersion     "1.5.17"
#define MyAppPublisher   "GP INC"
#define MyAppURL         ""
#define MyAppGuid        "{{37b59293-54b3-43f0-8166-ab23d5cf61ed}"
#define BundleName       "HandMadeGratingAddinVB.bundle"

; Path to the built output (relative to this .iss file)
#define ReleaseDir       "..\HandMadeGratingAddinVB\bin\Release\net48"

[Setup]
AppId={#MyAppGuid}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={commonappdata}\Autodesk\ApplicationPlugins\{#BundleName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\Output
OutputBaseFilename=MetalBarGrating_v{#MyAppVersion}_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\Contents\Resources\GratingIcon32.png
ArchitecturesInstallIn64BitMode=x64
WizardStyle=modern
; Force overwrite of previous version files
UsePreviousAppDir=yes
CloseApplications=force
RestartApplications=no
; Uncomment below if you add a license file:
; LicenseFile=..\LICENSE.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; PackageContents.xml goes at the bundle root
Source: "..\HandMadeGratingAddinVB\PackageContents.xml"; DestDir: "{app}"; Flags: ignoreversion
; Everything else goes under Contents\
Source: "{#ReleaseDir}\HandMadeGratingAddinVB.dll"; DestDir: "{app}\Contents"; Flags: ignoreversion restartreplace
Source: "{#ReleaseDir}\HandMadeGratingAddinVB.Inventor.addin"; DestDir: "{app}\Contents"; Flags: ignoreversion
Source: "{#ReleaseDir}\Resources\*"; DestDir: "{app}\Contents\Resources"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Code]
// Returns True when any supported Inventor install folder or registry key exists.
function InventorYearDirExists(const Pf, Pf32, Year: String): Boolean;
begin
  Result :=
    DirExists(Pf + '\Autodesk\Inventor ' + Year) or
    DirExists(Pf32 + '\Autodesk\Inventor ' + Year);
end;

function IsInventorInstalled(): Boolean;
var
  Pf, Pf32: String;
begin
  Pf := ExpandConstant('{pf}');
  Pf32 := ExpandConstant('{pf32}');

  // Year-suffixed install folders (2024+). Add new years here as releases ship.
  Result :=
    InventorYearDirExists(Pf, Pf32, '2024') or
    InventorYearDirExists(Pf, Pf32, '2025') or
    InventorYearDirExists(Pf, Pf32, '2026') or
    InventorYearDirExists(Pf, Pf32, '2027') or
    InventorYearDirExists(Pf, Pf32, '2028');

  if Result then
    Exit;

  // Generic fallback: HKLM\SOFTWARE\Autodesk\Inventor
  Result :=
    RegKeyExists(HKLM, 'SOFTWARE\Autodesk\Inventor') or
    RegKeyExists(HKLM64, 'SOFTWARE\Autodesk\Inventor');
end;

function InitializeSetup(): Boolean;
begin
  Result := True;

  if not IsInventorInstalled() then
  begin
    if MsgBox('Autodesk Inventor 2024 or later was not detected.' + #13#10 +
              'The add-in requires Inventor to function.' + #13#10#13#10 +
              'Continue installation anyway?',
              mbConfirmation, MB_YESNO) = IDNO then
    begin
      Result := False;
    end;
  end;
end;

// Close Inventor before install if running
function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  // Warn if Inventor is running
  if FindWindowByClassName('InventorMainFrame') <> 0 then
  begin
    if MsgBox('Autodesk Inventor is currently running.' + #13#10 +
              'Please close Inventor before continuing.' + #13#10#13#10 +
              'Click OK after closing Inventor, or Cancel to abort.',
              mbError, MB_OKCANCEL) = IDCANCEL then
    begin
      Result := 'Installation cancelled — Inventor must be closed first.';
    end;
  end;
end;

// Clean up old files on uninstall
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  BundlePath: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    BundlePath := ExpandConstant('{app}');
    if DirExists(BundlePath) then
      DelTree(BundlePath, True, True, True);
  end;
end;
