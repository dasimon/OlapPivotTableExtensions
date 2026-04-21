#define MyAppName "OLAP PivotTable Extensions"
#define MyAppVersion "0.9.8"
#define MyAppPublisher "dasimon"
#define MyAppURL "https://github.com/dasimon/OlapPivotTableExtensions"
#define MyAppExeName "OlapPivotTableExtensions.vsto"
#define BinDir "OlapPivotTableExtensions\bin\Release"

[Setup]
AppId={{E501A716-3610-450E-8D98-92A1E0003D1E}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={autopf}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer
OutputBaseFilename=OlapPivotTableExtensions-v{#MyAppVersion}-setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=6.1
UninstallDisplayIcon={app}\OlapPivotTableExtensions.dll

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Files]
; Main assembly and manifest
Source: "{#BinDir}\OlapPivotTableExtensions.dll";        DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\OlapPivotTableExtensions.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\OlapPivotTableExtensions.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\OlapPivotTableExtensions.vsto";       DestDir: "{app}"; Flags: ignoreversion

; NLog
Source: "{#BinDir}\NLog.dll";    DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\NLog.config"; DestDir: "{app}"; Flags: ignoreversion

; ADOMD + identity
Source: "{#BinDir}\Microsoft.AnalysisServices.AdomdClient.dll";      DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\Microsoft.AnalysisServices.SPClient.Interfaces.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\Microsoft.Identity.Client.dll";                   DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\Microsoft.IdentityModel.Abstractions.dll";        DestDir: "{app}"; Flags: ignoreversion

; Office tools runtime
Source: "{#BinDir}\Microsoft.Office.Tools.dll";                  DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\Microsoft.Office.Tools.Common.dll";           DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\Microsoft.Office.Tools.Excel.dll";            DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\Microsoft.Office.Tools.v4.0.Framework.dll";   DestDir: "{app}"; Flags: ignoreversion
Source: "{#BinDir}\Microsoft.VisualStudio.Tools.Applications.Runtime.dll"; DestDir: "{app}"; Flags: ignoreversion

; COM interop
Source: "{#BinDir}\adodb.dll"; DestDir: "{app}"; Flags: ignoreversion

; ADOMD localized resource DLLs
Source: "{#BinDir}\ar\*";    DestDir: "{app}\ar";    Flags: ignoreversion
Source: "{#BinDir}\bg\*";    DestDir: "{app}\bg";    Flags: ignoreversion
Source: "{#BinDir}\ca\*";    DestDir: "{app}\ca";    Flags: ignoreversion
Source: "{#BinDir}\cs\*";    DestDir: "{app}\cs";    Flags: ignoreversion
Source: "{#BinDir}\da\*";    DestDir: "{app}\da";    Flags: ignoreversion
Source: "{#BinDir}\de\*";    DestDir: "{app}\de";    Flags: ignoreversion
Source: "{#BinDir}\el\*";    DestDir: "{app}\el";    Flags: ignoreversion
Source: "{#BinDir}\es\*";    DestDir: "{app}\es";    Flags: ignoreversion
Source: "{#BinDir}\et\*";    DestDir: "{app}\et";    Flags: ignoreversion
Source: "{#BinDir}\eu\*";    DestDir: "{app}\eu";    Flags: ignoreversion
Source: "{#BinDir}\fi\*";    DestDir: "{app}\fi";    Flags: ignoreversion
Source: "{#BinDir}\fr\*";    DestDir: "{app}\fr";    Flags: ignoreversion
Source: "{#BinDir}\gl\*";    DestDir: "{app}\gl";    Flags: ignoreversion
Source: "{#BinDir}\he\*";    DestDir: "{app}\he";    Flags: ignoreversion
Source: "{#BinDir}\hi\*";    DestDir: "{app}\hi";    Flags: ignoreversion
Source: "{#BinDir}\hr\*";    DestDir: "{app}\hr";    Flags: ignoreversion
Source: "{#BinDir}\hu\*";    DestDir: "{app}\hu";    Flags: ignoreversion
Source: "{#BinDir}\id\*";    DestDir: "{app}\id";    Flags: ignoreversion
Source: "{#BinDir}\it\*";    DestDir: "{app}\it";    Flags: ignoreversion
Source: "{#BinDir}\ja\*";    DestDir: "{app}\ja";    Flags: ignoreversion
Source: "{#BinDir}\kk\*";    DestDir: "{app}\kk";    Flags: ignoreversion
Source: "{#BinDir}\ko\*";    DestDir: "{app}\ko";    Flags: ignoreversion
Source: "{#BinDir}\lt\*";    DestDir: "{app}\lt";    Flags: ignoreversion
Source: "{#BinDir}\lv\*";    DestDir: "{app}\lv";    Flags: ignoreversion
Source: "{#BinDir}\ms\*";    DestDir: "{app}\ms";    Flags: ignoreversion
Source: "{#BinDir}\nl\*";    DestDir: "{app}\nl";    Flags: ignoreversion
Source: "{#BinDir}\no\*";    DestDir: "{app}\no";    Flags: ignoreversion
Source: "{#BinDir}\pl\*";    DestDir: "{app}\pl";    Flags: ignoreversion
Source: "{#BinDir}\pt\*";    DestDir: "{app}\pt";    Flags: ignoreversion
Source: "{#BinDir}\pt-PT\*"; DestDir: "{app}\pt-PT"; Flags: ignoreversion
Source: "{#BinDir}\ro\*";    DestDir: "{app}\ro";    Flags: ignoreversion
Source: "{#BinDir}\ru\*";    DestDir: "{app}\ru";    Flags: ignoreversion
Source: "{#BinDir}\sk\*";    DestDir: "{app}\sk";    Flags: ignoreversion
Source: "{#BinDir}\sl\*";    DestDir: "{app}\sl";    Flags: ignoreversion
Source: "{#BinDir}\sr-Cyrl\*"; DestDir: "{app}\sr-Cyrl"; Flags: ignoreversion
Source: "{#BinDir}\sr-Latn\*"; DestDir: "{app}\sr-Latn"; Flags: ignoreversion
Source: "{#BinDir}\sv\*";    DestDir: "{app}\sv";    Flags: ignoreversion
Source: "{#BinDir}\th\*";    DestDir: "{app}\th";    Flags: ignoreversion
Source: "{#BinDir}\tr\*";    DestDir: "{app}\tr";    Flags: ignoreversion
Source: "{#BinDir}\uk\*";    DestDir: "{app}\uk";    Flags: ignoreversion

[Run]
Filename: "{commonpf32}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"; \
  Parameters: "/i ""{app}\OlapPivotTableExtensions.vsto"""; \
  StatusMsg: "Registering Excel add-in..."; \
  Flags: waituntilterminated

[UninstallRun]
Filename: "{commonpf32}\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"; \
  Parameters: "/u ""{app}\OlapPivotTableExtensions.vsto"""; \
  Flags: waituntilterminated

[Code]
function IsDotNet48Installed(): Boolean;
var
  release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM,
    'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
    'Release', release) and (release >= 528040);
end;

function InitializeSetup(): Boolean;
begin
  if not IsDotNet48Installed() then
  begin
    MsgBox(
      'OLAP PivotTable Extensions requires .NET Framework 4.8.' + #13#10 +
      'Please install it from https://dotnet.microsoft.com/download/dotnet-framework/net48 and run this installer again.',
      mbError, MB_OK);
    Result := False;
  end
  else
    Result := True;
end;
