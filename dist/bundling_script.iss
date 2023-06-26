; Bundling Script

#define MyDistributionFiles "E:My Stuff\My Project Center\Work\IntakeTool\pkging\dist"
#define MyBuildDir "E:My Stuff\My Project Center\Work\IntakeTool\pkging\Build"
#define MyAppName "IntakeTool"
#define MyAppVersion "v1.0.4"
; #define MyAppVersion GetEnv('APPVERSIONTEXT')
#define MyAppVersionFileName "v1_0_4"
; #define MyAppVersionFileName GetEnv('APPVERSIONFILE')
#define MyAppURL "https://github.com/Samlant/IntakeTool"
#define MyAppExeName "IntakeTool.exe"
#define MyAppIcoName "intake_tool.ico"
#define MyInstallIcoName "install.ico"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppVersionFileName}
AppPublisherURL={#MyAppURL}
DefaultDirName={localappdata}\work-tools
DefaultGroupName=Work-Tools
OutputDir={#MyBuildDir}
OutputBaseFilename={#MyAppName}-{#MyAppVersionFileName}-Setup
DisableDirPage=yes
SetupIconFile={#MyDistributionFiles}\{#MyInstallIcoName}
LicenseFile={#MyDistributionFiles}\LICENSE.txt
Compression=lzma
SolidCompression=yes
WizardStyle=modern
DisableStartupPrompt=yes
DisableReadyPage=yes
PrivilegesRequired=lowest

[Files]
Source: "{#MyDistributionFiles}\{#MyAppExeName}"; DestDir: "{app}"
Source: "{#MyDistributionFiles}\{#MyAppIcoName}"; DestDir: "{app}"
Source: "{#MyDistributionFiles}\README.html"; Flags: isreadme; DestDir: "{app}"

[Tasks]
Name: autoRunFile; Description: "Auto-run on Windows Start-up";
Name: desktopicon; Description: "Create a &desktop icon";
Name: startmenu; Description: "Create a Start Menu folder";
Name: quicklaunchicon; Description: "Create a &Quick Launch icon";

[Icons]
Name: "{group}\IntakeTool"; Filename: "{app}\IntakeTool.exe"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppIcoName}"; Tasks: desktopicon
Name: "{group}\{cm:ProgramOnTheWeb, {#MyAppName}}"; Filename: "{#MyAppURL}"; Tasks: startmenu
Name: "{group}\{cm:UninstallProgram, {#MyAppName}}"; Filename: "{uninstallexe}"; Tasks: startmenu
Name: "{group}\{cm:UninstallProgram, {#MyAppName}}"; Filename: "{uninstallexe}"; Tasks: startmenu
Name: "{userstartup}\IntakeTool"; Filename: "{app}\{#MyAppExeName}"; Tasks: autoRunFile






