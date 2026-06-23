; ============================================================
; MINI TOOL ADD-IN - INNO SETUP SCRIPT
; Author: Kane
; Version: 2.0 - Fixed for Inventor 2024+
; ============================================================

#define MyAppName "MiniTool Add-in"
#define MyAppVersion "2.0"
#define MyAppPublisher "Kane"
#define MyAppURL "mailto:replacefile.addin@gmail.com"

; Đường dẫn source (thay đổi theo máy của bạn)
#define SourcePath "D:\Setup\Ngoc\App cua Ngoc\Mini-Tool\Mini Tool\OpenCadDrawingAddin\bin\Release"

[Setup]
AppId={{65DB7978-5BCD-4388-BD7B-7EBA2C4B33D2}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Cài trực tiếp vào folder ApplicationPlugins (KHÔNG .bundle)
DefaultDirName={userappdata}\Autodesk\ApplicationPlugins\Mini-Tool
DisableDirPage=yes

DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

LicenseFile={#SourcePath}\license.txt

OutputDir={#SourcePath}\Installer
OutputBaseFilename=MiniToolAddin_Setup_v{#MyAppVersion}

Compression=lzma2/ultra64
SolidCompression=yes

WizardStyle=modern
PrivilegesRequired=lowest

UninstallDisplayIcon={app}\MiniTool.dll
UninstallDisplayName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Messages]
BeveledLabel=Mini Tool Add-in for Autodesk Inventor

[Files]
; Copy tất cả DLL (chính + phụ)
Source: "{#SourcePath}\*.dll"; DestDir: "{app}"; Flags: ignoreversion

; Copy file .addin - BẮT BUỘC tên đúng
Source: "{#SourcePath}\MiniTool.addin"; DestDir: "{app}"; Flags: ignoreversion

; Các file phụ
Source: "{#SourcePath}\*.config"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#SourcePath}\license.txt"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#SourcePath}\*.pdb"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

; KHÔNG copy PackageContents.xml nữa

[Icons]
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
  if FindWindowByClassName('InventorMainFrame') <> 0 then
  begin
    MsgBox('Please close Autodesk Inventor before installing.' + #13#10 +
           'Vui lòng đóng Autodesk Inventor trước khi cài đặt.', mbError, MB_OK);
    Result := False;
  end;
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  if FindWindowByClassName('InventorMainFrame') <> 0 then
  begin
    MsgBox('Please close Autodesk Inventor before uninstalling.' + #13#10 +
           'Vui lòng đóng Autodesk Inventor trước khi gỡ cài đặt.', mbError, MB_OK);
    Result := False;
  end;
end;

[UninstallDelete]
Type: filesandordirs; Name: "{app}"