; 图片批量整理工具安装脚本
; 使用Inno Setup制作

[Setup]
; 应用程序基本信息
AppName=图片批量整理工具
AppVersion=1.9
AppPublisher=开发者
AppPublisherURL=
AppSupportURL=
AppUpdatesURL=
DefaultDirName={autopf}\图片批量整理工具
DefaultGroupName=图片批量整理工具
AllowNoIcons=yes
LicenseFile=
OutputDir=installer_output
OutputBaseFilename=图片批量整理工具_安装包
SetupIconFile=logo.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

; 支持的Windows版本
MinVersion=6.1sp1

; 语言设置
[Languages]
Name: "chinesesimp"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

; 任务设置
[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

; 文件安装
[Files]
Source: "dist\图片批量整理工具.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "logo.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion
; 注意：如果有其他依赖文件，请在这里添加

; 图标创建
[Icons]
Name: "{group}\图片批量整理工具"; Filename: "{app}\图片批量整理工具.exe"; IconFilename: "{app}\logo.ico"
Name: "{group}\{cm:UninstallProgram,图片批量整理工具}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\图片批量整理工具"; Filename: "{app}\图片批量整理工具.exe"; IconFilename: "{app}\logo.ico"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\图片批量整理工具"; Filename: "{app}\图片批量整理工具.exe"; IconFilename: "{app}\logo.ico"; Tasks: quicklaunchicon

; 运行设置
[Run]
Filename: "{app}\图片批量整理工具.exe"; Description: "{cm:LaunchProgram,图片批量整理工具}"; Flags: nowait postinstall skipifsilent

; 卸载设置
[UninstallDelete]
Type: filesandordirs; Name: "{app}"

; 自定义消息
[Messages]
WelcomeLabel2=这将在您的计算机上安装 [name/ver]。%n%n图片批量整理工具是一个简单易用的图片管理软件，可以帮助您快速整理和复制图片文件。%n%n建议您在继续安装前关闭所有其他应用程序。
ClickNext=单击"下一步"继续，或单击"取消"退出安装程序。
BeveledLabel=图片批量整理工具 v1.9

[Code]
// 检查是否已安装
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

// 安装完成后的处理
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // 可以在这里添加安装后的处理逻辑
  end;
end;