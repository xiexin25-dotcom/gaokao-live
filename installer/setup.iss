; 吉林省高考志愿规划系统 · Inno Setup 安装脚本
; 需要先用 PyInstaller 生成 dist\gaokao\ 目录，再用 Inno Setup 编译此脚本

#define AppName      "吉林高考志愿规划"
#define AppVersion   "3.4"
#define AppPublisher "吉林高考志愿"
#define AppURL       "http://localhost:5000"
#define AppExeName   "gaokao.exe"
#define DataFile     "..\dist\gaokao"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={localappdata}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=.
OutputBaseFilename=吉林高考志愿规划_v3.4_安装包
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\{#AppExeName}
; 安装界面语言
ShowLanguageDialog=no

[Languages]
Name: "chs"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "在桌面创建快捷方式"; GroupDescription: "附加任务:"; Flags: unchecked

[Files]
; PyInstaller 打包产物（程序主体）
Source: "..\dist\gaokao\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; 数据文件（单独放，支持后续更新数据）
Source: "..\data\2026_jilin_gaokao_data.xlsx"; DestDir: "{app}\data"; Flags: ignoreversion onlyifdoesntexist

; 如果已有缓存一并打入（加速首次启动，可选）
; Source: "..\data\df_cache.pkl"; DestDir: "{app}\data"; Flags: ignoreversion onlyifdoesntexist

[Dirs]
; 确保 outputs 目录存在（可写）
Name: "{app}\outputs"
Name: "{app}\data"

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"
Name: "{group}\卸载 {#AppName}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
; 安装完成后自动启动
Filename: "{app}\{#AppExeName}"; Description: "立即启动程序"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 卸载时清理 outputs 和 cache（保留用户数据可注释掉）
Type: filesandordirs; Name: "{app}\outputs"
Type: files;         Name: "{app}\data\df_cache.pkl"
