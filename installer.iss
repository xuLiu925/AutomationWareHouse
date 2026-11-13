; FreightPilot 安装包（Inno Setup 6）。由 build.py 调用 ISCC 编译，不要手动改版本号：
;   ISCC /DAppVersion=2026.11.13.1200 /DSourceDir=dist\FreightPilot /DOutputDir=dist installer.iss
;
; 装到当前用户自己的目录（%LOCALAPPDATA%\Programs\FreightPilot），不需要管理员权限：
; 程序的数据文件（dachser_profile、pallet_params.json、飞书登录等）和在线更新都写在程序目录里，
; 装到 Program Files 普通用户写不进去。

#ifndef AppVersion
  #define AppVersion "0.0.0.0"
#endif
#ifndef SourceDir
  #define SourceDir "dist\FreightPilot"
#endif
#ifndef OutputDir
  #define OutputDir "dist"
#endif
#define AppName "FreightPilot"
#define AppExe "FreightPilot.exe"

[Setup]
; AppId 不能改：重新安装 / 装新版时靠它找到旧的安装，覆盖到同一个文件夹
AppId={{A8F1EBB3-7169-436D-98AA-D73F329C6125}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=Xu Liu
DefaultDirName={localappdata}\Programs\{#AppName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir={#OutputDir}
OutputBaseFilename={#AppName}-Setup-{#AppVersion}
SetupIconFile=assets\xlfp.ico
UninstallDisplayIcon={app}\{#AppExe}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no

[Languages]
; Inno Setup 自带的语言里没有简体中文：项目里放了 installer\ChineseSimplified.isl 就用它，否则英文界面
#ifexist AddBackslash(SourcePath) + "installer\ChineseSimplified.isl"
Name: "zh"; MessagesFile: "installer\ChineseSimplified.isl"
#elif FileExists(AddBackslash(CompilerPath) + "Languages\ChineseSimplified.isl")
Name: "zh"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
#else
Name: "en"; MessagesFile: "compiler:Default.isl"
#endif

[Tasks]
Name: "desktopicon"; Description: "在桌面创建快捷方式"; GroupDescription: "快捷方式:"

[Files]
; 旧名字的 exe 只给旧版在线更新过渡用，新安装不需要
Source: "{#SourceDir}\*"; DestDir: "{app}"; Excludes: "XL FreightPilot.exe,pallet_params.json"; \
  Flags: ignoreversion recursesubdirs createallsubdirs
; 产品参数：已有就不覆盖（程序运行时会从服务器同步最新的）
Source: "{#SourceDir}\pallet_params.json"; DestDir: "{app}"; Flags: onlyifdoesntexist uninsneveruninstall

[Dirs]
Name: "{app}\危险品文件"; Flags: uninsneveruninstall

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExe}"; Description: "立即运行 {#AppName}"; WorkingDir: "{app}"; \
  Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 在线更新后新增的程序文件不在安装记录里，卸载时一并删掉；
; 用户数据（dachser_profile、飞书登录/缓存、pallet_params.json、危险品文件、xlfp_config.json）保留
Type: filesandordirs; Name: "{app}\_update"
Type: files; Name: "{app}\*.exe"
Type: files; Name: "{app}\*.dll"
Type: files; Name: "{app}\*.pyd"
Type: files; Name: "{app}\version.txt"
Type: filesandordirs; Name: "{app}\playwright"
Type: filesandordirs; Name: "{app}\assets"
