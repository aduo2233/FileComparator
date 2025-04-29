[Setup]
AppName=文档重复率检测器
AppVersion=2.0
DefaultDirName={pf}\DocComparator
OutputDir=.\build
OutputBaseFilename=DocComparator_Setup
Compression=lzma2
SolidCompression=yes

[Files]
Source: "dist\DocComparator.exe"; DestDir: "{app}"
Source: "icon.ico"; DestDir: "{app}"

[Icons]
Name: "{commonprograms}\文档查重工具"; Filename: "{app}\DocComparator.exe"; IconFilename: "{app}\icon.ico"