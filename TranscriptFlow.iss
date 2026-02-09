; TranscriptFlow Inno Setup Script
; Creates a professional Windows installer with proper permissions and directory structure
; Supports both x86 and x64 architectures with automatic detection

#define MyAppName "TranscriptFlow Pro"
#define MyAppVersion "1.0.9"
#define MyAppPublisher "TranscriptFlow"
#define MyAppURL "https://transcriptflow.com"
#define MyAppExeName "TranscriptFlow.exe"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
AppId={{A7B8C9D0-1234-5678-90AB-CDEF12345678}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
; Automatically selects Program Files or Program Files (x86) based on architecture
; Use {commonpf} to force Program Files regardless of initial privilege level
DefaultDirName={commonpf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=
InfoBeforeFile=
InfoAfterFile=
OutputDir=installer_output
OutputBaseFilename=TranscriptFlow_Setup_{#MyAppVersion}
SetupIconFile=app_icon.ico
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
; CRITICAL: Set to 'none' to avoid noisy startup warnings, but disable 'dialog' overrides
; to suppress the "Select Install Mode" (AppData) dialog. 
; This results in a clean, forced Program Files installation with standard UAC.
PrivilegesRequired=none
PrivilegesRequiredOverridesAllowed=commandline
; Don't use previous directory to ensure we move away from AppData if needed
UsePreviousAppDir=no
; Architecture settings - supports both x86 and x64
; If you want x64 only, keep these lines. For both architectures, comment them out.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Uninstall settings
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
; Version info
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Installer
VersionInfoCopyright=Copyright (C) 2026 {#MyAppPublisher}
; Disable directory page if you want to force Program Files installation
DisableDirPage=no
; Restart if needed (usually not required for this app)
RestartIfNeededByRun=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked; OnlyBelowVersion: 6.1; Check: not IsAdminInstallMode

[Files]
; Main application files from PyInstaller output
Source: "dist\TranscriptFlow\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "HELP.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "splash.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "Jameel Noori Nastaleeq.ttf"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
var
  AppDataDirPage: TInputDirWizardPage;

function IsAppDataWritable(): Boolean;
var
  TestFile: String;
  AppDataDir: String;
begin
  Result := False;
  AppDataDir := ExpandConstant('{userappdata}\TranscriptFlow');
  TestFile := AppDataDir + '\write_test.tmp';
  
  try
    // Try to create a test file
    if SaveStringToFile(TestFile, 'test', False) then
    begin
      DeleteFile(TestFile);
      Result := True;
    end;
  except
    Result := False;
  end;
end;

procedure InitializeWizard;
begin
  // No additional wizard pages needed - AppData setup is automatic
  // Could add custom pages here if needed in the future
end;

function CreateDirectoryWithCheck(DirPath: String): Boolean;
var
  ErrorCode: Integer;
begin
  Result := True;
  
  if not DirExists(DirPath) then
  begin
    if not CreateDir(DirPath) then
    begin
      // Try with ForceDirectories
      if not ForceDirectories(DirPath) then
      begin
        MsgBox('Failed to create directory: ' + DirPath + #13#10 + 
               'Please ensure you have proper permissions.', 
               mbError, MB_OK);
        Result := False;
      end;
    end;
  end;
end;

procedure CopyDictionaryFiles(SourceDir, DestDir: String);
var
  FindRec: TFindRec;
  SourcePath, DestPath: String;
begin
  if not DirExists(SourceDir) then
    Exit;
    
  if FindFirst(SourceDir + '\*.*', FindRec) then
  begin
    try
      repeat
        if (FindRec.Name <> '.') and (FindRec.Name <> '..') then
        begin
          SourcePath := SourceDir + '\' + FindRec.Name;
          DestPath := DestDir + '\' + FindRec.Name;
          
          // Only copy if destination doesn't exist (preserve user's files)
          if not FileExists(DestPath) then
          begin
            if not FileCopy(SourcePath, DestPath, False) then
            begin
              Log('Warning: Failed to copy ' + SourcePath + ' to ' + DestPath);
            end;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  AppDataDir: String;
  BackupDir: String;
  DictsDir: String;
  CustomDictsDir: String;
  WaveformCacheDir: String;
  LogsDir: String;
  Success: Boolean;
begin
  if CurStep = ssPostInstall then
  begin
    Log('Creating AppData directory structure...');
    
    // Get AppData paths
    AppDataDir := ExpandConstant('{userappdata}\TranscriptFlow');
    BackupDir := AppDataDir + '\BACKUP';
    DictsDir := AppDataDir + '\dicts';
    CustomDictsDir := AppDataDir + '\custom_dicts';
    WaveformCacheDir := AppDataDir + '\waveform_cache';
    LogsDir := AppDataDir + '\logs';
    
    Success := True;
    
    // Create main AppData directory
    if not CreateDirectoryWithCheck(AppDataDir) then
      Success := False;
    
    // Create subdirectories
    if Success then
    begin
      if not CreateDirectoryWithCheck(BackupDir) then
        Success := False;
      if not CreateDirectoryWithCheck(DictsDir) then
        Success := False;
      if not CreateDirectoryWithCheck(CustomDictsDir) then
        Success := False;
      if not CreateDirectoryWithCheck(WaveformCacheDir) then
        Success := False;
      if not CreateDirectoryWithCheck(LogsDir) then
        Success := False;
    end;
    
    if Success then
    begin
      Log('AppData directories created successfully');
      
      // Copy bundled dictionaries to AppData (only if they don't exist)
      Log('Copying dictionary files...');
      CopyDictionaryFiles(ExpandConstant('{app}\dicts'), DictsDir);
      CopyDictionaryFiles(ExpandConstant('{app}\custom_dicts'), CustomDictsDir);
      
      // Verify write permissions
      if IsAppDataWritable() then
      begin
        Log('AppData directory has write permissions - OK');
      end
      else
      begin
        MsgBox('Warning: AppData directory may not have write permissions.' + #13#10 +
               'The application may not be able to save backups or download dictionaries.' + #13#10#13#10 +
               'Directory: ' + AppDataDir, 
               mbInformation, MB_OK);
      end;
    end
    else
    begin
      MsgBox('Failed to create some application directories.' + #13#10 +
             'The application may not function correctly.' + #13#10#13#10 +
             'Please check your permissions and try again.', 
             mbError, MB_OK);
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDataDir: String;
  DialogResult: Integer;
begin
  if CurUninstallStep = usUninstall then
  begin
    AppDataDir := ExpandConstant('{userappdata}\TranscriptFlow');
    
    // Ask user if they want to keep their data
    if DirExists(AppDataDir) then
    begin
      DialogResult := MsgBox(
        'Do you want to remove your personal data (backups, dictionaries, settings)?' + #13#10 + #13#10 +
        'Click Yes to remove all data.' + #13#10 +
        'Click No to keep your data for future installations.',
        mbConfirmation, MB_YESNO
      );
      
      if DialogResult = IDYES then
      begin
        DelTree(AppDataDir, True, True, True);
      end;
    end;
  end;
end;

[Registry]
; Associate .tflow files with TranscriptFlow
Root: HKCR; Subkey: ".tflow"; ValueType: string; ValueName: ""; ValueData: "TranscriptFlowFile"; Flags: uninsdeletevalue
Root: HKCR; Subkey: "TranscriptFlowFile"; ValueType: string; ValueName: ""; ValueData: "TranscriptFlow Document"; Flags: uninsdeletekey
Root: HKCR; Subkey: "TranscriptFlowFile\DefaultIcon"; ValueType: string; ValueName: ""; ValueData: "{app}\{#MyAppExeName},0"
Root: HKCR; Subkey: "TranscriptFlowFile\shell\open\command"; ValueType: string; ValueName: ""; ValueData: """{app}\{#MyAppExeName}"" ""%1"""

[UninstallDelete]
Type: filesandordirs; Name: "{app}"
