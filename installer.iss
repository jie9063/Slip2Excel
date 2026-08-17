#define MyAppName "Slip2Excel"
#ifndef MyAppVersion
  #define MyAppVersion "0.00"
#endif
#define MyAppPublisher "Slip2Excel"
#define MyAppExeName "Slip2Excel.exe"

[Setup]
AppId={{49C11E4A-7B53-4C5E-A01A-8C57E6D45201}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
VersionInfoVersion=0.0.0.0
VersionInfoTextVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DisableDirPage=no
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=release
OutputBaseFilename=Slip2Excel-v{#MyAppVersion}-Setup
SetupIconFile=assets\slip2excel-icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Files]
Source: "dist\Slip2Excel-v{#MyAppVersion}.exe"; DestDir: "{app}"; DestName: "{#MyAppExeName}"; Flags: ignoreversion

[Dirs]
Name: "{app}\OllamaModels"

[Registry]
Root: HKCU; Subkey: "Environment"; ValueType: expandsz; ValueName: "OLLAMA_MODELS"; ValueData: "{app}\OllamaModels"; Flags: preservestringtype

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional options:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Slip2Excel"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\{#MyAppExeName}"

[Code]
const
  OllamaSetupUrl = 'https://ollama.com/download/OllamaSetup.exe';
  OllamaModel = 'qwen2.5vl:7b';

var
  OllamaDownloadPage: TDownloadWizardPage;

function SetEnvironmentVariable(lpName, lpValue: String): Boolean;
  external 'SetEnvironmentVariableW@kernel32.dll stdcall';

function IsOllamaInstalled(): Boolean;
begin
  Result := FileExists(ExpandConstant('{app}\Ollama\ollama.exe')) or
    FileExists(ExpandConstant('{localappdata}\Programs\Ollama\ollama.exe'));
end;

function OllamaExecutable(): String;
begin
  if FileExists(ExpandConstant('{app}\Ollama\ollama.exe')) then begin
    Result := ExpandConstant('{app}\Ollama\ollama.exe');
  end else if FileExists(ExpandConstant('{localappdata}\Programs\Ollama\ollama.exe')) then begin
    Result := ExpandConstant('{localappdata}\Programs\Ollama\ollama.exe');
  end else begin
    Result := '';
  end;
end;

function InstallOllamaIfNeeded(): Boolean;
var
  ResultCode: Integer;
  OllamaSetupPath: String;
  OllamaInstallPath: String;
  OllamaModelsPath: String;
begin
  Result := False;
  if IsOllamaInstalled() then begin
    Log('Ollama is already installed.');
    Result := True;
    exit;
  end;
  if WizardSilent() then begin
    Log('Skipping Ollama download during a silent installation.');
    Result := False;
    exit;
  end;

  OllamaDownloadPage := CreateDownloadPage(
    'Downloading Ollama',
    'Downloading the official Ollama installer. This may take a few minutes.', nil);
  OllamaDownloadPage.Add(OllamaSetupUrl, 'OllamaSetup.exe', '');
  OllamaDownloadPage.Show;
  try
    OllamaDownloadPage.Download;
  except
    MsgBox(
      'Ollama could not be downloaded. You can install it later from https://ollama.com/download/windows',
      mbError, MB_OK);
    Result := False;
    exit;
  finally
    OllamaDownloadPage.Hide;
  end;

  OllamaSetupPath := ExpandConstant('{tmp}\OllamaSetup.exe');
  OllamaInstallPath := ExpandConstant('{app}\Ollama');
  OllamaModelsPath := ExpandConstant('{app}\OllamaModels');
  SetEnvironmentVariable('OLLAMA_MODELS', OllamaModelsPath);
  if not Exec(OllamaSetupPath, '/DIR="' + OllamaInstallPath + '"', '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then begin
    MsgBox('The official Ollama installer could not be started.', mbError, MB_OK);
  end else if ResultCode <> 0 then begin
    MsgBox('Ollama installation did not complete. You can try again later from https://ollama.com/download/windows', mbError, MB_OK);
  end else begin
    Result := FileExists(OllamaInstallPath + '\ollama.exe');
  end;
end;

function DownloadDefaultModel(): Boolean;
var
  ResultCode: Integer;
  Executable: String;
begin
  Result := False;
  if WizardSilent() then begin
    Log('Skipping model download during a silent installation.');
    exit;
  end;
  Executable := OllamaExecutable();
  if Executable = '' then begin
    MsgBox('Ollama was not found, so the recognition model could not be downloaded.', mbError, MB_OK);
    exit;
  end;
  MsgBox(
    'Slip2Excel is downloading its free recognition model (' + OllamaModel + ').' + #13#10 +
    'This is a one-time download of several GB. Keep this window open until it finishes.',
    mbInformation, MB_OK);
  if not Exec(Executable, 'pull ' + OllamaModel, '', SW_SHOWNORMAL, ewWaitUntilTerminated, ResultCode) then begin
    MsgBox('The recognition model download could not be started.', mbError, MB_OK);
  end else if ResultCode <> 0 then begin
    MsgBox('The recognition model could not be downloaded. Check your internet connection and run the installer again.', mbError, MB_OK);
  end else begin
    Result := True;
  end;
end;

function InitializeSetup(): Boolean;
begin
  MsgBox(
    'After installation, select your photo folder and blank Excel template in Settings.' + #13#10 + #13#10 +
    'The installer will automatically install Ollama when needed, then download the free recognition model. This requires an internet connection and several GB of disk space.',
    mbInformation, MB_OK);
  Result := True;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then begin
    if InstallOllamaIfNeeded() then begin
      DownloadDefaultModel();
    end;
  end;
end;
