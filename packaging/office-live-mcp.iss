; AppId is permanent: never change it between releases.
#ifndef AppVersion
  #error AppVersion must be passed by packaging/build.py
#endif
#ifndef SourceDir
  #error SourceDir must be passed by packaging/build.py
#endif
#define RepoURL "https://github.com/mikhalchankasm/office-live-mcp"

[Setup]
AppId={{47C3867B-C095-4C76-AACD-DA6A268CF854}
AppName=Office Live MCP
AppVersion={#AppVersion}
AppPublisher=mikhalchankasm
AppPublisherURL={#RepoURL}
AppSupportURL={#RepoURL}
DefaultDirName={code:DefaultRoot}
UsePreviousAppDir=yes
DisableDirPage=auto
DefaultGroupName=Office Live MCP
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
CloseApplications=yes
RestartApplications=no
DisableWelcomePage=no
LicenseFile=..\LICENSE
WizardStyle=modern
ShowLanguageDialog=auto
OutputBaseFilename=office-live-mcp-{#AppVersion}-setup
Compression=lzma2/max
SolidCompression=yes
UninstallDisplayIcon={app}\app\office-live-mcp.exe
SetupLogging=yes

[Languages]
Name: "en"; MessagesFile: "compiler:Default.isl"
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[CustomMessages]
en.Connect=Connect to agents
ru.Connect=Подключить к агентам
en.Doctor=Diagnostics (doctor)
ru.Doctor=Проверка (doctor)
en.Agents=Agents
ru.Agents=Агенты
en.AgentDescription=Select the agents to connect to Office Live MCP.
ru.AgentDescription=Выберите агентов для подключения к Office Live MCP.
en.AccessNote=Access applies only to added agents. Connected agents keep their settings. Uncheck an agent to disconnect it in all recorded scopes. Use the Start menu shortcut for VS Code / projects.
ru.AccessNote=Режим применяется только к добавляемым агентам. Подключённые сохраняют настройки. Снятая отметка отключает агента во всех учтённых областях. VS Code / проекты: через ярлык «Подключить к агентам».
en.Full=Full access
ru.Full=Полный доступ
en.Readonly=Read only
ru.Readonly=Только чтение
en.Detected=found on this computer
ru.Detected=найден на компьютере
en.Registered=connected
ru.Registered=подключён
en.OfficeMissing=Desktop Excel or Word is required. Neither is registered on this computer. Continue installing anyway?
ru.OfficeMissing=Нужен настольный Excel или Word. Ни один из них не зарегистрирован на этом компьютере. Продолжить установку?
en.ExcelOnly=Only Excel was found. Word tools will be unavailable.
ru.ExcelOnly=Найден только Excel. Инструменты Word будут недоступны.
en.WordOnly=Only Word was found. Excel tools will be unavailable.
ru.WordOnly=Найден только Word. Инструменты Excel будут недоступны.
en.OfficeBoth=Excel and Word are registered.
ru.OfficeBoth=Excel и Word зарегистрированы.
en.Result=Connection result
ru.Result=Результат подключения
en.Success=Office Live MCP is installed. The selected registrations are configured.
ru.Success=Office Live MCP установлен. Выбранные подключения настроены.
en.Problem=The program files remain installed, but configuration needs attention. Use Start > Office Live MCP > Connect to agents to retry.
ru.Problem=Файлы программы установлены, но настройка требует внимания. Повторите подключение: Пуск → Office Live MCP → Подключить к агентам.
en.ClientsFailed=Could not read the agent list. Existing registrations will be preserved.
ru.ClientsFailed=Не удалось прочитать список агентов. Существующие регистрации будут сохранены.
en.CommandFailed=Command failed (exit code %1). See the setup log for details.
ru.CommandFailed=Команда завершилась с ошибкой (код %1). Подробности в журнале установки.
en.BadResponse=The command did not return a valid JSON response. See the setup log.
ru.BadResponse=Команда не вернула корректный JSON-ответ. Подробности в журнале установки.
en.Finish=Restart your agents (Claude, Cursor, Codex...) and ask: show which workbooks are open in Excel.
ru.Finish=Перезапустите агентов (Claude, Cursor, Codex…) и попросите: покажи, какие книги открыты в Excel.
en.RemoveProblem=Some registrations could not be removed:%n%1%nUninstall anyway? Choose No to cancel and fix these entries first.
ru.RemoveProblem=Не удалось снять некоторые регистрации:%n%1%nУдалить всё равно? Выберите «Нет», чтобы отменить удаление и исправить записи.

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}\app"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\LICENSE"; DestDir: "{app}"; DestName: "LICENSE.txt"; Flags: ignoreversion

[InstallDelete]
Type: filesandordirs; Name: "{app}\app\_internal"

[Icons]
Name: "{group}\{cm:Connect}"; Filename: "{app}\app\office-live-mcp.exe"; Parameters: "setup --pause"; WorkingDir: "{app}"
Name: "{group}\{cm:Doctor}"; Filename: "{app}\app\office-live-mcp.exe"; Parameters: "doctor --pause"; WorkingDir: "{app}"

[UninstallDelete]
Type: files; Name: "{app}\office-live-mcp.install"
Type: files; Name: "{app}\app\office-live-mcp.files.json"
Type: filesandordirs; Name: "{app}\app.old-*"
Type: filesandordirs; Name: "{app}\app.new-*"
Type: dirifempty; Name: "{app}\app"
Type: dirifempty; Name: "{app}"

[Code]
#include "installer-json.iss"

var
  AgentsPage: TInputOptionWizardPage;
  ResultPage: TOutputMsgMemoWizardPage;
  FullRadio, ReadonlyRadio: TNewRadioButton;
  AccessNote: TNewStaticText;
  ClientIDs: TArrayOfString;
  Updating, ClientsLoaded, ClientsValid, FilesCopied, PostinstallDone: Boolean;
  PostinstallCode: Integer;
  ResultText, ClientProblem: String;

function DefaultRoot(Param: String): String;
var Data, Marker: AnsiString; Root: String;
begin
  Result := ExpandConstant('{localappdata}\Programs\office-live-mcp');
  if not LoadStringFromFile(ExpandConstant('{localappdata}\office-live-mcp\state.json'), Data) then Exit;
  Root := JsonText(Utf8Decode(Data), 'root');
  // Only absolute local paths are used as hints; invalid/ambiguous state falls back to the default.
  if Length(Root) < 3 then Exit;
  if (Root[2] <> ':') or (Root[3] <> '\') or (Pos('..', Root) > 0) then Exit;
  if not DirExists(Root) then Exit;
  if not LoadStringFromFile(AddBackslash(Root) + 'office-live-mcp.install', Marker) then Exit;
  if JsonText(Utf8Decode(Marker), 'product') = 'office-live-mcp' then Result := Root;
end;

function RunCLI(const Params: String; var Data, Problem: String): Integer;
var Output: TExecOutput; I, Code: Integer; Started: Boolean;
begin
  Result := 1; Data := ''; Problem := '';
  Log('office-live-mcp ' + Params);
  try
    Started := ExecAndCaptureOutput(ExpandConstant('{app}\app\office-live-mcp.exe'), Params,
      ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, Code, Output);
    if not Started then begin
      Problem := SysErrorMessage(Code); Log(Problem); Exit;
    end;
    for I := 0 to GetArrayLength(Output.StdErr) - 1 do Log(Output.StdErr[I]);
    for I := 0 to GetArrayLength(Output.StdOut) - 1 do begin
      Log(Output.StdOut[I]); Data := Data + Output.StdOut[I];
    end;
    if Output.Error then begin Problem := CustomMessage('BadResponse'); Exit; end;
    Result := Code;
    Problem := JsonProblems(Data);
    if (Result <> 0) and (Problem = '') then
      Problem := FmtMessage(CustomMessage('CommandFailed'), [IntToStr(Code)]);
  except
    Result := 1; Problem := GetExceptionMessage; Log(Problem);
  end;
end;

procedure InitializeWizard;
begin
  AgentsPage := CreateInputOptionPage(wpInstalling, CustomMessage('Agents'),
    CustomMessage('AgentDescription'), '', False, True);
  AgentsPage.CheckListBox.Top := 0;
  AgentsPage.CheckListBox.Height := ScaleY(124);
  FullRadio := TNewRadioButton.Create(AgentsPage);
  FullRadio.Parent := AgentsPage.Surface;
  FullRadio.SetBounds(0, ScaleY(132), AgentsPage.SurfaceWidth, ScaleY(22));
  FullRadio.Caption := CustomMessage('Full'); FullRadio.Checked := True;
  ReadonlyRadio := TNewRadioButton.Create(AgentsPage);
  ReadonlyRadio.Parent := AgentsPage.Surface;
  ReadonlyRadio.SetBounds(0, ScaleY(156), AgentsPage.SurfaceWidth, ScaleY(22));
  ReadonlyRadio.Caption := CustomMessage('Readonly');
  AccessNote := TNewStaticText.Create(AgentsPage);
  AccessNote.Parent := AgentsPage.Surface; AccessNote.AutoSize := False;
  AccessNote.SetBounds(0, ScaleY(188), AgentsPage.SurfaceWidth, AgentsPage.SurfaceHeight - ScaleY(188));
  AccessNote.WordWrap := True; AccessNote.Caption := CustomMessage('AccessNote');
  ResultPage := CreateOutputMsgMemoPage(AgentsPage.ID, CustomMessage('Result'), '', '', '');
end;

procedure LoadClients;
var Data, Raw, Problem, ID, LabelText: String; Items: TArrayOfString; I, N, Code: Integer;
begin
  ClientsLoaded := True; ClientsValid := False;
  Code := RunCLI('clients --json', Data, Problem);
  if (Code = 0) and JsonField(Data, 'clients', Raw) then begin
    if JsonArray(Raw, Items) then begin
      for I := 0 to GetArrayLength(Items) - 1 do begin
        ID := JsonText(Items[I], 'id');
        if (ID <> '') and (ID <> 'vscode') then begin
          N := GetArrayLength(ClientIDs); SetArrayLength(ClientIDs, N + 1); ClientIDs[N] := ID;
          LabelText := JsonText(Items[I], 'label');
          if JsonTrue(Items[I], 'registered') then LabelText := LabelText + ' (' + CustomMessage('Registered') + ')'
          else if JsonTrue(Items[I], 'detected') then LabelText := LabelText + ' (' + CustomMessage('Detected') + ')';
          AgentsPage.Add(LabelText);
          AgentsPage.Values[N] := JsonTrue(Items[I], 'registered') or ((not Updating) and JsonTrue(Items[I], 'detected'));
        end;
      end;
      ClientsValid := GetArrayLength(ClientIDs) > 0;
      ReadonlyRadio.Checked := JsonText(Data, 'access') = 'readonly';
      FullRadio.Checked := not ReadonlyRadio.Checked;
    end;
  end;
  if not ClientsValid then begin
    ClientProblem := CustomMessage('ClientsFailed') + #13#10 + Problem;
    AccessNote.Caption := ClientProblem;
    AgentsPage.CheckListBox.Enabled := False; FullRadio.Enabled := False; ReadonlyRadio.Enabled := False;
  end;
end;

function HasSwitch(const Name: String): Boolean;
var I: Integer;
begin
  Result := False;
  for I := 1 to ParamCount do
    if CompareText(ParamStr(I), Name) = 0 then begin Result := True; Exit; end;
end;

procedure ConfigureInstallation(const Clients, Access: String);
var Params, Data, Problem: String;
begin
  if PostinstallDone then Exit;
  PostinstallDone := True;
  // Values are passed as argv, never through cmd.exe/PowerShell.
  Params := 'postinstall --json --clients ' + AddQuotes(Clients) + ' --access ' + AddQuotes(Access);
  if HasSwitch('/NOLINKS') then Params := Params + ' --no-links';
  PostinstallCode := RunCLI(Params, Data, Problem);
  if (PostinstallCode = 0) and not JsonTrue(Data, 'ok') then begin
    PostinstallCode := 1; Problem := CustomMessage('BadResponse');
  end;
  if ClientProblem <> '' then begin
    PostinstallCode := 1; Problem := ClientProblem + #13#10 + Problem;
  end;
  if PostinstallCode = 0 then ResultText := CustomMessage('Success')
  else ResultText := CustomMessage('Problem') + #13#10#13#10 + Problem;
  Log(ResultText);
  if not WizardSilent then ResultPage.RichEditViewer.Text := ResultText;
end;

function OfficeStatus: String;
var ExcelFound, WordFound: Boolean;
begin
  ExcelFound := RegKeyExists(HKCR32, 'Excel.Application\CLSID') or RegKeyExists(HKCR64, 'Excel.Application\CLSID');
  WordFound := RegKeyExists(HKCR32, 'Word.Application\CLSID') or RegKeyExists(HKCR64, 'Word.Application\CLSID');
  if ExcelFound and WordFound then Result := CustomMessage('OfficeBoth')
  else if ExcelFound then Result := CustomMessage('ExcelOnly')
  else if WordFound then Result := CustomMessage('WordOnly')
  else Result := CustomMessage('OfficeMissing');
end;

function UpdateReadyMemo(Space, NewLine, MemoUserInfoInfo, MemoDirInfoInfo, MemoTypeInfoInfo,
  MemoComponentsInfoInfo, MemoGroupInfoInfo, MemoTasksInfoInfo: String): String;
begin
  Result := MemoDirInfoInfo + NewLine + NewLine + OfficeStatus;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var I: Integer; Clients, Access, Status: String;
begin
  Result := True;
  if CurPageID = wpReady then begin
    Status := OfficeStatus; Log(Status);
    if (Status = CustomMessage('OfficeMissing')) and not WizardSilent then begin
      Result := SuppressibleMsgBox(Status, mbConfirmation, MB_YESNO, IDYES) = IDYES;
      if not Result then WizardForm.Close;
    end;
  end;
  if (CurPageID = AgentsPage.ID) and not WizardSilent then begin
    Clients := '';
    for I := 0 to GetArrayLength(ClientIDs) - 1 do
      if AgentsPage.Values[I] then begin
        if Clients <> '' then Clients := Clients + ',';
        Clients := Clients + ClientIDs[I];
      end;
    if Clients = '' then Clients := 'none';
    if not ClientsValid then Clients := 'keep';
    Access := 'full'; if ReadonlyRadio.Checked then Access := 'readonly';
    ConfigureInstallation(Clients, Access);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var Clients, DefaultClients: String;
begin
  if CurStep = ssInstall then
    Updating := FileExists(ExpandConstant('{app}\office-live-mcp.install')) or FileExists(ExpandConstant('{app}\unins000.exe'));
  if CurStep = ssPostInstall then begin
    FilesCopied := True;
    if WizardSilent then begin
      DefaultClients := 'none'; if Updating then DefaultClients := 'keep';
      Clients := ExpandConstant('{param:CLIENTS|' + DefaultClients + '}');
      ConfigureInstallation(Clients, ExpandConstant('{param:ACCESS|full}'));
    end;
  end;
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := WizardSilent and ((PageID = AgentsPage.ID) or (PageID = ResultPage.ID));
end;

procedure CurPageChanged(CurPageID: Integer);
begin
  if CurPageID = AgentsPage.ID then begin
    if not ClientsLoaded then LoadClients;
    WizardForm.BackButton.Enabled := False;
    WizardForm.CancelButton.Enabled := False;
  end;
  if CurPageID = ResultPage.ID then WizardForm.BackButton.Enabled := False;
  if CurPageID = wpFinished then begin
    WizardForm.FinishedLabel.Caption := CustomMessage('Finish');
    if PostinstallCode <> 0 then WizardForm.FinishedLabel.Caption := ResultText + #13#10#13#10 + CustomMessage('Finish');
  end;
end;

procedure CancelButtonClick(CurPageID: Integer; var Cancel, Confirm: Boolean);
begin
  if FilesCopied and not PostinstallDone then begin Cancel := False; Confirm := False; end;
end;

function GetCustomSetupExitCode: Integer;
begin
  Result := PostinstallCode;
end;

function InitializeUninstall: Boolean;
var Data, Problem: String; Code: Integer;
begin
  Code := RunCLI('uninstall --keep-files --yes --json', Data, Problem);
  if (Code = 0) and not JsonTrue(Data, 'ok') then begin Code := 1; Problem := CustomMessage('BadResponse'); end;
  Result := True;
  if Code <> 0 then begin
    Log(Problem);
    if not UninstallSilent then
      Result := SuppressibleMsgBox(FmtMessage(CustomMessage('RemoveProblem'), [Problem]), mbConfirmation, MB_YESNO, IDYES) = IDYES;
  end;
end;
