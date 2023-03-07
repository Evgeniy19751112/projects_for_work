unit UnitMain;

{
Основной модуль (главная форма приложения. Обеспечивает основной диалог с
пользователем. Все настройки сохряняются в реестре. Если запустить с
параметром AUTO, то подключение и запуск скрипта выполняется автоматически,
но при условии полной первичной настройки. В случае ошибки вывести сообщение
в текстовый файл и открыть в блокноте. При наличии параметра QUIET совместно
с AUTO окно приложения не открывается, но сообщение об ошибке всё равно
откроется в блокноте. Подготовленные файлы отправляются на указанный адрес
электронной почты и сохраняется в отправленных письмах.

Начато: 2023-02-26
Завершено:
Автор: Тявкин Е.Н.

Ход работы:
- делаем интерфейс приложения и прописываем первичные действия кнопок.
- реализуем доступ к серверу MS SQL с получением списка баз данных.
- реализуем подключение к почтовому серверу для отправки и сохранения писем.
- реализуем поток для получения данных из БД.
- реализуем автомат последовательности действий подключить+создать файл+послать.
- Устраняем найденные проблемы при попытке переноса на рабочую сеть с тестовыми базами.
- реализуем анализ параметров командной строки.
- сохраняем полученный протокол в журнале событий Windows.
- делаем сборку для постоянной эксплуатации.
}

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.ExtCtrls, Vcl.AppEvnts, Data.DB, Data.Win.ADODB, Vcl.StdCtrls,
  Vcl.CheckLst, UnitQuery, Vcl.ImgList, IdIMAP4, IdMessage, IdAntiFreezeBase,
  Vcl.IdAntiFreeze, IdCustomTransparentProxy, IdSocks, IdIOHandler,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdBaseComponent,
  IdComponent, IdTCPConnection, IdTCPClient, IdExplicitTLSClientServerBase,
  IdMessageClient, IdSMTPBase, IdSMTP, Vcl.Buttons, IdAttachmentFile, UnitBases;

type
  TStageType = (stNone, stWaitConnect, stBeforeRunThreads, stAfterFinishThreads,
      stBeforeSendMail, stAfterSendMail, stFinishSuccess, stFinishWithError);

  TfrmASR = class(TForm)
    _IdIMAP: TIdIMAP4;
    _IdMsg: TIdMessage;
    _IdSMTP: TIdSMTP;
    _IdSocksInfo: TIdSocksInfo;
    _IdSSLIO_IMAP: TIdSSLIOHandlerSocketOpenSSL;
    _IdSSLIO_SMTP: TIdSSLIOHandlerSocketOpenSSL;
    ApplicationEvents1: TApplicationEvents;
    BalloonHint1: TBalloonHint;
    cmdRunSendReports: TBitBtn;
    cmdTestMail: TBitBtn;
    grbMail: TGroupBox;
    IdAntiFreeze1: TIdAntiFreeze;
    ImagesIndicator: TImageList;
    memLog: TMemo;
    memMailBody: TMemo;
    timAutomat: TTimer;
    timCallThread: TTimer;
    txtMailLogin: TLabeledEdit;
    txtMailPassword: TLabeledEdit;
    txtMailPortIMAP: TLabeledEdit;
    txtMailPortSMTP: TLabeledEdit;
    txtMailRecipient: TLabeledEdit;
    txtMailSender: TLabeledEdit;
    txtMailServerIMAP: TLabeledEdit;
    txtMailServerSMTP: TLabeledEdit;
    txtMailSubject: TLabeledEdit;
    imgSMTP: TImage;
    cmdConnectDB: TBitBtn;
    imgIMAP: TImage;
    cmbCharSet: TComboBox;
    lblCharSet: TLabel;
    lblSMTP_TLS: TLabel;
    lblIMAP_TLS: TLabel;
    cmbSMTP_TLS: TComboBox;
    cmbIMAP_TLS: TComboBox;
    lblSMTP_SSL: TLabel;
    cmbSMTP_SSL: TComboBox;
    lblIMAP_SSL: TLabel;
    cmbIMAP_SSL: TComboBox;
    txtTimeout: TLabeledEdit;
    chkHiddenExcel: TCheckBox;
    chkNoUpdateExcel: TCheckBox;
    imgProgress: TImage;
    procedure ApplicationEvents1Exception(Sender: TObject; E: Exception);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure cmdTestMailClick(Sender: TObject);
    procedure EndOfThread(Sender: TObject);
    procedure timAutomatTimer(Sender: TObject);
    procedure cmdRunSendReportsClick(Sender: TObject);
    procedure cmdConnectDBClick(Sender: TObject);
    procedure _IdSMTPConnected(Sender: TObject);
    procedure _IdSMTPDisconnected(Sender: TObject);
    procedure _IdSMTPFailedRecipient(Sender: TObject; const AAddress, ACode,
      AText: string; var VContinue: Boolean);
    procedure _IdIMAPConnected(Sender: TObject);
    procedure _IdIMAPDisconnected(Sender: TObject);
    procedure chkHiddenExcelClick(Sender: TObject);
  private
    FsFolderName: string;     // Путь к файлам приложения
    FsFolderTemp: string;     // Путь к файлам временным
    FStage: TStageType;       // Состояние автомата отправки отчётов
    procedure SetLog(const Value: string);
    procedure GetParams;      // Грузим из реестра параметры
    function GetUseTLS(combo: TComboBox): TIdUseTLS;
    function GetOpenSSL(combo: TComboBox): TIdSSLVersion;
  public
    function MailPrepareMessage: Boolean;  // Подготовить сообщение для отправки
    function FileListAtTempFolder: TStringList;  // Список файлов во временной папке
    procedure DoSendMail;  // Отправить сообщение получателю
    procedure PrepareProgress(iQuantity: Integer);
    procedure DrawLineProgress(iIndex, iValue, iLast: Integer);
    property log: string write SetLog;
    property Stage: TStageType read FStage write FStage;
    property sFolderTemp: string read FsFolderTemp;
  end;

var
  frmASR: TfrmASR;

implementation

uses
  System.StrUtils, System.Win.Registry, UN_DiskUtils, System.Win.ComObj,
  Winapi.ShellAPI, UnitFormBase, IdText, IdMessageBuilder;

{$R *.dfm}

const // Локальные константы
  csl_RegKey = '\SOFTWARE\TZN\EGISSO\send_report';

procedure TfrmASR.ApplicationEvents1Exception(Sender: TObject; E: Exception);
begin
  try
    log := 'Application Exception: Class ' +
        Sender.ClassName + ': ' + E.Message;
    FStage := stFinishWithError;
  except
  end;
end;

procedure TfrmASR.chkHiddenExcelClick(Sender: TObject);
begin
  if chkHiddenExcel.Checked then
    chkNoUpdateExcel.Checked := True;
end;

procedure TfrmASR.cmdConnectDBClick(Sender: TObject);
begin
  frmServerDB.ShowModal;
end;

procedure TfrmASR.cmdRunSendReportsClick(Sender: TObject);
begin
  timAutomat.Enabled := True;
end;

procedure TfrmASR.cmdTestMailClick(Sender: TObject);
var
  r: TRegistry;
  oldChar: Char;
  lst: TStringList;
begin
  // Сохряняем поля только после успешной отправки тестового сообщения
  // за исключением получателя, темы и тела письма. Они сохряняются
  // при уничтожении формы
  try
    log := 'Готовим тестовое письмо';
    // Отправить тестовое письмо. Для этого сохраним файл текстовый,
    // создадим письмо во временной папке, а затем выполним отправку письма
    lst := TStringList.Create;
    try
      lst.Add('Тестовое сообщение для проверки почты');
      lst.Add('Сейчас ' + FormatDateTime('dd.mm.yyyy hh:nn:ss', Now));
      lst.SaveToFile(FsFolderTemp + 'test_почты.txt');
    finally
      lst.Free;
    end;
    if not MailPrepareMessage then
      raise Exception.Create('Ошибка подготовки письма для тестировани!');

    // Тестовое письмо готово, теперь подключаемся, отправляем, отключаемся
    log := 'Посылаем тестовое письмо';
    DoSendMail;

    // Сделать сохранение
    log := 'Сохраняем настройки';
    r := TRegistry.Create;
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);
    try
      with txtMailLogin do r.WriteString(Name, Text);
      with txtMailPassword do r.WriteString(Name, Text);
      with txtMailPortIMAP do r.WriteString(Name, Text);
      with txtMailPortSMTP do r.WriteString(Name, Text);
      with txtMailServerIMAP do r.WriteString(Name, Text);
      with txtMailServerSMTP do r.WriteString(Name, Text);
      with txtMailRecipient do r.WriteString(Name, Text);
      with txtMailSubject do r.WriteString(Name, Text);
      with memMailBody do
      begin
        oldChar := Lines.Delimiter;
        Lines.Delimiter := '/';
        r.WriteString(Name, Lines.DelimitedText);
        Lines.Delimiter := oldChar;
      end;
      with txtMailSender do r.WriteString(Name, Text);
    finally
      r.Free;
    end;
  except
    on E: Exception do
      log := E.Message;
  end;
  log := 'Тест пройден';
end;

procedure TfrmASR.DoSendMail;
var
  lstFolders: TStringList;
begin
  // Непосредственная отправка письма
  FStage := stBeforeSendMail;

  // Устанавливаем параметры
  _IdSocksInfo.Port := StrToIntDef(txtMailPortSMTP.Text, 587);  // Порт по которому мы будем связываться
  _IdSocksInfo.Authentication := saNoAuthentication;
  _IdSocksInfo.Version := svNoSocks;

  // SMTP
  _IdSMTP.Host := txtMailServerSMTP.Text;
  _IdSMTP.Port := _IdSocksInfo.Port;
  _IdSMTP.Username := txtMailLogin.Text;
  _IdSMTP.Password := txtMailPassword.Text;
  _IdSMTP.HeloName := 'HelloNameForMail';
  _IdSMTP.ConnectTimeout := StrToIntDef(txtTimeout.Text, 30) * 1000; //optional
  _IdSMTP.IOHandler := _IdSSLIO_SMTP;
  _IdSMTP.UseTLS := GetUseTLS(cmbSMTP_TLS); // Выставляем значение - использовать неявный TSL
  _IdSMTP.UseEhlo := True;
  _IdSMTP.AuthType := satDefault; // Тип аутентификации: Login/Password
  _IdSSLIO_SMTP.SSLOptions.Method := GetOpenSSL(cmbSMTP_SSL);

  // IMAP
  _IdIMAP.Port := StrToIntDef(txtMailPortIMAP.Text, 143);
  _IdIMAP.Host := txtMailServerIMAP.Text;
  _IdIMAP.Username := txtMailLogin.Text;
  _IdIMAP.Password := txtMailPassword.Text;
  _IdIMAP.ConnectTimeout := StrToIntDef(txtTimeout.Text, 30) * 1000; //optional
  _IdIMAP.IOHandler := _IdSSLIO_IMAP;
  _IdIMAP.IOHandler.Host := _IdIMAP.Host;
  _IdIMAP.IOHandler.Port := _IdIMAP.Port;
  _IdIMAP.UseTLS := GetUseTLS(cmbIMAP_TLS);
  _IdSSLIO_IMAP.SSLOptions.Method := GetOpenSSL(cmbSMTP_SSL);

  // В IdMsg письмо уже лежит. Файл нужен только для контроля при отладке
  lstFolders := TStringList.Create;
  try
    _IdIMAP.Connect();
    try
      _IdIMAP.SelectMailBox('Sent');
      _IdSMTP.Connect;
      try
        _IdMsg.Flags := [];  // Сбросили все флаги перед отправкой
        _IdSMTP.Send(_IdMsg);
        _IdIMAP.AppendMsg('Sent', _IdMsg, [mfSeen]);  // А в отправленные добавляем "прочтённое"
      finally
        _IdSMTP.Disconnect();
      end;
    finally
      _IdIMAP.Disconnect();
    end;
  finally
    lstFolders.Free;
    FStage := stAfterSendMail;
  end;
end;

procedure TfrmASR.DrawLineProgress(iIndex, iValue, iLast: Integer);
var
  iTopLine: Integer;
  iDelimeter: Integer;
begin
  // Рисуем зелёную полосу прогресса
  try
    iTopLine := iIndex * imgProgress.Tag + 1;
    if iLast = 0 then
      iLast := 1;
    iDelimeter := imgProgress.Width * iValue div iLast;
    if iDelimeter < 1 then
      iDelimeter := 1;
    if iValue >= iLast then
      iDelimeter := imgProgress.Width - 2;
    if iDelimeter > imgProgress.Width - 2 then
      iDelimeter := imgProgress.Width - 2;
    with imgProgress.Picture.Bitmap.Canvas do
    begin
      Brush.Color := clRed;
      FillRect(Rect(iDelimeter, iTopLine, Width - 2, iTopLine + imgProgress.Tag));
      Brush.Color := clLime;
      FillRect(Rect(1, iTopLine, iDelimeter, iTopLine + imgProgress.Tag));
    end;
  except
  end;
end;

procedure TfrmASR.EndOfThread(Sender: TObject);
var
  thr: TMyThreadQuery;
begin
  // Вызывается из потока. Понижаем счётчик, а при нуле открываем папку
  // с файлами (для контроля).
  frmServerDB.DecThrCounter;
  try
    thr := Sender as TMyThreadQuery;
    if (FStage = stBeforeRunThreads) and (thr.HasError) then
      FStage := stFinishWithError
    else if (FStage = stBeforeRunThreads) and (frmServerDB.ThrCounter = 0) then
    begin
      FStage := stAfterFinishThreads;

      // Показать что в папке получилось
      ShellExecute(Handle, 'open', PChar(FsFolderTemp), nil, nil, SW_NORMAL);
    end;
  except
    on E: Exception do
      log := E.Message;
  end;
end;

function TfrmASR.FileListAtTempFolder: TStringList;
var
  sr: TSearchRec;
begin
  // Вывести список файлов во временной папке
  Result := TStringList.Create;
  if FindFirst(FsFolderTemp + '*.*', faArchive, sr) = 0 then
    repeat
      if (sr.Name <> '.') and (sr.Name <> '..') then
        Result.Add(sr.Name);
    until FindNext(sr) <> 0;
end;

procedure TfrmASR.FormCreate(Sender: TObject);
var
  iCounter: Integer;
  sTemp: string;
begin
  // Инициализация
  Constraints.MinHeight := Height;
  Constraints.MinWidth := Width;
  memLog.Lines.Clear;
  memMailBody.Lines.Clear;
  FStage := stNone;

  // Ставим красный выключенный индикатор
  imagesIndicator.GetBitmap(0, imgSMTP.Picture.Bitmap);
  imagesIndicator.GetBitmap(0, imgIMAP.Picture.Bitmap);

  // Готовим прогресс бар
  with imgProgress.Picture.Bitmap do
  begin
    Width := imgProgress.Width;
    Height := imgProgress.Height;
    PixelFormat := pf24bit;
    with Canvas do
    begin
      Brush.Color := clDkGray;
      FillRect(ClipRect);
    end;
  end;

  // Определяем что и где будет располагаться (каталоги)
  FsFolderName := ExtractFilePath(Application.ExeName);
  FsFolderTemp := GetEnvironmentVariable('TEMP');
  if FsFolderTemp = '' then
    FsFolderTemp := FsFolderName;
  FsFolderTemp := IncludeTrailingPathDelimiter(FsFolderTemp);
  iCounter := 0;
  repeat
    sTemp := FsFolderTemp + 'tzn' + IntToStr(iCounter);
    Inc(iCounter);
  until not DirectoryExists(sTemp);
  FsFolderTemp := IncludeTrailingPathDelimiter(sTemp);
  ForceDirectories(FsFolderTemp);

  // Грузим параметры
  GetParams;
  log := 'Загрузился';

  // Проверяем элементы автоматизации
  if (ParamCount > 0) and (UpperCase(ParamStr(1)) = 'AUTO') then
    PostMessage(cmdRunSendReports.Handle, BM_CLICK, 0, 0);
end;

procedure TfrmASR.FormDestroy(Sender: TObject);
var
  r: TRegistry;
  EventLog: THandle;
  MyMsg: array [0..2] of PChar;
  strMsg: string;
begin
  // Остановить таймер, вызывающий поток обработки
  timCallThread.Enabled := False;
  timCallThread.Tag := 0;

  // Сохраним в реестре имена отмеченных баз
  r := TRegistry.Create;
  try
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);
    with txtMailRecipient do r.WriteString(Name, Text);
    with txtMailSubject do r.WriteString(Name, Text);
    with memMailBody do
    begin
      Lines.Delimiter := '/';
      r.WriteString(Name, Lines.DelimitedText);
    end;
    with cmbCharSet do r.WriteString(Name, Text);
    with cmbSMTP_TLS do r.WriteString(Name, Text);
    with cmbIMAP_TLS do r.WriteString(Name, Text);
    with cmbSMTP_SSL do r.WriteString(Name, Text);
    with cmbIMAP_SSL do r.WriteString(Name, Text);
    with txtTimeout do r.WriteString(Name, Text);
    with chkHiddenExcel do r.WriteBool(Name, Checked);
    with chkNoUpdateExcel do r.WriteBool(Name, Checked);
  finally
    r.Free;
  end;

  // Запишемв журнал Windows сообщение из протокола
  try
    EventLog := RegisterEventSource(nil, PChar(Name));
    strMsg := memLog.Lines.DelimitedText;
    MyMsg[0] := 'A test event message';
    MyMsg[1] := nil;
    ReportEvent(EventLog, EVENTLOG_INFORMATION_TYPE, 0, 0, nil, 1, 0, @MyMsg, nil);
  except
  end;
end;

function TfrmASR.GetOpenSSL(combo: TComboBox): TIdSSLVersion;
begin
  // Вернуть значение по индексу
  Result := sslvSSLv23;
  case combo.ItemIndex of
    0: Result := sslvSSLv2;
    1: Result := sslvSSLv23;
    2: Result := sslvSSLv3;
    3: Result := sslvTLSv1;
    4: Result := sslvTLSv1_1;
    5: Result := sslvTLSv1_2;
  end;
end;

procedure TfrmASR.GetParams;
var
  r: TRegistry;
  oldChar: Char;
begin
  // Грузим парметры из реестра
  r := TRegistry.Create;
  try
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);

    // Остальные поля без дублей в переменных
    with txtMailLogin do Text := r.ReadString(Name);
    with txtMailPassword do Text := r.ReadString(Name);
    with txtMailPortIMAP do Text := r.ReadString(Name);
    with txtMailPortSMTP do Text := r.ReadString(Name);
    with txtMailRecipient do Text := r.ReadString(Name);
    with txtMailServerIMAP do Text := r.ReadString(Name);
    with txtMailServerSMTP do Text := r.ReadString(Name);
    with txtMailSubject do Text := r.ReadString(Name);
    with memMailBody do
    begin
      oldChar := Lines.Delimiter;
      Lines.Delimiter := '/';
      Lines.DelimitedText := r.ReadString(Name);
      Lines.Delimiter := oldChar;
    end;
    with txtMailSender do Text := r.ReadString(Name);
    with cmbCharSet do
      if r.ValueExists(Name) then
        Text := r.ReadString(Name);
    with cmbSMTP_TLS do
      if r.ValueExists(Name) then
        Text := r.ReadString(Name);
    with cmbIMAP_TLS do
      if r.ValueExists(Name) then
        Text := r.ReadString(Name);
    with cmbSMTP_SSL do
      if r.ValueExists(Name) then
        Text := r.ReadString(Name);
    with cmbIMAP_SSL do
      if r.ValueExists(Name) then
        Text := r.ReadString(Name);
    with txtTimeout do
      if r.ValueExists(Name) then
        Text := r.ReadString(Name);
    with chkHiddenExcel do Checked := r.ReadBool(Name);
    with chkNoUpdateExcel do Checked := r.ReadBool(Name);
  finally
    r.Free;
  end;
end;

function TfrmASR.GetUseTLS(combo: TComboBox): TIdUseTLS;
begin
  // Вернуть значение по индексу
  Result := utNoTLSSupport;
  case combo.ItemIndex of
    0: Result := utNoTLSSupport;
    1: Result := utUseImplicitTLS;
    2: Result := utUseRequireTLS;
    3: Result := utUseExplicitTLS;
  end;
end;

function TfrmASR.MailPrepareMessage: Boolean;
var
  att: TIdAttachmentFile;
  sExtFile, sTempEmlFile: string;
begin
  // Подготовить файл .EML для отправки получателю
  _IdMsg.Clear;
  _IdMsg.From.Text := txtMailSender.Text;  // От кого
  _IdMsg.ReceiptRecipient.Text := _IdMsg.From.Address;  // Кому ответ о прочтении
  _IdMsg.Recipients.EMailAddresses := txtMailRecipient.Text;  // Получатель
  _IdMsg.Subject := txtMailSubject.Text;  // Тема письма

  if cmbCharSet.ItemIndex < 0 then
    cmbCharSet.Text := cmbCharSet.Items.Strings[0];
  _IdMsg.CharSet := cmbCharSet.Text;
  _IdMsg.Body.Text := memMailBody.Lines.Text;

  // Флаги для письма для сохранения в файл
  _IdMsg.Flags := [mfDraft, mfSeen, mfRecent];

  // Добавить файлы если есть
  with FileListAtTempFolder do
  try
    while Count > 0 do
    begin
      att := TIdAttachmentFile.Create(_IdMsg.MessageParts,
          FsFolderTemp + Strings[0]);
      att.CharSet := _IdMsg.CharSet;
      sExtFile := LowerCase(ExtractFileExt(Strings[0]));
      if sExtFile = '.csv' then
        att.ContentType := 'text/csv'
      else if sExtFile = '.txt' then
        att.ContentType := 'text/plain'
      else if sExtFile = '.xls' then
        att.ContentType := 'application/vnd.ms-excel'
      else if sExtFile = '.xlsx' then
        att.ContentType := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      else if sExtFile = '.doc' then
        att.ContentType := 'application/msword'
      else if sExtFile = '.xls' then
        att.ContentType := 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      else
        att.ContentType := 'application/x-binary';
      Delete(0);
    end;
  finally
    Free;
  end;

  // Сохранить файл письма перед отправкой
  sTempEmlFile := FsFolderTemp + 'new_message.eml';
  if FileExists(sTempEmlFile) then
    UN_DeleteFiles(Handle, sTempEmlFile, False);
  _IdMsg.SaveToFile(sTempEmlFile);

  // Показать что в папке получилось
  ShellExecute(Handle, 'open', PChar(FsFolderTemp), nil, nil, SW_NORMAL);
  Result := True;
end;

procedure TfrmASR.PrepareProgress(iQuantity: Integer);
var
  iOneRowHeight: Integer;
  iCounter: Integer;
begin
  // Разделить область прогресс-бара на указаное кол-во полос, но не больше чем
  // высота минимальной полосы (1 точка полоса + 2 точки границы, соседние
  // границы совпадают)
  if iQuantity > (imgProgress.Height - 1) div 2 then
    iQuantity := (imgProgress.Height - 1) div 2
  else if iQuantity <= 0 then
    iQuantity := 1;
  iOneRowHeight := imgProgress.Height div iQuantity;
  imgProgress.Tag := iOneRowHeight;

  // Окрасить каждую полосу в тёмный красный цвет
  with imgProgress.Picture.Bitmap.Canvas do
  begin
    Pen.Color := clBlack;
    Pen.Width := 1;
    Brush.Color := clMaroon;
    FillRect(ClipRect);
    for iCounter := 0 to iQuantity - 1 do
      Rectangle(0, iCounter * iOneRowHeight, Width,
          (iCounter + 1) * iOneRowHeight);
  end;
end;

procedure TfrmASR.SetLog(const Value: string);
begin
  // Сообщения в протокол с временем события
  memLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) + '> ' + Value);
end;

procedure TfrmASR.timAutomatTimer(Sender: TObject);
begin
  // Контроль автомата поэтапно
  case FStage of
    stNone:
      begin
        // Первичное состояние. Посылаем "подключение" и ждём коннекта
        PostMessage(frmServerDB.cmdConnect.Handle, BM_CLICK, 0, 0);
      end;

    stWaitConnect:
      begin
        // Если есть связь с базой, то запуск потоков
        if frmServerDB.Database._SQLC.Connected then
          log := 'Потоков запущено: ' + IntToStr(frmServerDB.RunQueryAtThread);
      end;

    stAfterFinishThreads:
      begin
        // Потоки завершены без ошибок - послать письмо
        timAutomat.Enabled := False;
        if not MailPrepareMessage then
          raise Exception.Create('Ошибка подготовки письма для отправки!');
        log := 'Посылаем письмо';
        DoSendMail;
        timAutomat.Enabled := True;
      end;

    stAfterSendMail:
      begin
        // Отправка завершена. Удалить временную папку в корзину
        UN_DeleteFiles(Handle, FsFolderTemp, True);
        FStage := stFinishSuccess;
      end;

    stFinishSuccess, stFinishWithError:
      begin
        // Закрыть приложение
        timAutomat.Enabled := False;
        log := 'Полный цикл завершён.';

        // Проверяем элементы автоматизации
        if (ParamCount > 0) and (UpperCase(ParamStr(1)) = 'AUTO') then
          PostMessage(Handle, WM_CLOSE, 0, 0);
      end;
  end;
end;

procedure TfrmASR._IdIMAPConnected(Sender: TObject);
begin
  // Ставим зелёный включенный индикатор
  imagesIndicator.GetBitmap(6, imgIMAP.Picture.Bitmap);
  imgIMAP.Repaint;
  log := 'Подключен IMAP4';
end;

procedure TfrmASR._IdIMAPDisconnected(Sender: TObject);
begin
  // Ставим красный включенный индикатор
  imagesIndicator.GetBitmap(2, imgIMAP.Picture.Bitmap);
  imgIMAP.Repaint;
  log := 'Соединение IMAP4 завершено';
end;

procedure TfrmASR._IdSMTPConnected(Sender: TObject);
begin
  // Ставим зелёный включенный индикатор
  imagesIndicator.GetBitmap(6, imgSMTP.Picture.Bitmap);
  imgSMTP.Repaint;
  log := 'Подключен SMTP';
end;

procedure TfrmASR._IdSMTPDisconnected(Sender: TObject);
begin
  // Ставим красный включенный индикатор
  imagesIndicator.GetBitmap(2, imgSMTP.Picture.Bitmap);
  imgSMTP.Repaint;
  log := 'Соединение SMTP завершено';
end;

procedure TfrmASR._IdSMTPFailedRecipient(Sender: TObject; const AAddress, ACode,
  AText: string; var VContinue: Boolean);
var
  s: string;
begin
  // Сообщить об ошибке
  s := 'Произошла ошибка доставки письма: AAddress = "%s", ACode = "%s", ' +
      'AText = "%s", VContinue = "%s"';
  log := Format(s, [AAddress, ACode, AText, BoolToStr(VContinue, True)]);
  VContinue := True;
end;

end.
