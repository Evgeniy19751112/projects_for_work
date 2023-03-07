unit UnitFormBase;
{
  Модуль (форма) управления подключением к базе данных и настройки шаблонов

Начато: 2023-03-05
Завершено: 2023-03-05
Автор: Тявкин Е.Н.
}

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.Buttons, Vcl.ImgList, Vcl.StdCtrls, Vcl.CheckLst, Vcl.ExtCtrls,
  UnitBases, System.StrUtils, System.Win.Registry, UnitQuery;

type
  TfrmServerDB = class(TForm)
    grbParams: TGroupBox;
    imgIndicator: TImage;
    txtSrv: TLabeledEdit;
    txtLogin: TLabeledEdit;
    txtPass: TLabeledEdit;
    cmdConnect: TButton;
    grbDB: TGroupBox;
    lst_chk_Bases: TCheckListBox;
    grbReports: TGroupBox;
    lblReportBaseInfo: TLabel;
    txtPattern: TLabeledEdit;
    txtExample: TLabeledEdit;
    BalloonHint1: TBalloonHint;
    ImagesIndicator: TImageList;
    cmdSelectSQL: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure MozhnoZayti(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure cmdSelectSQLClick(Sender: TObject);
    procedure cmdConnectClick(Sender: TObject);
    procedure lst_chk_BasesClick(Sender: TObject);
    procedure txtPatternChange(Sender: TObject);
  private
    FDatabase: TMyDatabase;   // Управление доступом к базе данных
    FThrCounter: Byte;        // Счётчик потоков для контроля завершения всех
    function GetPatternString(var inStr: string): string;
    procedure SetLog(const Value: string);
    procedure CheckPattern;
  public
    function RunQueryAtThread: Integer;  // Запустить запросы в потоках (вернуть кол-во)
    property Database: TMyDatabase read FDatabase;
    procedure DecThrCounter;
    property log: string write SetLog;
    property ThrCounter: Byte read FThrCounter;
  end;

var
  frmServerDB: TfrmServerDB;

implementation

uses
  UnitMain;

const // Локальные константы
  csl_RegKey = '\SOFTWARE\TZN\EGISSO\send_report';

{$R *.dfm}

procedure TfrmServerDB.FormCreate(Sender: TObject);
begin
  // Инициализация
  FDatabase := TMyDatabase.Create(Self, frmASR.memLog, ImagesIndicator,
      imgIndicator, lst_chk_Bases);
  FThrCounter := 0;

  // Ставим красный выключенный индикатор
  ImagesIndicator.GetBitmap(0, imgIndicator.Picture.Bitmap);

  // Грузим параметры
  FDatabase.GetParams;
  txtSrv.Text := FDatabase.sHostName;
  txtLogin.Text := FDatabase.sLogin;
  txtPass.Text := FDatabase.sPassword;

  MozhnoZayti(nil);
  log := 'Загрузился';
end;

procedure TfrmServerDB.FormDestroy(Sender: TObject);
var
  r: TRegistry;
  s: string;
begin
  // Сохраним в реестре имена отмеченных баз
  r := TRegistry.Create;
  try
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);
    s := '';
    while lst_chk_Bases.Count > 0 do
    begin
      if lst_chk_Bases.Checked[0] then
      begin
        s := s + ' ' + lst_chk_Bases.Items.Strings[0];
        if Assigned(lst_chk_Bases.Items.Objects[0])  then
          r.WriteString(lst_chk_Bases.Items.Strings[0],
              (lst_chk_Bases.Items.Objects[0] as TStringList).DelimitedText);
      end;
      lst_chk_Bases.Items.Delete(0);
    end;
    s := Trim(s);
    r.WriteString('DB_Names', s);
  finally
    r.Free;
  end;

  // Закрыть все соединения
  FDatabase.Free;
end;

procedure TfrmServerDB.MozhnoZayti(Sender: TObject);
begin
  // Если пользователь ничего не ввёл, то и не пускать его
  cmdConnect.Enabled := True;
  if (Trim(txtSrv.Text) = '') or
      (Trim(txtLogin.Text) = '') or
      (Trim(txtPass.Text) = '')
  then
    cmdConnect.Enabled := False;
end;

procedure TfrmServerDB.cmdSelectSQLClick(Sender: TObject);
begin
  // Сделать тестовый запрос ко всем базам данных и сформировать файлы отчётов
  if not Database._SQLC.Connected then
    cmdConnect.Click;
  RunQueryAtThread;
end;

procedure TfrmServerDB.DecThrCounter;
begin
  if FThrCounter > 0 then
    Dec(FThrCounter);
end;

procedure TfrmServerDB.SetLog(const Value: string);
begin
  // Сообщения в протокол с временем события
  frmASR.memLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) +
      '> TfrmServerDB: ' + Value);
end;

function TfrmServerDB.RunQueryAtThread: Integer;
var
  j: Integer;
  sFileName, sPattern, sTemp: string;
begin
  // Запустить потоки на выполнение запроса к БД и формирования результата
  if not Database._SQLC.Connected then
    raise Exception.Create('Потоки не работают без подключения к базе');
  frmASR.Stage := stBeforeRunThreads;
  Result := 0;
  for j := 0 to lst_chk_Bases.Count - 1 do
  begin
    if lst_chk_Bases.Checked[j] then
    begin
      // Формируем полное имя файла
      sFileName := frmASR.sFolderTemp + lst_chk_Bases.Items.Strings[j];
      if Assigned(lst_chk_Bases.Items.Objects[j]) then
      try
        sTemp := (lst_chk_Bases.Items.Objects[j] as TStringList).Strings[1];
        sPattern := GetPatternString(sTemp);
        if sPattern <> '' then
          sFileName := Format(sTemp, [FormatDateTime(sPattern, Now)])
        else
          sFileName := sTemp;
        sFileName := frmASR.sFolderTemp + sFileName;
      except
      end;

      // Если такой файл есть, то убить его
      if FileExists(sFileName) then
        DeleteFile(sFileName);

      // Поток выгружается после завершения - нет смысла держать переменную
      TMyThreadQuery.Create(FDatabase._SQLC, lst_chk_Bases.Items.Strings[j],
          sFileName, frmASR.EndOfThread, frmASR.memLog.Lines, Result);

      Inc(FThrCounter);
      Inc(Result);
    end;
  end;
  frmASR.PrepareProgress(Result);
end;

procedure TfrmServerDB.lst_chk_BasesClick(Sender: TObject);
var
  j: Integer;
  lst: TStringList;
  oldEvent: TNotifyEvent;
begin
  // Выбор базы и вывод сведений о шаблоне для выбранной строки
  oldEvent := txtPattern.OnChange;
  txtPattern.OnChange := nil;
  j := lst_chk_Bases.ItemIndex;
  if j >= 0 then
  begin
    txtPattern.Text := '';
    txtExample.Text := '';
    lblReportBaseInfo.Caption := '';
    try
      if Assigned(lst_chk_Bases.Items.Objects[j]) then
      begin
        lst := lst_chk_Bases.Items.Objects[j] as TStringList;
        txtPattern.Text := lst.Strings[1];
        CheckPattern;
        lblReportBaseInfo.Caption := lst.Strings[0];
      end;
    except
      on E: Exception do
        log := E.Message;
    end;
  end;
  txtPattern.OnChange := oldEvent;
end;

procedure TfrmServerDB.CheckPattern;
var
  s, sPattern: string;
begin
  // Вывести по шаблону пример имени файла
  s := txtPattern.Text;
  sPattern := GetPatternString(s);
  if sPattern <> '' then
    txtExample.Text := Format(s, [FormatDateTime(sPattern, Now)])
  else
    txtExample.Text := s;
end;

procedure TfrmServerDB.txtPatternChange(Sender: TObject);
var
  lst: TStringList;
begin
  // При изменении записывать шаблон в выбранную строку
  if lst_chk_Bases.ItemIndex < 0 then
    Exit;
  with lst_chk_Bases do
    lst := Items.Objects[ItemIndex] as TStringList;
  if not Assigned(lst) then
    Exit;
  if lst.Count < 2 then
    Exit;
  lst.Strings[1] := txtPattern.Text;
  CheckPattern;
end;

procedure TfrmServerDB.cmdConnectClick(Sender: TObject);
begin
  // Подключить и сохранить (если не подключено)
  if FDatabase._SQLC.Connected then
  begin
    FDatabase._SQLC.Close;
    Exit;
  end;

  // Присвоить введёное пользователем в переменные
  FDatabase.SetConnectParams(txtSrv.Text, txtLogin.Text, txtPass.Text);
  frmASR.Stage := stWaitConnect;

  // Конектимся
  FDatabase.MyConnect;
end;

function TfrmServerDB.GetPatternString(var inStr: string): string;
var
  iPosL, iPosR: Integer;
  sTempL, sTempR: string;
begin
  // Извлекает фрагмент из строки с шаблоном между "<" & ">"
  // Во входной строке заменяет этот участок на %s (включая угловые скобки)
  // Возвращает чистый шаблон из участка между скобок.
  // Если шаблона нет, то вернёт пустую строку не меняя исходную
  Result := '';
  iPosL := PosEx('<', inStr);
  if iPosL <= 0 then
    Exit;
  sTempL := LeftStr(inStr, iPosL - 1);
  sTempR := RightStr(inStr, Length(inStr) - iPosL);
  iPosR := PosEx('>', sTempR);
  if iPosR <= 0 then
    Exit;
  Result := LeftStr(sTempR, iPosR - 1);
  Delete(sTempR, 1, iPosR);
  inStr := sTempL + '%s' + sTempR;
end;

end.
