unit UnitQuery;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.CheckLst, Data.DB, Data.Win.ADODB, Vcl.ExtCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.Menus, System.IniFiles, System.StrUtils,
  System.Win.Registry, UN_DiskUtils, System.Win.ComObj;

type  // Поток для запросов к базе данных. После обработки возврат в форму
  TMyThreadQuery = class(TThread)
  private
    // Входные данные
    FTable: string;         // Имя базы данных для доступа
    FSNILS: string;         // СНИЛС для поиска без разделителей
    FSQLC: TADOConnection;  // Связь с базой данных
    FRow, FCol: Integer;    // Строка и колонка для вывода результата
    FLog: TStrings;         // Куда выводить сообщения об ошибках (если будут)
    // Локальные свойства
    FSQLQ: TADOQuery;
    procedure SyncQuery;
  protected
    procedure Execute; override;
  public
    constructor Create(_SQLC: TADOConnection; _Table, _SNILS: string;
        _endEvent: TNotifyEvent; _Row, _Col: Integer; _log: TStrings);
    destructor Destroy; override;
    property iRow: Integer read FRow;
    property iCol: Integer read FCol;
    property oSQLQ: TADOQuery read FSQLQ;
    property sTable: string read FTable;
    property sSNILS: string read FSNILS;
  end;

implementation

{ TMyThreadQuery }

constructor TMyThreadQuery.Create(_SQLC: TADOConnection; _Table,
    _SNILS: string; _endEvent: TNotifyEvent; _Row, _Col: Integer;
    _log: TStrings);
begin
  // _SQLC: - подключение к базе данных
  // _Table, _SNILS: - база данных и СНИЛС для запроса (поиска)
  // _endEvent: - обработчик завершения потока
  // _Row, _Col: - строка и столбец в Excel для вывода результата
  // _log: - для вывода строк с результами
  inherited Create(True);

  // Подготовка потока
  FTable := _Table;
  FSNILS := _SNILS;
  FSQLC := _SQLC;
  FRow := _Row;
  FCol := _Col;
  OnTerminate := _endEvent;
  FLog := _log;

  FSQLQ := TADOQuery.Create(nil);
  FSQLQ.Connection := FSQLC;

  Priority := tpNormal;
  Resume;
end;

destructor TMyThreadQuery.Destroy;
begin
  // Убрать за собой
  try
    try
      if FSQLQ.Active then
        FSQLQ.Close;
    finally
      FSQLQ.Free;
    end;
  except
  end;
  inherited;
end;

procedure TMyThreadQuery.Execute;
begin
  inherited;
  FreeOnTerminate := True;
  //Synchronize(SyncQuery);
  SyncQuery;
  Terminate;
end;

procedure TMyThreadQuery.SyncQuery;
begin
  // Делаем запрос к базе данных
  try
    with FSQLQ.SQL do
    begin
      Clear;
      Add('SELECT TOP (100) [ID], [FAMIL], [IMJA], [OTCH], ' +
          'FORMAT([DROG], ''dd.MM.yyyy'') as DROG, [POL], ' +
          'REPLACE(REPLACE(NPS, ''-'', ''''), '' '', '''') AS NPS,[pku]');
      Add('FROM [' + FTable + '].[dbo].[F2]');
      Add('WHERE REPLACE(REPLACE(NPS, ''-'', ''''), '' '', '''') = ' +
          QuotedStr(FSNILS));
    end;
    FSQLQ.Open;
  except
    on E: Exception do
      FLog.Add('Ошибка потока для ' + FTable + ': ' + E.Message);
  end;
end;

end.
