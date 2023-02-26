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
    FMsg: string;           // Сообщение (инфо по заявке), возращаемое в форму
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
    property sMsg: string read FMsg;
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
  FMsg := '';

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
var
  qStatement: TADOQuery;
  iСitizen: Integer;
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
    iСitizen := FSQLQ.FieldByName('ID').AsInteger;

    // Делаем выборку по заявкам, где фигурирует этот гражданин

  except
    on E: Exception do
      FLog.Add('Ошибка потока для ' + FTable + ': ' + E.Message);
  end;
end;

end.

(*
/****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP (1000)
/*
[TName]
      ,[FName]
      ,[FType]
      ,[FLen1]
      ,[FLen2]
      ,[Nullable]
      ,[Def]
      ,[id]
      ,[Title]
*/
	  af.*
	  ,atbl.*
  FROM [ASP_TINDA_R].[dbo].[AspFields] as af
  LEFT JOIN [ASP_TINDA_R].[dbo].[AspTables] as atbl
  ON af.TName = atbl.TName
  WHERE --[TName] = 'F6'
  af.[FName] like '%F6%' and atbl.Title like '%сем%'

/****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP (1000) [ID]
      ,[F2_ID]
      ,[KOD]
      ,[DATWP]
      ,[DATRZ]
      ,[SUMMA]
      ,[SRDZ]
      ,[NZAPD]
      ,[NZAYV]
      ,[NDELA]
      ,[K_GSP]
      ,[STATE]
      ,[KISP]
      ,[TRUSTEE]
      ,[ATTRIBUTES]
      ,[TAB]
      ,[FRA_ID]
      ,[Source1]
      ,[Block]
      ,[BLOCK_WORK]
      ,[STATE2]
  FROM [ASP_TINDA_R].[dbo].[F6]
  WHERE [KOD] = '55110000000000' and [F2_ID] = 17247


/****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP (1000) [TName]
      ,[Title]
      ,[NotInStruct]
      ,[Type_ID]
      ,[id]
      ,[UNITE_ORDER]
      ,[ClearDB_SortOrder]
      ,[ASPFileGroups_ID]
      ,[N_ClearDBPartition_Part_ID]
      ,[ARH]
  FROM [ASP_TINDA_R].[dbo].[AspTables]
  WHERE [Title] like '%заяв%'

  /****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP (1000) [ID]
      ,[UR]
      ,[KOD]
      ,[NKOD]
      ,[NAIMP]
      ,[DOPFIELD]
      ,[DEACT_DATE]
      ,[DEACT_COMMENT]
      ,[SHORTNAME]
      ,[ISTIME]
      ,[S_TYPE_ID]
      ,[PRIOR]
  FROM [ASP_TINDA_R].[dbo].[S16]
  WHERE [NKOD] like '%265%'
  or [ID] = 56260000000000


  /****** Script for SelectTopNRows command from SSMS  ******/
SELECT TOP (1000) [ID]
      ,[PY]
      ,[NKAR]
      ,[FRA_ID]
      ,[FRA_STATUS]
      ,[FRA_REG_ID]
      ,[FRA_REG_STATUS]
      ,[DATZK]
      ,[FAMIL]
      ,[IMJA]
      ,[OTCH]
      ,[DROG]
      ,[NDETYCH]
      ,[PPROG]
      ,[KISP]
      ,[POL]
      ,[PSR]
      ,[PNM]
      ,[PDV]
      ,[PKV]
      ,[PRAB]
      ,[PROP]
      ,[DOPSV]
      ,[DPROP]
      ,[PWID]
      ,[NPS]
      ,[NATION_ID]
      ,[PPODR]
      ,[MROG]
      ,[pku]
      ,[puch]
      ,[NAPRV]
      ,[F17_ID]
      ,[ID_POLUCH]
      ,[CONSENT]
      ,[CELLULAR]
      ,[PIN_CODE]
      ,[PIN_CODE_ACTIVE]
      ,[PIN_CODE_DATE]
  FROM [ASP_TINDA_R].[dbo].[F2]
  WHERE [ID] = '17247'


  /****** Скрипт для команды SelectTopNRows из среды SSMS  ******/
SELECT TOP (1000) [ID]
      ,[F2_ID]
      ,[KOD]
      ,[DATWP]
      ,[DATRZ]
      ,[SUMMA]
      ,[SRDZ]
      ,[NZAPD]
      ,[NZAYV]
      ,[NDELA]
      ,[K_GSP]
      ,[STATE]
      ,[KISP]
      ,[TRUSTEE]
      ,[ATTRIBUTES]
      ,[TAB]
      ,[Source1]
      ,[Block]
      ,[BLOCK_WORK]
      ,[STATE2]
  FROM [ASP_TINDA_G].[dbo].[F6]
  WHERE [F2_ID] in (33954, 33955)

  select * from F_PENSFAM where F6IZM_ID = 148461

  /****** Скрипт для команды SelectTopNRows из среды SSMS  ******/
SELECT TOP (1000) [ID]
      ,[PY]
      ,[NKAR]
      ,[FRA_ID]
      ,[FRA_STATUS]
      ,[FRA_REG_ID]
      ,[FRA_REG_STATUS]
      ,[DATZK]
      ,[FAMIL]
      ,[IMJA]
      ,[OTCH]
      ,[DROG]
      ,[NDETYCH]
      ,[KISP]
      ,[POL]
      ,[PSR]
      ,[PNM]
      ,[PDV]
      ,[PKV]
      ,[PRAB]
      ,[DOPSV]
      ,[DPROP]
      ,[PWID]
      ,[NPS]
      ,[NATION_ID]
      ,[PPODR]
      ,[MROG]
      ,[pku]
      ,[puch]
      ,[NAPRV]
      ,[F17_ID]
      ,[ID_POLUCH]
      ,[CONSENT]
      ,[CELLULAR]
      ,[PIN_CODE]
      ,[PIN_CODE_ACTIVE]
      ,[PIN_CODE_DATE]
  FROM [ASP_TINDA_G].[dbo].[F2]
  WHERE ([FAMIL] = 'Соломатин' and [IMJA] = 'Евгений' and [OTCH] = 'Дмитриевич') or
  ([FAMIL] = 'Соломатина' and [IMJA] = 'Светлана' and [OTCH] = 'Евгеньевна')

  *)

