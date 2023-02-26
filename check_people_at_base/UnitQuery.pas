unit UnitQuery;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.CheckLst, Data.DB, Data.Win.ADODB, Vcl.ExtCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.Menus, System.IniFiles, System.StrUtils,
  System.Win.Registry, UN_DiskUtils, System.Win.ComObj;

type  // ����� ��� �������� � ���� ������. ����� ��������� ������� � �����
  TMyThreadQuery = class(TThread)
  private
    // ������� ������
    FTable: string;         // ��� ���� ������ ��� �������
    FSNILS: string;         // ����� ��� ������ ��� ������������
    FSQLC: TADOConnection;  // ����� � ����� ������
    FRow, FCol: Integer;    // ������ � ������� ��� ������ ����������
    FLog: TStrings;         // ���� �������� ��������� �� ������� (���� �����)
    // ��������� ��������
    FSQLQ: TADOQuery;
    FMsg: string;           // ��������� (���� �� ������), ����������� � �����
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
  // _SQLC: - ����������� � ���� ������
  // _Table, _SNILS: - ���� ������ � ����� ��� ������� (������)
  // _endEvent: - ���������� ���������� ������
  // _Row, _Col: - ������ � ������� � Excel ��� ������ ����������
  // _log: - ��� ������ ����� � ����������
  inherited Create(True);

  // ���������� ������
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
  // ������ �� �����
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
  i�itizen: Integer;
begin
  // ������ ������ � ���� ������
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
    i�itizen := FSQLQ.FieldByName('ID').AsInteger;

    // ������ ������� �� �������, ��� ���������� ���� ���������

  except
    on E: Exception do
      FLog.Add('������ ������ ��� ' + FTable + ': ' + E.Message);
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
  af.[FName] like '%F6%' and atbl.Title like '%���%'

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
  WHERE [Title] like '%����%'

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


  /****** ������ ��� ������� SelectTopNRows �� ����� SSMS  ******/
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

  /****** ������ ��� ������� SelectTopNRows �� ����� SSMS  ******/
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
  WHERE ([FAMIL] = '���������' and [IMJA] = '�������' and [OTCH] = '����������') or
  ([FAMIL] = '����������' and [IMJA] = '��������' and [OTCH] = '����������')

  *)

