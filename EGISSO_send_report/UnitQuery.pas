unit UnitQuery;

{
������ ���������� ����� ��� ���������� �������� � ��������� ������.
���������� ���������� ����������� � �����, ��� ��� ������ �� �������
�� �������� ����� (��� ������ ����/������ ����� ������ ���� ������).

������: 2023-02-26
���������: 2023-03-05
�����: ������ �.�.

}

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
    FFileName: string;      // ��� �����, ���� ��������� ���������
    FSQLC: TADOConnection;  // ����� � ����� ������
    FLog: TStrings;         // ���� �������� ��������� �� ������� (���� �����)
    // ��������� ��������
    FSQLQ: TADOQuery;
    FHasError: Boolean;
    FIndex: Integer;
    // ������ �������������� ����������, ������������ ���������
    FAppExcel: OleVariant;    // ���������� Excel
    FWB, FSheet: OleVariant;  // �������� ����� � �������� ���� (�����
                              // ���������� ������������)
    FFirstLine: TStringList;  // ������ ������ �� ������� (������ ����)
    iLastRow, iLastCol: Integer;
    arrColumnsSize: array of Integer;  // ������ �������� ��������
    procedure SyncQuery;
    procedure SyncSaveToFile;
    procedure AddInfoToLog;
    procedure SyncCreateExcelFile; // ������ ������ - ���������� Excel
    procedure SyncUpdateProgress;  // �������� �������� �� ������� �����
    procedure SyncFormatTable;  // �������� ������� ���������
  protected
    procedure Execute; override;
  public
    constructor Create(_SQLC: TADOConnection; _Table, _FileName: string;
        _endEvent: TNotifyEvent; _log: TStrings; iIndex: Integer);
    destructor Destroy; override;
    property oSQLQ: TADOQuery read FSQLQ;
    property sTable: string read FTable;
    property HasError: Boolean read FHasError;
  end;

implementation

uses
  UnitMain;

{ TMyThreadQuery }

procedure TMyThreadQuery.AddInfoToLog;
begin
  // ������ �������� � ���������� ������
  FLog.Add('����� ��� "' + FTable + '" �������� ���� ������');
end;

procedure TMyThreadQuery.SyncCreateExcelFile;
var
  j: Integer;
  iFileFormat: Variant;
begin
  // ��������� ���� � Excel'� � ��������� ���
  // ������������ � Excel (����� ��������� ��� ������������ ������)
  FFirstLine := TStringList.Create;
  FAppExcel := CreateOleObject('Excel.Application');

  // ��������� �������� � ������� ������
  FAppExcel.DisplayAlerts := False;
  FAppExcel.ScreenUpdating := not frmASR.chkNoUpdateExcel.Checked;
  FAppExcel.Visible := not frmASR.chkHiddenExcel.Checked;

  // ������� ����� ����� � ��������� �
  FWB := FAppExcel.Workbooks.Add;
  FLog.Add('�������� � "' + FFileName + '"');
  iFileFormat := 56;
  if LowerCase(ExtractFileExt(FFileName)) = '.xlsx' then
    iFileFormat := 51
  else if LowerCase(ExtractFileExt(FFileName)) = '.xlsm' then
    iFileFormat := 52;
  try
    FWB.SaveAs(FFileName, iFileFormat);  // 56 = XLS, 51 = XLSX, 52 = XLSM
  except
    on E: Exception do
      FLog.Add('������ ������ ��� "' + FTable +
          '": ������ ���������� �����: ' + E.Message);
  end;

  // ����������� � ������� ������ � ��������� ���
  FSheet := FWB.ActiveSheet;
  FSheet.Cells[1, 1] := '��������� ���������� �������';
  iLastCol := FSQLQ.FieldCount;
  SetLength(arrColumnsSize, iLastCol);
  for j := 0 to iLastCol - 1 do
  begin
    // ������ ��������� �������� � �������
    FFirstLine.Add(FSQLQ.FieldDefList.FieldDefs[j].Name);
    FSheet.Cells[2, j + 1] := FFirstLine.Strings[j];
    arrColumnsSize[j] := Length(FFirstLine.Strings[j]);
  end;
end;

procedure TMyThreadQuery.SyncFormatTable;
var
  j: Integer;
begin
  // ��������� �������
  // ��������� (������ 1)
  FSheet.Range[FSheet.Cells[1, 1], FSheet.Cells[1, iLastCol]].Select;
  FAppExcel.Selection.HorizontalAlignment := -4108;
  FAppExcel.Selection.VerticalAlignment := -4108;
  FAppExcel.Selection.WrapText := True;
  FAppExcel.Selection.Orientation := 0;
  FAppExcel.Selection.AddIndent := False;
  FAppExcel.Selection.IndentLevel := 0;
  FAppExcel.Selection.ShrinkToFit := False;
  FAppExcel.Selection.ReadingOrder := -5002;
  FAppExcel.Selection.MergeCells := True;
  FAppExcel.Selection.Font.Name := 'Times New Roman';
  FAppExcel.Selection.Font.FontStyle := 'Bold';
  FAppExcel.Selection.Font.Size := 13;
  FAppExcel.Selection.RowHeight := 45;
  FAppExcel.Selection.Borders[5].LineStyle := -4142;
  FAppExcel.Selection.Borders[6].LineStyle := -4142;
  FAppExcel.Selection.Borders[7].LineStyle := 1;
  FAppExcel.Selection.Borders[7].ColorIndex := 1;
  FAppExcel.Selection.Borders[7].Weight := -4138;
  FAppExcel.Selection.Borders[8].LineStyle := 1;
  FAppExcel.Selection.Borders[8].ColorIndex := 1;
  FAppExcel.Selection.Borders[8].Weight := -4138;
  FAppExcel.Selection.Borders[9].LineStyle := 1;
  FAppExcel.Selection.Borders[9].ColorIndex := 1;
  FAppExcel.Selection.Borders[9].Weight := -4138;
  FAppExcel.Selection.Borders[10].LineStyle := 1;
  FAppExcel.Selection.Borders[10].ColorIndex := 1;
  FAppExcel.Selection.Borders[10].Weight := -4138;
  FAppExcel.Selection.Borders[11].LineStyle := -4142;
  FAppExcel.Selection.Borders[12].LineStyle := -4142;

  // ��������� �������� [������ 2]
  FSheet.Range[FSheet.Cells[2, 1], FSheet.Cells[2, iLastCol]].Select;
  FAppExcel.Selection.HorizontalAlignment := -4108;
  FAppExcel.Selection.VerticalAlignment := -4108;
  FAppExcel.Selection.WrapText := True;
  FAppExcel.Selection.AddIndent := False;
  FAppExcel.Selection.IndentLevel := 0;
  FAppExcel.Selection.ShrinkToFit := False;
  FAppExcel.Selection.ReadingOrder := -5002;
  FAppExcel.Selection.Font.Name := 'Times New Roman';
  FAppExcel.Selection.Font.FontStyle := 'Bold';
  FAppExcel.Selection.Font.Size := 12;
  FAppExcel.Selection.RowHeight := 30;
  FAppExcel.Selection.Borders[5].LineStyle := -4142;
  FAppExcel.Selection.Borders[6].LineStyle := -4142;
  FAppExcel.Selection.Borders[7].LineStyle := 1;
  FAppExcel.Selection.Borders[7].ColorIndex := 1;
  FAppExcel.Selection.Borders[7].Weight := 1;
  FAppExcel.Selection.Borders[8].LineStyle := 1;
  FAppExcel.Selection.Borders[8].ColorIndex := 1;
  FAppExcel.Selection.Borders[8].TintAndShade := 0;
  FAppExcel.Selection.Borders[8].Weight := -4138;
  FAppExcel.Selection.Borders[9].LineStyle := 1;
  FAppExcel.Selection.Borders[9].ColorIndex := 1;
  FAppExcel.Selection.Borders[9].TintAndShade := 0;
  FAppExcel.Selection.Borders[9].Weight := 1;
  FAppExcel.Selection.Borders[10].LineStyle := 1;
  FAppExcel.Selection.Borders[10].ColorIndex := 1;
  FAppExcel.Selection.Borders[10].TintAndShade := 0;
  FAppExcel.Selection.Borders[10].Weight := 1;
  FAppExcel.Selection.Borders[11].LineStyle := 1;
  FAppExcel.Selection.Borders[11].ColorIndex := 1;
  FAppExcel.Selection.Borders[11].TintAndShade := 0;
  FAppExcel.Selection.Borders[11].Weight := 1;
  FAppExcel.Selection.Borders[12].LineStyle := 1;
  FAppExcel.Selection.Borders[12].ColorIndex := 1;
  FAppExcel.Selection.Borders[12].TintAndShade := 0;
  FAppExcel.Selection.Borders[12].Weight := 1;

  // ��������� ����� �������
  FSheet.Range[FSheet.Cells[3, 1], FSheet.Cells[iLastRow, iLastCol]].Select;
  FAppExcel.Selection.VerticalAlignment := -4108;
  FAppExcel.Selection.WrapText := False;
  FAppExcel.Selection.AddIndent := False;
  FAppExcel.Selection.IndentLevel := 0;
  FAppExcel.Selection.ShrinkToFit := False;
  FAppExcel.Selection.ReadingOrder := -5002;
  FAppExcel.Selection.Font.Name := 'Times New Roman';
  FAppExcel.Selection.Font.Size := 11;
  FAppExcel.Selection.RowHeight := 20;
  FAppExcel.Selection.Borders[5].LineStyle := -4142;
  FAppExcel.Selection.Borders[6].LineStyle := -4142;
  FAppExcel.Selection.Borders[7].LineStyle := 1;
  FAppExcel.Selection.Borders[7].ColorIndex := 1;
  FAppExcel.Selection.Borders[7].Weight := 1;
  FAppExcel.Selection.Borders[8].LineStyle := 1;
  FAppExcel.Selection.Borders[8].ColorIndex := 1;
  FAppExcel.Selection.Borders[8].TintAndShade := 0;
  FAppExcel.Selection.Borders[8].Weight := 1;
  FAppExcel.Selection.Borders[9].LineStyle := 1;
  FAppExcel.Selection.Borders[9].ColorIndex := 1;
  FAppExcel.Selection.Borders[9].TintAndShade := 0;
  FAppExcel.Selection.Borders[9].Weight := 1;
  FAppExcel.Selection.Borders[10].LineStyle := 1;
  FAppExcel.Selection.Borders[10].ColorIndex := 1;
  FAppExcel.Selection.Borders[10].TintAndShade := 0;
  FAppExcel.Selection.Borders[10].Weight := 1;
  FAppExcel.Selection.Borders[11].LineStyle := 1;
  FAppExcel.Selection.Borders[11].ColorIndex := 1;
  FAppExcel.Selection.Borders[11].TintAndShade := 0;
  FAppExcel.Selection.Borders[11].Weight := 1;
  FAppExcel.Selection.Borders[12].LineStyle := 1;
  FAppExcel.Selection.Borders[12].ColorIndex := 1;
  FAppExcel.Selection.Borders[12].TintAndShade := 0;
  FAppExcel.Selection.Borders[12].Weight := 1;

  // �������� ������ ��������
  FSheet.Cells[1, 1].Select;
  for j := 0 to iLastCol - 1 do
    FSheet.Columns[j + 1].ColumnWidth := arrColumnsSize[j] + 3;
end;

constructor TMyThreadQuery.Create(_SQLC: TADOConnection; _Table,
  _FileName: string; _endEvent: TNotifyEvent; _log: TStrings; iIndex: Integer);
begin
  // _SQLC: - ����������� � ���� ������
  // _Table, _FileName: - ���� ������ � ��� �����
  // _endEvent: - ���������� ���������� ������
  // _log: - ��� ������ ����� � ����������
  inherited Create(True);

  // ���������� ������
  FTable := _Table;
  FFileName := _FileName;
  FSQLC := _SQLC;
  OnTerminate := _endEvent;
  FLog := _log;
  FHasError := False;
  FIndex := iIndex;

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
var
  iUpdateAfterRows: Integer;  // ����� ������� ����� ���������
  j, jTemp: Integer;
  FData: Variant;  // ��� ������ � Excel �� �������
begin
  inherited;
  FreeOnTerminate := True;
  try
    SyncQuery;  // ������� ������

    // ���������� �������� ���������� ��������
    iUpdateAfterRows := FSQLQ.RecordCount div frmASR.Width;
    if iUpdateAfterRows <= 0 then
      iUpdateAfterRows := 1;  // ������ ���� ��� ����� ������� ������
    SyncCreateExcelFile;  // ������� ���� Excel

    // ������� ������ �� ������� � �����
    iLastRow := 2;
    FData := VarArrayCreate([1, iLastCol], varVariant);
    FSQLQ.First;
    while not FSQLQ.Eof do
    begin
      Inc(iLastRow);
      for j := 0 to iLastCol - 1 do
      begin
        FData[j + 1] := FSQLQ.FieldByName(FFirstLine.Strings[j]).AsString;
        if FSQLQ.FieldDefList.FieldDefs[j].DataType = ftString then
          // ��� ��������� �������� ������ ������ ������ ���������
          FSheet.Cells[iLastRow, j + 1].NumberFormat := '@';
        jTemp := Length(FSQLQ.FieldByName(FFirstLine.Strings[j]).AsString);
        if arrColumnsSize[j] < jTemp then
          arrColumnsSize[j] := jTemp;
      end;
      FSheet.Range[FSheet.Cells[iLastRow, 1],
          FSheet.Cells[iLastRow, iLastCol]].Value := FData;

      // ���� ����, �� �������� �������
      if FSQLQ.RecNo mod iUpdateAfterRows = 0 then
        Synchronize(SyncUpdateProgress);

      FSQLQ.Next;
    end;
    Synchronize(SyncUpdateProgress);  // ������ ���������� �����������

    SyncFormatTable;  // ��������� �������
    SyncSaveToFile;  // ��������� ���� � ������� Excel
  except
    on E: Exception do
    begin
      FLog.Add('������ ������ ��� "' + FTable + '": ' + E.Message);
      FHasError := True;
    end;
  end;
  if Assigned(FFirstLine) then
  try
    FFirstLine.Free;
  except
  end;
  FAppExcel := Unassigned;
  Synchronize(AddInfoToLog);
  Terminate;
end;

procedure TMyThreadQuery.SyncQuery;
begin
  // ������ ������ � ���� ������
  with FSQLQ.SQL do
  begin
    Clear;
    Add('USE ' + FTable);
    Add('select (select top 1 DB_NAME from DATABASEINFO) as "�����",');
    Add('dt as "����", fileName as "����", fType as "���",');
    Add('msgId as "���� ID", packageId as "�����",');
    Add('case when Smevresp is null then ''��� ������'' else ''��'' end as "���������",');
    Add('cntRecSucc as "�������", cntRecError as "������"');
    Add('from eg_lr_fact_smev');
  end;
  FSQLQ.Open;
end;

procedure TMyThreadQuery.SyncSaveToFile;
begin
  try
    // �������� ������������ ��������
    FAppExcel.DisplayAlerts := True;
    FAppExcel.ScreenUpdating := True;
    FAppExcel.Visible := True;
  except
    on E: Exception do
      FLog.Add('������ ������ ��� "' + FTable + '": ' + E.Message);
  end;
  try
    // ������� ����� �� Excel
    FAppExcel.ActiveWorkbook.Save;
    FAppExcel.ActiveWorkbook.Close(True);
    FAppExcel.Quit;
  except
    on E: Exception do
    begin
      FLog.Add('������ ������ ��� "' + FTable + '": ' + E.Message);
      FHasError := True;
    end;
  end;
end;

procedure TMyThreadQuery.SyncUpdateProgress;
begin
  frmASR.DrawLineProgress(FIndex, FSQLQ.RecNo, FSQLQ.RecordCount);
end;

end.
