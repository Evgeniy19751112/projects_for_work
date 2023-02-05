unit UnitMainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.CheckLst, Data.DB, Data.Win.ADODB, Vcl.ExtCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.Menus, System.IniFiles, UnitQuery, Vcl.AppEvnts;

type
  TfrmCPD = class(TForm)
    grbParams: TGroupBox;
    txtSrv: TLabeledEdit;
    txtLogin: TLabeledEdit;
    txtPass: TLabeledEdit;
    cmdConnect: TButton;
    grbDB: TGroupBox;
    memLog: TMemo;
    cmdGetData: TButton;
    cmdConnectExcel: TButton;
    _ds: TDataSource;
    _SQLC: TADOConnection;
    _SQLQ: TADOQuery;
    lst_chk_Bases: TCheckListBox;
    grbSvyazka: TGroupBox;
    sgr_Source: TStringGrid;
    _pm: TPopupMenu;
    ApplicationEvents1: TApplicationEvents;
    timCallThread: TTimer;
    N1: TMenuItem;
    N2: TMenuItem;
    grpOptions: TGroupBox;
    chkVisual: TCheckBox;
    chkUpdate: TCheckBox;
    txtThreadsMax: TLabeledEdit;
    txtUpdateInterval: TLabeledEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure cmdConnectClick(Sender: TObject);
    procedure txtChange(Sender: TObject);
    procedure cmdGetDataClick(Sender: TObject);
    procedure cmdConnectExcelClick(Sender: TObject);
    procedure PopupItemClick(Sender: TObject);
    procedure UpdateSheetAfterQuery(Sender: TObject);
    procedure ApplicationEvents1Exception(Sender: TObject; E: Exception);
    procedure timCallThreadTimer(Sender: TObject);
    procedure txtThreadsMaxChange(Sender: TObject);
    procedure txtUpdateIntervalChange(Sender: TObject);
  private
    FsHostName: string;       // ����
    FsLogin: string;          // �����
    FsPassword: string;       // ������
    FsFolderName: string;     // ���� � ������
    FAppExcel: OleVariant;    // ���������� Excel
    FWB, FSheet: OleVariant;  // �������� ����� � �������� ���� (�����
                              // ���������� ������������)
    FFirstLine: TStringList;  // ������ ������ �� ������� (������ ����)
    FColumns: array [0..5] of Integer;  // ��������������� ������ ����� � Excel
    FlstTables: TStringList;  // ������ ��� ������ ��� ������������� ���������
    FThreadsRunning: ShortInt; // ���-�� ���������� �������
    FThreadsMaximum: ShortInt; // �������� ���������� �������
    FUpdateInterval: Integer;  // �������� (���-�� �����) ����� �����������
    procedure SetLog(const Value: string);
    procedure MyConnect;      // ����������� � �� (� �.�. ��������)
    procedure GetParams;      // ������ �� ������� ���������
    procedure SetParams;      // ������ �������� � �������
    procedure MozhnoZayti(Sender: TObject);
    function OpenIniFile: TMemIniFile;  // ������� INI ����
    procedure SaveColumnRelationships;  // ��������� ��������� ���������
    procedure LoadColumnRelationships;  // ��������� ��������� ���������
  public
    property log: string write SetLog;
  end;

var
  frmCPD: TfrmCPD;

implementation

uses
  System.StrUtils, System.Win.Registry, UN_DiskUtils, System.Win.ComObj;

{$R *.dfm}

const // ��������� ���������
  csl_RegKey = '\Software\TZN\CheckPeopleInDB';

procedure TfrmCPD.ApplicationEvents1Exception(Sender: TObject; E: Exception);
begin
  try
    log := 'Application Exception: Class ' +
        Sender.ClassName + ': ' + E.Message;
  except
  end;
end;

procedure TfrmCPD.cmdConnectExcelClick(Sender: TObject);
var
  j, iLastCol: Integer;
  tmp_menu: TMenuItem;
begin
  // ������������ � Excel
  try
    //FAppExcel := CreateOleObject('Excel.Application');
    FAppExcel := GetActiveOleObject('Excel.Application');
    FWB := FAppExcel.ActiveWorkbook;
    FSheet := FWB.ActiveSheet;
    FSheet.Cells[1, 1].Select;
    FSheet.Cells.SpecialCells(11).Select;
    iLastCol := FSheet.Cells.SpecialCells(11).Column;
    for j := 1 to iLastCol + 5 do
    begin
      tmp_menu := TMenuItem.Create(_pm);
      tmp_menu.Caption := FAppExcel.Cells[1, j].Value;
      if Trim(tmp_menu.Caption) = '' then
      begin
        tmp_menu.Caption := Chr(Ord('A') + (j - 1) mod 26);
        if j > 26 then
          tmp_menu.Caption := Chr(Ord('A') + j div 26 mod 26 - 1) +
              tmp_menu.Caption;
      end;
      tmp_menu.Tag := j;
      tmp_menu.OnClick := PopupItemClick;
      _pm.Items.Add(tmp_menu);
      FFirstLine.Add(tmp_menu.Caption);
    end;
    LoadColumnRelationships;
  except
    on E: Exception do
      log := E.Message;
  end;
end;

procedure TfrmCPD.cmdConnectClick(Sender: TObject);
begin
  // ��������� � ����������
  SetParams;
  MyConnect;
end;

procedure TfrmCPD.cmdGetDataClick(Sender: TObject);
var
  j: Integer;
  bFlag: Boolean;
  sName: string;
begin
  // ����� �������� ����� � �� ��������� ���������� ���� ����� ���������
  bFlag := False;
  for j := 1 to sgr_Source.RowCount - 1 do
    // ��� ���� �������� ��������� ������, �.�. ����� �������� ��� ��
    if (j <> 4) and (sgr_Source.Cells[1, j] = '') then
    begin
      bFlag := True;
      log := '��� ��������� ��� ' + sgr_Source.Cells[0, j];
      sgr_Source.Row := j;
      Break;
    end;
  if bFlag then
    Exit;

  // ��������� ���������� ��������� � ���������� � ���������
  SaveColumnRelationships;
  if _SQLC.Connected then
  begin
    log := '������ ��������� ������';
    FAppExcel.DisplayAlerts := False;

    // �������������� ������� ���������
    FAppExcel.Visible := chkVisual.Checked;
    cmdConnect.Enabled := False;
    cmdGetData.Enabled := False;
    cmdConnectExcel.Enabled := False;
    if not chkVisual.Checked then
    begin
      FAppExcel.ScreenUpdating := False;
      FUpdateInterval := 0;
      chkUpdate.Checked := False;
    end;
    FAppExcel.ScreenUpdating := chkUpdate.Checked and (FUpdateInterval = 0);
    txtThreadsMax.Enabled := False;
    txtUpdateInterval.Enabled := False;
    chkVisual.Enabled := False;
    chkUpdate.Enabled := False;

    // ��� ���������� �������� ������ (��� ��������� �������� ������)
    for j := 0 to 5 do
    begin
      sName := sgr_Source.Cells[1, j + 1];
      FColumns[j] := FFirstLine.IndexOf(sName) + 1;
    end;
    FlstTables.Clear;
    for j := 0 to lst_chk_Bases.Count - 1 do
      if lst_chk_Bases.Checked[j] then
      begin
        FlstTables.Add(lst_chk_Bases.Items.Strings[j]);
        FSheet.Cells[1, FColumns[5] + j].Value := lst_chk_Bases.Items.Strings[j];
      end;
    if FlstTables.Count <= 0 then
      Abort;

    timCallThread.Tag := FSheet.Cells.SpecialCells(11).Row;
    timCallThread.Enabled := True;
  end
  else
    log := '��� ����������� � �������';
end;

procedure TfrmCPD.FormCreate(Sender: TObject);
begin
  // �������������
  Constraints.MinHeight := Height;
  Constraints.MinWidth := Width;
  memLog.Lines.Clear;
  FsFolderName := ExtractFilePath(Application.ExeName);
  FsHostName := '';
  FsLogin := '';
  FsPassword := '';
  FThreadsMaximum := 1;
  FUpdateInterval := 100;
  GetParams;
  MozhnoZayti(nil);
  FAppExcel := Null;
  FWB := Null;
  FSheet := Null;
  FFirstLine := TStringList.Create;
  FlstTables := TStringList.Create;
  log := '����������';
  if cmdConnect.Enabled then
    MyConnect;

  // �������� ����� ��� ������ �����
  with sgr_Source.Cols[0] do
  begin
    Strings[0] := '������';
    Strings[1] := '�������';
    Strings[2] := '���';
    Strings[3] := '��������';
    Strings[4] := '���� ��������';
    Strings[5] := '�����';
    Strings[6] := '����������';
  end;
  with sgr_Source.Rows[0] do
  begin
    Strings[1] := '��������� ���';
    Strings[2] := '�������� ��������';
  end;
end;

procedure TfrmCPD.FormDestroy(Sender: TObject);
var
  r: TRegistry;
  s: string;
begin
  // ���������� ������, ���������� ����� ���������
  timCallThread.Enabled := False;
  timCallThread.Tag := 0;

  // �������� � ������� ����� ���������� ���
  r := TRegistry.Create;
  try
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);
    s := '';
    while lst_chk_Bases.Count > 0 do
    begin
      if lst_chk_Bases.Checked[0] then
        s := s + ' ' + lst_chk_Bases.Items.Strings[0];
      lst_chk_Bases.Items.Delete(0);
    end;
    s := Trim(s);
    r.WriteString('DB_Names', s);
    r.WriteInteger('ThreadsMaximum', FThreadsMaximum);
    r.WriteInteger('UpdateInterval', FUpdateInterval);
    r.WriteBool(chkVisual.Name, chkVisual.Checked);
    r.WriteBool(chkUpdate.Name, chkUpdate.Checked);
  finally
    r.Free;
  end;

  // ������� ��� ����������
  try
    FAppExcel.Visible := True;
    FAppExcel.ScreenUpdating := True;
    FAppExcel.DisplayAlerts := True;
  except
  end;
  FSheet := Null;
  FWB := Null;
  FAppExcel := Null;
  try
    if _SQLQ.Active then _SQLQ.Close;
  except
  end;
  try
    if _SQLC.Connected then _SQLC.Close;
  except
  end;

  // ������ ��������� ���������� ����������
  FFirstLine.Free;
  FlstTables.Free;

  // ���������� ����������� ��� � ����� ������
  memLog.SelectAll;
  memLog.CopyToClipboard;
end;

procedure TfrmCPD.GetParams;
const // ��������� ����������
  csg_Login = 'sa';                      // �����
  csg_SrvName = 'localhost';             // ��� �������
var
  r: TRegistry;
  s: string;
  i: Integer;
begin
  // ������ �������� �� �������
  r := TRegistry.Create;
  try
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);

    FsHostName := r.ReadString('HostName');
    if FsHostName = '' then
      FsHostName := csg_SrvName;
    txtSrv.Text := FsHostName;

    FsLogin := r.ReadString('Login');
    if FsLogin = '' then
      FsLogin := csg_Login;
    txtLogin.Text := FsLogin;

    FsPassword := r.ReadString('Password');
    txtPass.Text := FsPassword;

    s := r.ReadString('DB_Names');
    lst_chk_Bases.Items.DelimitedText := s;
    for i := 0 to lst_chk_Bases.Count - 1 do
      lst_chk_Bases.Checked[i] := True;

    FsFolderName := Trim(r.ReadString('FolderName'));
    if FsFolderName = '' then
      FsFolderName := ExtractFilePath(Application.ExeName);

    if r.ValueExists('ThreadsMaximum') then
      FThreadsMaximum := r.ReadInteger('ThreadsMaximum');
    txtThreadsMax.Text := IntToStr(FThreadsMaximum);

    if r.ValueExists('UpdateInterval') then
      FUpdateInterval := r.ReadInteger('UpdateInterval');
    txtUpdateInterval.Text := IntToStr(FUpdateInterval);

    if r.ValueExists(chkVisual.Name) then
      chkVisual.Checked := r.ReadBool(chkVisual.Name);

    if r.ValueExists(chkUpdate.Name) then
      chkUpdate.Checked := r.ReadBool(chkUpdate.Name);
  finally
    r.Free;
  end;
end;

procedure TfrmCPD.LoadColumnRelationships;
var
  ini: TMemIniFile;
  j, p, k: Integer;
  sLast, sCurr, sTemp: string;
begin
  // ��� �������� ������ �������� ���������� ����������� �������� � �������
  // � ������ ������� ���������� �������� ���������� ������������� ��������
  ini := OpenIniFile;
  if Assigned(ini) then
  try
    for j := 1 to sgr_Source.RowCount - 1 do
    begin
      sCurr := sgr_Source.Cells[0, j];
      sLast := '';
      if ini.SectionExists(sCurr) then
        sLast := ini.ReadString(sCurr, 'LastColumnName', '');
      if sLast = '' then
        sLast := sCurr;
      p := FFirstLine.IndexOf(sLast);
      if p < 0 then
        sLast := ''  // ��� ����� ������ � �������
      else
        sgr_Source.Cells[2, j] := FSheet.Cells[2, p + 1].Value;

      // ��������� ��� ��������� �������� ��� ������ �������� � ����� �����
      p := 0;
      while ini.ValueExists(sCurr, 'Column' + IntToStr(p)) do
      begin
        sTemp := ini.ReadString(sCurr, 'Column' + IntToStr(p), '');
        k := FFirstLine.IndexOf(sTemp);
        if k >= 0 then
        begin
          sLast := sTemp;
          sgr_Source.Cells[2, j] := FSheet.Cells[2, k + 1].Value;
          Break;
        end;
        Inc(p);
      end;
      sgr_Source.Cells[1, j] := sLast;
    end;
  finally
    ini.Free;
  end;
  PopupItemClick(nil);
end;

procedure TfrmCPD.MozhnoZayti(Sender: TObject);
begin
  // ���� ������������ ������ �� ���, �� � �� ������� ���
  cmdConnect.Enabled := True;
  if Trim(txtSrv.Text) = '' then
    cmdConnect.Enabled := False;
  if Trim(txtLogin.Text) = '' then
    cmdConnect.Enabled := False;
  if Trim(txtPass.Text) = '' then
    cmdConnect.Enabled := False;
end;

procedure TfrmCPD.MyConnect;
var
  s: string;
begin
  // ����������� �����������
  log := '����������� � ������� � ��';

  // ��� ������ ����� ������ ��� ������
  try
    if _SQLC.Connected then
    begin
      log := '����������� ��� �������';
      Abort;
    end;
    try
      _SQLC.LoginPrompt := False;
      s := 'Provider=SQLOLEDB.1;Password=' + FsPassword +
          ';Persist Security Info=True;User ID=' + FsLogin +
          ';Initial Catalog=' + 'master' +
          ';Data Source=' + FsHostName;
      _SQLC.ConnectionString := s;
      _SQLC.Open;  // ������������ (��������)
      _SQLQ.Connection := _SQLC;
      _SQLQ.SQL.Text := 'SELECT dtb.name ' +
          'FROM [master].[sys].[databases] AS dtb ' +
          'WHERE (CAST(case when dtb.name in ' +
          '(''master'',''model'',''msdb'',''tempdb'') ' +
          'then 1 else dtb.is_distributor end AS bit) <> 1)';
      _SQLQ.Open;  // ������ ��� �����
      if _SQLQ.RecordCount > 0 then
        while not _SQLQ.Eof do
        begin
          s := _SQLQ.FieldByName('name').AsString;
          if lst_chk_Bases.Items.IndexOf(s) < 0 then
            lst_chk_Bases.Items.Add(s);
          _SQLQ.Next;
        end;
    finally
      _SQLQ.Close;
    end;
  except
    on E: Exception do
      log := E.Message;
  end;
end;

function TfrmCPD.OpenIniFile: TMemIniFile;
var
  s: string;
begin
  // ������� ���� INI, ��� ����� ������ ������������� ���������� �����,
  // � � ������ ������ ������������ ��� �������� ������������� ����
  s := Application.ExeName;
  s := LeftStr(s, Length(s) - 4) + '.ini';
  Result := TMemIniFile.Create(s, TEncoding.UTF8);
end;

procedure TfrmCPD.PopupItemClick(Sender: TObject);
var
  pm: TMenuItem;
  jRow: Integer;
  sName, sValue: string;
begin
  // ��������� �������� ��� ��������� ������
  if Assigned(Sender) then
  begin
    pm := Sender as TMenuItem;
    jRow := sgr_Source.Selection.Top;
    sName := '';
    sValue := '';
    if pm.Tag > 0 then
    begin
      sName := ReplaceStr(pm.Caption, '&', '');
      sValue := FAppExcel.Cells[2, pm.Tag].Value;
    end;
    sgr_Source.Cells[1, jRow] := sName;
    sgr_Source.Cells[2, jRow] := sValue;
  end;

  // ������� ������� �� ������������ ��������� � ����
  for jRow := 0 to _pm.Items.Count - 1 do
  begin
    pm := _pm.Items.Items[jRow];
    sName := ReplaceStr(pm.Caption, '&', '');
    if sgr_Source.Cols[1].IndexOf(sName) >= 0 then
      pm.Checked := True
    else
      pm.Checked := False;
  end;
end;

procedure TfrmCPD.SaveColumnRelationships;
var
  ini: TMemIniFile;
  j, p: Integer;
  sLast, sCurr: string;
begin
  // ��������� �������� �� ������������ ������������ ����� � �������
  // ����� ������� ��������� ���������� ��� ��������� ��������
  ini := OpenIniFile;
  if Assigned(ini) then
  try
    for j := 1 to sgr_Source.RowCount - 1 do
    begin
      sCurr := sgr_Source.Cells[0, j];
      sLast := sgr_Source.Cells[1, j];
      ini.WriteString(sCurr, 'LastColumnName', sLast);
      p := 0;
      while ini.ValueExists(sCurr, 'Column' + IntToStr(p)) do
      begin
        if ini.ReadString(sCurr, 'Column' + IntToStr(p), '') = sLast then
        begin
          sLast := '';
          Break;
        end;
        Inc(p);
      end;
      if sLast <> '' then
        ini.WriteString(sCurr, 'Column' + IntToStr(p), sLast);
    end;
    ini.UpdateFile;
  finally
    ini.Free;
  end;
end;

procedure TfrmCPD.SetLog(const Value: string);
begin
  // ��������� � �������� � �������� �������
  memLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) + '> ' + Value);
end;

procedure TfrmCPD.SetParams;
var
  r: TRegistry;
begin
  // �������� �������� � �������
  r := TRegistry.Create;
  try
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);

    FsHostName := txtSrv.Text;
    r.WriteString('HostName', FsHostName);

    FsLogin := txtLogin.Text;
    r.WriteString('Login', FsLogin);

    FsPassword := txtPass.Text;
    r.WriteString('Password', FsPassword);

    r.WriteString('FolderName', FsFolderName);
  finally
    r.Free;
  end;
end;

procedure TfrmCPD.timCallThreadTimer(Sender: TObject);
var
  j, iRecs: Integer;
  sSNILS: string;
  b: Boolean;
begin
  // ��������� ������ � ��� ��������� ������
  if (FThreadsMaximum > 0) and (FThreadsRunning >= FThreadsMaximum) then
    Exit;  // ������� ����� �������

  timCallThread.Enabled := False;
  j := timCallThread.Tag;
  if j < 2 then
  begin
    log := '����� ��������� ������';
    txtUpdateIntervalChange(nil);
    FAppExcel.Visible := True;
    FAppExcel.ScreenUpdating := True;
    FAppExcel.DisplayAlerts := True;
    txtThreadsMax.Enabled := True;
    txtUpdateInterval.Enabled := True;
    chkVisual.Enabled := True;
    chkUpdate.Enabled := True;
    cmdConnect.Enabled := True;
    cmdGetData.Enabled := True;
    cmdConnectExcel.Enabled := True;
    Exit;
  end;

  // �������� ������ �� ������� (��� � �����)
  b := FAppExcel.ScreenUpdating;
  if chkUpdate.Checked and (FUpdateInterval > 0) and
      (j mod FUpdateInterval = 0) then
    FAppExcel.ScreenUpdating := True;
  FSheet.Cells[j, FColumns[4]].Select;
  sSNILS := FSheet.Cells[j, FColumns[4]].Value;
  sSNILS := ReplaceStr(ReplaceStr(sSNILS, '-', ''), ' ', '');
  if b <> FAppExcel.ScreenUpdating then
    FAppExcel.ScreenUpdating := b;

  for iRecs := 0 to FlstTables.Count - 1 do
  begin
    TMyThreadQuery.Create(_SQLC, FlstTables.Strings[iRecs], sSNILS,
        UpdateSheetAfterQuery, j, FColumns[5] + iRecs, memLog.Lines);
    Inc(FThreadsRunning);
  end;
  Dec(j);
  timCallThread.Tag := j;
  timCallThread.Enabled := True;
end;

procedure TfrmCPD.txtChange(Sender: TObject);
begin
  MozhnoZayti(Sender);
end;

procedure TfrmCPD.txtThreadsMaxChange(Sender: TObject);
var
  j: Integer;
begin
  j := StrToIntDef(txtThreadsMax.Text, FThreadsMaximum);
  if j < 0 then
    j := 0;
  if j > 64 then
    j := 64;
  if j <> FThreadsMaximum then
  begin
    FThreadsMaximum := j;
    txtThreadsMax.Text := IntToStr(j);
  end;
end;

procedure TfrmCPD.txtUpdateIntervalChange(Sender: TObject);
var
  j: Integer;
begin
  j := StrToIntDef(txtUpdateInterval.Text, FUpdateInterval);
  if j < 0 then
    j := 0;
  if j > MAXSHORT then
    j := MAXSHORT;
  if j <> FUpdateInterval then
  begin
    FUpdateInterval := j;
    txtUpdateInterval.Text := IntToStr(j);
  end;
end;

procedure TfrmCPD.UpdateSheetAfterQuery(Sender: TObject);
var
  thr: TMyThreadQuery;
  q: TADOQuery;
  iRecs: Integer;
  sFa1, sIm1, sOt1, sDR1, sFa2, sIm2, sOt2, sDR2: string;
  sMsg: string;
begin
  // �������� ������ � ������� (��� � �����)
  try
    thr := Sender as TMyThreadQuery;
    q := thr.oSQLQ;
    if q.Active then
      try
        if q.RecordCount <= 0 then
          FSheet.Cells[thr.iRow, thr.iCol].Value := '����� �� ������'
        else
        begin
          sMsg := '';
          for iRecs := 1 to q.RecordCount do
          begin
            q.RecNo := iRecs;

            // ������� ��������� ������ �� ���������� ��� (����� �/� �� ���������)
            sFa1 := AnsiLowerCase(FSheet.Cells[thr.iRow, FColumns[0]].Value);
            sIm1 := AnsiLowerCase(FSheet.Cells[thr.iRow, FColumns[1]].Value);
            sOt1 := AnsiLowerCase(FSheet.Cells[thr.iRow, FColumns[2]].Value);

            sFa2 := AnsiLowerCase(q.FieldByName('FAMIL').AsString);
            sIm2 := AnsiLowerCase(q.FieldByName('IMJA').AsString);
            sOt2 := AnsiLowerCase(q.FieldByName('OTCH').AsString);

            if sFa1 <> sFa2 then
              sMsg := sMsg + '; ' + '������� � ���� "' + sFa2 + '"';

            if sIm1 <> sIm2 then
              sMsg := sMsg + '; ' + '��� � ���� "' + sIm2 + '"';

            if sOt1 <> sOt2 then
              sMsg := sMsg + '; ' + '�������� � ���� "' + sOt2 + '"';

            if FColumns[3] > 0 then
            begin
              // ��� ������ ���������� ��
              sDR1 := FSheet.Cells[thr.iRow, FColumns[3]].Value;
              sDR2 := AnsiLowerCase(q.FieldByName('DROG').AsString);

              if sDR1 <> sDR2 then
                sMsg := sMsg + '; ' + '���� �������� � ���� "' + sDR2 + '"';
            end;
          end;
          if LeftStr(sMsg, 2) = '; ' then
            Delete(sMsg, 1, 2);

          FSheet.Cells[thr.iRow, thr.iCol].Value := '� ���� ' +
              ' ������ ����� ' + thr.sSNILS + '.' +
              IfThen(sMsg <> '', ' ��: ' + sMsg);
        end;
      finally
        q.Close;
      end;
  except
    on E: Exception do
      log := E.Message;
  end;
  Dec(FThreadsRunning);
end;

end.

