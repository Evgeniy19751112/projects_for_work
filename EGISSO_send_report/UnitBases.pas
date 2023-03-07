unit UnitBases;
{
������ ��� ������ � ����� ������ (������� � ��������� ����� �������������� �
����� ������ ���.

������: 2023-03-04
���������: 2023-03-05
�����: ������ �.�.
}

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Data.Win.ADODB, Data.DB, Vcl.Forms, Vcl.StdCtrls,
  Vcl.Controls, Vcl.ExtCtrls, Vcl.CheckLst;

type
  TMyDatabase = class(TObject)
  private
    FSQLC: TADOConnection;
    FSQLQ: TADOQuery;
    FOwnerForm: TForm;
    FmemLog: TMemo;
    FimagesIndicator: TImageList;
    FimgIndicator: TImage;
    Flst_chk_Bases: TCheckListBox;
    FsHostName: string;       // ����
    FsLogin: string;          // �����
    FsPassword: string;       // ������
    FConnectChange: Boolean;  // ������� ��������� ������/������ (��� ����������)
    procedure _SQLCAfterConnect(Sender: TObject);
    procedure _SQLCAfterDisconnect(Sender: TObject);
    procedure _SQLCBeforeConnect(Sender: TObject);
    procedure SetLog(const Value: string);
    procedure SetParams;      // ������ �������� � �������
    property log: string write SetLog;
    property _SQLQ: TADOQuery read FSQLQ;
  public
    constructor Create(AOwner: TForm; ALog: TMemo; AimagesIndicator: TImageList;
        AimgIndicator: TImage; Alst_chk_Bases: TCheckListBox);
    destructor Destroy; override;
    procedure SetConnectParams(AHost, ALogin, APass: string);
    procedure MyConnect;      // ����������� � �� (� �.�. ��������)
    procedure GetParams;      // ������ �� ������� ���������
    property lst_chk_Bases: TCheckListBox read Flst_chk_Bases;
    property _SQLC: TADOConnection read FSQLC;
    property sHostName: string read FsHostName;
    property sLogin: string read FsLogin;
    property sPassword: string read FsPassword;
  end;

implementation

uses
  System.StrUtils, System.Win.Registry, UnitFormBase;

const // ��������� ���������
  csl_RegKey = '\SOFTWARE\TZN\EGISSO\send_report';

{ TMyDatabase }

constructor TMyDatabase.Create(AOwner: TForm; ALog: TMemo;
    AimagesIndicator: TImageList; AimgIndicator: TImage;
    Alst_chk_Bases: TCheckListBox);
begin
  inherited Create;
  // ������� �������� � ������
  FsHostName := '';
  FsLogin := '';
  FsPassword := '';
  FOwnerForm := AOwner;
  FmemLog := ALog;
  FimagesIndicator := AimagesIndicator;
  FimgIndicator := AimgIndicator;
  Flst_chk_Bases := Alst_chk_Bases;
  FConnectChange := False;
  FSQLC := TADOConnection.Create(FOwnerForm);
  FSQLC.LoginPrompt := False;
  FSQLC.AfterConnect := _SQLCAfterConnect;
  FSQLC.AfterDisconnect := _SQLCAfterDisconnect;
  FSQLC.BeforeConnect := _SQLCBeforeConnect;
  FSQLQ := TADOQuery.Create(FOwnerForm);
  FSQLQ.Connection := FSQLC;
end;

destructor TMyDatabase.Destroy;
begin
  // ��������� ���e ������
  try
    if _SQLQ.Active then _SQLQ.Close;
  except
  end;
  try
    if _SQLC.Connected then _SQLC.Close;
  except
  end;

  // ������� ������� �� ������ ���
  while lst_chk_Bases.Count > 0 do
  begin
    if Assigned(lst_chk_Bases.Items.Objects[0]) then
      try
        lst_chk_Bases.Items.Objects[0].Free;
      except
      end;
    lst_chk_Bases.Items.Delete(0);
  end;
  inherited;
end;

procedure TMyDatabase.GetParams;
const // ��������� ����������
  csg_Login = 'sa';                      // �����
  csg_SrvName = 'localhost';             // ��� �������
var
  r: TRegistry;
  s: string;
  i: Integer;
  lst: TStringList;
begin
  // ������ �������� �� �������
  r := TRegistry.Create;
  try
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);

    FsHostName := r.ReadString('ServerName');
    if FsHostName = '' then
      FsHostName := csg_SrvName;

    FsLogin := r.ReadString('ServerLogin');
    if FsLogin = '' then
      FsLogin := csg_Login;

    FsPassword := r.ReadString('ServerPassword');

    s := r.ReadString('DB_Names');
    lst_chk_Bases.Items.DelimitedText := s;
    for i := 0 to lst_chk_Bases.Count - 1 do
    begin
      lst_chk_Bases.Checked[i] := True;
      if r.ValueExists(lst_chk_Bases.Items.Strings[i]) then
      begin
        lst := TStringList.Create;
        lst.DelimitedText := r.ReadString(lst_chk_Bases.Items.Strings[i]);
        lst_chk_Bases.Items.Objects[i] := lst;
      end;
    end;
  finally
    r.Free;
  end;
end;

procedure TMyDatabase.MyConnect;
var
  s, sDBName, sDBNumber: string;
  j, idBase: Integer;
  lst: TStringList;
  tempQuery: TADOQuery;
begin
  // ����������� �����������
  log := '����������� � ������� � ��';
  if (FsHostName = '') or (FsLogin = '') or (FsPassword = '') then
    raise Exception.Create('�� ������� ��������� ����������� � �������!');

  // ��� ������ ����� ������ ��� ������
  tempQuery := TADOQuery.Create(FOwnerForm);
  try
    if _SQLC.Connected then
    begin
      log := '����������� ��� �������';
      Abort;
    end;
    try
      // ������������� ����������
      s := 'Provider=SQLOLEDB.1;Password=' + FsPassword +
          ';Persist Security Info=True;User ID=' + FsLogin +
          ';Initial Catalog=' + 'master' +
          ';Data Source=' + FsHostName;
      _SQLC.ConnectionString := s;
      _SQLC.Open;  // ������������ (��������)

      // �������� ������ ���������������� ��� ������
      tempQuery.Connection := _SQLC;
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

      // �������������� ������ ��� ���, ������� �� �� ��������� ���
      if _SQLQ.Active then
        _SQLQ.Close;
      _SQLQ.SQL.Clear;
      _SQLQ.SQL.Add('');
      _SQLQ.SQL.Add('select top 1 [DB_NAME], [BASE_NUMBER] from [DATABASEINFO]');
      for j := 0 to lst_chk_Bases.Count - 1 do
      begin
        _SQLQ.SQL.Strings[0] := 'use [' + lst_chk_Bases.Items.Strings[j] + ']';
        try
          try
            _SQLQ.Open;

            // ���� ���� ������ � ������� ��� ���� ������,
            // �� ��� ��������� ������� ����� ����
            sDBName := Trim(_SQLQ.FieldByName('DB_NAME').AsString);
            sDBNumber := Trim(_SQLQ.FieldByName('BASE_NUMBER').AsString);
            if (_SQLQ.RecordCount < 1) or (sDBName = '') or (sDBNumber = '') then
              Abort;
          finally
            _SQLQ.Close;
          end;

          // ���������� ������ ��� ��������� ����, ���� ��� ��� ���
          if Assigned(lst_chk_Bases.Items.Objects[j]) then
            lst := lst_chk_Bases.Items.Objects[j] as TStringList
          else
            lst := TStringList.Create;
          if lst.Count = 0 then
            lst.Add(sDBName)
          else if lst.Strings[0] <> sDBName then
          begin
            // �������� ���-�� �� ���, �������� � ��������
            log := '�����������! ' + QuotedStr(lst.Strings[0]) +
                ' <> ' + QuotedStr(sDBName);
            lst.Strings[0] := sDBName;
          end;
          if lst.Count < 2 then
          begin
            sDBNumber := RightStr('000' + sDBNumber, 3);
            idBase := j;
            try
              // �������� ������� ������ �� ������� "[master].[dbo].[ListDb]"
              tempQuery.SQL.Text := 'SELECT TOP (1) [dbid] ' +
                    'FROM [master].[dbo].[ListDb] ' +
                    'WHERE [alias] = "' + lst_chk_Bases.Items.Strings[j] + '"';
              tempQuery.Open;
              idBase := tempQuery.FieldByName('dbid').AsInteger;
              tempQuery.Close;
            except
            end;
            lst.Add(Format('%s_%s ResultExecSQL_0_%d_<ddmmyyyyhhnnsszzz>.xls',
                [sDBNumber, sDBName, idBase]));
          end;
          lst_chk_Bases.Items.Objects[j] := lst;
        except
          lst_chk_Bases.ItemEnabled[j] := False;
        end;
      end;
    finally
      if _SQLQ.Active then
        _SQLQ.Close;
    end;
  except
    on E: Exception do
    begin
      // ������ ������� ����������� ��������� � �������� �� ������
      FimagesIndicator.GetBitmap(0, FimgIndicator.Picture.Bitmap);
      FimgIndicator.Repaint;
      log := E.Message;
    end;
  end;
  tempQuery.Free;

  if _SQLC.Connected then
  begin
    if FConnectChange then SetParams;
  end
  else
    log := '�� ���� ��������� ��������� ����������� ��� ��������� �����������';
end;

procedure TMyDatabase.SetConnectParams(AHost, ALogin, APass: string);
begin
  if FsHostName <> AHost then
  begin
    FsHostName := AHost;
    FConnectChange := True;
  end;
  if FsLogin <> ALogin then
  begin
    FsLogin := ALogin;
    FConnectChange := False;
  end;
  if FsPassword <> APass then
  begin
    FConnectChange := False;
    FsPassword := APass;
  end;
end;

procedure TMyDatabase.SetLog(const Value: string);
begin
  // ��������� � �������� � �������� �������
  FmemLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) + '> TMyDatabase: ' + Value);
end;

procedure TMyDatabase.SetParams;
var
  r: TRegistry;
begin
  // �������� �������� � �������
  r := TRegistry.Create;
  try
    r.RootKey := HKEY_CURRENT_USER;
    r.OpenKey(csl_RegKey, True);

    r.WriteString('ServerName', FsHostName);
    r.WriteString('ServerLogin', FsLogin);
    r.WriteString('ServerPassword', FsPassword);
  finally
    r.Free;
  end;
  FConnectChange := False;
end;

procedure TMyDatabase._SQLCAfterConnect(Sender: TObject);
begin
  // ����� ����������� ������������� ��������� ������ � ������ ����� ������
  FimagesIndicator.GetBitmap(6, FimgIndicator.Picture.Bitmap);
  frmServerDB.cmdConnect.Caption := '�����������';
  frmServerDB.txtSrv.Enabled := False;
  frmServerDB.txtLogin.Enabled := False;
  frmServerDB.txtPass.Enabled := False;
  FimgIndicator.Repaint;
  log := '����������� � ������� ' + FsHostName;
end;

procedure TMyDatabase._SQLCAfterDisconnect(Sender: TObject);
begin
  // ����� ������� ���������� ������������� ��������� ������� � ������ ����� ������
  FimagesIndicator.GetBitmap(2, FimgIndicator.Picture.Bitmap);
  frmServerDB.cmdConnect.Caption := '������������';
  frmServerDB.txtSrv.Enabled := True;
  frmServerDB.txtLogin.Enabled := True;
  frmServerDB.txtPass.Enabled := True;
  FimgIndicator.Repaint;
  log := '����������� � ������� ' + FsHostName + ' ���������';
end;

procedure TMyDatabase._SQLCBeforeConnect(Sender: TObject);
begin
  // ����� ���������� ���������� ������������� ��������� ������ �����
  FimagesIndicator.GetBitmap(4, FimgIndicator.Picture.Bitmap);
  FimgIndicator.Repaint;
end;

end.
