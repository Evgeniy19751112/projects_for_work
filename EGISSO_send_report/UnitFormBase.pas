unit UnitFormBase;
{
  ������ (�����) ���������� ������������ � ���� ������ � ��������� ��������

������: 2023-03-05
���������: 2023-03-05
�����: ������ �.�.
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
    FDatabase: TMyDatabase;   // ���������� �������� � ���� ������
    FThrCounter: Byte;        // ������� ������� ��� �������� ���������� ����
    function GetPatternString(var inStr: string): string;
    procedure SetLog(const Value: string);
    procedure CheckPattern;
  public
    function RunQueryAtThread: Integer;  // ��������� ������� � ������� (������� ���-��)
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

const // ��������� ���������
  csl_RegKey = '\SOFTWARE\TZN\EGISSO\send_report';

{$R *.dfm}

procedure TfrmServerDB.FormCreate(Sender: TObject);
begin
  // �������������
  FDatabase := TMyDatabase.Create(Self, frmASR.memLog, ImagesIndicator,
      imgIndicator, lst_chk_Bases);
  FThrCounter := 0;

  // ������ ������� ����������� ���������
  ImagesIndicator.GetBitmap(0, imgIndicator.Picture.Bitmap);

  // ������ ���������
  FDatabase.GetParams;
  txtSrv.Text := FDatabase.sHostName;
  txtLogin.Text := FDatabase.sLogin;
  txtPass.Text := FDatabase.sPassword;

  MozhnoZayti(nil);
  log := '����������';
end;

procedure TfrmServerDB.FormDestroy(Sender: TObject);
var
  r: TRegistry;
  s: string;
begin
  // �������� � ������� ����� ���������� ���
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

  // ������� ��� ����������
  FDatabase.Free;
end;

procedure TfrmServerDB.MozhnoZayti(Sender: TObject);
begin
  // ���� ������������ ������ �� ���, �� � �� ������� ���
  cmdConnect.Enabled := True;
  if (Trim(txtSrv.Text) = '') or
      (Trim(txtLogin.Text) = '') or
      (Trim(txtPass.Text) = '')
  then
    cmdConnect.Enabled := False;
end;

procedure TfrmServerDB.cmdSelectSQLClick(Sender: TObject);
begin
  // ������� �������� ������ �� ���� ����� ������ � ������������ ����� �������
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
  // ��������� � �������� � �������� �������
  frmASR.memLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) +
      '> TfrmServerDB: ' + Value);
end;

function TfrmServerDB.RunQueryAtThread: Integer;
var
  j: Integer;
  sFileName, sPattern, sTemp: string;
begin
  // ��������� ������ �� ���������� ������� � �� � ������������ ����������
  if not Database._SQLC.Connected then
    raise Exception.Create('������ �� �������� ��� ����������� � ����');
  frmASR.Stage := stBeforeRunThreads;
  Result := 0;
  for j := 0 to lst_chk_Bases.Count - 1 do
  begin
    if lst_chk_Bases.Checked[j] then
    begin
      // ��������� ������ ��� �����
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

      // ���� ����� ���� ����, �� ����� ���
      if FileExists(sFileName) then
        DeleteFile(sFileName);

      // ����� ����������� ����� ���������� - ��� ������ ������� ����������
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
  // ����� ���� � ����� �������� � ������� ��� ��������� ������
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
  // ������� �� ������� ������ ����� �����
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
  // ��� ��������� ���������� ������ � ��������� ������
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
  // ���������� � ��������� (���� �� ����������)
  if FDatabase._SQLC.Connected then
  begin
    FDatabase._SQLC.Close;
    Exit;
  end;

  // ��������� ������� ������������� � ����������
  FDatabase.SetConnectParams(txtSrv.Text, txtLogin.Text, txtPass.Text);
  frmASR.Stage := stWaitConnect;

  // ����������
  FDatabase.MyConnect;
end;

function TfrmServerDB.GetPatternString(var inStr: string): string;
var
  iPosL, iPosR: Integer;
  sTempL, sTempR: string;
begin
  // ��������� �������� �� ������ � �������� ����� "<" & ">"
  // �� ������� ������ �������� ���� ������� �� %s (������� ������� ������)
  // ���������� ������ ������ �� ������� ����� ������.
  // ���� ������� ���, �� ����� ������ ������ �� ����� ��������
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
