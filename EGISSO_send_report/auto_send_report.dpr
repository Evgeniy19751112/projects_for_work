program auto_send_report;

uses
  Vcl.Forms,
  UnitMain in 'UnitMain.pas' {frmASR},
  UnitQuery in 'UnitQuery.pas',
  UN_DiskUtils in '..\..\UniProc\UN_DiskUtils.pas',
  UnitBases in 'UnitBases.pas',
  UnitFormBase in 'UnitFormBase.pas' {frmServerDB};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'EGISSO: auto send report';
  Application.CreateForm(TfrmASR, frmASR);
  Application.CreateForm(TfrmServerDB, frmServerDB);
  Application.Run;
end.
