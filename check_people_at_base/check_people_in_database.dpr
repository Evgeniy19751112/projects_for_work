program check_people_in_database;

uses
  Vcl.Forms,
  UnitMainForm in 'UnitMainForm.pas' {frmCPD},
  UN_DiskUtils in 'UN_DiskUtils.pas',
  UnitQuery in 'UnitQuery.pas';

{$R *.res}

begin
  try // Заверну, чтобы не порождать ошибку. Т.к. нет защиты от "дурака"
      // это пока необходимость.
    Application.Initialize;
    Application.MainFormOnTaskbar := True;
    Application.CreateForm(TfrmCPD, frmCPD);
    Application.Run;
  except
  end;
end.
