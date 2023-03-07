unit UnitQuery;

{
Модуль содержащий класс для выполнения запросов в отдельном потоке.
Результаты выполнения сохраняются в файле, где имя задано по шаблону
из основной формы (для каждой базы/потока можно задать свой шаблон).

Начато: 2023-02-26
Завершено: 2023-03-05
Автор: Тявкин Е.Н.

}

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
    FFileName: string;      // Имя файла, куда выводится результат
    FSQLC: TADOConnection;  // Связь с базой данных
    FLog: TStrings;         // Куда выводить сообщения об ошибках (если будут)
    // Локальные свойства
    FSQLQ: TADOQuery;
    FHasError: Boolean;
    FIndex: Integer;
    // Строки обрабатываются асинхронно, визуализация синхронно
    FAppExcel: OleVariant;    // Приложение Excel
    FWB, FSheet: OleVariant;  // Активная книга и активный лист (перед
                              // обращением активировать)
    FFirstLine: TStringList;  // Первая строка из таблицы (вместо кеша)
    iLastRow, iLastCol: Integer;
    arrColumnsSize: array of Integer;  // Массив размеров столбцов
    procedure SyncQuery;
    procedure SyncSaveToFile;
    procedure AddInfoToLog;
    procedure SyncCreateExcelFile; // Первая стадия - подготовка Excel
    procedure SyncUpdateProgress;  // Обновить прогресс на главной форме
    procedure SyncFormatTable;  // Оформить таблицу полностью
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
  // Просто сообщить о завершении потока
  FLog.Add('Поток для "' + FTable + '" завершил свою работу');
end;

procedure TMyThreadQuery.SyncCreateExcelFile;
var
  j: Integer;
  iFileFormat: Variant;
begin
  // Формируем файл в Excel'е и сохраняем его
  // Подключаемся к Excel (новый экземпляр для параллельной работы)
  FFirstLine := TStringList.Create;
  FAppExcel := CreateOleObject('Excel.Application');

  // Полностью работаем в скрытом режиме
  FAppExcel.DisplayAlerts := False;
  FAppExcel.ScreenUpdating := not frmASR.chkNoUpdateExcel.Checked;
  FAppExcel.Visible := not frmASR.chkHiddenExcel.Checked;

  // Создать новую книгу и сохранить её
  FWB := FAppExcel.Workbooks.Add;
  FLog.Add('Сохраняю в "' + FFileName + '"');
  iFileFormat := 56;
  if LowerCase(ExtractFileExt(FFileName)) = '.xlsx' then
    iFileFormat := 51
  else if LowerCase(ExtractFileExt(FFileName)) = '.xlsm' then
    iFileFormat := 52;
  try
    FWB.SaveAs(FFileName, iFileFormat);  // 56 = XLS, 51 = XLSX, 52 = XLSM
  except
    on E: Exception do
      FLog.Add('Ошибка потока для "' + FTable +
          '": Ошибка сохранения файла: ' + E.Message);
  end;

  // Связываемся с текущим листом и заполняем его
  FSheet := FWB.ActiveSheet;
  FSheet.Cells[1, 1] := 'Результат выполнения скрипта';
  iLastCol := FSQLQ.FieldCount;
  SetLength(arrColumnsSize, iLastCol);
  for j := 0 to iLastCol - 1 do
  begin
    // Создаём заголовки столбцов в таблице
    FFirstLine.Add(FSQLQ.FieldDefList.FieldDefs[j].Name);
    FSheet.Cells[2, j + 1] := FFirstLine.Strings[j];
    arrColumnsSize[j] := Length(FFirstLine.Strings[j]);
  end;
end;

procedure TMyThreadQuery.SyncFormatTable;
var
  j: Integer;
begin
  // Оформляем таблицу
  // Заголовок (строка 1)
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

  // Заголовки столбцов [строка 2]
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

  // Остальная часть таблицы
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

  // Изменить ширину столбцов
  FSheet.Cells[1, 1].Select;
  for j := 0 to iLastCol - 1 do
    FSheet.Columns[j + 1].ColumnWidth := arrColumnsSize[j] + 3;
end;

constructor TMyThreadQuery.Create(_SQLC: TADOConnection; _Table,
  _FileName: string; _endEvent: TNotifyEvent; _log: TStrings; iIndex: Integer);
begin
  // _SQLC: - подключение к базе данных
  // _Table, _FileName: - база данных и имя файла
  // _endEvent: - обработчик завершения потока
  // _log: - для вывода строк с результами
  inherited Create(True);

  // Подготовка потока
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
var
  iUpdateAfterRows: Integer;  // Через сколько строк обновляем
  j, jTemp: Integer;
  FData: Variant;  // Для записи в Excel по строкам
begin
  inherited;
  FreeOnTerminate := True;
  try
    SyncQuery;  // Сделали запрос

    // Определить интервал обновления прогреса
    iUpdateAfterRows := FSQLQ.RecordCount div frmASR.Width;
    if iUpdateAfterRows <= 0 then
      iUpdateAfterRows := 1;  // Каждый цикл при малых объёмах данных
    SyncCreateExcelFile;  // Создали файл Excel

    // Выводим строки из запроса в цикле
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
          // Для строковых значений делаем формат ячейки текстовый
          FSheet.Cells[iLastRow, j + 1].NumberFormat := '@';
        jTemp := Length(FSQLQ.FieldByName(FFirstLine.Strings[j]).AsString);
        if arrColumnsSize[j] < jTemp then
          arrColumnsSize[j] := jTemp;
      end;
      FSheet.Range[FSheet.Cells[iLastRow, 1],
          FSheet.Cells[iLastRow, iLastCol]].Value := FData;

      // Если пора, то обновить прогрес
      if FSQLQ.RecNo mod iUpdateAfterRows = 0 then
        Synchronize(SyncUpdateProgress);

      FSQLQ.Next;
    end;
    Synchronize(SyncUpdateProgress);  // Полное заполнение прогресбара

    SyncFormatTable;  // Оформляем таблицу
    SyncSaveToFile;  // Сохранить файл и закрыть Excel
  except
    on E: Exception do
    begin
      FLog.Add('Ошибка потока для "' + FTable + '": ' + E.Message);
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
  // Делаем запрос к базе данных
  with FSQLQ.SQL do
  begin
    Clear;
    Add('USE ' + FTable);
    Add('select (select top 1 DB_NAME from DATABASEINFO) as "Район",');
    Add('dt as "Дата", fileName as "Файл", fType as "Тип",');
    Add('msgId as "СМЭВ ID", packageId as "Пакет",');
    Add('case when Smevresp is null then ''Нет ответа'' else ''Да'' end as "Обработан",');
    Add('cntRecSucc as "Успешно", cntRecError as "Ошибок"');
    Add('from eg_lr_fact_smev');
  end;
  FSQLQ.Open;
end;

procedure TMyThreadQuery.SyncSaveToFile;
begin
  try
    // Включаем визуализацию полюбому
    FAppExcel.DisplayAlerts := True;
    FAppExcel.ScreenUpdating := True;
    FAppExcel.Visible := True;
  except
    on E: Exception do
      FLog.Add('Ошибка потока для "' + FTable + '": ' + E.Message);
  end;
  try
    // Попытка выйти из Excel
    FAppExcel.ActiveWorkbook.Save;
    FAppExcel.ActiveWorkbook.Close(True);
    FAppExcel.Quit;
  except
    on E: Exception do
    begin
      FLog.Add('Ошибка потока для "' + FTable + '": ' + E.Message);
      FHasError := True;
    end;
  end;
end;

procedure TMyThreadQuery.SyncUpdateProgress;
begin
  frmASR.DrawLineProgress(FIndex, FSQLQ.RecNo, FSQLQ.RecordCount);
end;

end.
