unit UN_DiskUtils;

interface

uses
  Windows;

function DriveReady(driveletter: Char): Boolean;

// ������� ������� � �����, ���� �� �����������
function fnUN_MakeFullPath(sPath: String): Boolean;

// ������ ���� ���� � ���� ������
function FileToString(sFileName: String): String;

// �����������/����������� ������
function UN_CopyFiles(Handle: HWND; Src: string; Dest: string; Move: Boolean;
    AutoRename: Boolean): Integer;

// �������� ������
function UN_DeleteFiles(Handle: HWND; Names: string; ToRecycle: Boolean): Integer;

implementation

uses
  SysUtils, StrUtils, Classes, ShellAPI;

function DriveReady(driveletter: Char): Boolean;
var
  OldErrorMode: Word;
  OldDirectory: String;
begin
  OldErrorMode := SetErrorMode(SEM_FAILCRITICALERRORS);
  GetDir(0, OldDirectory);
  Result := False;
  try
    {$I-}
    ChDir(driveletter + ':\');
    {$I+}
    Result := IOResult = 0;
  except
  end;
  ChDir(OldDirectory);
  SetErrorMode(OldErrorMode);
end; { DriveState }

// ������� ������� � �����, ���� �� ����������� (2022-02-12 �������� ��� �������������!!!)
function fnUN_MakeFullPath(sPath: String): Boolean;
begin
  Result := ForceDirectories(sPath);
end; { fnUN_MakeFullPath }

function FileToString(sFileName: String): String;
var
  aData: array [Word] of Char;
  f: TFileStream;
  iWS: Integer;
  iLen: Integer;
  sData: String;
begin
  iWS := SizeOf(aData);
  sData := '';
  f := nil;
  Result := '';
  if not FileExists(sFileName) then Exit;
  try // ������ ���� ���� � ���� ������
    f := TFileStream.Create(sFileName, fmOpenRead or fmShareDenyWrite);
    repeat
      FillChar(aData, iWS, #0);

      // ������ ���� � �����
      iLen := f.Read(aData, iWS);
      if iLen <= 0 then Break;

      // ����������� ����� � ������
      sData := sData + StrPas(aData);
      if sData = '' then Break;
    until False;
  finally
    if Assigned(f) then
      try
        f.Free;
      except
      end;
  end;
  Result := sData;
end; { FileToString }

function UN_CopyFiles(Handle: HWND; Src: string; Dest: string; Move: Boolean;
    AutoRename: Boolean): Integer;
var
  SHFileOpStruct: TSHFileOpStruct;
begin
  with SHFileOpStruct do
    begin
      Wnd := Handle;
      wFunc := FO_COPY;
      if Move then
        wFunc := FO_MOVE;
      pFrom := PChar(Src + #0#0);
      pTo := PChar(Dest);
      fFlags := FOF_SILENT or FOF_NOCONFIRMATION;
      if AutoRename then
        fFlags := fFlags or FOF_RENAMEONCOLLISION;
      fAnyOperationsAborted := False;
      hNameMappings := nil;
      lpszProgressTitle := nil;
    end;
  Result := SHFileOperation(SHFileOpStruct);
end;

function UN_DeleteFiles(Handle: HWND; Names: string; ToRecycle: Boolean): Integer;
var
  SHFileOpStruct: TSHFileOpStruct;
begin
  with SHFileOpStruct do
    begin
      Wnd := Handle;
      wFunc := FO_DELETE;
      pFrom := PChar(Names + #0#0);
      pTo := nil;
      fFlags := FOF_SILENT or FOF_NOCONFIRMATION;
      if ToRecycle then
        fFlags := fFlags or FOF_ALLOWUNDO;
      fAnyOperationsAborted := False;
      hNameMappings := nil;
      lpszProgressTitle := nil;
    end;
  Result := SHFileOperation(SHFileOpStruct);
end;

end.
