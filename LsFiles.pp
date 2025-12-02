unit LsFiles;

{$mode delphi}

interface

  uses
    LsParams;

  procedure EnumerateFiles(Parameters: TLsParams);

implementation

  uses
    LsUtils, ShlObj, Windows, ActiveX, ShlWapi, ShlFldr, PropSys, SysUtils, Variants;

  function GetDisplayName(Folder: IShellFolder2; Item: PItemIDList): String;
  var
    StrRet: TStrRet;
    DisplayName: array[0..MAX_PATH] of Char;
  begin
    Result := '';
    if Succeeded(Folder.GetDisplayNameOf(Item, 0, StrRet)) then begin
      if Succeeded(StrRetToBufA(@StrRet, Item, DisplayName, SizeOf(DisplayName))) then begin
        Result := DisplayName;
      end;
    end;
  end;

  function GetColumnIDs(Columns: TStringArray): TSHColumnIDArray;
  var
    Column: String;
    ColumnID: SHCOLUMNID;
    Key: PROPERTYKEY;
  begin
    for Column in Columns do begin

      if Column = 'Tags' then begin
        FillChar(ColumnID, SizeOf(ColumnID), 0);
        SetLength(Result, Length(Result) + 1);
        Result[High(Result)] := ColumnID;
      end else if GetPropertyKeyFromCanonicalName(Column, Key) then begin

        with ColumnID do begin
          fmtid := Key.fmtid;
          pid := Key.pid;
        end;

        SetLength(Result, Length(Result) + 1);
        Result[High(Result)] := ColumnID;

      end;

    end;
  end;

  function GetStreamTags(FilePath: String): String;
  var
    StreamPath: String;
    F: TextFile;
    Line: String;
    Tags: String;
  begin
    Result := '';
    StreamPath := FilePath + ':keywords.txt';
    
    if not FileExists(StreamPath) then begin
      Exit;
    end;

    Tags := '';
    AssignFile(F, StreamPath);
    Reset(F);
    while not Eof(F) do begin
      ReadLn(F, Line);
      Line := Trim(Line);
      if Line <> '' then begin
        if Tags <> '' then begin
          Tags := Tags + ',';
        end;
        Tags := Tags + Line;
      end;
    end;
    CloseFile(F);
    
    Result := Tags;
  end;

  function GetItemDetails(Folder: IShellFolder2; Item: PItemIDList; ColumnID: SHCOLUMNID; ColumnName, FilePath: String): String;
  var
    Value: OleVariant;
  begin
    Result := '';
    
    if ColumnName = 'Tags' then begin
      Result := GetStreamTags(FilePath);
    end else if Succeeded(Folder.GetDetailsEx(Item, @ColumnID, @Value)) then begin
      Result := VariantAsString(Value);
      if ColumnName = 'System.PerceivedType' then begin
        Result := PerceivedTypeToString(StrToIntDef(Result, -1));
      end;
    end;
  end;

  procedure EnumerateFiles;
  var
    Folder: IShellFolder2;
    List: IEnumIDList;
    Fetched: ULONG;
    Item: PItemIDList;
    DisplayName: String;
    ColumnIDs: TSHColumnIDArray;
    StringValue: String;
    Column: String;
    ColumnIndex: Integer;
    FilePath: String;
  begin
    CoInitialize(Nil);
    try

      SetLength(ColumnIDs, 0);

      Folder := GetShellFolder2(Parameters.FolderPath);

      ColumnIDs := GetColumnIDs(Parameters.Columns);

      for Column in Parameters.Columns do begin
        Write(Format('%s|', [Column]));
      end;
      WriteLn;

      if Succeeded(Folder.EnumObjects(0, FOLDERS or NONFOLDERS or HIDDEN {* or SUPERHIDDEN *}, List)) then begin

        While(List.Next(1, Item, Fetched) = S_OK) do begin

          try
            DisplayName := GetDisplayName(Folder, Item);

            if MatchesFileSpec(DisplayName, Parameters.FileSpec) then begin

              FilePath := Parameters.FolderPath + '\' + DisplayName;

              for ColumnIndex := 0 to Length(ColumnIDs) - 1 do begin
                StringValue := GetItemDetails(Folder, Item, ColumnIDs[ColumnIndex], Parameters.Columns[ColumnIndex], FilePath);
                Write(Format('%s|', [StringValue]));
              end;

              WriteLn;

            end;

          finally
            CoTaskMemFree(Item);
          end;

        end;

      end;

    finally
      CoUninitialize;
    end;
  end;

end.