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

      if GetPropertyKeyFromCanonicalName(Column, Key) then begin

        with ColumnID do begin
          fmtid := Key.fmtid;
          pid := Key.pid;
        end;

        SetLength(Result, Length(Result) + 1);
        Result[High(Result)] := ColumnID;

      end;

    end;
  end;

  function GetItemDetails(Folder: IShellFolder2; Item: PItemIDList; ColumnID: SHCOLUMNID): String;
  var
    Value: OleVariant;
    VarTypeDesc: string;
    I: Integer;

  begin
    Result := '';
    if Succeeded(Folder.GetDetailsEx(Item, @ColumnID, @Value)) then begin
      Result := VariantAsString(Value);
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
    ColumnID: SHCOLUMNID;
    Value: OleVariant;
    StringValue: String;
  begin
    CoInitialize(Nil);
    try

      SetLength(ColumnIDs, 0);

      Folder := GetShellFolder2(Parameters.FolderPath);

      if Succeeded(Folder.EnumObjects(0, FOLDERS or NONFOLDERS or HIDDEN {* or SUPERHIDDEN *}, List)) then begin

        While(List.Next(1, Item, Fetched) = S_OK) do begin

          DisplayName := GetDisplayName(Folder, Item);

          if MatchesFileSpec(DisplayName, Parameters.FileSpec) then begin

            if Length(ColumnIDs) = 0 then begin

              for ColumnID in GetColumnIDs(Parameters.Columns) do begin
                SetLength(ColumnIDs, Length(ColumnIDs) + 1);
                ColumnIDs[High(ColumnIDs)] := ColumnID;
              end;
            end;

            for ColumnID in ColumnIDs do begin
              StringValue := GetItemDetails(Folder, Item, ColumnID);
              Write(Format('%s|', [StringValue]));
            end;

            WriteLn;

          end;

        end;

      end;

    finally
      CoUninitialize;
    end;
  end;

end.