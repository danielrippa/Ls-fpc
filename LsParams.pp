unit LsParams;

{$mode delphi}

interface

  uses
    SysUtils;

  type

    TLsParams = record
      FolderPath: String;
      FileSpec: String;
      Columns: TStringArray;
    end;

  function ParseCommandLineParameters: TLsParams;
  function MergeParameters(DefaultParameters, ParsedParameters: TLsParams): TLsParams;

implementation

  uses
    LsUtils, StrUtils;

  function ParseCommandLineParameters;
  var
    CommandLine: String;
    i: Integer;
    Parameters: array of string;
    Parameter: String;
  begin

    CommandLine := '';

    if ParamCount > 0 then begin

      for i := 1 to ParamCount do begin

        if i > 1 then begin
          CommandLine := CommandLine + ' ';
        end;

        CommandLine := CommandLine + ParamStr(i);
      end;

    end else begin
      Exit;
    end;

    Parameters := SplitString(CommandLine, ',');

    CommandLine := '';

    for i := 0 to Length(Parameters) - 1 do begin

      Parameter := Trim(Parameters[i]);

      case i of
        0: Result.FolderPath := Parameter;
        1: Result.FileSpec   := Parameter;
        else CommandLine := CommandLine + ' ' + Parameter;
      end;

    end;

    with Result do begin
      for Parameter in SplitString(CommandLine, ' ') do begin

        if Trim(Parameter) <> '' then begin
          SetLength(Columns, Length(Columns) + 1);
          Columns[High(Columns)] := Trim(Parameter);
        end;

      end;
    end;
  end;

  function MergeParameters;
 var
    Column: String;
  begin

    with Result do begin

      if ParsedParameters.FolderPath <> '' then begin
        FolderPath := ParsedParameters.FolderPath;
      end else begin
        FolderPath := DefaultParameters.FolderPath;
      end;

      if ParsedParameters.FileSpec <> '' then begin
        FileSpec := ParsedParameters.FileSpec;
      end else begin
        FileSpec := DefaultParameters.FileSpec;
      end;

      Columns := DefaultParameters.Columns;

      for Column in ParsedParameters.Columns do begin
        AddItemIfNotExists(Result.Columns, Column);
      end;

    end;
  end;

end.