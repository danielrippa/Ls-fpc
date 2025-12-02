unit LsUtils;

{$mode delphi}

interface

  uses
    SysUtils;

  type
    TVariantArray = array of OleVariant;

  function MatchesMask(Filename, Mask: String): Boolean;
  function MatchesFileSpec(DisplayName, FileSpec: String): Boolean;

  procedure AddItemIfNotExists(var Items: TStringArray; Item: string);

  procedure Fail(Message: String = ''; ErrorLevel: Integer = 1);

  function VariantAsString(Value: OleVariant): String;
  function PerceivedTypeToString(Value: Integer): String;

implementation

  uses
    Variants, Windows, ShlWapi;

  function PerceivedTypeToString;
  begin
    case Value of
      -2: Result := 'custom';
      -1: Result := 'unspecified';
      0: Result := 'folder';
      1: Result := 'text';
      2: Result := 'application';
      3: Result := 'image';
      4: Result := 'audio';
      5: Result := 'video';
      6: Result := 'compressed';
      7: Result := 'document';
      8: Result := 'system';
      else Result := IntToStr(Value);
    end;
  end;

  function PathMatchSpecA(pszFile, pszSpec: PAnsiChar): BOOL; stdcall; external 'shlwapi.dll' name 'PathMatchSpecA';

  function MatchesMask;
  begin
    if (Mask = '*.*') and (Pos('.', Filename) = 0) then begin
      Exit(True);
    end;

    Result := PathMatchSpecA(PAnsiChar(AnsiString(Filename)), PAnsiChar(AnsiString(Mask)));
  end;

  function MatchesFileSpec;
  begin

    DisplayName := LowerCase(DisplayName);
    FileSpec :=    LowerCase(FileSpec);

    if (Pos('*', FileSpec) > 0) or (Pos('?', FileSpec) > 0) then begin
      Result := MatchesMask(DisplayName, FileSpec);
    end else begin

      Result := (Pos(FileSpec, DisplayName) > 0) or
        (Pos(FileSpec, ExtractFileExt(DisplayName)) > 0);

    end;
  end;

  procedure Fail;
  begin
    WriteLn(StdErr, Message);
    ExitCode := ErrorLevel;
    Halt;
  end;

  function IsItemInArray(Items: TStringArray; SomeItem: String): Boolean;
  var
    Item: String;
  begin

    for Item in Items do begin
      if Item = SomeItem then begin

        Result := True;
        Exit;

      end;
    end;

    Result := False;
  end;

  procedure AddItemIfNotExists;
  begin
    if not IsItemInArray(Items, Item) then begin
      SetLength(Items, Length(Items) + 1);
      Items[High(Items)] := Item;
    end;
  end;

  function VariantAsString;
  var
    I: Integer;
  begin

    case VarType(Value) of

      varEmpty: Result := 'Empty';
      varNull: Result := 'Null';
      varSmallint: Result := IntToStr(Value);
      varInteger: Result := IntToStr(Value);
      varByte: Result := IntToStr(Value);
      varWord: Result := IntToStr(Value);
      varLongWord: Result := IntToStr(Value);
      varInt64: Result := IntToStr(Value);
      varSingle: Result := FloatToStr(Value);
      varDouble: Result := FloatToStr(Value);

      varCurrency:
        begin
          GetLocaleFormatSettings(LOCALE_USER_DEFAULT, FormatSettings);
          Result := CurrToStrF(Value, ffNumber, 15, FormatSettings);
        end;

      varDate: Result := FormatDateTime('yyyy-mm-dd hh:nn:ss', Value);
      varOleStr: Result := VarToWideStr(Value);
      varDispatch: Result := 'IDispatch';
      varError: Result := 'Error';
      varBoolean: Result := BoolToStr(Value, True);
      varVariant: Result := 'Variant';
      varUnknown: Result := 'Unknown';
      varStrArg: Result := 'String Argument';
      varString: Result := VarToStr(Value);
      varAny: Result := 'Any';
      varTypeMask: Result := 'Type Mask';

      else

        begin

          if VarIsArray(Value) then begin

            Result := '';

            try
              for I := VarArrayLowBound(Value, 1) to VarArrayHighBound(Value, 1) do begin
                if VarType(Value[I]) = varOleStr then begin
                  if VarToWideStr(Value[I]) = '' then begin
                    Result := Result + 'Empty';
                  end else begin
                    Result := Result + VarToWideStr(Value[I]);
                  end;
                  if I < VarArrayHighBound(Value, 1) then begin
                    Result := Result + ', ';
                  end;
                end else begin
                  Result := Result + 'Unsupported array element type';
                  if I < VarArrayHighBound(Value, 1) then begin
                    Result := Result + ', ';
                  end;
                end;
              end;
            except
              on E: EVariantBadIndexError do
                begin
                  Writeln('Variant array bounds error: ', E.Message);
                  Result := 'Variant array bounds error';
                end;
            end;

          end else begin
            Result := 'Unsupported array type';
          end;

        end;

    end;

  end;

end.