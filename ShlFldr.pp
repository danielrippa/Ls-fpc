unit ShlFldr;

{$mode delphi}

interface

  uses
    ShlObj;

  function GetShellFolder2(Path: String): IShellFolder2;

implementation

  uses
    LsUtils, Windows, SysUtils;

  function GetShellFolder2;
  var
    Folder: IShellFolder;
    Folder2: IShellFolder2;
    Item: PItemIDList;

    function ParseDisplayName(aPath: String): Boolean;
    var
      Eaten, Attributes: ULONG;
      PPath: PWideChar;
    begin
      PPath := PWideChar(WideString(aPath));
      Result := Succeeded(Folder.ParseDisplayName(0, Nil, PPath, Eaten, Item, Attributes));
    end;

  begin

    if Succeeded(SHGetDesktopFolder(Folder)) then begin
      if not ParseDisplayName(Path) then begin
        if not ParseDisplayName(ExpandFileName(Path)) then begin
          Fail(Format('Path ''%s'' not found.', [Path]));
        end;
      end;
    end;

    if Succeeded(Folder.BindToObject(Item, Nil, IID_IShellFolder, Pointer(Folder))) then begin
      if Succeeded(Folder.QueryInterface(IID_IShellFolder2, Folder2)) then begin
        Result := Folder2;
        Exit;
      end;
    end;

  end;

end.