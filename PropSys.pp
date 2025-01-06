unit PropSys;

{$mode delphi}

interface

  uses
    ShlObj;

  const
    FOLDERS =     $00000020;
    NONFOLDERS =  $00000040;
    HIDDEN =      $00000080;
    SUPERHIDDEN = $00010000;

  type

    PROPERTYKEY = packed record
      fmtid: TGUID;
      pid: DWORD;
    end;

    TSHColumnIDArray = array of SHCOLUMNID;

  function GetPropertyKeyFromCanonicalName(aCanonicalName: WideString; var aPropertyKey: PROPERTYKEY): Boolean;

implementation

  uses
    ActiveX;

  function PSGetPropertyKeyFromName(pszName: PWideChar; var pkey: PROPERTYKEY): HRESULT; stdcall; external 'propsys.dll' name 'PSGetPropertyKeyFromName';

  function GetPropertyKeyFromCanonicalName;
  begin
    Result := Succeeded(PSGetPropertyKeyFromName(PWideChar(aCanonicalName), aPropertyKey));
  end;

end.