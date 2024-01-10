unit u02;

interface

type
TParamValue = class
  public
    paramName: string;
    paramValueInt: Integer;
    paramValueString: String;
    paramType: Integer; // 0 = INT , 1 = String  , 2 = Float   , 3 = Boolean
    paramValueBoolean: Boolean;
    paramValueFloat: double;
    procedure SetValueInt(value: Integer);
    function GetValueInt(): Integer;
  end;

function _FloatToStr(value: double):string;
function _StrToFloat(value: string):double;

const REPORT_INT = 0;
      REPORT_STRING = 1;
      REPORT_FLOAT = 2;
      REPORT_BOOLEAN = 3;



implementation

uses System.SysUtils;



procedure TParamValue.SetValueInt(value: Integer);
begin
  paramValueInt := value;
  paramValueString := IntToStr(value);
end;

function TParamValue.GetValueInt(): Integer;
begin
  result := paramValueInt;
end;

function _FloatToStr(value: double):string;
begin
  result := FloatToStr(value);
end;

function _StrToFloat(value: string):double;
begin
  result := StrToFloat(value);
end;

end.
