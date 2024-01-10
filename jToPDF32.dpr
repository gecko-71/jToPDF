program jToPDF32;

{$APPTYPE CONSOLE}

{$R *.res}

uses
  System.SysUtils,
  preparingreport in 'preparingreport.pas',
  u02 in 'u02.pas',
  u03 in 'u03.pas';



procedure GenRap(strIn: string; strOut: string);
var
  Preparingreport : TPreparingreport;
begin
  try
    Preparingreport := TPreparingreport.Create();
    try
      Preparingreport.reporDefName := strIn;
      Preparingreport.id_PDF_file_str := '';
      Preparingreport.pathToPdf := strOut;
      Preparingreport.FUpdateUIEvent := Preparingreport._UpdateUI;
      Preparingreport.FEndUpdateUIEvent := Preparingreport._EndUpdateUI;
      Preparingreport.Start;
    finally
      Preparingreport.free;
    end;
  except
     on  E: Exception do
     begin
       Writeln(E.Message);
     end;
   end;
end;

begin
  try
    var strIn: string := ParamStr(1);
    var strOut: string := ParamStr(2);
    if (Length(strIn) > 0) then
    begin
    if Length(strOut) = 0 then
       strOut := GetCurrentDir();
      GenRap(strIn,strOut);
    end;
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.
