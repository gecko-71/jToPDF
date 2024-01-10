unit preparingreport;

interface

uses System.Classes, System.JSON, System.Generics.Collections, SynPdf, Winapi.Windows,
     Vcl.Graphics, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.SQLite,
  FireDAC.Phys.SQLiteDef, FireDAC.Stan.ExprFuncs,
  FireDAC.Phys.SQLiteWrapper.Stat, FireDAC.VCLUI.Wait, Data.DB,
  FireDAC.Comp.Client, FireDAC.Stan.Param, FireDAC.DatS, FireDAC.DApt.Intf,
  FireDAC.DApt, FireDAC.Comp.DataSet, u02 ;

type
  TUpdateUIEvent = procedure(mposition: Integer; fRaportName: string) of object;
  TEndUpdateUIEvent = procedure(ertxt: string; er: Boolean; fpdf_file: string;
                                fRaportName: string) of object;

  TSUMobj = class
    typeSrc: integer;
    fieldName: string;
  end;

  TobjBand = class
    typeobj: Integer;
		name: string;
    Caption: string;
    Top: Integer;
    Left: Integer;
    PenWidth: Integer;
    Width: Integer;
    Height: Integer;
    OldHeight : integer;
    Align: Integer;
    FontSize: integer;
    FontStyleBold: integer;
    FontStyleItalic: integer;
    FontStyleUnderLine: integer;
    MarginRight: integer;
    Field: string;
    queryNo: integer;
    FormatValue: string;
    AutoSizeH: integer;
    sUMobjList: TList <TSUMobj>;
    Visible: string;
    AsVisible: Boolean;
    constructor Create();
    destructor Destroy; override;
  end;

  TBand = class
    typeBand: Integer;
		name: string;
    Height: Integer;
    OldHeight: Integer;
    objBandList: TList <TobjBand>;
    constructor Create();
    destructor Destroy; override;
  end;

  TSUMValue = class
  public
    paramValuePageFloat: Currency;
    paramValuePagePrevFloat: Currency;
    paramValueAllFloat: Currency;
    field: TField;
    procedure Sum();
    procedure NewPage();
  end;

  TReportQuery = class(TFDQuery)
  public
    mainQ: Integer;
    queryNo: Integer;
    sqlTxt: string;
    ParamValues: TList <TParamValue>;
    sUMValues: TList <TSUMValue>;
    procedure Sum();
    procedure NewPage();
    procedure Next;
    procedure BuildSUMLIst();
    constructor Create();
    destructor Destroy; override;
  end;

  TReportStruct = class
  public
    raportName: string;
    RaportFileName: string;
    PageOrientation: integer;
    bandList: TList <TBand>;
    HeightAllTop: integer;
    HeightAllBottom: integer;
    HeightAllDetail: integer;
    CountDetailInPage : integer;
    CurrentPageNo : integer;
    TopCount : integer;
    BottomCount: integer;
    DetailCount: integer;
    DetailCountMax: integer;
    MaxBandBottomCount: integer;
    ScaleX: double;
    ScaleY: double;
    PageHeight : integer;
    PageWidth : integer;
    ReportQueryList: TList <TReportQuery>;
    debug: Integer;
    globalVariables : TList <TParamValue>;
    STICK_BAND_BOTTOM: Integer;
    constructor Create();
    destructor Destroy; override;
  end;

  TPageStruct = class
     RaportFileName: string;
     ReportStruct: TReportStruct;
     constructor Create();
     destructor Destroy; override;
  end;

  TPreparingReport = class
  private
    fposition: Integer;
    ferror_txt: string;
    ferror: Boolean;
    pdf_file: string;
    //ReportStruct: TReportStruct;
    PageStructList: TList <TPageStruct>;
    FDConnection: TFDConnection;

    mHWND : HWND;
    procedure UpdateUI();
    procedure EndUpdateUI();
    procedure GeneratePages();
    procedure SetGlobalVariable(ReportStruct: TReportStruct);
    procedure CopyGlobalVariableToParam(ReportStruct: TReportStruct);
    function GeneratePage(lPdf :TPdfDocumentGDI; ReportStruct: TReportStruct): boolean;
    procedure UpdateGC();
    function BuildPageDef(): Boolean;
    function BuildReportDef(ReportStruct: TReportStruct; JSONValue: TJSONValue): Boolean;
    function RednderBandTop(fCanvas: TCanvas;ReportStruct: TReportStruct): boolean;
    function RednderBandBottom(fCanvas: TCanvas;ReportStruct: TReportStruct): boolean;
    function RednderBandDetail(fCanvas: TCanvas{; var EndRap: boolean};ReportStruct: TReportStruct): boolean;
    function BandRectangleTop(fCanvas: TCanvas; band: Tband;ReportStruct: TReportStruct): Boolean;
    function BandRectangleDetail(fCanvas: TCanvas; band: Tband;ReportStruct: TReportStruct): Boolean;
    function BandRectangleBottom(fCanvas: TCanvas; band: Tband;ReportStruct: TReportStruct): Boolean;
    function CreateSQL(ReportStruct: TReportStruct):TReportQuery;
    function GetMaxDetaiCount( ReportStruct: TReportStruct): integer;
    function GetValueParam( ReportStruct: TReportStruct;fName: string): Variant;
    procedure RenderOBJ(ftop:integer;fCanvas: TCanvas; fobjBand: TobjBand;ReportStruct: TReportStruct);
    function GetValueFromQuery(fobjBand: TobjBand;ReportStruct: TReportStruct): string;
    procedure DrawVert(fCanvas: TCanvas; Box: TRect; const Text: string);
    function  ParsujStr(Caption : string; globalVariables : TList <TParamValue>): string;
    function SetHDetailBanf(fCanvas: TCanvas; band:Tband; ReportStruct:TReportStruct): integer;
    function GetGlobalValueInt(globalVariables : TList <TParamValue>; fname: string;
                           var ParamValue: TParamValue; var return: Boolean): Integer;
    function SetGlobalValueInt(globalVariables : TList <TParamValue>; fname: string; value: Integer): Boolean;
    function GetSumValue(fobjBand: TobjBand;ReportStruct: TReportStruct): string; //sUMobjList
    procedure AddUpdateParamRapVarBoolean(globalVariables : TList <TParamValue>;fparamName: string; fparamValue: Boolean);
    function GetMaxBandBottomCount( ReportStruct: TReportStruct): integer;
  public
    reporDefName: string;
    pathToPdf: string;
    DriverName: string;
    Database: string;
    FUpdateUIEvent: TUpdateUIEvent;
    FEndUpdateUIEvent: TEndUpdateUIEvent;
    ParamValues: TList <TParamValue>;
    id_PDF_file_str: string;
    procedure _UpdateUI(mposition: Integer;  fRaportName: string);
    procedure _EndUpdateUI(ertxt: string;er: boolean; fpdf_file: string;
                             fRaportName: string);
    constructor Create();
    destructor Destroy; override;
    procedure Start;
    procedure AddParamSQLINT(fparamName: string; fparamValue: Integer);
  end;

implementation

uses Vcl.Forms, System.SysUtils, System.IOUtils, System.StrUtils,  System.Variants, u03 ;

const REPORT_BAND_TOP = 0;
      PAGE_BAND_TOP = 1;
      DATAIL_BAND = 2;
      REPORT_BAND_BOTTOM = 3;
      PAGE_BAND_BOTTOM = 4;

      TYPE_TEXT_OBJ = 0;
      TYPE_FIELD_OBJ = 1;
      TYPE_BOX_OBJ = 2;
      TYPE_FIELD_BOX_OBJ = 3;
      TYPE_FIELD_SUM_BOX_OBJ = 4 ;

      CENTER_TEXT_OBJ = 1;
      LEFT_TEXT_OBJ = 0;
      RIGHT_TEXT_OBJ = 2;
      MAIN_SQL = 1;

procedure TSUMValue.NewPage();
begin
   paramValuePagePrevFloat := paramValuePageFloat;
   paramValuePageFloat := 0;
end;

procedure TSUMValue.Sum();
begin
  if not Field.IsNull then
  begin
    paramValueAllFloat := paramValueAllFloat + Field.Value;
    paramValuePageFloat := paramValuePageFloat +  + Field.Value;
  end;
end;

constructor TobjBand.Create();
begin
  sUMobjList := TList <TSUMobj>.Create();
  AsVisible := True;
end;

destructor TobjBand.Destroy;
var
  i: Integer;
begin
  for i := 0 to sUMobjList.Count -1 do
    sUMobjList[i].Free;
  sUMobjList.Free;
  inherited;
end;

destructor TPageStruct.Destroy();
var
  i: Integer;
begin
  ReportStruct.Free;
  inherited;
end;

constructor TPageStruct.Create();
begin
  inherited;
  ReportStruct := TReportStruct.Create();
end;

constructor TBand.Create();
begin
  inherited;
  objBandList := TList <TobjBand>.Create();
end;

destructor TBand.Destroy();
var
  i: Integer;
begin
  for i := 0 to objBandList.Count -1 do
  begin
     objBandList[i].Free;
  end;
  objBandList.Free;
  inherited;
end;

constructor TReportQuery.Create( );
begin
  inherited Create (nil);
  ParamValues := TList <TParamValue>.Create();
  sUMValues := TList <TSUMValue>.Create();
end;

destructor TReportQuery.Destroy();
var
  i: Integer;
begin
  for i := 0 to ParamValues.Count -1 do
  begin
     ParamValues[i].Free;
  end;
  ParamValues.Free;
  for i := 0 to sUMValues.Count -1 do
  begin
     sUMValues[i].Free;
  end;
  sUMValues.Free;
  inherited;
end;

procedure TReportQuery.BuildSUMLIst();
var
  i: integer;
  sumv: TsUMValue;
begin
  for i := 0 to Fields.Count -1 do
  begin
    if (Fields[i].DataType = ftFloat) or (Fields[i].DataType = ftLargeint)  or
        (Fields[i].DataType = ftInteger)  or
         (Fields[i].DataType = ftCurrency) then
    begin
      sumv := TsUMValue.Create();
      sumv.field := Fields[i];
      sumv.paramValuePageFloat := 0;
      sumv.paramValuePagePrevFloat := 0;
      sumv.paramValueAllFloat := 0;
      sUMValues.Add(sumv);
    end;
  end;
end;


procedure TReportQuery.Sum();
var
  i: integer;
begin
  for i := 0 to sUMValues.Count -1 do
    sUMValues[i].Sum();
end;

procedure TReportQuery.NewPage();
var
  i: integer;
begin
  for i := 0 to sUMValues.Count -1 do
    sUMValues[i].NewPage();
end;


procedure TReportQuery.Next;
begin
  inherited;
end;

constructor TReportStruct.Create();
begin
  inherited;
  bandList := TList <TBand>.Create();
  ReportQueryList := TList <TReportQuery>.Create();
  globalVariables := TList <TParamValue>.Create();
end;

destructor TReportStruct.Destroy();
var
  i: Integer;
begin
  for i := 0 to globalVariables.Count -1 do
  begin
     globalVariables[i].Free;
  end;
  globalVariables.Free;

  for i := 0 to bandList.Count -1 do
  begin
     bandList[i].Free;
  end;
  bandList.Free;

  for i := 0 to ReportQueryList.Count -1 do
  begin
     ReportQueryList[i].Free;
  end;
  ReportQueryList.Free;
  inherited;
end;

procedure TPreparingreport._UpdateUI(mposition: Integer;  fRaportName: string);
begin

end;

procedure TPreparingreport._EndUpdateUI(ertxt: string;er: boolean; fpdf_file: string;
                             fRaportName: string);
begin
  Writeln(ertxt + #13#10);
  Writeln(fpdf_file + #13#10);
  Writeln(fRaportName + #13#10);
end;

constructor TPreparingreport.Create();
begin
  ferror_txt := '';
  ferror := false;
  ParamValues := TList <TParamValue>.Create();
end;

destructor TPreparingreport.Destroy();
var
  i: Integer;
begin
  for i := 0 to ParamValues.Count -1 do
  begin
     ParamValues[i].Free;
  end;
  ParamValues.Free;
  inherited;
end;

procedure TPreparingreport.AddParamSQLINT(fparamName: string; fparamValue: Integer);
var
  ParamValue: TParamValue;
begin
  ParamValue := TParamValue.Create();
  ParamValue.paramName := fparamName;
  ParamValue.paramValueInt := fparamValue;
  ParamValue.paramType := REPORT_INT;
  ParamValues.Add(ParamValue);
end;

procedure TPreparingreport.AddUpdateParamRapVarBoolean(globalVariables : TList <TParamValue>;fparamName: string; fparamValue: Boolean);
var
  ParamValue: TParamValue;
  ii: integer;
begin
  for ii := 0 to globalVariables.Count - 1 do
  begin
     if globalVariables[ii].paramName = fparamName then
     begin
       globalVariables[ii].paramType := REPORT_BOOLEAN;
       globalVariables[ii].paramValueBoolean := fparamValue;
       exit;
     end;
  end;
  ParamValue := TParamValue.Create();
  ParamValue.paramName := fparamName;
  ParamValue.paramValueBoolean := fparamValue;
  ParamValue.paramType := REPORT_BOOLEAN;
  globalVariables.Add(ParamValue);
end;

procedure TPreparingreport.Start;
var
  i: Integer;
begin

  try
    PageStructList := TList <TPageStruct>.Create;
    try
      if BuildPageDef() then
      begin
        EndUpdateUI();
        exit;
      end;
      UpdateUI();
      GeneratePages();
      EndUpdateUI();
      for i :=0 to PageStructList.Count -1 do
          PageStructList[i].Free;
    finally
      PageStructList.free;
    end;
  except
    on E: Exception do
    begin
      ferror_txt := E.Message;
      ferror := True;
    end;
  end;
end;

procedure TPreparingreport.UpdateGC();
begin
  mHWND := GetDC(0);
end;

function TPreparingreport.BuildReportDef(ReportStruct: TReportStruct; JSONValue: TJSONValue): Boolean;
var
  pathToRepDef, inputstr: string;
  JV: TJSONValue;
  inputfile : TStringList;
  bandArray: TJSONArray;
  bandElementList : TJSONValue;
  bandElement : TJSONValue;
  band: TBand ;
  objArray: TJSONArray;
  objArrayList : TJSONValue;
  objArrayElement : TJSONValue;
  objBand :TobjBand;

  fontArray: TJSONArray;
  fontArrayList : TJSONValue;
  fontArrayElement : TJSONValue;
  i: integer;
  paramValues : TParamValue;
  paramArray: TJSONArray;
  paramElementList : TJSONValue;
  paramElement : TJSONValue;
  sQLparamValues : TParamValue;
  sQLparamArray: TJSONArray;
  sQLparamElementList : TJSONValue;
  sQLparamElement : TJSONValue;
  ReportQuery: TReportQuery;
  globalVarValues : TParamValue;
  globalVarArray: TJSONArray;
  globalVarElementList : TJSONValue;
  globalVarElement : TJSONValue;

  //sumVarValues : TParamValue;
  sumVarArray: TJSONArray;
  sumVarElementList : TJSONValue;
  sumVarElement : TJSONValue;

  sQLtxtArray: TJSONArray;
  ii: integer;
  ParamValueG : TParamValue;
  sUMobj: TSUMobj;
begin
  result:= false;
    try
        JV := (JsonValue as TJSONObject).Get('RaportName').JSONValue;
        ReportStruct.RaportName := JV.Value;
    except
      on E: Exception do
      begin
        ferror_txt := E.Message;
        ferror := True;
        result := True;
        exit;
      end;
    end;
    try
      JV := (JsonValue as TJSONObject).Get('PageOrientation').JSONValue;
      ReportStruct.PageOrientation := JV.GetValue<Integer>;
    except
      ReportStruct.PageOrientation := 0;
    end;
    try
      JV := (JsonValue as TJSONObject).Get('RaportFileName').JSONValue;
      ReportStruct.RaportFileName := JV.Value;
    except
      ReportStruct.RaportFileName := '';
    end;
    try
      JV := (JsonValue as TJSONObject).Get('PageHeight').JSONValue;
      ReportStruct.PageHeight := JV.GetValue<Integer>;
    except
      ReportStruct.PageHeight := 0;
    end;
    try
      JV := (JsonValue as TJSONObject).Get('PageWidth').JSONValue;
      ReportStruct.PageWidth := JV.GetValue<Integer>;
    except
      ReportStruct.PageWidth := 0;
    end;
    try
      JV := (JsonValue as TJSONObject).Get('debug').JSONValue;
      ReportStruct.debug := JV.GetValue<Integer>;
    except
      ReportStruct.debug := 0;
    end;

    try
      JV := (JsonValue as TJSONObject).Get('STICK_BAND_BOTTOM').JSONValue;
      ReportStruct.STICK_BAND_BOTTOM := JV.GetValue<Integer>;
    except
      ReportStruct.STICK_BAND_BOTTOM := 0;
    end;

    try
          if JsonValue.TryGetValue('globalVariables',  globalVarArray)  then
          begin
            globalVarArray := (JsonValue as TJSONObject).Get('globalVariables').JSONValue as TJSONArray;
            for globalVarElementList in globalVarArray do
            begin
              ParamValueG := TParamValue.Create();
              globalVarElement := (globalVarElementList as TJSONObject).Get('paramName').JSONValue;
              ParamValueG.paramName := globalVarElement.Value;
              globalVarElement := (globalVarElementList as TJSONObject).Get('paramType').JSONValue;
              ParamValueG.paramType := globalVarElement.GetValue<Integer>;
              globalVarElement := (globalVarElementList as TJSONObject).Get('paramValue').JSONValue;
              if ParamValueG.paramType = REPORT_STRING then
                ParamValueG.paramValueString := globalVarElement.Value
              else if ParamValueG.paramType = REPORT_INT then
                ParamValueG.paramValueInt := StrToInt(globalVarElement.Value);
              ReportStruct.globalVariables.Add(ParamValueG);
            end;
          end;
    except
      on E: Exception do
      begin
        ferror_txt := E.Message;
        ferror := True;
        result := True;
      end;
    end;

    try
      if JsonValue.TryGetValue('SQL',  sQLparamArray)  then
      begin
        sQLparamArray := (JsonValue as TJSONObject).Get('SQL').JSONValue as TJSONArray;
        for sQLparamElementList in sQLparamArray do
        begin
          ReportQuery := TReportQuery.Create();

          ReportQuery.sqlTxt := '';
          sQLtxtArray := (sQLparamElementList as TJSONObject).Get('sqlTxt').JSONValue as TJSONArray;
          if sQLtxtArray <> nil then
          begin
            for ii := 0 to sQLtxtArray.Count-1 do
            begin
              ReportQuery.sqlTxt := ReportQuery.sqlTxt + sQLtxtArray[ii].Value;
            end;
          end;


          sQLparamElement := (sQLparamElementList as TJSONObject).Get('queryNo').JSONValue;
          ReportQuery.queryNo := sQLparamElement.GetValue<Integer>;
          sQLparamElement := (sQLparamElementList as TJSONObject).Get('mainQ').JSONValue;
          ReportQuery.mainQ := sQLparamElement.GetValue<Integer>;
          if sQLparamElementList.TryGetValue('sqlParam',  paramArray)  then
          begin
            paramArray := (sQLparamElementList as TJSONObject).Get('sqlParam').JSONValue as TJSONArray;
            for paramElementList in paramArray do
            begin
              paramValues := TParamValue.Create();
              paramElement := (paramElementList as TJSONObject).Get('paramName').JSONValue;
              paramValues.paramName := paramElement.Value;
              paramElement := (paramElementList as TJSONObject).Get('paramType').JSONValue;
              paramValues.paramType := paramElement.GetValue<Integer>;
              ReportQuery.ParamValues.Add(paramValues);
            end;
          end;
          ReportStruct.ReportQueryList.Add(ReportQuery);
        end;
      end;
    except
      on E: Exception do
      begin
        ferror_txt := E.Message;
        ferror := True;
        result := True;
      end;
    end;
    //sqlParam
   //
    try
      if JsonValue.TryGetValue('band',  bandArray)  then
      begin
        bandArray := (JsonValue as TJSONObject).Get('band').JSONValue as TJSONArray;
        for bandElementList in bandArray do
        begin
          bandElement := (bandElementList as TJSONObject).Get('typeBand').JSONValue;
          band := TBand.Create();
          band.typeBand := bandElement.GetValue<Integer>;
          bandElement := (bandElementList as TJSONObject).Get('name').JSONValue;
          band.name := bandElement.Value;
          bandElement := (bandElementList as TJSONObject).Get('Height').JSONValue;
          band.Height := bandElement.GetValue<Integer>;
          band.OldHeight := band.Height;

          objArray := (bandElementList as TJSONObject).Get('objBand').JSONValue as TJSONArray;
          for objArrayList in objArray do
          begin
            objArrayElement := (objArrayList as TJSONObject).Get('typeobj').JSONValue;
            objBand := TobjBand.Create;
            objBand.typeobj := objArrayElement.GetValue<Integer>;

            if objArrayList.TryGetValue('sum',  sumVarArray)  then
            begin
              sumVarArray := ( objArrayList as TJSONObject).Get('sum').JSONValue as TJSONArray;
              for sumVarElementList in sumVarArray do
              begin
                sUMobj := TSUMobj.Create;
                if sumVarElementList.TryGetValue('typeSrc', sUMobj.typeSrc)  then
                begin
                    sumVarElement := (sumVarElementList as TJSONObject).Get('typeSrc').JSONValue;
                    sUMobj.typeSrc := sumVarElement.GetValue<Integer>;
                end;
                if sumVarElementList.TryGetValue('field', sUMobj.fieldName)  then
                begin
                    sumVarElement := (sumVarElementList as TJSONObject).Get('field').JSONValue;
                    sUMobj.fieldName := sumVarElement.Value;
                end;
                objBand.sUMobjList.Add(sUMobj) ;
              end;
            end;

            if objArrayList.TryGetValue('AutoSizeH', objBand.Align)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('AutoSizeH').JSONValue;
              objBand.AutoSizeH := objArrayElement.GetValue<Integer>;
            end;

            if objArrayList.TryGetValue('MarginRight', objBand.MarginRight)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('MarginRight').JSONValue;
              objBand.MarginRight := objArrayElement.GetValue<Integer>;
            end;

            if objArrayList.TryGetValue('FormatValue', objBand.FormatValue)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('FormatValue').JSONValue;
              objBand.FormatValue := objArrayElement.Value;
            end else
              objBand.FormatValue := '';

            objArrayElement := (objArrayList as TJSONObject).Get('Name').JSONValue;
            objBand.name := objArrayElement.Value;
            if objArrayList.TryGetValue('Caption', objBand.Caption)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('Caption').JSONValue;
              objBand.Caption := objArrayElement.Value;
            end;

            objArrayElement := (objArrayList as TJSONObject).Get('Top').JSONValue;
            objBand.Top := objArrayElement.GetValue<Integer>;
            if objArrayList.TryGetValue('Align', objBand.Align)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('Align').JSONValue;
              objBand.Align := objArrayElement.GetValue<Integer>;
            end;
            if objArrayList.TryGetValue('Field', objBand.Field)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('Field').JSONValue;
              objBand.Field :=  objArrayElement.Value;
            end;

            if objArrayList.TryGetValue('Left', objBand.Left)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('Left').JSONValue;
              objBand.Left := objArrayElement.GetValue<Integer>;
            end;
            if objArrayList.TryGetValue('PenWidth', objBand.PenWidth)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('PenWidth').JSONValue;
              objBand.PenWidth := objArrayElement.GetValue<Integer>;
            end;
            if objArrayList.TryGetValue('Width', objBand.Width)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('Width').JSONValue;
              objBand.Width := objArrayElement.GetValue<Integer>;
            end;
            if objArrayList.TryGetValue('Height', objBand.Height)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('Height').JSONValue;
              objBand.Height := objArrayElement.GetValue<Integer>;
            end;

            if objArrayList.TryGetValue('queryNo', objBand.queryNo)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('queryNo').JSONValue;
              objBand.queryNo := objArrayElement.GetValue<Integer>;
            end else objBand.queryNo := -1;

            if objArrayList.TryGetValue('Visible', objBand.Visible)  then
            begin
              objArrayElement := (objArrayList as TJSONObject).Get('Visible').JSONValue;
              objBand.Visible := objArrayElement.Value;
              if objBand.Visible = '' then
              begin
                objBand.Visible := 'T.';
                objBand.AsVisible := True;
              end else if objBand.Visible = 'T.' then
              begin
                objBand.AsVisible := True;
              end else if objBand.Visible = 'F.' then
              begin
                objBand.AsVisible := False;
              end else
              begin
                objBand.AsVisible := False;
              end;
            end else
            begin
              objBand.Visible := '';
              objBand.AsVisible := True;
            end;

            if objArrayList.TryGetValue('font',  fontArray)  then
            begin
              fontArray := (objArrayList as TJSONObject).Get('font').JSONValue as TJSONArray;
              for fontArrayList in fontArray do
              begin
                fontArrayElement := (fontArrayList as TJSONObject).Get('size').JSONValue;
                objBand.FontSize := fontArrayElement.GetValue<Integer>;
                fontArrayElement := (fontArrayList as TJSONObject).Get('bold').JSONValue;
                objBand.FontStyleBold := fontArrayElement.GetValue<Integer>;
                fontArrayElement := (fontArrayList as TJSONObject).Get('italic').JSONValue;
                objBand.FontStyleItalic := fontArrayElement.GetValue<Integer>;
                fontArrayElement := (fontArrayList as TJSONObject).Get('UnderLine').JSONValue;
                objBand.FontStyleUnderLine := fontArrayElement.GetValue<Integer>;
              end;
            end;
            band.objBandList.Add(objBand);
          end;

          ReportStruct.bandList.Add(band);
        end;
      end;
      ReportStruct.HeightAllTop := 0;
      ReportStruct.HeightAllBottom := 0;
      for i := 0 to ReportStruct.bandList.Count -1 do
      begin
        if (ReportStruct.bandList[i].typeBand = REPORT_BAND_TOP) or
           (ReportStruct.bandList[i].typeBand = PAGE_BAND_TOP) then
          ReportStruct.HeightAllTop := ReportStruct.bandList[i].Height +ReportStruct.HeightAllTop;
        if (ReportStruct.bandList[i].typeBand = REPORT_BAND_BOTTOM) or
           (ReportStruct.bandList[i].typeBand = PAGE_BAND_BOTTOM) then
          ReportStruct.HeightAllBottom := ReportStruct.bandList[i].Height + ReportStruct.HeightAllBottom;
        if (ReportStruct.bandList[i].typeBand = DATAIL_BAND) then
          ReportStruct.HeightAllDetail := ReportStruct.bandList[i].Height + ReportStruct.HeightAllDetail;
      end;

    except
      on E: Exception do
      begin
        ferror_txt := E.Message;
        ferror := True;
        result := True;
      end;
    end;
end;

function TPreparingreport.BuildPageDef(): Boolean;
var
  pathToRepDef, inputstr: string;
  JSONValue: TJSONValue;
  inputfile : TStringList;

  pageValues : TParamValue;
  pageArray: TJSONArray;
  pageElementList : TJSONValue;
  pageElement : TJSONValue;
  PageStruct: TPageStruct;
  globalVarValues : TParamValue;
  globalVarArray: TJSONArray;
  globalVarElementList : TJSONValue;
  globalVarElement : TJSONValue;

  ParamValue : TParamValue ;
  JsonRapObj: TJSONValue;
begin
  result:= false;
  JsonValue := nil;
  try
    pathToRepDef:= reporDefName ;
    inputfile := TStringList.Create;
    try
      inputfile.LoadFromFile(pathToRepDef, TEncoding.UTF8);
      inputstr := inputfile.Text;
    finally
      inputfile.Free;
    end;
    try
      JSONValue := TJSONObject.ParseJSONValue(inputstr);
    except
      on E: Exception do
      begin
        ferror_txt := E.Message;
        ferror := True;
      end;
    end;
    //PageStructList

    if JsonValue.TryGetValue('Page',  PageArray)  then
    begin
        PageArray := (JsonValue as TJSONObject).Get('Page').JSONValue as TJSONArray;
        for pageElementList in pageArray do
        begin
          PageStruct := TPageStruct.Create();
          PageStructList.Add(PageStruct);

          pageElement := (pageElementList as TJSONObject).Get('RaportFileName').JSONValue;
          PageStruct.RaportFileName := pageElement.Value;
          JsonRapObj := (pageElementList as TJSONObject).Get('RapDef').JSONValue;
          if BuildReportDef(PageStruct.ReportStruct, JsonRapObj)   then
            exit;
        end;
    end else
    begin
      ferror_txt := 'empty page definition';
      ferror := True;
      result := True;
    end;
  except
    on E: Exception do
    begin
      ferror_txt := E.Message;
      ferror := True;
      result := True;
    end;
  end;
end;

procedure TPreparingreport.CopyGlobalVariableToParam(ReportStruct: TReportStruct);
var
 i,ii: integer;
 ParamValue: TParamValue;
begin
  for i := 0 to ParamValues.Count - 1 do
  begin
     for ii := 0 to ReportStruct.globalVariables.Count - 1 do
     begin
       if ReportStruct.globalVariables[ii].paramName = ParamValues[i].paramName then
       begin
         if ReportStruct.globalVariables[ii].paramType = ParamValues[i].paramType then
         begin
           if ReportStruct.globalVariables[ii].paramType = REPORT_STRING then
              ParamValues[i].paramValueString := ReportStruct.globalVariables[ii].paramValueString
           else if ReportStruct.globalVariables[ii].paramType = REPORT_INT then
           begin
             ParamValues[i].paramValueInt := ReportStruct.globalVariables[ii].paramValueInt;
             ParamValues[i].paramValueInt := StrToInt(ReportStruct.globalVariables[ii].paramValueString);
           end;
         end;
       end;
     end;
  end;

end;

procedure TPreparingreport.SetGlobalVariable(ReportStruct: TReportStruct);
var
 i,ii: integer;
 ParamValue: TParamValue;
 yes: boolean;
begin
  for i := 0 to ParamValues.Count - 1 do
  begin
     yes := false;
     for ii := 0 to ReportStruct.globalVariables.Count - 1 do
     begin
       if ReportStruct.globalVariables[ii].paramName = ParamValues[i].paramName then
       begin
         yes := True;
         if ReportStruct.globalVariables[ii].paramType = ParamValues[i].paramType then
         begin
           if ReportStruct.globalVariables[ii].paramType = REPORT_STRING then
                ReportStruct.globalVariables[ii].paramValueString := ParamValues[i].paramValueString
           else if ReportStruct.globalVariables[ii].paramType = REPORT_INT then
           begin
             ReportStruct.globalVariables[ii].paramValueInt := ParamValues[i].paramValueInt;
             ReportStruct.globalVariables[ii].paramValueString := IntToStr(ParamValues[i].paramValueInt);
           end;
         end;
         break;
       end;
     end;
     if Yes = false then
     begin
       ParamValue := TParamValue.Create;
       ParamValue.paramValueString := ParamValues[i].paramValueString;
       if ParamValues[i].paramtype = REPORT_INT then
         ParamValue.paramValueString := IntToStr(ParamValues[i].paramValueInt);
       ParamValue.paramValueInt := ParamValues[i].paramValueInt;
       ParamValue.paramName := ParamValues[i].paramName;
       ParamValue.paramtype := ParamValues[i].paramtype;
       ReportStruct.globalVariables.Add(ParamValue);
     end;
  end;
end;

function TPreparingreport.GeneratePage(lPdf :TPdfDocumentGDI; ReportStruct: TReportStruct): boolean;
var
  endRap: boolean;
  //Metafile : TMetafile;
  //mCanvas : TCanvas;
  mCanvas : TCanvas;
  i: Integer;
  bmp: TBitmap;
  ParamValue: TParamValue;
  return: Boolean;
  mainSQL :TReportQuery;
begin
  result := false;
  endRap := false;
  mainSQL := nil;
  try
    //if ReportStruct.globalVariables.Count > 0 then
    if Assigned( ReportStruct.globalVariables) then
    begin
      AddUpdateParamRapVarBoolean(ReportStruct.globalVariables, 'ENDRAP' , false);
      SetGlobalVariable(ReportStruct);
    end;
    if ReportStruct.ReportQueryList.Count > 0 then
       mainSQL := CreateSQL(ReportStruct);
    try
      if ReportStruct.PageOrientation = 1 then
         lPdf.DefaultPageLandscape := true
      else
         lPdf.DefaultPageLandscape := false;

      if ReportStruct.HeightAllDetail > 0 then
         ReportStruct.CountDetailInPage := trunc((Integer(lPDF.DefaultPageHeight) - ReportStruct.HeightAllTop
                                           - ReportStruct.HeightAllBottom) / ReportStruct.HeightAllDetail);
      ReportStruct.CurrentPageNo := 0;
      GetGlobalValueInt(ReportStruct.globalVariables,'PAGE_NO',ParamValue, return);
      try
        //Metafile := TMetafile.Create;
          //Metafile.SetSize(ReportStruct.PageWidth, ReportStruct.PageHeight);
          //mCanvas := TCanvas.Create(Metafile, 0);
            //mCanvas.Lock;
            while True do
            begin
              //endRap := true;
              AddUpdateParamRapVarBoolean(ReportStruct.globalVariables, 'ENDRAP' , True);
              lPDF.AddPage;
              if Assigned(mainSQL) then
                 mainSQL.NewPage();
              if Assigned(ParamValue) then
              begin
                // add page no
                ParamValue.SetValueInt(ParamValue.GetValueInt() + 1);
              end;

              bmp := TBitmap.Create;
              try
                bmp.Width := ReportStruct.PageWidth;
                bmp.Height := ReportStruct.PageHeight;
                mCanvas := bmp.Canvas;
                ReportStruct.ScaleX :=  lPdf.VCLCanvasSize.Width / ReportStruct.PageWidth ;
                ReportStruct.ScaleY :=  lPdf.VCLCanvasSize.Height /ReportStruct.PageHeight ;
                ReportStruct.CurrentPageNo := ReportStruct.CurrentPageNo + 1;
                ReportStruct.TopCount := 0;
                ReportStruct.BottomCount := ReportStruct.PageHeight+1;
                ReportStruct.DetailCountMax :=  GetMaxDetaiCount(ReportStruct);
                ReportStruct.MaxBandBottomCount :=  GetMaxBandBottomCount(ReportStruct);

                if RednderBandTop(mCanvas, ReportStruct) then
                   break;
                if RednderBandDetail(mCanvas,{ endRap,}ReportStruct) then
                   break;
                if RednderBandBottom(mCanvas, ReportStruct) then
                   break;
                endRap := GetValueParam(ReportStruct, 'ENDRAP');
                if endRap then
                   break;
              finally
                //mCanvas.UnLock;
                lPDF.VCLCanvas.CopyRect(Rect(0,0,lPdf.VCLCanvasSize.Width,lPdf.VCLCanvasSize.Height) ,bmp.Canvas,Rect(0,0,bmp.Width,bmp.Height));
                //mCanvas.Free;
                //lPDF.Canvas.RenderMetaFile(Metafile, ReportStruct.ScaleX, ReportStruct.ScaleY);
                bmp.Free;
              end;
            end;
      finally
          //Metafile.Free;
      end;
    finally
      if ReportStruct.ReportQueryList.Count > 0 then
        for i:= 0 to ReportStruct.ReportQueryList.Count -1 do
          ReportStruct.ReportQueryList[i].Close;
    end;
    CopyGlobalVariableToParam(ReportStruct);
  except
    on E: Exception do
    begin
      ferror_txt := E.Message;
      ferror := True;
      result := True;
    end;
  end;

end;

procedure TPreparingreport.GeneratePages();
var
  lPdf   : TPdfDocumentGDI;
  //endRap: boolean;
  //Metafile : TMetafile;
  //mCanvas : TCanvas;
  mCanvas : TCanvas;
  i: Integer;
  bmp: TBitmap;
  pageIndx: Integer;
begin
  //pathToPdf := ExtractFilePath(Application.ExeName) + 'ReportFileOut';
  TDirectory.CreateDirectory(pathToPdf);
  if not DirectoryExists(pathToPdf) then
  begin
    ferror_txt := 'Brak katalogu ' + pathToPdf;
    ferror := True;
    exit;
  end;
  if PageStructList.Count = 0 then
  begin
    ferror_txt := 'Brak katalogu ' + pathToPdf;
    ferror := True;
    exit;
  end;

  pdf_file := pathToPdf + '\' + PageStructList[0].RaportFileName + id_PDF_file_str + '.pdf';
  //ReportStruct
  try
    FDConnection := TFDConnection.Create(nil);
    try
          lPdf := TPdfDocumentGDI.Create;
          try
            lPdf.Info.Author        := '';
            lPdf.Info.CreationDate  := Now;
            lPdf.Info.Creator       := '';
            lPdf.DefaultPaperSize   := psA4;
            lPdf.UseUniscribe := true;
            lPDF.Canvas.SetLineWidth(0.1);
            AddParamSQLINT('PAGE_NO', 0);
            for pageIndx := 0 to PageStructList.Count -1 do
            begin
               if GeneratePage(lPDF, PageStructList[pageindx].ReportStruct ) then
                 exit;
            end;
            lPdf.SaveToFile(pdf_file)

          finally
            lPdf.Free;
          end;
    finally
      FDConnection.Free;
    end;
  except
    on E: Exception do
    begin
      ferror_txt := E.Message;
      ferror := True;
    end;
  end;
end;


function TPreparingreport.SetHDetailBanf(fCanvas: TCanvas; band:Tband; ReportStruct:TReportStruct): integer;
var
  i: Integer;
  fobjBand: TobjBand;
  txt: string;
  TextRect: TRect;
  maxSizeh: Integer;
begin
  maxSizeh:= 0;
  for i := 0 to band.objBandList.Count -1 do
  begin
        fobjBand := band.objBandList[i];
        if (fobjBand.typeobj = TYPE_TEXT_OBJ) or (fobjBand.typeobj = TYPE_FIELD_OBJ) then
        begin

        end else
        if fobjBand.typeobj = TYPE_BOX_OBJ then
        begin
        end else
        if fobjBand.typeobj = TYPE_FIELD_BOX_OBJ then
        begin
          if fobjBand.AutoSizeH = 1 then
          begin
            txt := GetValueFromQuery(fobjBand, ReportStruct);
            fCanvas.Font.Size := fobjBand.FontSize;
            fCanvas.Font.Style := [];
            if fobjBand.FontStyleBold = 1 then
              fCanvas.Font.Style := fCanvas.Font.Style + [fsBold];
            if fobjBand.FontStyleItalic = 1 then
              fCanvas.Font.Style := fCanvas.Font.Style + [ fsItalic];
            if fobjBand.FontStyleUnderLine = 1 then
              fCanvas.Font.Style := fCanvas.Font.Style + [fsUnderline];

            TextRect.Top   :=  fobjBand.Top ;//+ fCanvas.TextHeight('W');// div 3;
            TextRect.Left  := fobjBand.Left ;// div 3;
            TextRect.Bottom := fobjBand.Height ;//- 50;//div 3;
            TextRect.Right := fobjBand.Left + fobjBand.Width - fCanvas.TextWidth('W') - fobjBand.MarginRight;
            fCanvas.TextRect(TextRect, txt, [tfCalcRect, tfNoClip, tfWordBreak]);
            if maxSizeh < TextRect.Bottom - TextRect.Top  then
              maxSizeh := TextRect.Bottom - TextRect.Top + fCanvas.TextHeight('¯¥') div 2;
          end else
          begin

            if fobjBand.Top + fobjBand.Height > band.Height then
               fobjBand.Height := band.Height - fobjBand.Top;
          end;
        end;
  end;
  if band.OldHeight > maxSizeh then
     maxSizeh := band.OldHeight;
  result := maxSizeh;
end;

const
   MAIN_QUERY =1 ;

function TPreparingreport.RednderBandDetail(fCanvas: TCanvas{; var EndRap: boolean};ReportStruct: TReportStruct): boolean;
var
  i: integer;
  reportQuery: TReportQuery;
  hDetail:Integer;
  break_for: boolean;
begin
  result := false;
  reportQuery := nil;
  for i:= 0 to ReportStruct.ReportQueryList.Count -1 do
  begin
    if ReportStruct.ReportQueryList[i].mainQ = MAIN_QUERY then
    begin
      reportQuery :=  ReportStruct.ReportQueryList[i];
      break;
    end;
  end;
  ReportStruct.DetailCount := ReportStruct.TopCount;
  //EndRap := False;
  AddUpdateParamRapVarBoolean(ReportStruct.globalVariables, 'ENDRAP' , false);
  while True do
  begin
    hDetail := 0;
    {for i := 0 to ReportStruct.bandList.Count -1 do
      if (ReportStruct.bandList[i].typeBand = DATAIL_BAND) then
        hDetail := hDetail + ReportStruct.bandList[i].Height; //SetHDetailBanf(fCanvas, ReportStruct, ReportStruct.bandList[i]);

    if ReportStruct.DetailCount + hDetail > ReportStruct.DetailCountMax then
    begin
      EndRap := False;
      break;
    end; }
    break_for := false;
    for i := 0 to ReportStruct.bandList.Count -1 do
    begin
      if (ReportStruct.bandList[i].typeBand = DATAIL_BAND) then
      begin
        hDetail := SetHDetailBanf(fCanvas, ReportStruct.bandList[i], ReportStruct );
        if ReportStruct.DetailCount + hDetail > ReportStruct.DetailCountMax then
        begin
          //EndRap := False;
          AddUpdateParamRapVarBoolean(ReportStruct.globalVariables, 'ENDRAP' , false);
          break_for := true;
          break;
        end;
        ReportStruct.bandList[i].Height := hDetail;
        BandRectangleDetail(fCanvas,ReportStruct.bandList[i], ReportStruct);
        if Assigned(reportQuery) then
           reportQuery.Sum();
      end;
    end;
    if  break_for then
        break;
    //EndRap := true;
    //  break;
    if reportQuery <> nil  then
    begin
      reportQuery.Next;
      if reportQuery.Eof then
      begin
        //EndRap := True;
        AddUpdateParamRapVarBoolean(ReportStruct.globalVariables, 'ENDRAP' , True);
        break;
      end;
    end else
    begin
      //EndRap := True;
      AddUpdateParamRapVarBoolean(ReportStruct.globalVariables, 'ENDRAP' , True);
      break;
    end;
    //if ReportStruct.DetailCount > ReportStruct.
  end;
end;

function TPreparingreport.BandRectangleDetail(fCanvas: TCanvas; band: Tband;ReportStruct: TReportStruct): Boolean;
var
  i: Integer;
  objBand: TobjBand;
begin
  result := false;
  fCanvas.Pen.Width := 1;
  if ReportStruct.debug = 1 then
  begin
    fCanvas.Rectangle(0, ReportStruct.DetailCount , ReportStruct.PageWidth, ReportStruct.DetailCount + band.Height);
    fCanvas.TextOut( 10, ReportStruct.DetailCount + 10,  band.name);
  end;
  for i := 0 to band.objBandList.Count -1 do
  begin
    objBand := band.objBandList[i];
    if objBand.AutoSizeH = 1 then
       objBand.Height := band.Height;
    RenderOBJ(ReportStruct.DetailCount, fCanvas, objBand, ReportStruct);
  end;
  ReportStruct.DetailCount := ReportStruct.DetailCount +  band.Height-1;
end;

function TPreparingreport.CreateSQL(ReportStruct: TReportStruct): TReportQuery;
var
  i, ii: Integer;
  FDParam: TFDParam;
begin
  result := nil;
  FDConnection.Params.Clear;
  FDConnection.Params.Add('DriverID=' + DriverName);
  FDConnection.Params.Add('Database=' + Database);
  FDConnection.Params.Add('LockingMode=Normal');
  FDConnection.UpdateOptions.AssignedValues := [uvEDelete, uvEInsert, uvEUpdate];
  FDConnection.UpdateOptions.EnableDelete := False;
  FDConnection.UpdateOptions.EnableInsert := False;
  FDConnection.UpdateOptions.EnableUpdate := False;
  FDConnection.LoginPrompt := False;
  FDConnection.Connected := True;
  for i:= 0 to ReportStruct.ReportQueryList.Count -1 do
  begin
    ReportStruct.ReportQueryList[i].Connection := FDConnection;
    ReportStruct.ReportQueryList[i].ResourceOptions.ParamCreate := false;
    ReportStruct.ReportQueryList[i].SQL.Text := ReportStruct.ReportQueryList[i].sqlTxt;
    for ii := 0 to ReportStruct.ReportQueryList[i].ParamValues.Count - 1 do
    begin
      FDParam := ReportStruct.ReportQueryList[i].Params.Add;
      FDParam.Name := ReportStruct.ReportQueryList[i].ParamValues[ii].paramName;
      if ReportStruct.ReportQueryList[i].ParamValues[ii].paramType = REPORT_INT  then
         FDParam.DataType := ftInteger;
      FDParam.ParamType := ptInput;
      FDParam.Value := GetValueParam( ReportStruct,FDParam.Name);
    end;
    if ReportStruct.ReportQueryList[i].mainQ = MAIN_SQL then
       result := ReportStruct.ReportQueryList[i];
    ReportStruct.ReportQueryList[i].Open();
    if ReportStruct.ReportQueryList[i].mainQ = MAIN_SQL then
       ReportStruct.ReportQueryList[i].BuildSUMLIst();
  end;
end;

function TPreparingreport.GetValueParam( ReportStruct: TReportStruct;fName: string): Variant;
var
  i: Integer;
begin
  result := Null;
  for i := 0 to ReportStruct.globalVariables.Count -1 do
  begin
     if fName = ReportStruct.globalVariables[i].paramName  then
       if ReportStruct.globalVariables[i].paramType = REPORT_INT  then
          result := ReportStruct.globalVariables[i].paramValueInt
       else if ReportStruct.globalVariables[i].paramType = REPORT_BOOLEAN  then
         result := ReportStruct.globalVariables[i].paramValueBoolean;
  end;
end;

function TPreparingreport.RednderBandTop(fCanvas: TCanvas;ReportStruct: TReportStruct): boolean;
var
  i: integer;
begin
  result := false;
  for i := 0 to ReportStruct.bandList.Count -1 do
  begin
    if (ReportStruct.bandList[i].typeBand = PAGE_BAND_TOP) then
    begin
      BandRectangleTop(fCanvas,ReportStruct.bandList[i], ReportStruct);
    end else
    if (ReportStruct.bandList[i].typeBand = REPORT_BAND_TOP) and
           (ReportStruct.CurrentPageNo = 1) then
    begin
      BandRectangleTop(fCanvas,ReportStruct.bandList[i], ReportStruct);
    end;
  end;
end;




function TPreparingreport.BandRectangleTop(fCanvas: TCanvas; band: Tband;ReportStruct: TReportStruct): Boolean;
var
  i: Integer;
  objBand: TobjBand;
begin
  result := false;
  fCanvas.Pen.Width := 1;
  if ReportStruct.debug = 1 then
  begin
    fCanvas.Rectangle(0, ReportStruct.TopCount , ReportStruct.PageWidth, ReportStruct.TopCount + band.Height);
    fCanvas.TextOut( 10, ReportStruct.TopCount + 10,  band.name);
  end;
  for i := 0 to band.objBandList.Count -1 do
  begin
    objBand := band.objBandList[i];
    RenderOBJ(ReportStruct.TopCount,fCanvas, objBand, ReportStruct);
  end;
  ReportStruct.TopCount := ReportStruct.TopCount +  band.Height-1;
end;

procedure TPreparingreport.RenderOBJ(ftop:integer;fCanvas: TCanvas; fobjBand: TobjBand;ReportStruct: TReportStruct);
var
  textW: integer;
  txt: Variant;
  textRect1: TRect;
  R: TRect;
  str: string;
  outVar: Variant;
begin
    textRect1 := TRect.Empty();
    if (fobjBand.typeobj = TYPE_TEXT_OBJ) or (fobjBand.typeobj = TYPE_FIELD_OBJ) then
    begin
      fCanvas.Font.Size := fobjBand.FontSize;
      fCanvas.Font.Style := [];
      if fobjBand.FontStyleBold = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [fsBold];
      if fobjBand.FontStyleItalic = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [ fsItalic];
      if fobjBand.FontStyleUnderLine = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [fsUnderline];

      if fobjBand.typeobj = TYPE_TEXT_OBJ then
      begin
        str := ParsujStr(fobjBand.Caption, ReportStruct.globalVariables);
        if fobjBand.Align = CENTER_TEXT_OBJ then
        begin
          textW := fCanvas.TextWidth(str);
          textW := Trunc((ReportStruct.PageWidth - textW) / 2);
        end else textW := fobjBand.Left;
        fCanvas.TextOut( textW, ftop +  fobjBand.Top ,  str);
      end else if fobjBand.typeobj = TYPE_FIELD_OBJ then
      begin
        txt := GetValueFromQuery(fobjBand, ReportStruct);
        if fobjBand.Align = CENTER_TEXT_OBJ then
        begin
          textW := fCanvas.TextWidth(txt);
          textW := Trunc((ReportStruct.PageWidth - textW) / 2);
        end else textW := fobjBand.Left;
        fCanvas.TextOut( textW, ftop +  fobjBand.Top , txt);
      end;
    end else
    if fobjBand.typeobj = TYPE_BOX_OBJ then
    begin
      fCanvas.Font.Size := fobjBand.FontSize;
      fCanvas.Font.Style := [];
      if fobjBand.FontStyleBold = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [fsBold];
      if fobjBand.FontStyleItalic = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [ fsItalic];
      if fobjBand.FontStyleUnderLine = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [fsUnderline];
      fCanvas.Pen.Width := fobjBand.PenWidth;
      if fobjBand.PenWidth > 0 then
         fCanvas.Rectangle(fobjBand.Left, ftop +  fobjBand.Top ,
                         fobjBand.Left + fobjBand.Width, ftop +  fobjBand.Top + fobjBand.Height );

      TextRect1.Top   := ftop +  fobjBand.Top;
      TextRect1.Left  := fobjBand.Left;
      TextRect1.Bottom := ftop +  fobjBand.Top + fobjBand.Height;
      TextRect1.Right := fobjBand.Left + fobjBand.Width;
      str := fobjBand.Caption;

      if fobjBand.Align = LEFT_TEXT_OBJ then
      begin
        TextRect1.Top   := TextRect1.Top + fCanvas.TextHeight('¯¥') div 6;
        TextRect1.Left  := TextRect1.Left + fCanvas.TextWidth('W') div 2;
        TextRect1.Bottom := TextRect1.Bottom + fCanvas.TextHeight('¯¥') div 2;
        TextRect1.Right := TextRect1.Right - fCanvas.TextWidth('W') div 2;
        fCanvas.TextRect(TextRect1, str, [tfLeft, tfWordBreak]) ;
      end else if fobjBand.Align = RIGHT_TEXT_OBJ then
      begin
        TextRect1.Top   := TextRect1.Top + fCanvas.TextHeight('¯¥') div 6;
        TextRect1.Left  := TextRect1.Left + fCanvas.TextWidth('W') div 2;
        TextRect1.Bottom := TextRect1.Bottom + fCanvas.TextHeight('¯¥') div 2;
        TextRect1.Right := TextRect1.Right - fCanvas.TextWidth('W') div 2;
        fCanvas.TextRect(TextRect1, str, [tfRight, tfWordBreak]) ;
      end
      else
         DrawVert(fCanvas, TextRect1, str) ;
    end else
    if (fobjBand.typeobj = TYPE_FIELD_BOX_OBJ) or (fobjBand.typeobj = TYPE_FIELD_SUM_BOX_OBJ) then
    begin
      if fobjBand.typeobj = TYPE_FIELD_SUM_BOX_OBJ then
      begin
         str := GetSumValue(fobjBand, ReportStruct); //sUMobjList
      end else if (fobjBand.typeobj = TYPE_FIELD_BOX_OBJ) then
         str := GetValueFromQuery(fobjBand, ReportStruct);
      fCanvas.Font.Size := fobjBand.FontSize;
      fCanvas.Font.Style := [];
      if fobjBand.FontStyleBold = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [fsBold];
      if fobjBand.FontStyleItalic = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [ fsItalic];
      if fobjBand.FontStyleUnderLine = 1 then
        fCanvas.Font.Style := fCanvas.Font.Style + [fsUnderline];
      fCanvas.Pen.Width := fobjBand.PenWidth;
      if fobjBand.PenWidth > 0 then
      begin
        fCanvas.Rectangle(fobjBand.Left, ftop +  fobjBand.Top ,
                         fobjBand.Left + fobjBand.Width, ftop +  fobjBand.Top + fobjBand.Height );
      end;

      if fobjBand.Visible <> '' then
      begin
        if Expression(fobjBand.Visible, ReportStruct.globalVariables, outVar) then
          fobjBand.AsVisible := outVar
        else
          fobjBand.AsVisible := True;
      end;
      if fobjBand.AsVisible then
      begin
        if fobjBand.Align = LEFT_TEXT_OBJ then
        begin
          TextRect1.Top   := ftop +  fobjBand.Top + fCanvas.TextHeight('¯¥') div 6;
          TextRect1.Left  := fobjBand.Left + fCanvas.TextWidth('W') div 2;
          TextRect1.Bottom := ftop +  fobjBand.Top + fobjBand.Height - fCanvas.TextHeight('¯¥') div 6;
          TextRect1.Right := fobjBand.Left + fobjBand.Width - fCanvas.TextWidth('W') div 2 - fobjBand.MarginRight;
          fCanvas.TextRect(TextRect1, str, [tfLeft, tfWordBreak]);
        end else if fobjBand.Align = RIGHT_TEXT_OBJ then
        begin
          TextRect1.Top   := ftop +  fobjBand.Top + fCanvas.TextHeight('¯¥') div 6;;
          TextRect1.Left  := fobjBand.Left + fCanvas.TextWidth('W') div 2;
          TextRect1.Bottom := ftop +  fobjBand.Top + fobjBand.Height - fCanvas.TextHeight('¯¥') div 6;
          TextRect1.Right := fobjBand.Left + fobjBand.Width - fCanvas.TextWidth('W') div 2 - fobjBand.MarginRight;
          fCanvas.TextRect(TextRect1, str, [tfRight,tfWordBreak]);
        end else
        begin
          //
          //fobjBand.Visible
          TextRect1.Top   := ftop +  fobjBand.Top + fCanvas.TextHeight('¯¥') div 6;
          TextRect1.Left  := fobjBand.Left + fCanvas.TextWidth('W') div 2;
          TextRect1.Bottom := ftop +  fobjBand.Top + fobjBand.Height - fCanvas.TextHeight('¯¥') div 6;
          TextRect1.Right := fobjBand.Left + fobjBand.Width - fCanvas.TextWidth('W') div 2 - fobjBand.MarginRight;
          fCanvas.TextRect(TextRect1, str, [tfCenter,tfWordBreak]);
        end;
      end;
    end;
end;

function  TPreparingreport.ParsujStr(Caption : string; globalVariables : TList <TParamValue>): string;
var
  i: integer;
  val : string;
begin
  result := Caption;
  for i := 0 to globalVariables.Count -1 do
  begin
     val := '';
     if globalVariables[i].paramType = REPORT_STRING then
         val := globalVariables[i].paramValueString
     else if globalVariables[i].paramType = REPORT_INT then
         val := IntToStr(globalVariables[i].paramValueInt);
     Caption := StringReplace(Caption, '$[' + globalVariables[i].paramName + ']$' , val,[rfReplaceAll]);
  end;
  result := Caption;
end;

function  TPreparingreport.GetGlobalValueInt(globalVariables : TList <TParamValue>; fname: string;
                           var ParamValue: TParamValue; var return: Boolean): Integer;
var
  i: integer;
  val : integer;
begin
  result := 0;
  val := 0;
  return := false;
  ParamValue := nil;
  for i := 0 to globalVariables.Count -1 do
  begin
     val := 0;
     if globalVariables[i].paramName = fname then
     begin
       if globalVariables[i].paramType = REPORT_INT then
       begin
         val := globalVariables[i].paramValueInt;
         ParamValue := globalVariables[i];
         return := True;
         break;
       end;
     end;
  end;
  result := val;
end;

function  TPreparingreport.SetGlobalValueInt(globalVariables : TList <TParamValue>; fname: string; value: Integer): Boolean;
var
  i: integer;
begin
  result := false;
  for i := 0 to globalVariables.Count -1 do
  begin
     if globalVariables[i].paramName = fname then
     begin
       if globalVariables[i].paramType = REPORT_INT then
       begin
         globalVariables[i].paramValueInt := value;
         globalVariables[i].paramValueString := IntToStr(value);
         result := True;
         break;
       end;
     end;
  end;
end;

procedure TPreparingreport.DrawVert(fCanvas: TCanvas; Box: TRect; const Text: string);
var
  R: TRect;
  s: string;
begin
  s := Text ;
  R := TRect.Empty();

  R.top := Box.top + 4;
  R.Left := Box.Left + 4;
  R.Bottom := Box.Bottom - 4;
  R.Right := Box.Right -4  ;
  fCanvas.TextRect(r, s, [tfCalcRect, tfNoClip, tfWordBreak]);
  Box.Left := Box.Left + 4 ;
  Box.Top := (Box.Top + ((Box.Bottom - Box.top) - (r.Bottom - r.top + 4)) div 2)  ;
  Box.Right := Box.Right- 4;
  Box.Bottom := Box.Bottom -4;
  fCanvas.TextRect(box, s, [tfCenter, tfWordBreak]);
end;

function TPreparingreport.GetSumValue(fobjBand: TobjBand;ReportStruct: TReportStruct): string; //sUMobjList
var
  i, ii, iii: Integer;
  val: Currency;
begin
  result := '';
  val := 0;
  for i:= 0 to fobjBand.sUMobjList.Count -1 do
  begin
      for ii:= 0 to ReportStruct.ReportQueryList.Count -1 do
      begin
        if ReportStruct.ReportQueryList[ii].mainQ = MAIN_SQL then
        begin
          for iii := 0 to ReportStruct.ReportQueryList[ii].sUMValues.Count -1 do
          begin
            if fobjBand.sUMobjList[i].fieldName = ReportStruct.ReportQueryList[ii].sUMValues[iii].field.FieldName then
            begin
              if fobjBand.sUMobjList[i].typeSrc = 0 then
                 val := ReportStruct.ReportQueryList[ii].sUMValues[iii].paramValuePageFloat
              else if fobjBand.sUMobjList[i].typeSrc = 1 then
                 val := ReportStruct.ReportQueryList[ii].sUMValues[iii].paramValuePagePrevFloat
              else if fobjBand.sUMobjList[i].typeSrc = 2 then
                 val := ReportStruct.ReportQueryList[ii].sUMValues[iii].paramValueAllFloat;
              if fobjBand.FormatValue <> '' then
                 result := FormatFloat(fobjBand.FormatValue, val);
              exit;
            end;
          end;
        end;
      end;
  end;
end;

function TPreparingreport.GetValueFromQuery(fobjBand: TobjBand;ReportStruct: TReportStruct): string;
var
  i: Integer;
  val: Double;
begin
  result := '';
  for i:= 0 to ReportStruct.ReportQueryList.Count -1 do
  begin
    if ReportStruct.ReportQueryList[i].queryNo = fobjBand.queryNo then
    begin
      var Field : TField;
      Field := ReportStruct.ReportQueryList[i].FieldByName(fobjBand.Field);

      if (Field.DataType = ftFloat) or (Field.DataType = ftLargeint)  or
         (Field.DataType = ftCurrency) then
      begin
        val := ReportStruct.ReportQueryList[i].FieldByName(fobjBand.Field).AsFloat;
        if fobjBand.FormatValue <> '' then
           result := FormatFloat(fobjBand.FormatValue, val);
        exit;
      end else
      begin
        if ReportStruct.ReportQueryList[i].FieldByName(fobjBand.Field).IsNull then
        begin
          result := '';
          exit;
        end;
        result := ReportStruct.ReportQueryList[i].FieldByName(fobjBand.Field).Value;
        exit;
      end;
    end;
  end;
end;

function TPreparingreport.BandRectangleBottom(fCanvas: TCanvas; band: Tband;ReportStruct: TReportStruct): Boolean;
var
  i: Integer;
  objBand: TobjBand;
begin
  result := false;
  fCanvas.Pen.Width := 1;
  if ReportStruct.debug = 1 then
  begin
    fCanvas.Rectangle(0, ReportStruct.BottomCount - band.Height, ReportStruct.PageWidth, ReportStruct.BottomCount- 3);
    fCanvas.TextOut( 10, ReportStruct.BottomCount - band.Height + 10,  band.name);
  end;

  for i := 0 to band.objBandList.Count -1 do
  begin
    objBand := band.objBandList[i];
    RenderOBJ(ReportStruct.BottomCount - band.Height,fCanvas, objBand, ReportStruct);
  end;

  ReportStruct.BottomCount := ReportStruct.BottomCount -  band.Height+1;
end;

{
function TPreparingreport.BandRectangleTop(fCanvas: TCanvas; band: Tband;ReportStruct: TReportStruct): Boolean;
var
  i: Integer;
  objBand: TobjBand;
begin
  fCanvas.Pen.Width := 1;
  if ReportStruct.debug = 1 then
  begin
    fCanvas.Rectangle(0, ReportStruct.TopCount , ReportStruct.PageWidth, ReportStruct.TopCount + band.Height);
    fCanvas.TextOut( 10, ReportStruct.TopCount + 10,  band.name);
  end;
  for i := 0 to band.objBandList.Count -1 do
  begin
    objBand := band.objBandList[i];
    RenderOBJ(ReportStruct.TopCount,fCanvas, objBand, ReportStruct);
  end;
  ReportStruct.TopCount := ReportStruct.TopCount +  band.Height-1;
end;  }

function TPreparingreport.RednderBandBottom(fCanvas: TCanvas;ReportStruct: TReportStruct): boolean;
var
  i: integer;
  endRap: Boolean;
begin
  result := false;
  if ReportStruct.STICK_BAND_BOTTOM = 1 then
    ReportStruct.BottomCount := ReportStruct.DetailCount + ReportStruct.MaxBandBottomCount - 1;

  endRap := GetValueParam(ReportStruct, 'ENDRAP');
  for i := 0 to ReportStruct.bandList.Count -1 do
  begin
    if (ReportStruct.bandList[i].typeBand = PAGE_BAND_BOTTOM) then
    begin
      BandRectangleBottom(fCanvas,ReportStruct.bandList[i], ReportStruct);
    end else
    if (ReportStruct.bandList[i].typeBand = REPORT_BAND_BOTTOM) and
           (endRap = true) then
    begin
      BandRectangleBottom(fCanvas,ReportStruct.bandList[i], ReportStruct);
    end;
  end;
end;

function TPreparingreport.GetMaxDetaiCount( ReportStruct: TReportStruct): integer;
var
  i: integer;
  maxDetail: integer;
begin
  maxDetail := ReportStruct.PageHeight+1;
  for i := 0 to ReportStruct.bandList.Count -1 do
  begin
    if (ReportStruct.bandList[i].typeBand = PAGE_BAND_BOTTOM) then
      maxDetail := maxDetail -  ReportStruct.bandList[i].Height+1
    else
    if (ReportStruct.bandList[i].typeBand = REPORT_BAND_BOTTOM) then
       maxDetail := maxDetail -  ReportStruct.bandList[i].Height+1;
  end;
  result := maxDetail;
end;

function TPreparingreport.GetMaxBandBottomCount( ReportStruct: TReportStruct): integer;
var
  i: integer;
  maxDetail: integer;
begin
  maxDetail := 0;
  for i := 0 to ReportStruct.bandList.Count -1 do
  begin
    if (ReportStruct.bandList[i].typeBand = PAGE_BAND_BOTTOM) then
      maxDetail := maxDetail +  ReportStruct.bandList[i].Height+1
    else
    if (ReportStruct.bandList[i].typeBand = REPORT_BAND_BOTTOM) then
       maxDetail := maxDetail +  ReportStruct.bandList[i].Height+1;
  end;
  result := maxDetail;
end;

procedure TPreparingreport.UpdateUI();
begin
  if Assigned(FUpdateUIEvent) then
    if PageStructList.Count > 0 then
      FUpdateUIEvent(fposition, PageStructList[0].RaportFileName);
end;

procedure TPreparingreport.EndUpdateUI();
begin
  if Assigned(FEndUpdateUIEvent) then
    if PageStructList.Count > 0 then
      FEndUpdateUIEvent(ferror_txt, ferror, pdf_file, PageStructList[0].RaportFileName)
    else
      FEndUpdateUIEvent(ferror_txt, ferror, pdf_file, '');
end;




end.
