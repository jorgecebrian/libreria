unit uLibreria;
interface
  uses
    System.SysUtils, System.StrUtils, System.Classes, System.DateUtils,
    System.Math, System.Character, System.RegularExpressions,
    Vcl.StdCtrls, Vcl.Controls, Vcl.Forms,
    Winapi.Windows, Winapi.ShellAPI, Messages, cxGrid, cxGridExportLink,
    Data.Win.ADODB,  Data.DB, FireDAC.Comp.Client, FireDAC.Stan.Option,
    MVCFramework.Commons, Uni, JsonDataObjects, JSON, VirtualQuery;

// LÓGICA PROGRAMA
  function DaTextoTipoTrafico( const pLetraTipoTrafico:String):String;
  function CodificaVehiculo(const pTipo_veh, pNeurona, pEspecialidad:string):string;

// TABLAS
      // Copia contenido tabla UNI en tabla Firedac MEM
  procedure CopiaContenidoTablas( Torigen:TuniTable; TDestino:TFDMemTable); overload;
  procedure CopiaContenidoTablas( Torigen:TFDMemTable; TDestino:TuniTable; const borraDestino:Boolean); overload;
  procedure ActualizaContenidoTabla( Torigen:TFDMemTable; TDestino:TuniTable);
  procedure AceleraTabla( pTabla:TFDMemTable);
  function  GeneraListaIdBorrarRESTdeMemTabla( pTMem:TFDMemTable):String;
// JSON
      // Crea JSON para INSERT y UPDATE,+lista de Id para borrado
  procedure ExportaDatasetParaModificarREST(elDataset: TDataSet; var BodyI, BodyU:String);
      // Crea JSON desde un Dataset
  function ExportaDataset_enJSON(elDataset: TDataSet): string;
      // Prepara SQL Update desde JSON
  function MontaUpdates(principioSQL, elBody:String):String;
      // Prepara SQL INSERT desde JSON
  function DeJSONrellenaValoresParaInsertSQL( pBody:String):String;

// PASSWORD
      // Devuelve desencriptade la clave
  function DesencriptaClave(const pClaveEncrip:String):String;
      // Devuelve clave encriptada o mensaje con error "Error: el carácter X no se admite en el password."
  function EncriptaClave(const pClave:String):String;
    // Dado un pwd Virosque mezclado extraigo solo pwd
  function ExtraePwd (const pwdMezcla:String):String;
    // Dado un pwd Virosque mezclado extraigo solo Version
  function ExtraeVersion (const pwdMezcla:String):String;
    // Dado un pwd Virosque mezclado extraigo solo opción
  function ExtraeModulo (const pwdMezcla:String):Integer;
    // Encriptar y desencriptar texto
  function EncryptStr(const pS :WideString; pKey: Word): String;
  function DecryptStr(const pS: String; pKey: Word): String;

// TEXTO y FECHA y NUMERO #12 salto de página, #10 salto de línea
// #13 retorno de carro, #9 tabulador horizontal, #11 tabulador vertical
  function FechaToTextoFechaTrimble( const pFecha:TDateTime):String;// 29/10/2022 11:11:11 fecha  -- > 2022-10-29T11:11:11.000
  function tDateTimetoSQLserverDate(const pFecha: TDateTime):string; // fecha -> 25/10/2022 14:28:30
  function TDateTimetoSQLdate( const pFecha:TDateTime):string; //fechaTime -->  2022-10-25 14:28:30
//  function TDatetoSQLfecha( const pFecha:TDateTime):string; //fecha -->  2022-10-25
  function FechaHoraToSQLtexto(const pFecha:TDateTime):String; //  fecha tDate 14/01/2023 22:33:22-->  CONVERT(DATETIME, '2023-01-14 22:33:22', 102)
  function FechaToSQLtexto(const pFecha:TDate):String; //  fecha tDate 14/01/2023 -->  CONVERT(DATETIME, '2020-01-14', 102)
  function TDatetoSQLdate( const pFecha:TDate):string; //  fecha tDate 14/01/2023 --> 20230114
  function TDateToStringDate(const pFecha:TDateTime):string; // fecha --> 25/11/2022
  function TDateTimeToStringTime(const pFecha:TDateTime):string;// DE TdateTime a  string con hora:minuto // 14:28:30
  function FechaToYearDia( const pFecha:Tdate):String;   // fecha 3-2-2023  -->  2023034
  function TDatetoPrincipioMes( const pFecha: TDateTime):string;  // 25/10/2022 14:28:30  --> 01/10/2022
//    Texto a Fecha
  function FechaTextoTrimbleToDateTime( const pFeTr:String):TDateTime;  //  2022-10-29T11:11:11.000  -->  29/10/2022 11:11:11
  function FechaTextoTrimbleToDate( const pFeTr:String):TDate; //  2022-10-29  -->  29/10/2022
  function FechaTextoToDate( const pFTxt:String):TDate; //  2022-10-29 o 29-10-2022 -->  29/10/2022
     // Convierte una cadena en fecha  20230517   -> 17/05/2023
  Function SQLdate_to_Date(pFechaString : String) : TDate;
     // Convierte una cadena en fecha   17/05/2023   ->  17/05/2023
  Function STRING_to_DATA(pFechaString : String) : tDateTime;
     // Convierte una cadena en fecha hora
  Function STRING_to_DATETIME(pFechaString : String) : tDateTime;

       //2023-03-29 00:00 y 1899-12-31 11:11:11.000 --> 29/03/2023 11:11:11
  function FechayHoraToFechaHora( const pFecha, pHora:TDateTime):TDateTime; // Une Fecha y Hora de Trans en una FechaHora
  function LastDay(const pFecha: TDateTime): TDateTime;  // 25/10/2022 --> 30/11/2022
          // DE TdateTime a  string con fecha completa   2022-10-25 14:28:30
  function SumaMesTexto( const pFecha:String):string;   // 25/10/2022 --> 25/11/2022
       // 07/05/2023 14:25:33 --> 01/05/2023 00:00:00
  function PonMesIni(const pFecha:TDateTime):TDateTime;
       // 07/05/2023 14:25:33 --> 31/05/2023 23:59:59
  function PonMesFin(const pFecha:TDateTime):TDateTime;
       // Pon Fecha con Time 00:00:00
  function PonMedianocheIni(const pFecha:TDateTime):TDateTime;
      // Pon Fecha con Time 23:59:59
  function PonMedianocheFin(const pFecha:TDateTime):TDateTime;
       //  2023,  034   -->   fecha 3-2-2023
  function YearDiaToFecha(const pYear, pDiaYear: Word): TDateTime;

     // Devuelve nombre mes   3 --> Marzo
  function obten_nombre_mes(const pNumMes:integer):string;
     // Convierte string Hora:Min en string min Ejemplo 03:25  -->  205
  function hora_a_minutos(const pHora:string):string;
      // Si me pasan 13:30 (13 horas y 30 minutos), debe devolver 13,5
  function devuelve_horas(const pHora:string):real;

  function EScapeString(const pCadena: string): string;
      // Copia el contenido de un fichero en un string
  function LeeTextoDesdeFichero(const pRutaNombreFichero: string): string;
      // Devuelve la posicion del datoAbuscar en el array
  function IndiceArray(pLista: array of String; const pDatoAbuscar: String): Integer;
  function NumVecesEsta_Subcadena_EnCadena( const pSubcadena, pCadena:string):Integer;
     // Quita de una cadena todos los caracteres X. Ej: cadena, '['
  function LimpiarCadenaDeCaracter(const pCadena:string; const pCaracter: Char): string;
     // Quita de una cadena todos los carácteres que no sean letras o numeros
  function LimpiarCadenaSoloLetrasNumeros(const vCadena: string): string;
     // Quita de una cadena todos los carácteres que no sean numeros
  function LimpiarCadenaSoloNumeros(const vCadena: string): string;
     // Genera una cadena con letras aleatorias (tamaño longitud)
  function PalabraAleatoria(const pLongitud: integer): string;
     // Cambiar en cadena una subcadena por otra
  function CambiaEn(const pCadena, pOri, pDes:String):String;
  function FloatToSQLnum(const pNum:Double):String; // 23,5 ---> "23.5"
     // Posición de Primer carácter no numérico
  function PosNoNum(const pTexto:String):Integer;
  function DameNumDeTexto(const pText: string): Integer;
      // Da Ruta de archivo   c:\proy\cod\gest\hola.exe --> c:\proy\cod\gest\
  function DaRutaDelArchivo(pArchivo:String):String;
  function gradosARadianes(const pGrados:Double):Double;
    // Kilometros entre dos coordenadas.
  function calcularDistanciaEntreDosCoordenadas( pLat1, pLat2, pLong1, pLong2:Double):Integer;
    // Busca entero en lista de enteros
  function BuscarEntero(const pLista: array of Integer; const pEnteroBuscado: Integer): Boolean;
     // Convierte un TStringList en un String con elementos entre comillas y separados por coma
  function TStringListToString(const pLis: TStringList): String;
     // Lo mismo sin QuotedStr
  function TStringListToStringNum(const pLis: TStringList): String;
    // Busca hora en string
  function BuscaHoraEnString(const pTex: String): String;


  // API Windows
     // Devuelve el nombre de la máquina y del usuario que ejecuta.
  function  GetCurrentUserName(out DomainName, UserName: string): Boolean;
  procedure CierraMiAplicacion( const NombreAplicacion:PWideChar);
  procedure CambiaENTERporTABenKeyPress(var Key:PChar);
  procedure CopiaPos_y_Tam( pTeditOrig:Tedit; pTcomboDest:TCombobox);
  procedure SetCursorFin;
  procedure SetCursorIni(Cursor: TCursor = crHourGlass);
  function  GeneraNombreArchivoTemporal( const pLargo:Integer; const pExtens:String):string;

//  IIF
     // como el IIF de C o SQL
  function IIF(esverdad: boolean; vVerdad, vFalso: integer): integer; overload;
  function IIF(esverdad: boolean; vVerdad, vFalso: Extended): Extended; overload;
  function IIF(esverdad: boolean; vVerdad, vFalso: string): string; overload;
  function IIF(esverdad: boolean; vVerdad, vFalso: TObject): TObject; overload;
  function IIF( esverdad:boolean; vVerdad, vFalso: variant): variant; overload;

//  SQL
     // Ejecuta SELECT proporcionada y devuelve RecordCount, no cierra
  function LeeSQL( textoSQL:String; var objQuery:TADOQuery):integer; overload;
  function LeeSQL( textoSQL:String; var objQuery:TVirtualQuery):integer; overload;
  function LeeSQL( textoSQL:String; var objQuery:TUniQuery):integer; overload;
     // Ejecuta SQL de escritura proporcionada, cierra, no devuelve nada.
  procedure EscribeSQL( textoSQL:String; objQuery:TADOQuery); overload;
  procedure EscribeSQL( textoSQL:String; objQuery:TVirtualQuery); overload;
  procedure EscribeSQL( textoSQL:String; objQuery:TUniQuery); overload;

  type    // Para funcion GetCurrentUserName
    PTokenUser = ^TTokenUser;
    TTokenUser = packed record
      User: SID_AND_ATTRIBUTES;
    end;
const
  MESES: Array [1..12] Of String =
   ('Enero','Febrero',     'Marzo',  'Abril',     'Mayo',  'Junio',
    'Julio', 'Agosto','Septiembre','Octubre','Noviembre','Diciembre');
  DIAS: Array [1..7] Of String =
   ('Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo');
  PI: Double = 3.14159265358979323846;
  RADIO_TIERRA_EN_KILOMETROS: Double = 6371.0;
  TIPOS_VEHICULO: Array [1..9] Of String =
   ('FURGONETA', 'TURISMO', 'TRACTORA', 'LONA', 'PORTABOBINAS', 'ASTILLERA',
    'FRIGO', 'GONDOLA', 'RIGIDO');
  TIPO_TRAZA_CARGA: String = ',01,02,10,11,12,13,14,15,16,20,21,22,23,24,25,';
  CKEY1 = 53761;  CKEY2 = 32618;
  NUMEROS: String = '1234567890';
//                   172.16.1.118
  URLobra: String = 'http://188.85.81.2:8091/apiU/Obra';
  URLpubl: String = 'http://188.85.81.2:8091/api';
  URLpriv: String = 'http://188.85.81.2:8091/apiU/privado';
  URLtorre:String = 'http://188.85.81.2:8091/apiU/torre';
{  URLobra: String = 'http://127.0.0.1:8091/apiU/Obra';
  URLpubl: String = 'http://127.0.0.1:8091/api';
  URLpriv: String = 'http://127.0.0.1:8091/apiU/privado';
  URLtorre:String = 'http://172.16.1.118:8091/apiU/torre';}
implementation

// LÓGICA APLICACIÓN
function DaTextoTipoTrafico( const pLetraTipoTrafico:String):String;
begin
  Case IndexStr(pLetraTipoTrafico, ['P', 'A', 'E', 'H']) of
    0: result := 'Propio';
    1: result := 'Agencia';
    2: result := 'Enganche';
    3: result := 'Habitual';
  else
    result := 'Otros '+pLetraTipoTrafico;
  End;
end;

function CodificaVehiculo(const pTipo_veh, pNeurona, pEspecialidad:string):string;
var      vCod:string;
begin//TIPOS_VEHICULO: Array [1..9] Of String=('FURGONETA','TURISMO','TRACTORA',
//         'LONA', 'PORTABOBINAS', 'ASTILLERA',  'FRIGO', 'GONDOLA', 'RIGIDO');
  Case IndexStr(pTipo_veh, TIPOS_VEHICULO) of
    1: vCod := 'U'+pNeurona;  //FURGONETA
    2: vCod := 'C'+pNeurona;  //TURISMO
    3: if pEspecialidad='Internacional'   //TRACTORA
         then vCod:= 'T'+pNeurona+'1' // Internacional
         else vCod:= 'T'+pNeurona+'0'; // Nacional
    4: vCod:= 'L'+pNeurona+'0'; // LONA
    5: vCod:= 'L'+pNeurona+'1'; // PORTABOBINAS
    6: vCod:= 'L'+pNeurona+'2'; // ASTILLERA
    7: if pEspecialidad='Monotemperatura'         then vCod:= 'F'+pNeurona+'0' // FRIGO
         else if pEspecialidad='Bitemperatura'    then vCod:= 'F'+pNeurona+'1'
         else if pEspecialidad='Multitemperatura' then vCod:= 'F'+pNeurona+'2';
    8: vCod:= 'G'+pNeurona;     // GONDOLA
    9: vCod:= 'R'+pNeurona;     // RIGIDO
  else
       vCod:= pTipo_veh;
  end;
  result := LeftStr(vCod+'000',3);
end;

// TABLA
procedure CopiaContenidoTablas( Torigen:TuniTable; TDestino:TFDMemTable); overload;
var
  i, j:Integer;
begin
  Torigen.DisableControls;
  TDestino.DisableControls;
  Tdestino.Open;
  Tdestino.First;
  TDestino.EmptyDataSet;
  while not Tdestino.Eof do
    Tdestino.Delete;
  if TOrigen.Active = False then
     TOrigen.Open;
  Torigen.First;
  while not Torigen.Eof do begin
    TDestino.Append;
    j := 0;
    for i:=0 to Torigen.fields.count-1 do begin
      if (Torigen.Fields[i].fieldname='Id') and (Tdestino.Fields[i].FieldName <> 'Id') then
        j := i-1
      else
        if Tdestino.Fields[j].FieldName = Torigen.Fields[i].fieldname then
          Tdestino.Fieldbyname(Tdestino.Fields[j].FieldName).Value := Torigen.Fields[i].Value;
      Inc(j);
    End;
    Tdestino.post;
    Torigen.Next;
  end;
  Torigen.First;
  TDestino.First;
  Torigen.EnableControls;
  TDestino.EnableControls;
end;

procedure CopiaContenidoTablas( Torigen:TFDMemTable; TDestino:TuniTable; const borraDestino:Boolean); overload;
var
  i, j:Integer;
begin
  Torigen.DisableControls;
  TDestino.DisableControls;
  if Tdestino.Active = False then Tdestino.Open;
  if borraDestino then Tdestino.EmptyTable;
  if Torigen.Active = False then
    Torigen.Open;
  Torigen.First;
  while not Torigen.Eof do begin
    TDestino.Append;
    j := 0;
    for i:=0 to Torigen.fields.count-1 do begin
      if (Torigen.Fields[i].fieldname<>'Id') and (Tdestino.Fields[j].FieldName = 'Id') then
        j := i+1;
      if Tdestino.Fields[j].FieldName = Torigen.Fields[i].fieldname then
         Tdestino.Fieldbyname(Tdestino.Fields[j].FieldName).Value := Torigen.Fields[i].Value;
      Inc(j);
    End;
    Tdestino.Post;
    Torigen.Next;
  end;
  Torigen.First;      Torigen.EnableControls;
  TDestino.First;     TDestino.EnableControls;
end;

procedure ActualizaContenidoTabla( Torigen:TFDMemTable; TDestino:TuniTable);
var    // Personalizado para Captura trazas trimble, tabla resumen
  i, j:Integer;  vFiltro:String;
begin
  Torigen.DisableControls;
  TDestino.DisableControls;
  if Tdestino.Active = False then
     Tdestino.Open;
  Tdestino.First;
  if Torigen.Active = False then
     Torigen.Open;
  Torigen.First;
  vFiltro := LeftStr( Torigen.FieldByName('Anyo').AsString, 4);
  Tdestino.Filtered := False;
  Tdestino.Filter   := 'Anyo ='+vFiltro;
  Tdestino.Filtered := True;
  while not Torigen.Eof do begin
    if (vFiltro <> LeftStr( Torigen.FieldByName('Anyo').AsString, 4)) then begin
      vFiltro := LeftStr( Torigen.FieldByName('Anyo').AsString, 4);
      Tdestino.Filtered := False;
      Tdestino.Filter   := 'Anyo ='+vFiltro;
      Tdestino.Filtered := True;
    end;
    if TDestino.Locate('Matricula', Torigen.FieldByName('Matricula').AsString,[] )
      then TDestino.Edit
      else TDestino.Append;
    j := 0;
    for i:=0 to Torigen.fields.count-1 do begin
      if (Torigen.Fields[i].fieldname<>'Id') and (Tdestino.Fields[i].FieldName = 'Id') then
        j := i-1
      else
        if Tdestino.Fields[i].FieldName = Torigen.Fields[j].fieldname then begin
          if Tdestino.Fieldbyname(Tdestino.Fields[i].FieldName).IsNull
             or ((Torigen.Fields[j].DataType = ftInteger) and (Torigen.Fields[j].AsInteger >0) ) then
            Tdestino.Fieldbyname(Tdestino.Fields[i].FieldName).Value := Torigen.Fields[j].Value;
        end;
      Inc(j);
    End;
    Tdestino.Post;
    Torigen.Next;
  end;
  Torigen.First;
  TDestino.First;
  Torigen.EnableControls;
  TDestino.EnableControls;
end;

procedure AceleraTabla( pTabla:TFDMemTable);
begin
//  if pTabla.Active = False then pTabla.Active = True;
  pTabla.LogChanges := False;
  pTabla.FetchOptions.RecsMax := 10000;
  pTabla.ResourceOptions.SilentMode := True;
  pTabla.UpdateOptions.LockMode := lmNone;
  pTabla.UpdateOptions.LockPoint := lpDeferred;
  pTabla.UpdateOptions.FetchGeneratorsPoint := gpImmediate;
  pTabla.DisableConstraints;
  pTabla.DisableControls;
end;

function GeneraListaIdBorrarRESTdeMemTabla( pTMem:TFDMemTable):String;
begin
  result := '';
  pTMem.First;
  while (not pTMem.eof) do begin
    result := result + pTMem.FieldByName('Id').AsString +',';
    pTMem.Next;
  end;
  pTMem.EmptyDataSet;
  if Length( result)>1 then begin
    result := LeftStr(result, Length( result)-1);
  end;
end;
// JSON
procedure ExportaDatasetParaModificarREST(elDataset: TDataSet; var BodyI, BodyU:String);
var
  elRegistro, cambiado: string;
  elCampo: TField;
  i: integer;
begin
  if (not elDataset.Active) or (elDataset.IsEmpty) then
    Exit;
  BodyI := '[';
  BodyU := '[';
  elDataset.DisableControls;
  elDataset.First;
  while not elDataset.Eof do begin
    cambiado := elDataset.FieldByName('Cambiado').AsString;
    elregistro := '{"';
    if (cambiado ='I') or (cambiado ='U') then
      for i := 0 to elDataset.FieldCount - 1 do begin
        elCampo := elDataset.Fields[i];
        if (elCampo.FieldName<>'Cambiado') then
          if (elCampo.FieldName='Id') and (cambiado ='I') then
            elRegistro := elRegistro
          else begin
            if (elRegistro <> '{"') then
              elRegistro := elRegistro + ',"';
            elRegistro := elRegistro + elCampo.FieldName + '":';
            if (elCampo.DataType = ftObject) or (elCampo.Text='')  then
              elRegistro := elRegistro + 'null'
            else
              elRegistro := elRegistro + IIF(elCampo.DataType = ftString,
                                    '"' + elCampo.Text + '"', elCampo.Text);
          end;
        if i = elDataset.FieldCount - 1 then begin
          elRegistro := elRegistro + '}';
          if (cambiado ='I') then
            BodyI := IIF( BodyI = '[', BodyI + elRegistro, BodyI + ',' + elRegistro);
          if (cambiado ='U') then
            BodyU := IIF( BodyU = '[', BodyU + elRegistro, BodyU + ',' + elRegistro);
          elRegistro := '';
        end;
      end;
    elDataset.Next;
  end;
  elDataset.EnableControls;
  BodyI := BodyI + ']';
  BodyU := BodyU + ']';
end;

function ExportaDataSet_enJSON(elDataset: TDataSet): string;
var
  elRegistro: string;
  elCampo: TField;
  i: integer;
begin
  Result := '';
  if (not elDataset.Active) or (elDataset.IsEmpty)
    then Exit;
  Result := '[';
  elDataset.DisableControls;
  elDataset.First;
  while not elDataset.Eof do begin
    for i := 0 to elDataset.FieldCount - 1 do begin
      elCampo := elDataset.Fields[i];
      if elRegistro = ''
        then elRegistro := '{"' + elCampo.FieldName + '":"' + elCampo.Text + '"'
        else elRegistro := elRegistro + ',"' + elCampo.FieldName + '":"' + elCampo.Text + '"';
      if i = elDataset.FieldCount - 1 then begin
        elRegistro := elRegistro + '}';
        if Result = '['
          then Result := Result + elRegistro
          else Result := Result + ',' + elRegistro;
        elRegistro := '';
      end;
    end;
    elDataset.Next;
  end;
  elDataset.EnableControls;
  Result := Result + ']';
end;

function MontaUpdates(principioSQL, elBody:String):String;
var     //UPDATE omElementosDisponibles
//SET Obra_o_Montaje = 'M', Descripcion = 'PANELES dssssr', Id_tdbTiposGastos = 123
//WHERE Id = 3;
  vSQL, whereSQL:String;
  vJSON: TJDOJsonObject;
  i, j, NumeroObjetos:Integer;
begin
  elBody := LimpiarCadenaDeCaracter( elBody, '[');
  NumeroObjetos := NumVecesEsta_Subcadena_EnCadena( '}', elBody);
  whereSQL := ' WHERE Id = -1;';
  vSQL := '';
  for i:=1 to NumeroObjetos do begin
    vSQL := vSQL + principioSQL;
    vJSON := StrToJSONObject(elBody);
    for j := 0 to vJSON.Count-1 do begin
      if (UpperCase( vJSON.Names[j]) = 'ID') then begin
        whereSQL := ' WHERE Id = '+vJSON.S[vJSON.Names[j]]+';';
        break;// PONER EN TODOS LOS OBJETOS LOS CAMPOS CALCULADOS DESPUES DE ID
      end;
      vSQL := vSQL + vJSON.Names[j]+' = ';
      if (vJSON.Types[vJSON.Names[j]]= jdtObject) or (vJSON.S[vJSON.Names[j]]='') then
         vSql := vSQL+'null'
      else
         vSql := vSQL+IIF(vJSON.Types[vJSON.Names[j]]= jdtString, QuotedStr(vJSON.S[vJSON.Names[j]]), vJSON.S[vJSON.Names[j]]);
      if j<vJSON.Count-1 then vSql := vSQL+', ';
    end;
    if RightStr(vSQL, 2)=', ' then
      vSQL:= LeftStr(vSQL, Length(vSQL)-2);
    vSQL := vSQL+whereSQL;
    if i<NumeroObjetos then
      vSQL := vSQL+'; ';
    Delete(elbody, 1, Pos('}', elBody)+1);
  end;
  vJSON.Free;
  result := vSQL;
end;

function DeJSONrellenaValoresParaInsertSQL( pBody:String):String;
var
  vSQL:String;
  vJSON: TJDOJsonObject;
  i, j, vNumeroObjetos:Integer;
begin
  vSQL := ' ';    // vJSON.Create;
  pBody := LimpiarCadenaDeCaracter( pBody, '[');
  vNumeroObjetos := NumVecesEsta_Subcadena_EnCadena( '}', pBody);
  for i:=1 to vNumeroObjetos do begin
    vJSON := StrToJSONObject( pBody);
    vSql := vSQL+' (';
    for j := 0 to vJSON.Count-1 do
      if (UpperCase( vJSON.Names[j]) = 'ID') then // PONER EN TODOS LOS OBJETOS LOS CAMPOS CALCULADOS DESPUES DE ID
        break
      else begin
        if (vJSON.Types[vJSON.Names[j]]= jdtObject) or (vJSON.S[vJSON.Names[j]]='')
          then vSql := vSQL+'null'
          else vSql := vSQL+IIF(vJSON.Types[vJSON.Names[j]]= jdtString, QuotedStr(vJSON.S[vJSON.Names[j]]), vJSON.S[vJSON.Names[j]]);
        if j<vJSON.Count-1 then vSql := vSQL+', ';
      end;
    if RightStr(vSQL,2) =', ' then vSQL := LeftStr(vSQL, Length(vSQL)-2);
    vSQL := vSQL+')';
    if i < vNumeroObjetos then vSql := vSQL+', ';
    Delete( pbody, 1, Pos('}', pBody)+1);
  end;
  vJSON.Free;
  result := vSQL;
end;

// PASSWORD
function DesencriptaClave(const pClaveEncrip:String):String;
var
  vLetrasOri, vLetrasDest:String;
  i:Integer;
begin
  vLetrasOri  := 'L9A-Zza+Yb*rKs.Wd,oI8;_V&0Cl=){RyT[fXc]?¿iE(G$3HÑJ%1M|x4NOP}QSuU!F6g2jk#mn·ñBpe:qt@7hDv¡5w';
  vLetrasDest := 'I¡Bl4!e5ñH9AgV·Ck.FUmjb+Nq)Kw{G?P6x0:r7=(fhÑOQR]¿EDno@8J,;3L2M1Sz_y-v*u}$t[s|p%i&ZaYcXdW#T';
  result := '';
  for i := 1 to pClaveEncrip.Length do
     result := result + vLetrasOri[Pos( pClaveEncrip[i], vLetrasDest)];
end;

function EncriptaClave(const  pClave:String):String;
var  // Devuelve clave encriptada o mensaje con error "Error: el carácter X no se admite en el password."
  vLetrasOri, vLetrasDest:String;
  i, j:Integer;
begin
  vLetrasOri  := 'L9A-Zza+Yb*rKs.Wd,oI8;_V&0Cl=){RyT[fXc]?¿iE(G$3HÑJ%1M|x4NOP}QSuU!F6g2jk#mn·ñBpe:qt@7hDv¡5w';
  vLetrasDest := 'I¡Bl4!e5ñH9AgV·Ck.FUmjb+Nq)Kw{G?P6x0:r7=(fhÑOQR]¿EDno@8J,;3L2M1Sz_y-v*u}$t[s|p%i&ZaYcXdW#T';
  result := '';
  for i := 1 to pClave.Length do begin
     j := Pos(pClave[i], vLetrasOri);
     if j=-1 then begin
       result := 'Error: el carácter '+ pClave[i]+' no se admite en el password.';
       exit;
     end else
       result := result + vLetrasDest[j]
  end;
end;

function ExtraePwd (const pwdMezcla:String):String;
var           // 2022-08-A#·Virosque2022·#1
  i:Integer;
  pwd:String;
begin
  pwd := pwdMezcla;
  if Pos('·#', pwdMezcla)>0 then begin
    i := Length(pwdMezcla);
    while (pwdMezcla[i] <> '·') do
      i := i-1;
    pwd := Copy(pwdMezcla, 1, i-1);
  end;
  i := Pos('#·', pwd);
  if i>0 then
    pwd := Copy(pwd, i+2, Length(pwd));
  result := pwd;
end;

function ExtraeVersion (const pwdMezcla:String):String;
begin                   // 2022-08-A#·Virosque2022·#1
  result := IIF( Pos('#·', pwdMezcla)>0, Copy(pwdMezcla, 1, Pos('#·', pwdMezcla)-1), '');
end;

function ExtraeModulo (const pwdMezcla:String):Integer;
var                     // 2022-08-A#·Virosque2022·#1
  i:Integer;
begin
  result := 0;
  if Pos('·#',pwdMezcla)>0 then begin
    i := Length(pwdMezcla);
    while (pwdMezcla[i] <> '·') do
      i := i-1;
    try
      result := StrtoInt(Copy(pwdMezcla, i+2, Length(pwdMezcla)-i));
    except
    end;
  end;
end;

function EncryptStr(const pS :WideString; pKey: Word): String;
var   i    :Integer;
      vRStr :RawByteString;
      vRStrB:TBytes Absolute vRStr;
begin
  Result:= '';
  vRStr:= UTF8Encode(pS);
  for i := 0 to Length(vRStr)-1 do begin
    vRStrB[i] := vRStrB[i] xor (pKey shr 8);
    pKey      := (vRStrB[i] + pKey) * CKEY1 + CKEY2;
  end;
  for i := 0 to Length(vRStr)-1 do begin
    Result:= Result + IntToHex(vRStrB[i], 2);
  end;
end;

function DecryptStr(const pS: String; pKey: Word): String;
var   i, vTmpKey  :Integer;
      vRStr       :RawByteString;
      vRStrB      :TBytes Absolute vRStr;
      vTmpStr     :string;
begin
  i:= 1;          vTmpStr:= UpperCase(pS);
  SetLength(vRStr, Length(vTmpStr) div 2);
  try
    while (i < Length(vTmpStr)) do begin
      vRStrB[i div 2]:= StrToInt('$' + vTmpStr[i] + vTmpStr[i+1]);
      Inc(i, 2);
    end;
  except
    Result:= '';        Exit;
  end;
  for i := 0 to Length(vRStr)-1 do begin
    vTmpKey   := vRStrB[i];
    vRStrB[i] := vRStrB[i] xor (pKey shr 8);
    pKey      := (vTmpKey + pKey) * CKEY1 + CKEY2;
  end;
  Result:= UTF8Decode(vRStr);
end;

// TEXTO Y FECHA
   //CONVERT(DATETIME, '2020-05-15', 102)
function FechaToTextoFechaTrimble( const pFecha:TDateTime):String;
var      vYear, vMes, vDia, vHora, vMin, vSec, vMSec: Word;
begin   // fecha  -- > 2022-10-29T11:11:11.000
  DecodeTime(pFecha, vHora, vMin, vSec, vMSec);
  DecodeDate(pFecha, vYear, vMes, vDia);
  Result := floattostr(vyear)+'-'+RightStr('0'+floattostr(vmes),2)+'-'+RightStr('0'+floattostr(vDia),2)+'T'
    +RightStr('0'+floattostr(vHora),2)+':'+RightStr('0'+floattostr(vMin),2)+':'+ RightStr('0'+floattostr(vSec),2);
end;

function TDateTimetoSQLserverDate( const pFecha: TDateTime):string;
var             // 25/10/2022 14:28:30
  vYear, vMes, vDia, vHora, vMin, vSec, vMSec: Word;
begin
  DecodeTime(pFecha, vHora, vMin, vSec, vMSec);
  DecodeDate(pFecha, vYear, vMes, vDia);
  Result := RightStr('0'+floattostr(vDia),2)+'/'+RightStr('0'+floattostr(vmes),2)+'/'+ floattostr(vyear)+
  ' '+RightStr('0'+floattostr(vHora),2)+':'+RightStr('0'+floattostr(vMin),2)+':'+ RightStr('0'+floattostr(vSec),2);
end;

function TDateTimetoSQLdate(const pFecha:TDateTime):string;
var            // 2022-10-25 14:28:30
  vYear, vMes, vDia, vHora, vMin, vSec, vMSec: Word;
begin
  DecodeTime(pFecha, vHora, vMin, vSec, vMSec);
  DecodeDate(pFecha, vYear, vMes, vDia);
  Result := IntToStr(vYear)+'-'+RightStr('0'+IntToStr(vMes),2)+'-'+ RightStr('0'+IntToStr(vDia),2)+
        ' '+RightStr('0'+IntToStr(vHora),2)+':'+RightStr('0'+IntToStr(vMin),2)+':'+RightStr('0'+IntToStr(vSec),2);
end;

{function TDatetoSQLfecha(const pFecha:TDate):string;
var      vYear, vMes, vDia: Word;
begin  //  fecha tDate 14/01/2023 -->  2023-01-14
  DecodeDate(pFecha, vYear, vMes, vDia);
  Result := IntToStr(vYear)+'-'+RightStr('0'+IntToStr(vMes),2)
                           +'-'+RightStr('0'+IntToStr(vDia),2);
end;}

function FechaToSQLtexto(const pFecha:TDate):String;
var      vYear, vMes, vDia: Word;                  //
begin  //  fecha tDate 14/01/2023 -->  CONVERT(DATETIME, '2020-01-14', 102)
  DecodeDate(pFecha, vYear, vMes, vDia);
  Result := 'CONVERT(DATETIME, '+QuotedStr(IntToStr(vYear)+'-'+RightStr('0'
             +IntToStr(vMes),2)+'-'+RightStr('0'+IntToStr(vDia),2))+', 102)';
end;

function FechaHoraToSQLtexto(const pFecha:TDateTime):String;
var      vYear, vMes, vDia, vHora, vMin, vSec, vMSec: Word;   vTex:String;
begin  //  fecha tDate 14/01/2023 22:33:22-->  CONVERT(DATETIME, '2023-01-14 22:33:22', 102)
  DecodeDate(pFecha, vYear, vMes, vDia);
  DecodeTime(pFecha, vHora, vMin, vSec, vMSec);
  vTex := 'CONVERT(DATETIME, '+QuotedStr(IntToStr(vYear)+'-'+
      RightStr('0'+IntToStr(vMes),2)+'-'+RightStr('0'+IntToStr(vDia),2)+' '+
      RightStr('0'+IntToStr(vHora),2)+':'+RightStr('0'+IntToStr(vMin),2)+':'+
      RightStr('0'+IntToStr(vSec),2))+', 102)';
  Result := vTex;
end;

function TDatetoSQLdate( const pFecha:TDate):string;
var      vYear, vMes, vDia: Word;
begin  //  fecha tDate 14/01/2023 -->  20230114
  DecodeDate(pFecha, vYear, vMes, vDia);
  Result := IntToStr(vYear)+RightStr('0'+IntToStr(vMes),2)
                           +RightStr('0'+IntToStr(vDia),2);
end;

function TDateToStringDate(const pFecha:TDateTime):string;
var            // fecha  -->  25/10/2022
  vYear, vMes, vDia: Word;
begin
  DecodeDate(pFecha, vYear, vMes, vDia);
  Result := RightStr('0'+IntToStr(vDia),2)+'/'+RightStr('0'+IntToStr(vmes),2)+'/'+ IntToStr(vyear);
end;

function TDateTimeToStringTime(const pFecha:TDateTime):string;
var            // 14:28:30
  vHora, vMin, vSec, vMSec: Word;
begin
  DecodeTime(pFecha, vHora, vMin, vSec, vMSec);
  Result := RightStr('0'+IntToStr(vHora),2)+':'+RightStr('0'+IntToStr(vMin),2);
end;

function FechaToYearDia( const pFecha:Tdate):String;
begin    // fecha 3-2-2023  -->  2023034
  Result := IntToStr( YearOf(pFecha))+RightStr('00'+IntToStr( DayOfTheYear(pFecha)),3);
end;

function YearDiaToFecha(const pYear, pDiaYear: Word): TDateTime;
begin    //  2022,  034   -->   fecha 3-2-2023
  Result := EncodeDate(pYear, 1, 1) + pDiaYear - 1;
end;

function TDatetoPrincipioMes( const pFecha: TDateTime):string;
begin     // 25/10/2022 14:28:30  --> 01/10/2022
  result := '01/'+copy(TDateTimetoSQLserverDate(pFecha), 4, 7);
end;

function FechaTextoTrimbleToDateTime( const pFeTr:String):TDateTime;
begin   //  2022-10-29T11:11:11.000  -->  29/10/2022 11:11:11
  Result := StrToDateTime( Copy(pFeTr,9,2)+'/'+Copy(pFeTr,6,2)+'/'+Copy(pFeTr,1,4)+' '
                          +IIF(Length(pFeTr)>14, Copy(pFeTr,12,8), '00:00:00'));
end;

function FechaTextoTrimbleToDate( const pFeTr:String):TDate;
begin   //  2022-10-29  -->  29/10/2022
  Result := StrToDate( Copy(pFeTr,9,2)+'/'+Copy(pFeTr,6,2)+'/'+Copy(pFeTr,1,4));
end;

function FechaTextoToDate( const pFTxt:String):TDate;
begin   //  2022-10-29 o 29-10-2022 -->  29/10/2022
  try       Result := StrToDate( Copy(pFTxt,1,2)+'/'+Copy(pFTxt,4,2)+'/'+Copy(pFTxt,7,4));
  except
    try     Result := StrToDate( Copy(pFTxt,9,2)+'/'+Copy(pFTxt,6,2)+'/'+Copy(pFTxt,1,4));
    except  Result := StrToDate( '01/01/1900');
    end;
  end;
end;

Function SQLdate_to_Date(pFechaString : String) : TDate;
begin                            //   20230517   -> 17/05/2023
  result := EncodeDate(StrToInt(copy(pFechaString, 1, 4)), StrToInt(copy(pFechaString, 5, 2)), strToInt(copy(pFechaString, 7, 2)));
end;

Function STRING_to_DATA(pFechaString : String) : tDateTime;
var                     // 17/05/2023   ->  17/05/2023
  vc1, vc2, vc3 : String;
  vy, vm, vd : word;
begin
  if copy(pFechaString,1,1) >= '0' then begin
    vc1 := copy(pFechaString, 1, 2);
    vc2 := copy(pFechaString, 4, 2);
    vc3 := copy(pFechaString, 7, 4);
    vy := StrToInt(vc3);
    vm := StrToInt(vc2);
    vd := strToInt(vc1);
    result := EncodeDate(vy, vm, vd);
  end;
end;

Function STRING_to_DATETIME(pFechaString : String) : tDateTime;
var
  vc1, vc2, vc3, vc4, vc5, vc6 : String;
  vn4, vn5, vn6 : integer;
  vy, vm, vd : word;
begin
  if copy(pFechaString,1,1) >= '0' then begin
    vc1 := copy(pFechaString, 1, 2);         vc2 := copy(pFechaString, 4, 2);
    vc3 := copy(pFechaString, 7, 4);         vc4 := copy(pFechaString, 12, 2);
    vc5 := copy(pFechaString, 15, 2);        vc6 := copy(pFechaString, 18, 2);
    TryStrToInt( vc4, vn4);                  TryStrToInt(vc5, vn5);
    TryStrToInt( vc6, vn6);
    vy := StrToInt(vc3);                     vm := StrToInt( vc2);
    vd := strToInt(vc1);
    result := EncodeDateTime(vy, vm, vd, vn4,vn5, vn6,0);
  end;
end;




function FechayHoraToFechaHora( const pFecha, pHora:TDateTime):TDateTime;
begin   //2023-03-29 00:00 y 1899-12-31 11:11:11.000 --> 29/03/2023 11:11:11
  Result := StrToDateTime( Copy(DateTimeToStr(pFecha),1,11)+Copy(DateTimeToStr(pHora),12,8));
end;  // Une Fecha y Hora de Trans en una FechaHora

function LastDay(const pFecha: TDateTime): TDateTime;
var      viDia, viMes, viAno: word;
begin
    DecodeDate(pFecha, viAno, viMes, viDia);
    viMes := viMes + 1;       //Avanzamos un mes.
    if viMes = 13 then begin     //Si nos hemos pasado, avanzamos el año.
      viMes := 1;
      viAno := viAno + 1;
    end;    //Y devolvemos el día anterior al primer día del mes siguiente.
    result := EncodeDate(viAno, viMes, 1) - 1;
end;

function SumaMesTexto( const pFecha:String):string;
begin     // 25/10/2022 --> 25/11/2022
  result := TDateToStringDate( IncMonth( StrToDate(pFecha), 1));
end;

function PonMesIni(const pFecha:TDateTime):TDateTime;
var    // 07/05/2023 14:25:33 --> 01/05/2023 00:00:00
  vYear, vMes, vDia: Word;
begin
  DecodeDate(pFecha, vYear, vMes, vDia);
  result := StrToDateTime( '01/'+RightStr('0'+InttoStr(vMes),2)+'/'+InttoStr(vYear)+' 00:00:00.000');
end;

function PonMesFin(const pFecha:TDateTime):TDateTime; // 07/05/2023 14:25:33 --> 31/05/2023 23:59:59
begin
  result := PonMedianocheFin(IncDay( PonMesIni(IncMonth(pFecha,1)),-1));
end;

function PonMedianocheIni(const pFecha:TDateTime):TDateTime;
var      vFechaTex:String;    vYear, vMes, vDia, vHor, vMin, vSec, vMse: Word;
begin
{  vFechaTex := FechaToTextoFechaTrimble( pFecha);
  vFechaTex := LeftStr(vFechaTex,11)+'00:00:00.000';
  result    := FechaTextoTrimbleToDateTime( vFechaTex);}
  DecodeDateTime(pFecha, vYear, vMes, vDia, vHor, vMin, vSec, vMse);
  result := EncodeDateTime(vYear, vMes, vDia, 0, 0, 0, 0);
end;

function PonMedianocheFin(const pFecha:TDateTime):TDateTime;
var      vFechaTex:String;    vYear, vMes, vDia, vHor, vMin, vSec, vMse: Word;
begin
{  vFechaTex := FechaToTextoFechaTrimble( pFecha);
  vFechaTex := LeftStr(vFechaTex,11)+'23:59:59.999';
  result    := FechaTextoTrimbleToDateTime( vFechaTex);   IncMilliSecond()}
  DecodeDateTime(pFecha, vYear, vMes, vDia, vHor, vMin, vSec, vMse);
  result :=  EncodeDateTime(vYear, vMes, vDia, 23, 59, 59, 999);
end;

function EScapeString(const pCadena: string): string;
var      I: Integer;
begin
  Result := '';
  for I := 1 to Length(pCadena) do
    if pCadena[I] in [ '''', '\', '/','"', ';','&']
      then Result := Result + 'X' + pCadena[I]
      else Result := Result + pCadena[I];
end;

function obten_nombre_mes(const pNumMes : integer):string;
begin    result := MESES[pNumMes];                                end;

function LeeTextoDesdeFichero(const pRutaNombreFichero: string): string;
var  // Ejemplo uso:   MyTexto := LeeTextoDesdeFichero('C:\Temp\File.txt');
  SL: TStringList;
begin
  Result := '';
  if FileExists(pRutaNombreFichero) then
    try
      SL := TStringList.Create;
      SL.LoadFromFile( pRutaNombreFichero);
      Result := SL.Text;
    finally
      SL.Free;
    end;
end;

function hora_a_minutos(const pHora:string):string;
var      i, vResult: integer;   // Convierte string Hora:Min en string min
         vTmp:string;           // Ejemplo 03:25 o 3:25  -->  205
begin
  i := Pos(':',pHora);
  if i>1 then begin
    vTmp := Copy(pHora,1,i-1);
    vResult := StrToIntDef(vTmp, 0)*60;
    inc(i);
    vTmp := Copy(pHora,i,2);
    vResult := vResult + StrToIntDef(vTmp, 0);
    Result  := IntToStr(vResult);
  end else result := '';

end;

function devuelve_horas(const pHora:string):real;
var  // Si me pasan 13:30 (13 horas y 30 minutos), debe devolver 13,5
  i, vHor, vMin: integer;
begin
  i := Pos(':',pHora);
  result := 0;
  if i>1 then begin
    try
      vHor := StrToInt(Copy(pHora, 1,i-1));
      vMin := strtoint(Copy(pHora, i+1,length(pHora)-i+1));
    except exit;
    end;
    result :=  vHor+( vMin /60.0);
  end;
end;

function IndiceArray(pLista: array of String; const pDatoAbuscar: String): Integer;
var i:Integer;
begin
  result := 0;
  for i := 0 to Length( pLista)-1 do
    if pLista[i] = pDatoAbuscar then begin
      result := i;
      exit;
    end;
end;

function NumVecesEsta_Subcadena_EnCadena( const pSubcadena, pCadena:string):Integer;
var
  veces, i : Integer;
begin
  veces := 0;
  if pCadena = pSubcadena
  then veces := 1
  else
    for i:=1 to length(pCadena)-length(pSubcadena)+1 do
        if (Copy(pCadena,i,length(pSubcadena)) = pSubcadena) then inc(veces);
  result := veces;
end;

function LimpiarCadenaDeCaracter(const pCadena:string; const pCaracter: Char): string;
var
  PosCar: Integer;
begin
  Result := pCadena;
  repeat
    PosCar := Pos(pCaracter, Result);
    if PosCar <> 0 then
      Delete(Result, PosCar, 1);
  until PosCar = 0;
end;

function LimpiarCadenaSoloLetrasNumeros(const vCadena: string): string;
var      I: Integer;
begin
  Result := '';
  for I := 1 to Length(vCadena) do
    if vCadena[i] in ['a'..'z', '0'..'9', 'A'..'Z'] then
      Result := Result + vCadena[i];
end;

function LimpiarCadenaSoloNumeros(const vCadena: string): string;
var      I: Integer;
begin
  Result := '';
  for I := 1 to Length(vCadena) do
    if vCadena[i] in ['0'..'9'] then
      Result := Result + vCadena[i];
end;

function CambiaEn(const pCadena, pOri, pDes:String):String;
begin    Result := ReplaceStr(pCadena, pOri, pDes);                 end;

function FloatToSQLnum(const pNum:Double):String; // 23,5 ---> "23.5"
begin    Result := ReplaceStr(FloatToStr(pNum), ',', '.');          end;

function PosNoNum(const pTexto:String):Integer;
var      i: Integer; // Posición del primer carácter no numérico de un string
begin
  for I := 1 to Length(pTexto) do
    if not IsNumber(pTexto[i]) then begin
      result := i;
      break;
    end;
end;

function DaRutaDelArchivo(pArchivo:String):String;
var      i: Integer; // Da Ruta de archivo   c:\proy\cod\gest\hola.exe --> c:\proy\cod\gest\
begin
  for I := Length(pArchivo) downto 1 do
    if pArchivo[i]='\' then break;
  result := LeftStr(pArchivo, i);
end;


function PalabraAleatoria(const pLongitud: integer): string;
const  Letras = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var    i : integer;
begin
    Result:='';
    for i:=1 to pLongitud do Result:=Result+Letras[1+Random(Length(Letras))];
end;

function DameNumDeTexto(const pText: string): Integer;
var
  vRegex: TRegEx;   vMatch: TMatch;
begin             // Define la expresión regular para buscar números en el texto
  vRegex := TRegEx.Create('\d+');
  // Busca la primera coincidencia de la expresión regular en el texto
  vMatch := vRegex.Match(pText);
  if vMatch.Success then Result := StrToInt(vMatch.Value)
                    else Result := 0;
end;


function gradosARadianes(const pGrados:Double):Double;
begin
  result := pGrados * PI / 180.0;
end;

function calcularDistanciaEntreDosCoordenadas( pLat1, pLat2, pLong1, pLong2:Double):Integer;
var
  vDifLat, vDifLong, vTemp, vResult:Double;
begin
  pLat1   := gradosARadianes(pLat1);      pLong1   := gradosARadianes(pLong1);
  pLat2   := gradosARadianes(pLat2);      pLong2   := gradosARadianes(pLong2);
  vDifLat := pLat2 - pLat1;               vDifLong := pLong2 - pLong1;
  vTemp   := Power(sin(vDifLat / 2.0), 2)+ cos(pLat1)* cos(pLat2)* Power(sin(vDifLong / 2.0), 2);
  vResult := 2.0 * ArcTan2(sqrt(vTemp), sqrt(1.0 - vTemp));
  result  := Round( RADIO_TIERRA_EN_KILOMETROS * vResult);
end;

function BuscarEntero(const pLista: array of Integer; const pEnteroBuscado: Integer): Boolean;
var      i: Integer;
begin  // Recorrer la lista y verificar si el entero buscado se encuentra en ella
  result := False;
  for i := 0 to Length(pLista) - 1 do
    if pLista[i] = pEnteroBuscado then  begin
      Result := True;
      Exit;
    end;
end;

function TStringListToString(const pLis: TStringList): String;
var      i: Integer;
begin  // Convierte un TStringList en un String con elementos entre comillas y separados por coma
  result := '';
  for i := 0 to pLis.Count - 1 do
    result := result +QuotedStr(pLis[i])+', ';
  result := LeftStr(result, Length(result)-2);
end;

function TStringListToStringNum(const pLis: TStringList): String;
var      i: Integer;
begin  // Convierte un TStringList en un String con elementos separados por coma
  result := '';
  for i := 0 to pLis.Count - 1 do
    result := result +pLis[i]+', ';
  result := LeftStr(result, Length(result)-2);
end;

function BuscaHoraEnString(const pTex: String): String;
var      i, vPos: Integer;  vHora, vTex:String;
begin  // Busca hora en String
  result := '';  vHora :=''; vTex := pTex;
  for i:= 1 to NumVecesEsta_Subcadena_EnCadena(':',pTex) do begin
    vPos := Pos(':',vTex);
    if vPos>2 then
      if Length(vTex)>(vPos+1) then 
       if IsNumber(vTex[vPos-2]) AND IsNumber(vTex[vPos-1]) AND
          IsNumber(vTex[vPos+1]) AND IsNumber(vTex[vPos+2]) then
            try
              vHora := Copy(vTex, vPos-2, 5);
              if Length(vTex)>(vPos+4) then
                if (vTex[vPos+3]=':') AND IsNumber(vTex[vPos+4]) AND IsNumber(vTex[vPos+5]) 
                  then vHora := vHora + Copy(vTex, vPos+3, 3); 
              if Length(vHora)=5 then vHora := vHora +':00';
              result := vHora;
              exit;
            except  
            end;
    vTex := RightStr(vTex, Length(vTex)-vPos);        
  end;
end;

// API WINDOWS
function GetCurrentUserName(out DomainName, UserName: string): Boolean;
var       // Devuelve el nombre de la máquina y del usuario que ejecuta.
  Token: THandle;
  InfoSize, UserNameSize, DomainNameSize: Cardinal;
  User: PTokenUser;
  Use: SID_NAME_USE;
  _DomainName, _UserName: array[0..255] of Char;
begin
  Result := False;
  DomainName := '';
  UserName := '';
  Token := 0;
  if not OpenThreadToken(GetCurrentThread, TOKEN_QUERY, True, Token) then begin
    if GetLastError = ERROR_NO_TOKEN then begin// current thread is not impersonating, try process token
      if not OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, Token) then Exit;
    end else Exit;
  end;
  try
    GetTokenInformation(Token, TokenUser, nil, 0, InfoSize);
    User := AllocMem(InfoSize * 2);
    try
      if GetTokenInformation(Token, TokenUser, User, InfoSize * 2, InfoSize) then begin
        DomainNameSize := SizeOf(_DomainName);
        UserNameSize := SizeOf(_UserName);
        Result := LookupAccountSid(nil, User^.User.Sid, _UserName, UserNameSize, _DomainName, DomainNameSize, Use);
        if Result then begin
          SetString(DomainName, _DomainName, StrLen(_DomainName));
          SetString(UserName, _UserName, StrLen(_UserName));
        end;
      end;
    finally  FreeMem(User);
    end;
  finally    CloseHandle(Token);
  end;
end;

procedure CierraMiAplicacion( const NombreAplicacion:PWideChar);
var
  h: HWND;
begin
  h := FindWindow(nil, NombreAplicacion);
  if h <> 0 then
    PostMessage(h, WM_CLOSE, 0, 0);
end;

procedure CopiaPos_y_Tam( pTeditOrig:Tedit; pTcomboDest:TCombobox);
begin
  pTcomboDest.Top    := pTeditOrig.Top;
  pTcomboDest.Left   := pTeditOrig.Left;
  pTcomboDest.Height := pTeditOrig.Height;
  pTcomboDest.Width  := pTeditOrig.Width;
end;

procedure CambiaENTERporTABenKeyPress(var Key:PChar);
 begin
{   if Key = #13 then begin
       Key := #0;
       Perform(WM_NEXTDLGCTL, 0, 0);
   end}
 end;

procedure SetCursorIni(Cursor: TCursor = crHourGlass);
begin     Screen.Cursor := Cursor;                               end;

procedure SetCursorFin;
begin     Screen.Cursor := crDefault;                            end;

function  GeneraNombreArchivoTemporal( const pLargo:Integer; const pExtens:String):string;
begin
  SetCursorIni;
  Application.ProcessMessages;
  if DirectoryExists('c:\temp') then result := 'c:\temp'
                                else result := GetEnvironmentVariable('TEMP');
  result := result+ '\'+PalabraAleatoria(pLargo)+'.'+pExtens;
  SetCursorFin;
end;

// IIF

function IIF(esverdad: boolean; vVerdad, vFalso: integer): integer; overload;
begin  { IF inmediato para enteros }
  if esverdad then Result := vVerdad
              else Result := vFalso;
end;

function IIF(esverdad: boolean; vVerdad, vFalso: Extended): Extended; overload;
begin   { IF inmediato para floats }
  if esverdad then Result := vVerdad
              else Result := vFalso;
end;

function IIF(esverdad: boolean; vVerdad, vFalso: string): string; overload;
begin   { IF inmediato para cadenas }
  if esverdad then Result := vVerdad
              else Result := vFalso;
end;

function IIF(esverdad: boolean; vVerdad, vFalso: TObject): TObject; overload;
begin   { IF inmediato para objetos }
  if esverdad then Result := vVerdad
              else Result := vFalso;
end;

Function IIF( esverdad:boolean; vVerdad,vFalso:variant):variant; overload;
begin
 if esverdad then Result := vVerdad
             else Result := vFalso;
end;

// SQL
function LeeSQL( textoSQL:String; var objQuery:TADOQuery):integer; overload;
begin
  if objQuery.Active then objQuery.Active := False;
  objQuery.SQL.Text := textoSQL;
  try       objQuery.Active := True;
            objQuery.First;
  except    on Exception do
  end;
  result := objQuery.RecordCount;
end;

function LeeSQL( textoSQL:String; var objQuery:TVirtualQuery):integer; overload;
begin
  if objQuery.Active then objQuery.Active := False;
  objQuery.SQL.Text := textoSQL;
  try       objQuery.Active := True;
            objQuery.First;
  except    on Exception do
  end;
  result := objQuery.RecordCount;
end;

function LeeSQL( textoSQL:String; var objQuery:TUniQuery):integer; overload;
begin
  if objQuery.Active then objQuery.Active := False;
  objQuery.SQL.Text := textoSQL;
  try       objQuery.Active := True;
            objQuery.First;
  except    on Exception do
  end;
  result := objQuery.RecordCount;
end;

procedure EscribeSQL( textoSQL:String; objQuery:TADOQuery); overload;
begin
  if objQuery.Active then objQuery.Active := False;
  objQuery.SQL.Text := textoSQL;
  try       objQuery.ExecSQL;
  except    on Exception do
  end;
  objQuery.Close;
end;

procedure EscribeSQL( textoSQL:String; objQuery:TVirtualQuery); overload;
begin
  if objQuery.Active then objQuery.Active := False;
  objQuery.SQL.Text := textoSQL;
  try       objQuery.ExecSQL;
  except    on Exception do
  end;
  objQuery.Close;
end;

procedure EscribeSQL( textoSQL:String; objQuery:TUniQuery); overload;
begin
  if objQuery.Active then objQuery.Active := False;
  objQuery.SQL.Text := textoSQL;
  try       objQuery.ExecSQL;
  except    on Exception do
  end;
  objQuery.Close;
end;


end.
