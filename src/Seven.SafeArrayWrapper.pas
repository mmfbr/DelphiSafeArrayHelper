unit Seven.SafeArrayWrapper;

{$IFDEF FPC}
  {$MODE Delphi}
{$ENDIF}

interface

uses
  {$IFDEF MSWINDOWS}
  Windows, ActiveX,
  {$ENDIF}
  SysUtils, Variants, ComObj;

type
  /// <summary>
  ///   Interface que encapsula e simplifica o uso de SafeArray COM
  /// </summary>
  /// <remarks>
  ///   ISafeArrayWrapper fornece uma API amigável e type-safe para trabalhar
  ///   com SafeArrays, especialmente útil para interoperabilidade COM/.NET.
  ///   O gerenciamento de memória é automático através de reference counting.
  /// </remarks>
  ISafeArrayWrapper = interface
    ['{8F4B3C2A-1E5D-4A8B-9C7E-2F3A4B5C6D7E}']
    
    /// <summary>
    ///   Retorna o número total de elementos no array
    /// </summary>
    /// <returns>Total de elementos em todas as dimensões</returns>
    function GetCount: Integer;
    
    /// <summary>
    ///   Retorna o número de dimensões do array
    /// </summary>
    /// <returns>Número de dimensões (1 para array unidimensional, 2 para bidimensional, etc)</returns>
    function GetDimensions: Integer;
    
    /// <summary>
    ///   Retorna o limite inferior de uma dimensão específica
    /// </summary>
    /// <param name="ADimension">Dimensão a consultar (1-based)</param>
    /// <returns>Índice do limite inferior</returns>
    function GetLBound(ADimension: Integer = 1): Integer;
    
    /// <summary>
    ///   Retorna o limite superior de uma dimensão específica
    /// </summary>
    /// <param name="ADimension">Dimensão a consultar (1-based)</param>
    /// <returns>Índice do limite superior</returns>
    function GetUBound(ADimension: Integer = 1): Integer;
    
    /// <summary>
    ///   Retorna o tamanho em bytes de cada elemento
    /// </summary>
    /// <returns>Tamanho em bytes de um elemento</returns>
    function GetElementSize: Integer;
    
    /// <summary>
    ///   Retorna o tipo Variant dos elementos do array
    /// </summary>
    /// <returns>TVarType indicando o tipo dos elementos</returns>
    function GetVarType: TVarType;
    
    /// <summary>
    ///   Retorna o ponteiro PSafeArray subjacente
    /// </summary>
    /// <returns>Ponteiro para a estrutura SafeArray nativa</returns>
    function GetSafeArray: PSafeArray;
    
    /// <summary>
    ///   Obtém o valor de um elemento em array unidimensional
    /// </summary>
    /// <param name="AIndex">Índice do elemento</param>
    /// <returns>Valor do elemento como Variant</returns>
    function GetItem(AIndex: Integer): Variant; overload;
    
    /// <summary>
    ///   Obtém o valor de um elemento em array multidimensional
    /// </summary>
    /// <param name="AIndices">Array com os índices para cada dimensão</param>
    /// <returns>Valor do elemento como Variant</returns>
    function GetItem(const AIndices: array of Integer): Variant; overload;
    
    /// <summary>
    ///   Define o valor de um elemento em array unidimensional
    /// </summary>
    /// <param name="AIndex">Índice do elemento</param>
    /// <param name="AValue">Valor a ser definido</param>
    procedure SetItem(AIndex: Integer; const AValue: Variant); overload;
    
    /// <summary>
    ///   Define o valor de um elemento em array multidimensional
    /// </summary>
    /// <param name="AIndices">Array com os índices para cada dimensão</param>
    /// <param name="AValue">Valor a ser definido</param>
    procedure SetItem(const AIndices: array of Integer; const AValue: Variant); overload;
    
    /// <summary>
    ///   Limpa todos os elementos do array (define como Null)
    /// </summary>
    /// <remarks>Disponível apenas para arrays unidimensionais</remarks>
    procedure Clear;
    
    /// <summary>
    ///   Redimensiona o array
    /// </summary>
    /// <param name="ANewSize">Novo tamanho do array</param>
    /// <remarks>Disponível apenas para arrays unidimensionais</remarks>
    procedure Resize(ANewSize: Integer);
    
    /// <summary>
    ///   Adiciona um elemento ao final do array
    /// </summary>
    /// <param name="AValue">Valor a ser adicionado</param>
    /// <remarks>Disponível apenas para arrays unidimensionais</remarks>
    procedure Append(const AValue: Variant);
    
    /// <summary>
    ///   Converte o SafeArray para um Variant array
    /// </summary>
    /// <returns>Variant contendo uma cópia do array</returns>
    function ToVariantArray: Variant;
    
    /// <summary>
    ///   Converte o SafeArray para um array nativo de strings
    /// </summary>
    /// <returns>TArray&lt;string&gt; com os elementos convertidos para string</returns>
    /// <remarks>Disponível apenas para arrays unidimensionais</remarks>
    function ToStringArray: TArray<string>;
    
    /// <summary>
    ///   Converte o SafeArray para um array nativo de inteiros
    /// </summary>
    /// <returns>TArray&lt;Integer&gt; com os elementos</returns>
    /// <remarks>Disponível apenas para arrays unidimensionais</remarks>
    function ToIntegerArray: TArray<Integer>;
    
    /// <summary>
    ///   Converte o SafeArray para um array nativo de doubles
    /// </summary>
    /// <returns>TArray&lt;Double&gt; com os elementos</returns>
    /// <remarks>Disponível apenas para arrays unidimensionais</remarks>
    function ToDoubleArray: TArray<Double>;
    
    // Propriedades
    /// <summary>Número total de elementos</summary>
    property Count: Integer read GetCount;
    /// <summary>Número de dimensões do array</summary>
    property Dimensions: Integer read GetDimensions;
    /// <summary>Acesso indexado aos elementos (propriedade default)</summary>
    property Items[AIndex: Integer]: Variant read GetItem write SetItem; default;
    /// <summary>Tipo Variant dos elementos</summary>
    property VarType: TVarType read GetVarType;
    /// <summary>Ponteiro para o SafeArray nativo</summary>
    property SafeArray: PSafeArray read GetSafeArray;
  end;
  
  /// <summary>
  ///   Implementação concreta da interface ISafeArrayWrapper
  /// </summary>
  TSafeArrayWrapper = class(TInterfacedObject, ISafeArrayWrapper)
  private
    FSafeArray: PSafeArray;
    FOwnsData: Boolean;
    FVarType: TVarType;
    FDimensions: Integer;
    
    /// <summary>Valida se um índice está dentro dos limites válidos</summary>
    procedure CheckIndex(AIndex: Integer);
    /// <summary>Valida se os índices estão dentro dos limites válidos para arrays multidimensionais</summary>
    procedure CheckIndices(const AIndices: array of Integer);
    /// <summary>Garante que o SafeArray não seja nil</summary>
    procedure EnsureNotNil;
  public
    /// <summary>
    ///   Cria um novo SafeArray unidimensional
    /// </summary>
    /// <param name="AVarType">Tipo dos elementos (varInteger, varDouble, etc)</param>
    /// <param name="ALBound">Limite inferior do array</param>
    /// <param name="AUBound">Limite superior do array</param>
    constructor Create(AVarType: TVarType; ALBound, AUBound: Integer); overload;
    
    /// <summary>
    ///   Cria um novo SafeArray multidimensional
    /// </summary>
    /// <param name="AVarType">Tipo dos elementos</param>
    /// <param name="ABounds">Array com pares de (LBound, UBound) para cada dimensão</param>
    constructor Create(AVarType: TVarType; const ABounds: array of Integer); overload;
    
    /// <summary>
    ///   Encapsula um SafeArray existente
    /// </summary>
    /// <param name="ASafeArray">Ponteiro para o SafeArray existente</param>
    /// <param name="AOwnsData">Se True, o wrapper destruirá o SafeArray ao ser liberado</param>
    constructor Create(ASafeArray: PSafeArray; AOwnsData: Boolean = False); overload;
    
    /// <summary>
    ///   Cria um wrapper a partir de um Variant array
    /// </summary>
    /// <param name="AVariant">Variant contendo um array</param>
    constructor CreateFromVariant(const AVariant: Variant); overload;
    
    /// <summary>
    ///   Destrutor - libera o SafeArray se FOwnsData for True
    /// </summary>
    destructor Destroy; override;
    
    // ISafeArrayWrapper implementation
    function GetCount: Integer;
    function GetDimensions: Integer;
    function GetLBound(ADimension: Integer = 1): Integer;
    function GetUBound(ADimension: Integer = 1): Integer;
    function GetElementSize: Integer;
    function GetVarType: TVarType;
    function GetSafeArray: PSafeArray;
    
    function GetItem(AIndex: Integer): Variant; overload;
    function GetItem(const AIndices: array of Integer): Variant; overload;
    procedure SetItem(AIndex: Integer; const AValue: Variant); overload;
    procedure SetItem(const AIndices: array of Integer; const AValue: Variant); overload;
    
    procedure Clear;
    procedure Resize(ANewSize: Integer);
    procedure Append(const AValue: Variant);
    function ToVariantArray: Variant;
    function ToStringArray: TArray<string>;
    function ToIntegerArray: TArray<Integer>;
    function ToDoubleArray: TArray<Double>;
  end;
  
  // Funções auxiliares
  
  /// <summary>
  ///   Cria um novo SafeArray unidimensional com índices de 0 a ACount-1
  /// </summary>
  /// <param name="AVarType">Tipo dos elementos</param>
  /// <param name="ACount">Número de elementos</param>
  /// <returns>Interface ISafeArrayWrapper para o novo array</returns>
  function CreateSafeArray(AVarType: TVarType; ACount: Integer): ISafeArrayWrapper; overload;
  
  /// <summary>
  ///   Cria um novo SafeArray multidimensional
  /// </summary>
  /// <param name="AVarType">Tipo dos elementos</param>
  /// <param name="ABounds">Array com pares de (LBound, UBound) para cada dimensão</param>
  /// <returns>Interface ISafeArrayWrapper para o novo array</returns>
  function CreateSafeArray(AVarType: TVarType; const ABounds: array of Integer): ISafeArrayWrapper; overload;
  
  /// <summary>
  ///   Encapsula um SafeArray existente
  /// </summary>
  /// <param name="ASafeArray">Ponteiro para o SafeArray</param>
  /// <param name="AOwnsData">Se True, o wrapper destruirá o SafeArray</param>
  /// <returns>Interface ISafeArrayWrapper</returns>
  function WrapSafeArray(ASafeArray: PSafeArray; AOwnsData: Boolean = False): ISafeArrayWrapper;
  
  /// <summary>
  ///   Converte um Variant array em ISafeArrayWrapper
  /// </summary>
  /// <param name="AVariant">Variant contendo um array</param>
  /// <returns>Interface ISafeArrayWrapper</returns>
  function VariantToSafeArray(const AVariant: Variant): ISafeArrayWrapper;

implementation

uses
  VarUtils;

{ TSafeArrayWrapper }

constructor TSafeArrayWrapper.Create(AVarType: TVarType; ALBound, AUBound: Integer);
var
  Bounds: TSafeArrayBound;
begin
  inherited Create;
  FVarType := AVarType;
  FDimensions := 1;
  FOwnsData := True;
  
  Bounds.lLbound := ALBound;
  Bounds.cElements := AUBound - ALBound + 1;
  
  FSafeArray := SafeArrayCreate(AVarType, 1, @Bounds);
  if FSafeArray = nil then
    raise EOleError.Create('Falha ao criar SafeArray');
end;

constructor TSafeArrayWrapper.Create(AVarType: TVarType; const ABounds: array of Integer);
var
  I: Integer;
  Bounds: array of TSafeArrayBound;
begin
  inherited Create;
  FVarType := AVarType;
  FDimensions := Length(ABounds) div 2;
  FOwnsData := True;
  
  if (Length(ABounds) mod 2) <> 0 then
    raise EArgumentException.Create('ABounds deve conter pares de (LBound, UBound)');
  
  SetLength(Bounds, FDimensions);
  for I := 0 to FDimensions - 1 do
  begin
    Bounds[I].lLbound := ABounds[I * 2];
    Bounds[I].cElements := ABounds[I * 2 + 1] - ABounds[I * 2] + 1;
  end;
  
  FSafeArray := SafeArrayCreate(AVarType, FDimensions, @Bounds[0]);
  if FSafeArray = nil then
    raise EOleError.Create('Falha ao criar SafeArray');
end;

constructor TSafeArrayWrapper.Create(ASafeArray: PSafeArray; AOwnsData: Boolean);
begin
  inherited Create;
  FSafeArray := ASafeArray;
  FOwnsData := AOwnsData;
  
  if FSafeArray <> nil then
  begin
    FDimensions := SafeArrayGetDim(FSafeArray);
    SafeArrayGetVartype(FSafeArray, FVarType);
  end;
end;

constructor TSafeArrayWrapper.CreateFromVariant(const AVariant: Variant);
begin
  inherited Create;
  FOwnsData := False;
  
  if not VarIsArray(AVariant) then
    raise EArgumentException.Create('Variant não é um array');
  
  FSafeArray := VarArrayAsPSafeArray(AVariant);
  FDimensions := SafeArrayGetDim(FSafeArray);
  SafeArrayGetVartype(FSafeArray, FVarType);
end;

destructor TSafeArrayWrapper.Destroy;
begin
  if FOwnsData and (FSafeArray <> nil) then
    SafeArrayDestroy(FSafeArray);
  inherited;
end;

procedure TSafeArrayWrapper.CheckIndex(AIndex: Integer);
var
  LBound, UBound: Integer;
begin
  EnsureNotNil;
  SafeArrayGetLBound(FSafeArray, 1, LBound);
  SafeArrayGetUBound(FSafeArray, 1, UBound);
  
  if (AIndex < LBound) or (AIndex > UBound) then
    raise ERangeError.CreateFmt('Índice %d fora do intervalo [%d..%d]', [AIndex, LBound, UBound]);
end;

procedure TSafeArrayWrapper.CheckIndices(const AIndices: array of Integer);
var
  I, LBound, UBound: Integer;
begin
  EnsureNotNil;
  
  if Length(AIndices) <> FDimensions then
    raise EArgumentException.CreateFmt('Número de índices (%d) não corresponde às dimensões (%d)', 
      [Length(AIndices), FDimensions]);
  
  for I := 0 to High(AIndices) do
  begin
    SafeArrayGetLBound(FSafeArray, I + 1, LBound);
    SafeArrayGetUBound(FSafeArray, I + 1, UBound);
    
    if (AIndices[I] < LBound) or (AIndices[I] > UBound) then
      raise ERangeError.CreateFmt('Índice %d na dimensão %d fora do intervalo [%d..%d]', 
        [AIndices[I], I + 1, LBound, UBound]);
  end;
end;

procedure TSafeArrayWrapper.EnsureNotNil;
begin
  if FSafeArray = nil then
    raise EOleError.Create('SafeArray não inicializado');
end;

function TSafeArrayWrapper.GetCount: Integer;
var
  I, LBound, UBound: Integer;
begin
  Result := 1;
  EnsureNotNil;
  
  for I := 1 to FDimensions do
  begin
    SafeArrayGetLBound(FSafeArray, I, LBound);
    SafeArrayGetUBound(FSafeArray, I, UBound);
    Result := Result * (UBound - LBound + 1);
  end;
end;

function TSafeArrayWrapper.GetDimensions: Integer;
begin
  Result := FDimensions;
end;

function TSafeArrayWrapper.GetLBound(ADimension: Integer): Integer;
begin
  EnsureNotNil;
  SafeArrayGetLBound(FSafeArray, ADimension, Result);
end;

function TSafeArrayWrapper.GetUBound(ADimension: Integer): Integer;
begin
  EnsureNotNil;
  SafeArrayGetUBound(FSafeArray, ADimension, Result);
end;

function TSafeArrayWrapper.GetElementSize: Integer;
begin
  EnsureNotNil;
  Result := SafeArrayGetElemsize(FSafeArray);
end;

function TSafeArrayWrapper.GetVarType: TVarType;
begin
  Result := FVarType;
end;

function TSafeArrayWrapper.GetSafeArray: PSafeArray;
begin
  Result := FSafeArray;
end;

function TSafeArrayWrapper.GetItem(AIndex: Integer): Variant;
var
  HR: HRESULT;
begin
  CheckIndex(AIndex);
  VarClear(Result);
  
  case FVarType of
    varSmallint:
      begin
        var Value: SmallInt;
        HR := SafeArrayGetElement(FSafeArray, @AIndex, @Value);
        OleCheck(HR);
        Result := Value;
      end;
    varInteger:
      begin
        var Value: Integer;
        HR := SafeArrayGetElement(FSafeArray, @AIndex, @Value);
        OleCheck(HR);
        Result := Value;
      end;
    varSingle:
      begin
        var Value: Single;
        HR := SafeArrayGetElement(FSafeArray, @AIndex, @Value);
        OleCheck(HR);
        Result := Value;
      end;
    varDouble:
      begin
        var Value: Double;
        HR := SafeArrayGetElement(FSafeArray, @AIndex, @Value);
        OleCheck(HR);
        Result := Value;
      end;
    varOleStr:
      begin
        var Value: PWideChar;
        HR := SafeArrayGetElement(FSafeArray, @AIndex, @Value);
        OleCheck(HR);
        Result := WideString(Value);
      end;
    varDispatch:
      begin
        var Value: IDispatch;
        HR := SafeArrayGetElement(FSafeArray, @AIndex, @Value);
        OleCheck(HR);
        Result := Value;
      end;
    varVariant:
      begin
        HR := SafeArrayGetElement(FSafeArray, @AIndex, @TVarData(Result));
        OleCheck(HR);
      end;
  else
    raise EOleError.CreateFmt('Tipo não suportado: %d', [FVarType]);
  end;
end;

function TSafeArrayWrapper.GetItem(const AIndices: array of Integer): Variant;
var
  HR: HRESULT;
  Indices: array of Integer;
  I: Integer;
begin
  CheckIndices(AIndices);
  VarClear(Result);
  
  SetLength(Indices, Length(AIndices));
  for I := 0 to High(AIndices) do
    Indices[I] := AIndices[I];
  
  case FVarType of
    varVariant:
      begin
        HR := SafeArrayGetElement(FSafeArray, @Indices[0], @TVarData(Result));
        OleCheck(HR);
      end;
  else
    raise EOleError.Create('GetItem multidimensional implementado apenas para varVariant');
  end;
end;

procedure TSafeArrayWrapper.SetItem(AIndex: Integer; const AValue: Variant);
var
  HR: HRESULT;
  Value: Variant;
begin
  CheckIndex(AIndex);
  Value := AValue;
  
  case FVarType of
    varSmallint:
      begin
        var IntValue: SmallInt := Value;
        HR := SafeArrayPutElement(FSafeArray, @AIndex, @IntValue);
      end;
    varInteger:
      begin
        var IntValue: Integer := Value;
        HR := SafeArrayPutElement(FSafeArray, @AIndex, @IntValue);
      end;
    varSingle:
      begin
        var FloatValue: Single := Value;
        HR := SafeArrayPutElement(FSafeArray, @AIndex, @FloatValue);
      end;
    varDouble:
      begin
        var FloatValue: Double := Value;
        HR := SafeArrayPutElement(FSafeArray, @AIndex, @FloatValue);
      end;
    varOleStr:
      begin
        var StrValue: WideString := Value;
        HR := SafeArrayPutElement(FSafeArray, @AIndex, PWideChar(StrValue));
      end;
    varDispatch:
      begin
        var DispValue: IDispatch := Value;
        HR := SafeArrayPutElement(FSafeArray, @AIndex, @DispValue);
      end;
    varVariant:
      begin
        HR := SafeArrayPutElement(FSafeArray, @AIndex, @TVarData(Value));
      end;
  else
    raise EOleError.CreateFmt('Tipo não suportado: %d', [FVarType]);
  end;
  
  OleCheck(HR);
end;

procedure TSafeArrayWrapper.SetItem(const AIndices: array of Integer; const AValue: Variant);
var
  HR: HRESULT;
  Indices: array of Integer;
  I: Integer;
  Value: Variant;
begin
  CheckIndices(AIndices);
  Value := AValue;
  
  SetLength(Indices, Length(AIndices));
  for I := 0 to High(AIndices) do
    Indices[I] := AIndices[I];
  
  case FVarType of
    varVariant:
      begin
        HR := SafeArrayPutElement(FSafeArray, @Indices[0], @TVarData(Value));
        OleCheck(HR);
      end;
  else
    raise EOleError.Create('SetItem multidimensional implementado apenas para varVariant');
  end;
end;

procedure TSafeArrayWrapper.Clear;
var
  LBound, UBound, I: Integer;
begin
  EnsureNotNil;
  
  if FDimensions = 1 then
  begin
    SafeArrayGetLBound(FSafeArray, 1, LBound);
    SafeArrayGetUBound(FSafeArray, 1, UBound);
    
    for I := LBound to UBound do
      SetItem(I, Null);
  end
  else
    raise EOleError.Create('Clear implementado apenas para arrays unidimensionais');
end;

procedure TSafeArrayWrapper.Resize(ANewSize: Integer);
var
  NewBound: TSafeArrayBound;
  HR: HRESULT;
begin
  EnsureNotNil;
  
  if FDimensions > 1 then
    raise EOleError.Create('Resize implementado apenas para arrays unidimensionais');
  
  NewBound.lLbound := GetLBound(1);
  NewBound.cElements := ANewSize;
  
  HR := SafeArrayRedim(FSafeArray, NewBound);
  OleCheck(HR);
end;

procedure TSafeArrayWrapper.Append(const AValue: Variant);
var
  CurrentSize: Integer;
begin
  EnsureNotNil;
  
  if FDimensions > 1 then
    raise EOleError.Create('Append implementado apenas para arrays unidimensionais');
  
  CurrentSize := GetUBound(1) - GetLBound(1) + 1;
  Resize(CurrentSize + 1);
  SetItem(GetLBound(1) + CurrentSize, AValue);
end;

function TSafeArrayWrapper.ToVariantArray: Variant;
var
  V: Variant;
begin
  EnsureNotNil;
  
  // Cria um variant array com as mesmas características
  TVarData(V).VType := varArray or FVarType;
  TVarData(V).VArray := FSafeArray;
  
  // Copia o conteúdo
  Result := VarArrayCopy(V);
  
  // Limpa o variant temporário sem destruir o SafeArray
  TVarData(V).VArray := nil;
  VarClear(V);
end;

function TSafeArrayWrapper.ToStringArray: TArray<string>;
var
  I, LBound, UBound: Integer;
begin
  EnsureNotNil;
  
  if FDimensions > 1 then
    raise EOleError.Create('ToStringArray implementado apenas para arrays unidimensionais');
  
  SafeArrayGetLBound(FSafeArray, 1, LBound);
  SafeArrayGetUBound(FSafeArray, 1, UBound);
  
  SetLength(Result, UBound - LBound + 1);
  for I := LBound to UBound do
    Result[I - LBound] := VarToStr(GetItem(I));
end;

function TSafeArrayWrapper.ToIntegerArray: TArray<Integer>;
var
  I, LBound, UBound: Integer;
begin
  EnsureNotNil;
  
  if FDimensions > 1 then
    raise EOleError.Create('ToIntegerArray implementado apenas para arrays unidimensionais');
  
  SafeArrayGetLBound(FSafeArray, 1, LBound);
  SafeArrayGetUBound(FSafeArray, 1, UBound);
  
  SetLength(Result, UBound - LBound + 1);
  for I := LBound to UBound do
    Result[I - LBound] := GetItem(I);
end;

function TSafeArrayWrapper.ToDoubleArray: TArray<Double>;
var
  I, LBound, UBound: Integer;
begin
  EnsureNotNil;
  
  if FDimensions > 1 then
    raise EOleError.Create('ToDoubleArray implementado apenas para arrays unidimensionais');
  
  SafeArrayGetLBound(FSafeArray, 1, LBound);
  SafeArrayGetUBound(FSafeArray, 1, UBound);
  
  SetLength(Result, UBound - LBound + 1);
  for I := LBound to UBound do
    Result[I - LBound] := GetItem(I);
end;

{ Funções auxiliares }

function CreateSafeArray(AVarType: TVarType; ACount: Integer): ISafeArrayWrapper;
begin
  Result := TSafeArrayWrapper.Create(AVarType, 0, ACount - 1);
end;

function CreateSafeArray(AVarType: TVarType; const ABounds: array of Integer): ISafeArrayWrapper;
begin
  Result := TSafeArrayWrapper.Create(AVarType, ABounds);
end;

function WrapSafeArray(ASafeArray: PSafeArray; AOwnsData: Boolean): ISafeArrayWrapper;
begin
  Result := TSafeArrayWrapper.Create(ASafeArray, AOwnsData);
end;

function VariantToSafeArray(const AVariant: Variant): ISafeArrayWrapper;
begin
  Result := TSafeArrayWrapper.CreateFromVariant(AVariant);
end;

end.