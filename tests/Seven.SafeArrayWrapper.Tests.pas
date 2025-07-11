unit Seven.SafeArrayWrapper.Tests;

interface

uses
  DUnitX.TestFramework,
  System.SysUtils,
  System.Variants,
  Winapi.ActiveX,
  Seven.SafeArrayWrapper;

type
  /// <summary>
  ///   Suite de testes para SafeArrayWrapper
  /// </summary>
  [TestFixture]
  TSafeArrayWrapperTests = class
  private
    procedure CheckArraysEqual(const Expected, Actual: TArray<Integer>); overload;
    procedure CheckArraysEqual(const Expected, Actual: TArray<Double>); overload;
    procedure CheckArraysEqual(const Expected, Actual: TArray<string>); overload;
  public
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    // Testes de criação
    [Test]
    procedure TestCreateUnidimensional;
    
    [Test]
    procedure TestCreateMultidimensional;
    
    [Test]
    procedure TestCreateFromVariant;
    
    [Test]
    procedure TestWrapExistingSafeArray;
    
    [Test]
    procedure TestCreateWithCustomBounds;
    
    // Testes de acesso a elementos
    [Test]
    procedure TestGetSetInteger;
    
    [Test]
    procedure TestGetSetDouble;
    
    [Test]
    procedure TestGetSetString;
    
    [Test]
    procedure TestGetSetVariant;
    
    [Test]
    procedure TestGetSetDispatch;
    
    [Test]
    procedure TestMultidimensionalAccess;
    
    // Testes de validação e erros
    [Test]
    procedure TestIndexOutOfBounds;
    
    [Test]
    procedure TestNilSafeArray;
    
    [Test]
    procedure TestInvalidDimensions;
    
    // Testes de manipulação
    [Test]
    procedure TestResize;
    
    [Test]
    procedure TestAppend;
    
    [Test]
    procedure TestClear;
    
    // Testes de conversão
    [Test]
    procedure TestToVariantArray;
    
    [Test]
    procedure TestToStringArray;
    
    [Test]
    procedure TestToIntegerArray;
    
    [Test]
    procedure TestToDoubleArray;
    
    // Testes de propriedades
    [Test]
    procedure TestCount;
    
    [Test]
    procedure TestDimensions;
    
    [Test]
    procedure TestBounds;
    
    [Test]
    procedure TestElementSize;
    
    [Test]
    procedure TestVarType;
    
    // Testes de gerenciamento de memória
    [Test]
    procedure TestMemoryManagement;
    
    [Test]
    procedure TestOwnership;
    
    // Testes de cenários reais
    [Test]
    procedure TestCOMInterop;
    
    [Test]
    procedure TestLargeArray;
    
    [Test]
    procedure TestMixedTypes;
  end;

implementation

uses
  System.Win.ComObj;

{ TSafeArrayWrapperTests }

procedure TSafeArrayWrapperTests.Setup;
begin
  // Inicializa COM para os testes
  CoInitialize(nil);
end;

procedure TSafeArrayWrapperTests.TearDown;
begin
  // Finaliza COM
  CoUninitialize;
end;

procedure TSafeArrayWrapperTests.CheckArraysEqual(const Expected, Actual: TArray<Integer>);
var
  I: Integer;
begin
  Assert.AreEqual(Length(Expected), Length(Actual), 'Arrays têm tamanhos diferentes');
  for I := 0 to High(Expected) do
    Assert.AreEqual(Expected[I], Actual[I], Format('Elemento [%d] diferente', [I]));
end;

procedure TSafeArrayWrapperTests.CheckArraysEqual(const Expected, Actual: TArray<Double>);
var
  I: Integer;
begin
  Assert.AreEqual(Length(Expected), Length(Actual), 'Arrays têm tamanhos diferentes');
  for I := 0 to High(Expected) do
    Assert.AreEqual(Expected[I], Actual[I], 0.0001, Format('Elemento [%d] diferente', [I]));
end;

procedure TSafeArrayWrapperTests.CheckArraysEqual(const Expected, Actual: TArray<string>);
var
  I: Integer;
begin
  Assert.AreEqual(Length(Expected), Length(Actual), 'Arrays têm tamanhos diferentes');
  for I := 0 to High(Expected) do
    Assert.AreEqual(Expected[I], Actual[I], Format('Elemento [%d] diferente', [I]));
end;

procedure TSafeArrayWrapperTests.TestCreateUnidimensional;
var
  SafeArr: ISafeArrayWrapper;
begin
  // Cria array com 10 elementos
  SafeArr := CreateSafeArray(varInteger, 10);
  
  Assert.IsNotNull(SafeArr, 'SafeArray não foi criado');
  Assert.AreEqual(10, SafeArr.Count);
  Assert.AreEqual(1, SafeArr.Dimensions);
  Assert.AreEqual(0, SafeArr.GetLBound);
  Assert.AreEqual(9, SafeArr.GetUBound);
  Assert.AreEqual(varInteger, Integer(SafeArr.VarType));
end;

procedure TSafeArrayWrapperTests.TestCreateMultidimensional;
var
  SafeArr: ISafeArrayWrapper;
begin
  // Cria array 3x4 (bounds: [0..2, 0..3])
  SafeArr := CreateSafeArray(varVariant, [0, 2, 0, 3]);
  
  Assert.IsNotNull(SafeArr);
  Assert.AreEqual(12, SafeArr.Count); // 3 * 4
  Assert.AreEqual(2, SafeArr.Dimensions);
  Assert.AreEqual(0, SafeArr.GetLBound(1));
  Assert.AreEqual(2, SafeArr.GetUBound(1));
  Assert.AreEqual(0, SafeArr.GetLBound(2));
  Assert.AreEqual(3, SafeArr.GetUBound(2));
end;

procedure TSafeArrayWrapperTests.TestCreateFromVariant;
var
  VarArr: Variant;
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  // Cria um Variant array
  VarArr := VarArrayCreate([0, 4], varInteger);
  for I := 0 to 4 do
    VarArr[I] := I * 10;
  
  // Converte para wrapper
  SafeArr := VariantToSafeArray(VarArr);
  
  Assert.IsNotNull(SafeArr);
  Assert.AreEqual(5, SafeArr.Count);
  Assert.AreEqual(varInteger, Integer(SafeArr.VarType));
  
  // Verifica valores
  for I := 0 to 4 do
    Assert.AreEqual(I * 10, Integer(SafeArr[I]));
end;

procedure TSafeArrayWrapperTests.TestWrapExistingSafeArray;
var
  Bounds: TSafeArrayBound;
  PSA: PSafeArray;
  SafeArr: ISafeArrayWrapper;
begin
  // Cria SafeArray nativo
  Bounds.lLbound := 0;
  Bounds.cElements := 5;
  PSA := SafeArrayCreate(varDouble, 1, @Bounds);
  
  try
    // Wrap sem ownership
    SafeArr := WrapSafeArray(PSA, False);
    
    Assert.IsNotNull(SafeArr);
    Assert.AreEqual(5, SafeArr.Count);
    Assert.AreEqual(varDouble, Integer(SafeArr.VarType));
    Assert.AreEqual(PSA, SafeArr.SafeArray);
  finally
    // Limpa manualmente já que wrapper não tem ownership
    SafeArrayDestroy(PSA);
  end;
end;

procedure TSafeArrayWrapperTests.TestCreateWithCustomBounds;
var
  SafeArr: ISafeArrayWrapper;
begin
  // Cria array com bounds customizados [5..10]
  SafeArr := TSafeArrayWrapper.Create(varString, 5, 10);
  
  Assert.AreEqual(6, SafeArr.Count); // 10 - 5 + 1
  Assert.AreEqual(5, SafeArr.GetLBound);
  Assert.AreEqual(10, SafeArr.GetUBound);
end;

procedure TSafeArrayWrapperTests.TestGetSetInteger;
var
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varInteger, 5);
  
  // Define valores
  for I := 0 to 4 do
    SafeArr[I] := I * 100;
  
  // Verifica valores
  for I := 0 to 4 do
    Assert.AreEqual(I * 100, Integer(SafeArr[I]));
end;

procedure TSafeArrayWrapperTests.TestGetSetDouble;
var
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varDouble, 5);
  
  // Define valores
  for I := 0 to 4 do
    SafeArr[I] := I * 3.14;
  
  // Verifica valores
  for I := 0 to 4 do
    Assert.AreEqual(I * 3.14, Double(SafeArr[I]), 0.0001);
end;

procedure TSafeArrayWrapperTests.TestGetSetString;
var
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varOleStr, 5);
  
  // Define valores
  for I := 0 to 4 do
    SafeArr[I] := Format('String %d', [I]);
  
  // Verifica valores
  for I := 0 to 4 do
    Assert.AreEqual(Format('String %d', [I]), string(SafeArr[I]));
end;

procedure TSafeArrayWrapperTests.TestGetSetVariant;
var
  SafeArr: ISafeArrayWrapper;
begin
  SafeArr := CreateSafeArray(varVariant, 4);
  
  // Testa diferentes tipos em Variant
  SafeArr[0] := 123;
  SafeArr[1] := 'Test String';
  SafeArr[2] := 3.14;
  SafeArr[3] := True;
  
  Assert.AreEqual(123, Integer(SafeArr[0]));
  Assert.AreEqual('Test String', string(SafeArr[1]));
  Assert.AreEqual(3.14, Double(SafeArr[2]), 0.0001);
  Assert.AreEqual(True, Boolean(SafeArr[3]));
end;

procedure TSafeArrayWrapperTests.TestGetSetDispatch;
var
  SafeArr: ISafeArrayWrapper;
  Obj1, Obj2: IDispatch;
begin
  SafeArr := CreateSafeArray(varDispatch, 2);
  
  // Cria objetos COM
  Obj1 := CreateOleObject('Scripting.Dictionary');
  Obj2 := CreateOleObject('Scripting.FileSystemObject');
  
  // Armazena no array
  SafeArr[0] := Obj1;
  SafeArr[1] := Obj2;
  
  // Verifica
  Assert.IsNotNull(IDispatch(SafeArr[0]));
  Assert.IsNotNull(IDispatch(SafeArr[1]));
  Assert.AreNotEqual(IDispatch(SafeArr[0]), IDispatch(SafeArr[1]));
end;

procedure TSafeArrayWrapperTests.TestMultidimensionalAccess;
var
  SafeArr: ISafeArrayWrapper;
  I, J: Integer;
  Value: string;
begin
  // Cria matriz 3x3
  SafeArr := CreateSafeArray(varVariant, [0, 2, 0, 2]);
  
  // Preenche
  for I := 0 to 2 do
    for J := 0 to 2 do
      SafeArr[[I, J]] := Format('[%d,%d]', [I, J]);
  
  // Verifica
  for I := 0 to 2 do
    for J := 0 to 2 do
    begin
      Value := SafeArr[[I, J]];
      Assert.AreEqual(Format('[%d,%d]', [I, J]), Value);
    end;
end;

procedure TSafeArrayWrapperTests.TestIndexOutOfBounds;
var
  SafeArr: ISafeArrayWrapper;
begin
  SafeArr := CreateSafeArray(varInteger, 5); // índices 0..4
  
  // Testa acesso fora dos limites
  Assert.WillRaise(
    procedure
    begin
      SafeArr[5] := 100; // índice inválido
    end,
    ERangeError
  );
  
  Assert.WillRaise(
    procedure
    begin
      var Value := SafeArr[-1]; // índice inválido
    end,
    ERangeError
  );
end;

procedure TSafeArrayWrapperTests.TestNilSafeArray;
var
  SafeArr: ISafeArrayWrapper;
begin
  // Cria wrapper com SafeArray nil
  SafeArr := WrapSafeArray(nil, False);
  
  // Qualquer operação deve lançar exceção
  Assert.WillRaise(
    procedure
    begin
      var Count := SafeArr.Count;
    end,
    EOleError
  );
end;

procedure TSafeArrayWrapperTests.TestInvalidDimensions;
var
  SafeArr: ISafeArrayWrapper;
begin
  // Testa passar número ímpar de bounds
  Assert.WillRaise(
    procedure
    begin
      SafeArr := CreateSafeArray(varInteger, [0, 5, 0]); // 3 elementos (ímpar)
    end,
    EArgumentException
  );
  
  // Testa acessar multidimensional como unidimensional
  SafeArr := CreateSafeArray(varVariant, [0, 2, 0, 2]);
  Assert.WillRaise(
    procedure
    begin
      SafeArr[0] := 100; // precisa de 2 índices
    end,
    EArgumentException
  );
end;

procedure TSafeArrayWrapperTests.TestResize;
var
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varInteger, 5);
  
  // Preenche
  for I := 0 to 4 do
    SafeArr[I] := I;
  
  // Redimensiona para maior
  SafeArr.Resize(10);
  Assert.AreEqual(10, SafeArr.Count);
  
  // Valores antigos preservados
  for I := 0 to 4 do
    Assert.AreEqual(I, Integer(SafeArr[I]));
  
  // Redimensiona para menor
  SafeArr.Resize(3);
  Assert.AreEqual(3, SafeArr.Count);
end;

procedure TSafeArrayWrapperTests.TestAppend;
var
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varInteger, 3);
  
  // Valores iniciais
  for I := 0 to 2 do
    SafeArr[I] := I;
  
  // Append
  SafeArr.Append(100);
  SafeArr.Append(200);
  
  Assert.AreEqual(5, SafeArr.Count);
  Assert.AreEqual(100, Integer(SafeArr[3]));
  Assert.AreEqual(200, Integer(SafeArr[4]));
end;

procedure TSafeArrayWrapperTests.TestClear;
var
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varVariant, 5);
  
  // Preenche
  for I := 0 to 4 do
    SafeArr[I] := I * 10;
  
  // Clear
  SafeArr.Clear;
  
  // Verifica que todos são Null
  for I := 0 to 4 do
    Assert.IsTrue(VarIsNull(SafeArr[I]));
end;

procedure TSafeArrayWrapperTests.TestToVariantArray;
var
  SafeArr: ISafeArrayWrapper;
  VarArr: Variant;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varInteger, 5);
  
  // Preenche
  for I := 0 to 4 do
    SafeArr[I] := I * 10;
  
  // Converte
  VarArr := SafeArr.ToVariantArray;
  
  Assert.IsTrue(VarIsArray(VarArr));
  Assert.AreEqual(0, VarArrayLowBound(VarArr, 1));
  Assert.AreEqual(4, VarArrayHighBound(VarArr, 1));
  
  // Verifica valores
  for I := 0 to 4 do
    Assert.AreEqual(I * 10, Integer(VarArr[I]));
end;

procedure TSafeArrayWrapperTests.TestToStringArray;
var
  SafeArr: ISafeArrayWrapper;
  StrArray: TArray<string>;
  Expected: TArray<string>;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varOleStr, 4);
  
  SafeArr[0] := 'First';
  SafeArr[1] := 'Second';
  SafeArr[2] := 'Third';
  SafeArr[3] := 'Fourth';
  
  StrArray := SafeArr.ToStringArray;
  Expected := ['First', 'Second', 'Third', 'Fourth'];
  
  CheckArraysEqual(Expected, StrArray);
end;

procedure TSafeArrayWrapperTests.TestToIntegerArray;
var
  SafeArr: ISafeArrayWrapper;
  IntArray: TArray<Integer>;
  Expected: TArray<Integer>;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varInteger, 5);
  
  for I := 0 to 4 do
    SafeArr[I] := I * 2;
  
  IntArray := SafeArr.ToIntegerArray;
  Expected := [0, 2, 4, 6, 8];
  
  CheckArraysEqual(Expected, IntArray);
end;

procedure TSafeArrayWrapperTests.TestToDoubleArray;
var
  SafeArr: ISafeArrayWrapper;
  DblArray: TArray<Double>;
  Expected: TArray<Double>;
  I: Integer;
begin
  SafeArr := CreateSafeArray(varDouble, 4);
  
  for I := 0 to 3 do
    SafeArr[I] := I * 1.5;
  
  DblArray := SafeArr.ToDoubleArray;
  Expected := [0.0, 1.5, 3.0, 4.5];
  
  CheckArraysEqual(Expected, DblArray);
end;

procedure TSafeArrayWrapperTests.TestCount;
var
  SafeArr: ISafeArrayWrapper;
begin
  // Unidimensional
  SafeArr := CreateSafeArray(varInteger, 10);
  Assert.AreEqual(10, SafeArr.Count);
  
  // Multidimensional
  SafeArr := CreateSafeArray(varVariant, [0, 4, 0, 3]); // 5x4
  Assert.AreEqual(20, SafeArr.Count);
  
  // Com bounds customizados
  SafeArr := TSafeArrayWrapper.Create(varInteger, 5, 9); // [5..9]
  Assert.AreEqual(5, SafeArr.Count);
end;

procedure TSafeArrayWrapperTests.TestDimensions;
var
  SafeArr: ISafeArrayWrapper;
begin
  // 1D
  SafeArr := CreateSafeArray(varInteger, 10);
  Assert.AreEqual(1, SafeArr.Dimensions);
  
  // 2D
  SafeArr := CreateSafeArray(varVariant, [0, 5, 0, 5]);
  Assert.AreEqual(2, SafeArr.Dimensions);
  
  // 3D
  SafeArr := CreateSafeArray(varVariant, [0, 2, 0, 3, 0, 4]);
  Assert.AreEqual(3, SafeArr.Dimensions);
end;

procedure TSafeArrayWrapperTests.TestBounds;
var
  SafeArr: ISafeArrayWrapper;
begin
  // Bounds padrão
  SafeArr := CreateSafeArray(varInteger, 10);
  Assert.AreEqual(0, SafeArr.GetLBound);
  Assert.AreEqual(9, SafeArr.GetUBound);
  
  // Bounds customizados
  SafeArr := TSafeArrayWrapper.Create(varInteger, -5, 5);
  Assert.AreEqual(-5, SafeArr.GetLBound);
  Assert.AreEqual(5, SafeArr.GetUBound);
  Assert.AreEqual(11, SafeArr.Count); // -5 até 5 = 11 elementos
end;

procedure TSafeArrayWrapperTests.TestElementSize;
var
  SafeArr: ISafeArrayWrapper;
begin
  // Integer (4 bytes)
  SafeArr := CreateSafeArray(varInteger, 1);
  Assert.AreEqual(4, SafeArr.GetElementSize);
  
  // Double (8 bytes)
  SafeArr := CreateSafeArray(varDouble, 1);
  Assert.AreEqual(8, SafeArr.GetElementSize);
  
  // Variant (16 bytes no Windows de 32 bits, pode variar)
  SafeArr := CreateSafeArray(varVariant, 1);
  Assert.IsTrue(SafeArr.GetElementSize >= 16);
end;

procedure TSafeArrayWrapperTests.TestVarType;
var
  SafeArr: ISafeArrayWrapper;
begin
  SafeArr := CreateSafeArray(varInteger, 1);
  Assert.AreEqual(varInteger, Integer(SafeArr.VarType));
  
  SafeArr := CreateSafeArray(varDouble, 1);
  Assert.AreEqual(varDouble, Integer(SafeArr.VarType));
  
  SafeArr := CreateSafeArray(varOleStr, 1);
  Assert.AreEqual(varOleStr, Integer(SafeArr.VarType));
end;

procedure TSafeArrayWrapperTests.TestMemoryManagement;
var
  SafeArr1, SafeArr2: ISafeArrayWrapper;
  PSA: PSafeArray;
begin
  // Teste de reference counting
  SafeArr1 := CreateSafeArray(varInteger, 10);
  PSA := SafeArr1.SafeArray;
  
  // Segunda referência
  SafeArr2 := SafeArr1;
  
  // Libera primeira referência
  SafeArr1 := nil;
  
  // SafeArray ainda deve ser válido
  Assert.IsNotNull(SafeArr2);
  Assert.AreEqual(10, SafeArr2.Count);
  
  // Após liberar todas as referências, SafeArray é destruído
  SafeArr2 := nil;
  // PSA agora aponta para memória inválida (não podemos testar isso diretamente)
end;

procedure TSafeArrayWrapperTests.TestOwnership;
var
  Bounds: TSafeArrayBound;
  PSA: PSafeArray;
  SafeArr: ISafeArrayWrapper;
  IsValid: Boolean;
begin
  // Cria SafeArray nativo
  Bounds.lLbound := 0;
  Bounds.cElements := 5;
  PSA := SafeArrayCreate(varInteger, 1, @Bounds);
  
  // Wrap COM ownership
  SafeArr := WrapSafeArray(PSA, True);
  SafeArr := nil; // Deve destruir o SafeArray
  
  // Cria outro
  PSA := SafeArrayCreate(varInteger, 1, @Bounds);
  
  // Wrap SEM ownership
  SafeArr := WrapSafeArray(PSA, False);
  SafeArr := nil; // NÃO deve destruir o SafeArray
  
  // Verifica que ainda é válido
  IsValid := SafeArrayGetDim(PSA) = 1;
  Assert.IsTrue(IsValid);
  
  // Limpa manualmente
  SafeArrayDestroy(PSA);
end;

procedure TSafeArrayWrapperTests.TestCOMInterop;
var
  Dict: OleVariant;
  SafeArr: ISafeArrayWrapper;
  Keys, Items: OleVariant;
  I: Integer;
begin
  // Cria Dictionary COM
  Dict := CreateOleObject('Scripting.Dictionary');
  
  // Adiciona itens
  Dict.Add('Key1', 'Value1');
  Dict.Add('Key2', 'Value2');
  Dict.Add('Key3', 'Value3');
  
  // Obtém arrays
  Keys := Dict.Keys;
  Items := Dict.Items;
  
  // Wrap os arrays
  var KeysArr := VariantToSafeArray(Keys);
  var ItemsArr := VariantToSafeArray(Items);
  
  Assert.AreEqual(3, KeysArr.Count);
  Assert.AreEqual(3, ItemsArr.Count);
  
  // Verifica conteúdo
  for I := 0 to 2 do
  begin
    Assert.AreEqual(Format('Key%d', [I + 1]), string(KeysArr[I]));
    Assert.AreEqual(Format('Value%d', [I + 1]), string(ItemsArr[I]));
  end;
end;

procedure TSafeArrayWrapperTests.TestLargeArray;
const
  SIZE = 10000;
var
  SafeArr: ISafeArrayWrapper;
  I: Integer;
  Sum: Int64;
begin
  // Cria array grande
  SafeArr := CreateSafeArray(varInteger, SIZE);
  
  // Preenche
  for I := 0 to SIZE - 1 do
    SafeArr[I] := I;
  
  // Verifica alguns valores
  Assert.AreEqual(0, Integer(SafeArr[0]));
  Assert.AreEqual(SIZE div 2, Integer(SafeArr[SIZE div 2]));
  Assert.AreEqual(SIZE - 1, Integer(SafeArr[SIZE - 1]));
  
  // Calcula soma para verificar integridade
  Sum := 0;
  for I := 0 to SIZE - 1 do
    Sum := Sum + Integer(SafeArr[I]);
  
  Assert.AreEqual(Int64(SIZE) * (SIZE - 1) div 2, Sum);
end;

procedure TSafeArrayWrapperTests.TestMixedTypes;
var
  SafeArr: ISafeArrayWrapper;
  Today: TDateTime;
begin
  SafeArr := CreateSafeArray(varVariant, 6);
  
  // Armazena diferentes tipos
  SafeArr[0] := 123;           // Integer
  SafeArr[1] := 'Hello World'; // String
  SafeArr[2] := 3.14159;       // Double
  SafeArr[3] := True;          // Boolean
  SafeArr[4] := Now;           // DateTime
  SafeArr[5] := Null;          // Null
  
  // Verifica tipos
  Assert.AreEqual(varInteger, VarType(SafeArr[0]) and varTypeMask);
  Assert.AreEqual(varUString, VarType(SafeArr[1]) and varTypeMask);
  Assert.AreEqual(varDouble, VarType(SafeArr[2]) and varTypeMask);
  Assert.AreEqual(varBoolean, VarType(SafeArr[3]) and varTypeMask);
  Assert.AreEqual(varDate, VarType(SafeArr[4]) and varTypeMask);
  Assert.IsTrue(VarIsNull(SafeArr[5]));
  
  // Verifica valores
  Assert.AreEqual(123, Integer(SafeArr[0]));
  Assert.AreEqual('Hello World', string(SafeArr[1]));
  Assert.AreEqual(3.14159, Double(SafeArr[2]), 0.00001);
  Assert.AreEqual(True, Boolean(SafeArr[3]));
  
  Today := Date;
  Assert.IsTrue(Abs(TDateTime(SafeArr[4]) - Now) < 1.0); // Menos de 1 dia de diferença
end;

initialization
  TDUnitX.RegisterTestFixture(TSafeArrayWrapperTests);

end.