program ExemploSafeArrayWrapper1;

{$APPTYPE CONSOLE}

uses
  System.SysUtils,
  System.Variants,
  System.Win.ComObj,
  ActiveX,
  mscorlib_TLB, // Import da TypeLib do .NET
  SafeArrayWrapper; // Nossa unit wrapper

// Exemplo 1: Criando e manipulando um SafeArray simples
procedure ExemploBasico;
var
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  WriteLn('=== Exemplo Básico ===');
  
  // Cria um SafeArray de inteiros com 10 elementos (índices 0..9)
  SafeArr := CreateSafeArray(varInteger, 10);
  
  // Preenche o array
  for I := 0 to 9 do
    SafeArr[I] := I * 10;
  
  // Lê e exibe os valores
  Write('Valores: ');
  for I := 0 to 9 do
    Write(SafeArr[I], ' ');
  WriteLn;
  
  // Adiciona um novo elemento
  SafeArr.Append(100);
  WriteLn('Após Append: Count = ', SafeArr.Count);
  WriteLn('Último elemento: ', SafeArr[10]);
  
  WriteLn;
end;

// Exemplo 2: Trabalhando com strings
procedure ExemploStrings;
var
  SafeArr: ISafeArrayWrapper;
  StringArray: TArray<string>;
  I: Integer;
begin
  WriteLn('=== Exemplo com Strings ===');
  
  // Cria um SafeArray de strings
  SafeArr := CreateSafeArray(varOleStr, 5);
  
  // Preenche com strings
  SafeArr[0] := 'Primeira';
  SafeArr[1] := 'Segunda';
  SafeArr[2] := 'Terceira';
  SafeArr[3] := 'Quarta';
  SafeArr[4] := 'Quinta';
  
  // Converte para array nativo do Delphi
  StringArray := SafeArr.ToStringArray;
  
  Write('Strings: ');
  for I := 0 to High(StringArray) do
    Write(StringArray[I], ' ');
  WriteLn;
  
  WriteLn;
end;

// Exemplo 3: Interoperabilidade COM com .NET
procedure ExemploCOMDotNet;
var
  DotNetType: _Type;
  AppDomain: _AppDomain;
  Assembly: _Assembly;
  SafeArr: ISafeArrayWrapper;
  Args: OleVariant;
  Result: OleVariant;
  MethodInfo: _MethodInfo;
  Methods: PSafeArray;
  I: Integer;
begin
  WriteLn('=== Exemplo COM com .NET ===');
  
  try
    // Obtém o AppDomain atual
    AppDomain := CoAppDomain.CreateRemote('');
    
    // Carrega mscorlib
    Assembly := AppDomain.Load_2('mscorlib');
    
    // Obtém o tipo System.Math
    DotNetType := Assembly.GetType_2('System.Math');
    
    // Cria um SafeArray para passar argumentos para métodos .NET
    SafeArr := CreateSafeArray(varVariant, 2);
    SafeArr[0] := 16.0;  // Argumento para Math.Sqrt
    SafeArr[1] := Null;  // Segundo argumento (não usado neste caso)
    
    // Converte para Variant para passar ao método COM
    Args := SafeArr.ToVariantArray;
    
    // Obtém informações sobre os métodos
    Methods := DotNetType.GetMethods;
    if Methods <> nil then
    begin
      var MethodsWrapper: ISafeArrayWrapper;
      MethodsWrapper := WrapSafeArray(Methods, False);
      WriteLn('Número de métodos em System.Math: ', MethodsWrapper.Count);
      
      // Lista alguns métodos (primeiros 5)
      Write('Alguns métodos: ');
      for I := 0 to 4 do
      begin
        if I < MethodsWrapper.Count then
        begin
          MethodInfo := IUnknown(MethodsWrapper[I]) as _MethodInfo;
          if Assigned(MethodInfo) then
            Write(MethodInfo.Name, ' ');
        end;
      end;
      WriteLn;
    end;
    
    // Chama Math.Sqrt(16) usando reflexão
    MethodInfo := DotNetType.GetMethod_6('Sqrt');
    if Assigned(MethodInfo) then
    begin
      Result := MethodInfo.Invoke_3(Null, Args);
      WriteLn('Math.Sqrt(16) = ', Double(Result));
    end;
    
  except
    on E: Exception do
      WriteLn('Erro COM: ', E.Message);
  end;
  
  WriteLn;
end;

// Exemplo 4: Array multidimensional
procedure ExemploMultidimensional;
var
  SafeArr: ISafeArrayWrapper;
  I, J: Integer;
begin
  WriteLn('=== Exemplo Multidimensional ===');
  
  // Cria um array 3x3 (bounds: [0..2, 0..2])
  SafeArr := CreateSafeArray(varVariant, [0, 2, 0, 2]);
  
  WriteLn('Dimensões: ', SafeArr.Dimensions);
  WriteLn('Total de elementos: ', SafeArr.Count);
  
  // Preenche a matriz
  for I := 0 to 2 do
    for J := 0 to 2 do
      SafeArr[[I, J]] := Format('Cell[%d,%d]', [I, J]);
  
  // Exibe a matriz
  WriteLn('Matriz:');
  for I := 0 to 2 do
  begin
    for J := 0 to 2 do
      Write(SafeArr[[I, J]], #9);
    WriteLn;
  end;
  
  WriteLn;
end;

// Exemplo 5: Conversão de Variant Array existente
procedure ExemploConversaoVariant;
var
  VarArr: Variant;
  SafeArr: ISafeArrayWrapper;
  I: Integer;
begin
  WriteLn('=== Exemplo Conversão de Variant ===');
  
  // Cria um Variant Array tradicional
  VarArr := VarArrayCreate([0, 4], varDouble);
  for I := 0 to 4 do
    VarArr[I] := (I + 1) * 3.14;
  
  // Converte para nosso wrapper
  SafeArr := VariantToSafeArray(VarArr);
  
  WriteLn('Tipo do array: varDouble (', varDouble, ')');
  WriteLn('Elementos:');
  for I := 0 to SafeArr.Count - 1 do
    WriteLn(Format('  [%d] = %.2f', [I, Double(SafeArr[I])]));
  
  // Modifica através do wrapper
  SafeArr[2] := 99.99;
  
  // O Variant original também foi modificado (mesma memória)
  WriteLn('Valor no Variant original após modificação: ', VarArr[2]);
  
  WriteLn;
end;

// Exemplo 6: Passando SafeArray para método COM
procedure ExemploPassandoParaCOM;
var
  ArrayList: OleVariant;
  SafeArr: ISafeArrayWrapper;
  Count: Integer;
begin
  WriteLn('=== Exemplo Passando para COM ===');
  
  try
    // Cria um ArrayList do .NET
    ArrayList := CreateOleObject('System.Collections.ArrayList');
    
    // Cria nosso SafeArray com alguns valores
    SafeArr := CreateSafeArray(varVariant, 5);
    SafeArr[0] := 'Item 1';
    SafeArr[1] := 'Item 2';
    SafeArr[2] := 'Item 3';
    SafeArr[3] := 'Item 4';
    SafeArr[4] := 'Item 5';
    
    // Adiciona cada item ao ArrayList
    // (ArrayList.AddRange espera um ICollection, então adicionamos um por um)
    for Count := 0 to SafeArr.Count - 1 do
      ArrayList.Add(SafeArr[Count]);
    
    WriteLn('Items adicionados ao ArrayList: ', ArrayList.Count);
    
    // Recupera como array
    var ReturnedArray: OleVariant;
    ReturnedArray := ArrayList.ToArray();
    
    // Wrap o array retornado
    var ReturnedSafeArr: ISafeArrayWrapper;
    ReturnedSafeArr := VariantToSafeArray(ReturnedArray);
    
    WriteLn('Items recuperados:');
    for Count := 0 to ReturnedSafeArr.Count - 1 do
      WriteLn('  ', ReturnedSafeArr[Count]);
    
  except
    on E: Exception do
      WriteLn('Erro: ', E.Message);
  end;
  
  WriteLn;
end;

begin
  CoInitialize(nil);
  try
    WriteLn('Demonstração do SafeArray Wrapper');
    WriteLn('=================================');
    WriteLn;
    
    ExemploBasico;
    ExemploStrings;
    ExemploCOMDotNet;
    ExemploMultidimensional;
    ExemploConversaoVariant;
    ExemploPassandoParaCOM;
    
    WriteLn('Pressione ENTER para sair...');
    ReadLn;
  finally
    CoUninitialize;
  end;
end.