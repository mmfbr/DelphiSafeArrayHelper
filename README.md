# DelphiSafeArrayHelper

Uma biblioteca Delphi/FreePascal que fornece um wrapper amigável e type-safe para trabalhar com SafeArrays COM, especialmente útil para interoperabilidade COM/.NET.

## 🎯 Motivação

Trabalhar com SafeArrays em Delphi pode ser complexo e propenso a erros. Esta biblioteca simplifica o uso de SafeArrays através de uma interface intuitiva com gerenciamento automático de memória.

## ✨ Características

- **Interface-based**: Gerenciamento automático de memória através de reference counting
- **API Intuitiva**: Acesso aos elementos como arrays nativos do Delphi
- **Type-safe**: Validação de índices e tipos em tempo de execução
- **Conversões Fáceis**: Métodos para converter entre SafeArray e arrays nativos
- **Suporte Multidimensional**: Trabalhe com arrays de múltiplas dimensões
- **Documentação Completa**: XMLDoc em todo o código
- **Bem Testado**: Suite completa de testes com DUnitX

## 📋 Requisitos

- Delphi 2010 ou superior
- FreePascal 3.0 ou superior (modo Delphi)
- Windows (usa Windows API para SafeArray)

## 🚀 Instalação

1. Clone o repositório:
```bash
git clone https://github.com/mmfbr/DelphiSafeArrayHelper.git
```

2. Adicione o caminho da pasta `src` ao Library Path do seu projeto

3. Adicione `SafeArrayWrapper` à cláusula uses:
```pascal
uses
  SafeArrayWrapper;
```

## 📖 Uso Básico

### Criando um SafeArray

```pascal
var
  SafeArr: ISafeArrayWrapper;
begin
  // Array de inteiros com 10 elementos (índices 0..9)
  SafeArr := CreateSafeArray(varInteger, 10);
  
  // Array com bounds customizados [5..15]
  SafeArr := TSafeArrayWrapper.Create(varInteger, 5, 15);
  
  // Array multidimensional 3x4
  SafeArr := CreateSafeArray(varVariant, [0, 2, 0, 3]);
end;
```

### Acessando Elementos

```pascal
// Array unidimensional
SafeArr[0] := 100;
Value := SafeArr[0];

// Array multidimensional
SafeArr[[1, 2]] := 'Hello';
Value := SafeArr[[1, 2]];
```

### Conversões

```pascal
// De Variant para SafeArray wrapper
var
  VarArr: Variant;
  SafeArr: ISafeArrayWrapper;
begin
  VarArr := VarArrayCreate([0, 4], varDouble);
  SafeArr := VariantToSafeArray(VarArr);
end;

// De SafeArray wrapper para arrays nativos
var
  IntArray: TArray<Integer>;
  StrArray: TArray<string>;
begin
  IntArray := SafeArr.ToIntegerArray;
  StrArray := SafeArr.ToStringArray;
end;
```

### Interoperabilidade COM/.NET

```pascal
uses
  mscorlib_TLB,
  SafeArrayWrapper;

var
  DotNetArray: OleVariant;
  SafeArr: ISafeArrayWrapper;
  Args: OleVariant;
begin
  // Preparar argumentos para método .NET
  SafeArr := CreateSafeArray(varVariant, 2);
  SafeArr[0] := 'Hello';
  SafeArr[1] := 'World';
  
  // Converter para passar ao método
  Args := SafeArr.ToVariantArray;
  DotNetObject.MethodThatExpectsArray(Args);
  
  // Receber array de método .NET
  DotNetArray := DotNetObject.GetNames();
  SafeArr := VariantToSafeArray(DotNetArray);
  
  // Processar elementos
  for I := 0 to SafeArr.Count - 1 do
    ProcessName(SafeArr[I]);
end;
```

## 🧪 Executando os Testes

O projeto inclui uma suite completa de testes usando DUnitX:

1. Abra `tests/SafeArrayWrapperTests.dproj` no Delphi
2. Compile e execute o projeto
3. Verifique os resultados no console ou GUI do DUnitX

## 📚 Documentação da API

### Interface ISafeArrayWrapper

#### Propriedades
- `Count: Integer` - Número total de elementos
- `Dimensions: Integer` - Número de dimensões
- `Items[Index]: Variant` - Acesso indexado (propriedade default)
- `VarType: TVarType` - Tipo dos elementos
- `SafeArray: PSafeArray` - Ponteiro para o SafeArray nativo

#### Métodos
- `GetLBound(Dimension: Integer = 1): Integer` - Limite inferior
- `GetUBound(Dimension: Integer = 1): Integer` - Limite superior
- `Clear` - Limpa todos os elementos
- `Resize(NewSize: Integer)` - Redimensiona o array
- `Append(Value: Variant)` - Adiciona elemento ao final
- `ToVariantArray: Variant` - Converte para Variant array
- `ToStringArray: TArray<string>` - Converte para array de strings
- `ToIntegerArray: TArray<Integer>` - Converte para array de inteiros
- `ToDoubleArray: TArray<Double>` - Converte para array de doubles

### Funções Auxiliares

- `CreateSafeArray(VarType: TVarType; Count: Integer): ISafeArrayWrapper`
- `WrapSafeArray(SafeArray: PSafeArray; OwnsData: Boolean = False): ISafeArrayWrapper`
- `VariantToSafeArray(Variant: Variant): ISafeArrayWrapper`

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor:

1. Faça um Fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto está licenciado sob a MIT License - veja o arquivo [LICENSE](LICENSE) para detalhes.

## 🙏 Agradecimentos

- Comunidade Delphi por anos de conhecimento compartilhado sobre COM e SafeArrays
- Contribuidores do projeto DUnitX por fornecer um excelente framework de testes

## 📞 Contato

Seu Nome - [@mmfbr77](https://twitter.com/mmfbr77)

Link do Projeto: [https://github.com/mmfbr/DelphiSafeArrayHelper](https://github.com/mmfbr/DelphiSafeArrayHelper)