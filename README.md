# Funções VBA para Recuperação de Dados da API Datagro

Este repositório contém várias funções VBA desenvolvidas para recuperar dados da API Datagro. Estas funções permitem buscar, filtrar e processar dados de maneira eficiente, utilizando parâmetros como TICKER, DATA, MEDIDA, entre outros.

## Funcionalidades

1. **RegisterGetData**: Registra a macro `GetDataPoint` com descrição e categorias personalizadas.
2. **GetInfo**: Recupera informações gerais para um TICKER fornecido.
3. **GetData**: Busca dados para um TICKER dentro de um intervalo de datas especificado e com medidas específicas.
4. **GetDataPoint**: Recupera um ponto de dado específico para um TICKER, data e medida fornecidos.
5. **GetFields**: Obtém os campos disponíveis para um TICKER fornecido, excluindo campos predefinidos.
6. **GetAverage**: Calcula a média de uma medida específica para um TICKER dentro de um intervalo de datas.
7. **GetSum**: Calcula a soma de uma medida específica para um TICKER dentro de um intervalo de datas.

## Pré-requisitos

1. **MSXML2**: Certifique-se de ter a referência à biblioteca Microsoft XML, v6.0 ativada no VBA.
2. **JsonConverter**: Baixe e importe `JsonConverter.bas` do repositório [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) para permitir a análise de respostas JSON.

## Instalação

1. Clone este repositório para o seu ambiente local:
    ```sh
    git clone https://github.com/seu-usuario/seu-repositorio.git
    ```

2. Abra o arquivo VBA no seu editor VBA (Excel, por exemplo).

3. Adicione uma referência à biblioteca MSXML2:
    - Vá para `Ferramentas` > `Referências`
    - Marque `Microsoft XML, v6.0`

4. Importe `JsonConverter.bas`:
    - Vá para `Arquivo` > `Importar arquivo`
    - Selecione o arquivo `JsonConverter.bas` baixado do [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)

## Uso

### RegisterGetData

Registra a macro `GetDataPoint` com uma descrição e categorias personalizadas.

```vba
Public Sub RegisterGetData()
    Application.MacroOptions _
        Macro:="GetDataPoint", _
        Description:="Busca e filtra dados da fonte de dados com base no TICKER, ANO e MEDIDA fornecidos.", _
        Category:="Funções Personalizadas", _
        ArgumentDescriptions:=Array("TICKER - O identificador para o conjunto de dados.", _
                                    "ANO - O ano para o qual os dados devem ser recuperados. Se não especificado, busca os dados mais recentes.", _
                                    "MEDIDA - A medida específica a ser recuperada dos dados. Se não especificado, retorna todos os dados disponíveis.")
End Sub
```

### GetInfo

Recupera informações gerais para um TICKER fornecido.

```vba
Public Function GetInfo(ticker As String) As String
    ' Recupera informações gerais para um TICKER fornecido da fonte de dados.
    ' Parâmetros:
    ' TICKER (String): O identificador único para o conjunto de dados para recuperar informações.
    ' Retorna:
    ' Uma string contendo as informações "longo" recuperadas ou uma mensagem de erro se os dados não puderem ser recuperados.
End Function
```

### GetData

Busca dados para um TICKER dentro de um intervalo de datas especificado e com medidas específicas.

```vba
Public Function GetData(ticker As String, startDate As String, endDate As String, ParamArray MEASURES() As Variant) As Variant
    ' Busca dados para um TICKER dentro de um intervalo de datas especificado e com medidas específicas.
    ' Parâmetros:
    ' ticker (String): O identificador único para o conjunto de dados.
    ' startDate (String): Data de início no formato YYYY-MM-DD.
    ' endDate (String): Data de término no formato YYYY-MM-DD.
    ' MEASURES (Array): Array de medidas específicas a serem recuperadas.
    ' Retorna:
    ' Um array de dados ou uma mensagem de erro se os dados não puderem ser recuperados.
End Function
```

### GetDataPoint

Recupera um ponto de dado específico para um TICKER, data e medida fornecidos.

```vba
Public Function GetDataPoint(ticker As String, dateInput As String, measure As String, Optional extraParam As String = "") As Variant
    ' Recupera um ponto de dado específico para um TICKER, data e medida fornecidos.
    ' Parâmetros:
    ' ticker (String): O identificador único para o conjunto de dados.
    ' dateInput (String): Data no formato YYYY-MM-DD.
    ' measure (String): A medida específica a ser recuperada.
    ' extraParam (String, opcional): Parâmetro adicional para a solicitação.
    ' Retorna:
    ' O valor específico da medida ou uma mensagem de erro se os dados não puderem ser recuperados.
End Function
```

### GetFields

Obtém os campos disponíveis para um TICKER fornecido, excluindo campos predefinidos.

```vba
Public Function GetFields(cod As String) As String
    ' Obtém os campos disponíveis para um TICKER fornecido, excluindo campos predefinidos.
    ' Parâmetros:
    ' cod (String): O identificador único para o conjunto de dados.
    ' Retorna:
    ' Uma string formatada com os campos disponíveis ou uma mensagem de erro se os dados não puderem ser recuperados.
End Function
```

### GetAverage

Calcula a média de uma medida específica para um TICKER dentro de um intervalo de datas.

```vba
Public Function GetAverage(ticker As String, measure As String, startDate As String, endDate As String) As Variant
    ' Calcula a média de uma medida específica para um TICKER dentro de um intervalo de datas.
    ' Parâmetros:
    ' ticker (String): O identificador único para o conjunto de dados.
    ' measure (String): A medida específica a ser calculada.
    ' startDate (String): Data de início no formato YYYY-MM-DD.
    ' endDate (String): Data de término no formato YYYY-MM-DD.
    ' Retorna:
    ' A média da medida ou uma mensagem de erro se os dados não puderem ser recuperados.
End Function
```

### GetSum

Calcula a soma de uma medida específica para um TICKER dentro de um intervalo de datas.

```vba
Public Function GetSum(ticker As String, measure As String, startDate As String, endDate As String) As Variant
    ' Calcula a soma de uma medida específica para um TICKER dentro de um intervalo de datas.
    ' Parâmetros:
    ' ticker (String): O identificador único para o conjunto de dados.
    ' measure (String): A medida específica a ser somada.
    ' startDate (String): Data de início no formato YYYY-MM-DD.
    ' endDate (String): Data de término no formato YYYY-MM-DD.
    ' Retorna:
    ' A soma da medida ou uma mensagem de erro se os dados não puderem ser recuperados.
End Function
```
