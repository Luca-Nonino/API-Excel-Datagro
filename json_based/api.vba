
'''' README:

''You MUST have the JsonConverter modula alongside this api module
' VBA-JSON v2.3.1
' (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
' JSON Converter for VBA

Public Sub RegisterGetData()
    Application.MacroOptions _
        Macro:="GetDataPoint", _
        Description:="Busca e filtra dados da fonte de dados com base no TICKER, ANO e MEDIDA fornecidos.", _
        Category:="Funções Personalizadas", _
        ArgumentDescriptions:=Array("TICKER - O identificador para o conjunto de dados.", _
                                    "ANO - O ano para o qual os dados devem ser recuperados. Se não especificado, busca os dados mais recentes.", _
                                    "MEDIDA - A medida específica a ser recuperada dos dados. Se não especificado, retorna todos os dados disponíveis.")
End Sub

Public Function GetInfo(ticker As String) As String
    ' Recupera informações gerais para um TICKER fornecido da fonte de dados.
    '
    ' Parâmetros:
    ' TICKER (String): O identificador único para o conjunto de dados para recuperar informações.
    '
    ' Retorna: 
    ' Uma string contendo as informações "longo" recuperadas ou uma mensagem de erro se os dados não puderem ser recuperados.
    Dim userEmail As String
    Dim userPassword As String
    userEmail = ThisWorkbook.Names("userEmail").RefersTo
    userPassword = ThisWorkbook.Names("userPassword").RefersTo
    Dim httpRequest As Object
    Dim url As String
    Dim jsonResponse As Object
    
    ' Constrói a URL com as credenciais do usuário para autenticação
    url = "https://precos.api.datagro.com/dados/?a=" & ticker & "&x=j&nome=" & userEmail & "&senha=" & userPassword

    ' Cria e envia a solicitação HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.Send

    ' Verifica o status da solicitação HTTP
    If httpRequest.Status <> 200 Then
        GetInfo = "Erro ao buscar dados: " & httpRequest.Status
        Exit Function
    End If

    ' Analisa a resposta JSON
    Set jsonResponse = JsonConverter.ParseJson(httpRequest.responseText)

    ' Extrai e retorna as informações "longo"
    If jsonResponse.count > 0 Then
        GetInfo = jsonResponse(1)("longo")
    Else
        GetInfo = "Nenhum dado disponível"
    End If
End Function

Public Function GetData(ticker As String, startDate As String, endDate As String, ParamArray MEASURES() As Variant) As Variant
    Dim userEmail As String
    Dim userPassword As String
    userEmail = ThisWorkbook.Names("userEmail").RefersTo
    userPassword = ThisWorkbook.Names("userPassword").RefersTo
    Dim httpRequest As Object
    Dim url As String
    Dim jsonResponse As Object
    Dim result As Collection
    Dim outputArray As Variant
    Dim headers() As String
    Dim measureHeaders() As String
    Dim measureData() As Variant
    Dim i As Long, j As Long, m As Long

    ' Verifica os parâmetros obrigatórios
    If ticker = "" Or startDate = "" Or endDate = "" Then
        GetData = "Ticker, startDate e endDate são parâmetros obrigatórios."
        Exit Function
    End If

    ' Valida os formatos de data (YYYY-MM-DD)
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        GetData = "Formato de data inválido. Por favor, use YYYY-MM-DD."
        Exit Function
    End If

    ' Converte datas para o formato YYYYMMDD para a URL
    Dim startDateFormatted As String
    Dim endDateFormatted As String
    startDateFormatted = Replace(startDate, "-", "")
    endDateFormatted = Replace(endDate, "-", "")

    ' Constrói a URL com base nos parâmetros
    url = "https://precos.api.datagro.com/dados/?a=" & ticker & "&x=j&nome=" & userEmail & "&senha=" & userPassword
    url = url & "&i=" & startDateFormatted & "&f=" & endDateFormatted

    ' Cria e envia a solicitação HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.Send

    ' Verifica o status da solicitação HTTP
    If httpRequest.Status <> 200 Then
        GetData = "Erro ao buscar dados: " & httpRequest.Status & " - " & httpRequest.statusText
        Exit Function
    End If

    ' Analisa a resposta JSON
    Set jsonResponse = JsonConverter.ParseJson(httpRequest.responseText)

    ' Prepara os cabeçalhos para o array de saída
    ' Sempre inclui "cod" e "dia"
    ReDim headers(1 To 2 + UBound(MEASURES) - LBound(MEASURES))
    headers(1) = "cod"
    headers(2) = "dia"
    For i = LBound(MEASURES) To UBound(MEASURES)
        headers(i + 2) = LCase(MEASURES(i))
    Next i

    ' Inicializa a coleção de resultados
    Set result = New Collection
    result.Add headers ' Adiciona cabeçalhos como a primeira linha

    ' Extrai os dados desejados da resposta JSON
    For Each Item In jsonResponse
        ' Prepara a linha para cada item
        ReDim measureData(1 To UBound(headers))
        measureData(1) = Item("cod")
        measureData(2) = Item("dia")
        For m = LBound(MEASURES) To UBound(MEASURES)
            measureHeaders = Split(LCase(MEASURES(m)), ",")
            ' Verifica cada medida e adiciona à linha
            For j = LBound(measureHeaders) To UBound(measureHeaders)
                measureData(m + 2) = IIf(Item.Exists(measureHeaders(j)), Item(measureHeaders(j)), "N/A")
            Next j
        Next m
        result.Add measureData ' Adiciona os dados da linha à coleção de resultados
    Next Item

    ' Converte a coleção para um array para saída
    If result.count > 1 Then
        ReDim outputArray(1 To result.count, 1 To UBound(headers))
        For i = 1 To result.count
            For j = 1 To UBound(headers)
                outputArray(i, j) = result(i)(j)
            Next j
        Next i
        GetData = outputArray
    Else
        GetData = "Nenhum dado encontrado para as medidas ou intervalo de datas especificados."
    End If
End Function

Public Function GetDataPoint(ticker As String, dateInput As String, measure As String, Optional extraParam As String = "") As Variant
    Dim httpRequest As Object
    Dim url As String
    Dim jsonResponse As Object
    Dim dateStr As String
    Dim result As Collection
    Dim outputArray As Variant
    Dim headers() As String
    Dim measureData() As Variant
    Dim i As Long, j As Long
    Dim userEmail As String
    Dim userPassword As String
    Dim periodo As String
    Dim uniqueParam As String
    userEmail = ThisWorkbook.Names("userEmail").RefersTo
    userPassword = ThisWorkbook.Names("userPassword").RefersTo
    
    On Error GoTo handler
    
    ' Imprime para depuração a recuperação das credenciais do usuário
    Debug.Print "Email do Usuário: " & userEmail
    Debug.Print "Senha do Usuário: " & userPassword
    
    ' Verifica os parâmetros obrigatórios
    If ticker = "" Or dateInput = "" Or measure = "" Then
        GetDataPoint = "Todos os parâmetros são obrigatórios."
        Exit Function
    End If

    ' Valida o formato da data (YYYY-MM-DD)
    If Not IsDate(dateInput) Then
        GetDataPoint = "Formato de data inválido. Por favor, use YYYY-MM-DD."
        Exit Function
    End If

    ' Converte a data para o formato YYYYMMDD para a URL
    dateStr = Replace(dateInput, "-", "")

    ' Determina o período com base no comprimento de dateStr
    Select Case Len(dateStr)
        Case 8 ' Dados diários
            periodo = "d"
        Case Else
            GetDataPoint = "Formato de data inválido."
            Exit Function
    End Select
    
    ' Gera um parâmetro único para evitar cache
    uniqueParam = "&timestamp=" & CStr(Timer) ' Adiciona o tempo atual em segundos desde a meia-noite
    
    ' Constrói a URL com base nos parâmetros
    url = "https://precos.api.datagro.com/dados/?a=" & ticker & "&i=" & dateStr & "&x=j&nome=" & userEmail & "&senha=" & userPassword & "&p=" & periodo & uniqueParam
    If extraParam <> "" Then
        url = url & "&b=" & extraParam
    End If

    ' Imprime para depuração a URL construída
    Debug.Print "URL: " & url

    ' Cria e envia a solicitação HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.setRequestHeader "Cache-Control", "no-cache, no-store, must-revalidate"
    httpRequest.setRequestHeader "Pragma", "no-cache"
    httpRequest.setRequestHeader "Expires", "0"

    httpRequest.Send

    ' Verifica o status da solicitação HTTP
    If httpRequest.Status <> 200 Then
        GetDataPoint = "Erro ao buscar dados: " & httpRequest.Status & " - " & httpRequest.statusText
        Exit Function
    End If

    ' Analisa a resposta JSON
    Set jsonResponse = JsonConverter.ParseJson(httpRequest.responseText)

    ' Inicializa a coleção de resultados
    Set result = New Collection

    ' Extrai os dados desejados da resposta JSON
    For Each Item In jsonResponse
        ' Prepara a linha para cada item
        ReDim measureData(1 To 2) ' Supondo que apenas "cod" e a medida são necessários
        measureData(1) = Item("cod")
        measureData(2) = IIf(Item.Exists(measure), Item(measure), "N/A")
        ' Imprime para depuração os dados extraídos
        Debug.Print "Dados da medida: " & measureData(1) & ", " & measureData(2)
        result.Add measureData ' Adiciona os dados da linha à coleção de resultados
    Next Item

    If result.count > 0 Then
        ReDim outputArray(1 To result.count, 1 To 2)
        For i = 1 To result.count
            For j = 1 To 2
                outputArray(i, j) = result(i)(j)
                ' Imprime para depuração o conteúdo do array
                Debug.Print "Conteúdo do array em (" & i & ", " & j & "): " & outputArray(i, j)
            Next j
        Next i
        ' Retorna um valor específico em vez do array
        ' Por exemplo, retorna apenas a medida do primeiro item
        GetDataPoint = outputArray(1, 2) ' Ajuste os índices conforme necessário
    Else
handler:
        GetDataPoint = "Nenhum dado disponível para a data especificada."
    End If

End Function

' Certifique-se de ter adicionado uma referência a Microsoft XML, v6.0 (MSXML2) e importado JsonConverter.bas
' de https://github.com/VBA-tools/VBA-JSON no seu projeto para análise JSON.

Public Function GetFields(cod As String) As String
    Dim userEmail As String
    Dim userPassword As String
    userEmail = ThisWorkbook.Names("userEmail").RefersTo
    userPassword = ThisWorkbook.Names("userPassword").RefersTo
    Dim httpRequest As Object
    Dim jsonResponse As Object
    Dim field As Variant
    Dim result As String
    Dim url As String
    Dim skipFields As Object
    Set skipFields = CreateObject("Scripting.Dictionary")
    
    ' Lista de campos a serem ignorados
    skipFields.Add "idbolsa", True
    skipFields.Add "nome", True
    skipFields.Add "longo", True
    skipFields.Add "subproduto", True
    skipFields.Add "decimais", True
    skipFields.Add "correlatos", True
    skipFields.Add "rep", True
    skipFields.Add "freq", True
    skipFields.Add "flex", True
    skipFields.Add "bolsa", True
    skipFields.Add "cod", True
    skipFields.Add "dia", True
    skipFields.Add "chart1", True
    skipFields.Add "units", True

    ' Constrói a URL sem a data fixa para buscar todos os dados
     url = "https://precos.api.datagro.com/dados/?a=" & cod & "&x=j&nome=" & userEmail & "&senha=" & userPassword
    
    ' Cria e envia a solicitação HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.Send
    
    ' Analisa a resposta JSON
    Set jsonResponse = JsonConverter.ParseJson(httpRequest.responseText)
    
    ' Extraindo os nomes dos campos dinamicamente e ignorando os especificados
    If jsonResponse.count > 0 Then
        result = ""
        For Each field In jsonResponse(1).Keys
            If Not skipFields.Exists(field) Then
                If result <> "" Then
                    result = result & ","
                End If
                result = result & "'" & field & "'"
            End If
        Next field
    Else
        result = "Nenhum dado disponível"
    End If
    
    ' Retorna a string formatada
    GetFields = "[" & result & "]"
End Function

Public Function GetAverage(ticker As String, measure As String, startDate As String, endDate As String) As Variant
    Dim httpRequest As Object
    Dim url As String
    Dim jsonResponse As Object
    Dim total As Double
    Dim count As Long
    Dim userEmail As String
    Dim userPassword As String
    Dim dateStr As String
    Dim periodo As String

    userEmail = ThisWorkbook.Names("userEmail").RefersTo
    userPassword = ThisWorkbook.Names("userPassword").RefersTo

    On Error GoTo handler

    ' Valida entrada
    If ticker = "" Or measure = "" Or startDate = "" Or endDate = "" Then
        GetAverage = "Ticker, measure, startDate e endDate são parâmetros obrigatórios."
        Exit Function
    End If

    ' Valida os formatos de data (YYYY-MM-DD)
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        GetAverage = "Formato de data inválido. Por favor, use YYYY-MM-DD."
        Exit Function
    End If

    ' Converte datas para o formato YYYYMMDD para a URL
    Dim startDateFormatted As String
    Dim endDateFormatted As String
    startDateFormatted = Replace(startDate, "-", "")
    endDateFormatted = Replace(endDate, "-", "")

    ' Constrói a URL com base nos parâmetros
    url = "https://precos.api.datagro.com/dados/?a=" & ticker & "&i=" & startDateFormatted & "&f=" & endDateFormatted & "&x=j&nome=" & userEmail & "&senha=" & userPassword

    ' Cria e envia a solicitação HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.Send

    ' Verifica o status da solicitação HTTP
    If httpRequest.Status <> 200 Then
        GetAverage = "Erro ao buscar dados: " & httpRequest.Status & " - " & httpRequest.statusText
        Exit Function
    End If

    ' Analisa a resposta JSON
    Set jsonResponse = JsonConverter.ParseJson(httpRequest.responseText)

    ' Calcula o total e a contagem para a média
    total = 0
    count = 0
    For Each Item In jsonResponse
        If Item.Exists(measure) Then
            total = total + Item(measure)
            count = count + 1
        End If
    Next Item

    If count > 0 Then
        GetAverage = total / count
    Else
        GetAverage = "Nenhum dado disponível para a medida ou intervalo de datas especificados."
    End If

    Exit Function

handler:
    GetAverage = "Ocorreu um erro ao buscar dados."
End Function

Public Function GetSum(ticker As String, measure As String, startDate As String, endDate As String) As Variant
    Dim httpRequest As Object
    Dim url As String
    Dim jsonResponse As Object
    Dim total As Double
    Dim userEmail As String
    Dim userPassword As String

    userEmail = ThisWorkbook.Names("userEmail").RefersTo
    userPassword = ThisWorkbook.Names("userPassword").RefersTo

    On Error GoTo handler

    ' Valida entrada
    If ticker = "" Or measure = "" Or startDate = "" Or endDate = "" Then
        GetSum = "Ticker, measure, startDate e endDate são parâmetros obrigatórios."
        Exit Function
    End If

    ' Valida os formatos de data (YYYY-MM-DD)
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        GetSum = "Formato de data inválido. Por favor, use YYYY-MM-DD."
        Exit Function
    End If

    ' Converte datas para o formato YYYYMMDD para a URL
    Dim startDateFormatted As String
    Dim endDateFormatted As String
    startDateFormatted = Replace(startDate, "-", "")
    endDateFormatted = Replace(endDate, "-", "")

    ' Constrói a URL com base nos parâmetros
    url = "https://precos.api.datagro.com/dados/?a=" & ticker & "&i=" & startDateFormatted & "&f=" & endDateFormatted & "&x=j&nome=" & userEmail & "&senha=" & userPassword

    ' Cria e envia a solicitação HTTP
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", url, False
    httpRequest.Send

    ' Verifica o status da solicitação HTTP
    If httpRequest.Status <> 200 Then
        GetSum = "Erro ao buscar dados: " & httpRequest.Status & " - " & httpRequest.statusText
        Exit Function
    End If

    ' Analisa a resposta JSON
    Set jsonResponse = JsonConverter.ParseJson(httpRequest.responseText)

    ' Calcula o total
    total = 0
    For Each Item In jsonResponse
        If Item.Exists(measure) Then
            total = total + Item(measure)
        End If
    Next Item

    If total <> 0 Then
        GetSum = total
    Else
        GetSum = "Nenhum dado disponível para a medida ou intervalo de datas especificados."
    End If

    Exit Function

handler:
    GetSum = "Ocorreu um erro ao buscar dados."
End Function
