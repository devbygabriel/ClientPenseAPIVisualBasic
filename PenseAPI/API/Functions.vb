Imports System.IO
Imports System.Net
Imports System.Text
Imports Newtonsoft.Json

Module Functions

    Public baseUrl As String = ""
    Public accessToken As String = ""
    Public tokenValidity As Double
    Public accessKey As String = ""
    Public clientId As String = ""
    Public returnWithError As String = ""
    Public qrCodeUrl As String
    Public paymentId As Integer
    Public paymentStatus As String
    Public updateAt As DateTime

    Public Function Authentication(accessKey As String, clientId As String)

        Dim dictData As New Dictionary(Of String, String)

        'Construção do body que será enviado na requisição
        dictData.Add("accessKey", accessKey)
        dictData.Add("clientId", clientId)

        Dim webClient As New WebClient()
        Dim requestByte As Byte()
        Dim request As String
        Dim requestString() As Byte

        Try

            'Definição do tipo de requisição que será feita
            webClient.Headers("content-type") = "application/json"

            'Serialização dos dados
            requestString = Encoding.Default.GetBytes(JsonConvert.SerializeObject(dictData, Formatting.Indented))
            'Fazendo a requisição
            requestByte = webClient.UploadData(baseUrl & "/api/Auth", "post", requestString)
            'Pegando o retorno em string
            request = Encoding.Default.GetString(requestByte)

            'Desserializar retorno Json
            Dim getReturn As PenseAPI.ReturnAuthentication = JsonConvert.DeserializeObject(Of PenseAPI.ReturnAuthentication)(request)

            webClient.Dispose()

            'Obtendo o token retornado no Json
            accessToken = getReturn.accessToken

        Catch ex As WebException

            'Tratamento de erro
            Using stream = ex.Response.GetResponseStream()

                Using sr = New StreamReader(stream)
                    returnWithError = sr.ReadToEnd()
                End Using

            End Using

            Return False
            Exit Function

        End Try

        'Salva em milisegundos a data e hora da última atualização do token 
        tokenValidity = CLng(DateTime.UtcNow.Subtract(New DateTime(1970, 1, 1)).TotalMilliseconds)

        Return True

    End Function

    Private Function RefreshAuthentication()

        If tokenValidity >= (tokenValidity + (12 * 60000)) Then
            If Authentication(accessKey, clientId) = False Then
                Return False
                Exit Function
            End If
        End If

        Return True

    End Function

    Public Function CreateStore(open As String, close As String, companyName As String, document As String, externalReference As String, streetNumber As String, streetName As String, cityName As String, stateName As String, latitude As String, longitude As String, reference As String, tradeName As String)

        'Verifica se já se passaram 12 minutos ou mais desde a última requisição do token
        If RefreshAuthentication() = False Then
            Return False
            Exit Function
        End If

        'Construção do body que será enviado na requisição
        Dim quotationMarks As String = """"
        Dim parameters As String = ""

        parameters = "{
" & quotationMarks & "businessHours" & quotationMarks & ": " & "{
" & quotationMarks & "monday" & quotationMarks & ": " & "[
{
" & quotationMarks & "open" & quotationMarks & ": " & quotationMarks & open & quotationMarks & ",
" & quotationMarks & "close" & quotationMarks & ": " & quotationMarks & close & quotationMarks & "
}
],
" & quotationMarks & "tuesday" & quotationMarks & ": " & "[
{
" & quotationMarks & "open" & quotationMarks & ": " & quotationMarks & open & quotationMarks & ",
" & quotationMarks & "close" & quotationMarks & ": " & quotationMarks & close & quotationMarks & "
}
],
" & quotationMarks & "wednesday" & quotationMarks & ": " & "[
{
" & quotationMarks & "open" & quotationMarks & ": " & quotationMarks & open & quotationMarks & ",
" & quotationMarks & "close" & quotationMarks & ": " & quotationMarks & close & quotationMarks & "
}
],
" & quotationMarks & "sunday" & quotationMarks & ": " & "[
{
" & quotationMarks & "open" & quotationMarks & ": " & quotationMarks & open & quotationMarks & ",
" & quotationMarks & "close" & quotationMarks & ": " & quotationMarks & close & quotationMarks & "
}
],
" & quotationMarks & "saturday" & quotationMarks & ": " & "[
{
" & quotationMarks & "open" & quotationMarks & ": " & quotationMarks & open & quotationMarks & ",
" & quotationMarks & "close" & quotationMarks & ": " & quotationMarks & close & quotationMarks & "
}
],
" & quotationMarks & "friday" & quotationMarks & ": " & "[
{
" & quotationMarks & "open" & quotationMarks & ": " & quotationMarks & open & quotationMarks & ",
" & quotationMarks & "close" & quotationMarks & ": " & quotationMarks & close & quotationMarks & "
}
],
" & quotationMarks & "thursday" & quotationMarks & ": " & "[
{
" & quotationMarks & "open" & quotationMarks & ": " & quotationMarks & open & quotationMarks & ",
" & quotationMarks & "close" & quotationMarks & ": " & quotationMarks & close & quotationMarks & "
}
]
},
" & quotationMarks & "companyName" & quotationMarks & ": " & quotationMarks & companyName & quotationMarks & ",
" & quotationMarks & "document" & quotationMarks & ": " & quotationMarks & document & quotationMarks & ",
" & quotationMarks & "externalReference" & quotationMarks & ": " & quotationMarks & externalReference & quotationMarks & ",
" & quotationMarks & "location" & quotationMarks & ": " & "{
" & quotationMarks & "streetNumber" & quotationMarks & ": " & quotationMarks & streetNumber & quotationMarks & ",
" & quotationMarks & "streetName" & quotationMarks & ": " & quotationMarks & streetName & quotationMarks & ",
" & quotationMarks & "cityName" & quotationMarks & ": " & quotationMarks & cityName & quotationMarks & ",
" & quotationMarks & "stateName" & quotationMarks & ": " & quotationMarks & stateName & quotationMarks & ",
" & quotationMarks & "latitude" & quotationMarks & ": " & quotationMarks & latitude & quotationMarks & ",
" & quotationMarks & "longitude" & quotationMarks & ": " & quotationMarks & longitude & quotationMarks & ",
" & quotationMarks & "reference" & quotationMarks & ": " & quotationMarks & reference & quotationMarks & "
},
" & quotationMarks & "tradeName" & quotationMarks & ": " & quotationMarks & tradeName & quotationMarks & "
}"

        parameters = parameters.Replace(vbCrLf, "")

        Dim webClient As New WebClient()
        Dim requestByte As Byte()
        Dim request As String
        Dim requestString() As Byte

        Try

            'Definição do tipo de requisição que será feita
            webClient.Headers("content-type") = "application/json"
            'Passagem do token obtido na autenticação
            webClient.Headers("authorization") = "Bearer " & accessToken

            'Convertendo o body para bytes
            requestString = Encoding.Default.GetBytes(parameters)
            'Fazendo a requisição
            requestByte = webClient.UploadData(baseUrl & "/api/Config/Store", "post", requestString)
            'Pegando o retorno em string
            request = Encoding.Default.GetString(requestByte)

            webClient.Dispose()

        Catch ex As WebException

            'Tratamento de erro
            Using stream = ex.Response.GetResponseStream()

                Using sr = New StreamReader(stream)
                    returnWithError = sr.ReadToEnd()
                End Using

            End Using

            Return False
            Exit Function

        End Try

        Return True

    End Function

    Public Function RegisterPOS(externalReference As String, externalReferenceStore As String, name As String)

        'Verifica se já se passaram 12 minutos ou mais desde a última requisição do token
        If RefreshAuthentication() = False Then
            Return False
            Exit Function
        End If

        Dim dictData As New Dictionary(Of String, String)

        'Construção do body que será enviado na requisição
        dictData.Add("externalReference", externalReference)
        dictData.Add("externalReferenceStore", externalReferenceStore)
        dictData.Add("name", name)

        Dim webClient As New WebClient()
        Dim requestByte As Byte()
        Dim request As String
        Dim requestString() As Byte

        Try

            'Definição do tipo de requisição que será feita
            webClient.Headers("content-type") = "application/json"
            'Passagem do token obtido na autenticação
            webClient.Headers("authorization") = "Bearer " & accessToken

            'Serialização dos dados
            requestString = Encoding.Default.GetBytes(JsonConvert.SerializeObject(dictData, Formatting.Indented))
            'Fazendo a requisição
            requestByte = webClient.UploadData(baseUrl & "/api/Config/Pdv", "post", requestString)
            'Pegando o retorno em string
            request = Encoding.Default.GetString(requestByte)

            webClient.Dispose()

        Catch ex As WebException

            'Tratamento de erro
            Using stream = ex.Response.GetResponseStream()

                Using sr = New StreamReader(stream)
                    returnWithError = sr.ReadToEnd()
                End Using

            End Using

            Return False
            Exit Function

        End Try

        Return True

    End Function

    Public Function Payment(description As String, externalReference As String, value As String, wallet As String, store As String, pdv As String, items As String)

        'Verifica se já se passaram 12 minutos ou mais desde a última requisição do token
        If RefreshAuthentication() = False Then
            Return False
            Exit Function
        End If

        'Construção do body que será enviado na requisição
        Dim quotationMarks As String = """"
        Dim parameters As String = ""

        parameters = "{
" & quotationMarks & "description" & quotationMarks & ": " & quotationMarks & description & quotationMarks & ",
" & quotationMarks & "externalReference" & quotationMarks & ": " & quotationMarks & externalReference & quotationMarks & ",
" & quotationMarks & "value" & quotationMarks & ": " & quotationMarks & value & quotationMarks & ",
" & quotationMarks & "wallet" & quotationMarks & ": " & quotationMarks & wallet & quotationMarks & ",
" & items & "
" & quotationMarks & "store" & quotationMarks & ": " & quotationMarks & store & quotationMarks & ",
" & quotationMarks & "pdv" & quotationMarks & ": " & quotationMarks & pdv & quotationMarks & "
}"

        parameters = parameters.Replace(vbCrLf, "")

        Dim webClient As New WebClient()
        Dim requestByte As Byte()
        Dim request As String
        Dim requestString() As Byte

        Try

            'Definição do tipo de requisição que será feita
            webClient.Headers("content-type") = "application/json"
            'Passagem do token obtido na autenticação
            webClient.Headers("authorization") = "Bearer " & accessToken

            'Convertendo o body em bytes
            requestString = Encoding.Default.GetBytes(parameters)
            'Fazendo a requisição
            requestByte = webClient.UploadData(baseUrl & "/api/Payment", "post", requestString)
            'Pegando o retorno em string
            request = Encoding.Default.GetString(requestByte)

            'Desserializar retorno Json
            Dim getReturn As PenseAPI.ReturnPayment = JsonConvert.DeserializeObject(Of PenseAPI.ReturnPayment)(request)

            webClient.Dispose()

            'Obtendo o qrcode e o id da venda
            qrCodeUrl = getReturn.qrCodeUrl
            paymentId = getReturn.id

        Catch ex As WebException

            'Tratamento de erro
            Using stream = ex.Response.GetResponseStream()

                Using sr = New StreamReader(stream)
                    returnWithError = sr.ReadToEnd()
                End Using

            End Using

            Return False
            Exit Function

        End Try

        Return True

    End Function

    Public Function GetPaymentStatus(paymentId As Integer)

        'Verifica se já se passaram 12 minutos ou mais desde a última requisição do token
        If RefreshAuthentication() = False Then
            Return False
            Exit Function
        End If

        Dim webClient As New WebClient()
        Dim requestByte As Byte()
        Dim request As String

        Try

            'Definição do tipo de requisição que será feita
            webClient.Headers("content-type") = "application/json"
            'Passagem do token obtido na autenticação
            webClient.Headers("authorization") = "Bearer " & accessToken

            'Fazendo a requisição
            requestByte = webClient.DownloadData(baseUrl & "/api/Payment/" & paymentId)
            'Pegando o retorno em string
            request = Encoding.Default.GetString(requestByte)

            'Desserializar retorno Json
            Dim getReturn As PenseAPI.ReturnPaymentStatus = JsonConvert.DeserializeObject(Of PenseAPI.ReturnPaymentStatus)(request)

            'Obtendo status, qrcode e data de atualização da venda
            paymentStatus = getReturn.status
            qrCodeUrl = getReturn.qrCodeUrl
            updateAt = getReturn.updateAt

            webClient.Dispose()

        Catch ex As WebException

            'Tratamento de erro
            Using stream = ex.Response.GetResponseStream()

                Using sr = New StreamReader(stream)
                    returnWithError = sr.ReadToEnd()
                End Using

            End Using

            Return False
            Exit Function

        End Try

        Return True

    End Function

    Public Function GetPaymentStatusByExternalReference(externalReference As Integer)

        'Verifica se já se passaram 12 minutos ou mais desde a última requisição do token
        If RefreshAuthentication() = False Then
            Return False
            Exit Function
        End If

        Dim webClient As New WebClient()
        Dim requestByte As Byte()
        Dim request As String

        Try

            'Definição do tipo de requisição que será feita
            webClient.Headers("content-type") = "application/json"
            'Passagem do token obtido na autenticação
            webClient.Headers("authorization") = "Bearer " & accessToken

            'Fazendo a requisição
            requestByte = webClient.DownloadData(baseUrl & "/api/Payment/ByExternalReference/" & externalReference)
            'Pegando o retorno em string
            request = Encoding.Default.GetString(requestByte)

            'Desserializar retorno Json
            Dim getReturn As PenseAPI.ReturnPaymentStatus = JsonConvert.DeserializeObject(Of PenseAPI.ReturnPaymentStatus)(request)

            'Obtendo status, qrcode e data de atualização da venda
            paymentStatus = getReturn.status
            qrCodeUrl = getReturn.qrCodeUrl
            updateAt = getReturn.updateAt

            webClient.Dispose()

        Catch ex As WebException

            'Tratamento de erro
            Using stream = ex.Response.GetResponseStream()

                Using sr = New StreamReader(stream)
                    returnWithError = sr.ReadToEnd()
                End Using

            End Using

            Return False
            Exit Function

        End Try

        Return True

    End Function

    Public Function CancelPayment(paymentId As Integer)

        'Verifica se já se passaram 12 minutos ou mais desde a última requisição do token
        If RefreshAuthentication() = False Then
            Return False
            Exit Function
        End If

        Dim webClient As New WebClient()
        Dim request As String

        Try

            'Definição do tipo de requisição que será feita
            webClient.Headers("content-type") = "application/json"
            'Passagem do token obtido na autenticação
            webClient.Headers("authorization") = "Bearer " & accessToken

            'Fazendo a requisição
            request = webClient.UploadString(baseUrl & "/api/Payment/" & paymentId, "DELETE", "Cancelamento")

            webClient.Dispose()

        Catch ex As WebException

            'Tratamento de erro
            Using stream = ex.Response.GetResponseStream()

                Using sr = New StreamReader(stream)
                    returnWithError = sr.ReadToEnd()
                End Using

            End Using

            Return False
            Exit Function

        End Try

        Return True

    End Function

End Module