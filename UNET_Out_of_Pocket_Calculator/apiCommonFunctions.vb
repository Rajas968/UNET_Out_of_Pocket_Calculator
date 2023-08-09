Imports Newtonsoft.Json
Imports System.Text
Imports System.Net
Imports System.IO
Imports System.DirectoryServices.ActiveDirectory

''' <summary>
''' Version 1.0
'''Requires Newtonsoft.Json from NuGet 
'''Namespace: ggmap.eas.aru.automation
''' </summary>
Public Module apiCommonFunctions
    Public tManager As New tokenManager

    Public Class tokenManager

        Private Tokens As New List(Of Token)
        Private Const grant As String = "client_credentials"
        Private Const ap_id As String = "b80df20c-7c91-79d0-df4a-a9b1d7dfedbe"
        Private Const dom_id As String = "MS"

        Private Class Token
            Property AccessToken As String
            Property TokenOrigin As String
            Property CreationTime As Date
            Property ExpirationTime As Date
            Sub New(at As String, url As String, expires As Integer)
                AccessToken = at
                TokenOrigin = url
                ExpirationTime = Now.AddSeconds(expires)
                'CreationTime = Now
            End Sub
        End Class

        Public Function FetchToken(tokenURL As String) As String

            Dim t As Token

            Try

                t = Tokens.Single(Function(x) x.TokenOrigin = tokenURL)

                If Now > t.CreationTime.AddMinutes(29) Then
                    Tokens.Remove(t)
                    t = GetNewToken(tokenURL)
                    Tokens.Add(t)
                    Return t.AccessToken
                Else
                    Return t.AccessToken
                End If

            Catch ex As Exception
                t = GetNewToken(tokenURL)
                't = GetDOC360Token(tokenURL)
                Tokens.Add(t)
                Return t.AccessToken
            End Try
        End Function


        Private Function GetNewToken(url As String) As Token

            Dim tokenURI As New Uri(url)
            Dim myToken As New TOKEN_POST With {
                .client_id = "WXqiz85rSAG5RYHtnBlxroQqO7FZHZAc",
                .client_secret = "h93oPbIL23Dz9QT5OVp4wfifkFv8uJLw",
                .grant_type = "client_credentials"}

            'Pass our token object into the json serializer (This will create our JSON string for us)
            Dim myTokenOutput As String = JsonConvert.SerializeObject(myToken)

            'Send the token object as the POST data for the request and retrieve the response (must be UTF8 encoded)
            Dim myResponse As String = sendApiRequest(tokenURI, Encoding.UTF8.GetBytes(myTokenOutput), "application/json", "POST")

            Dim TokenObj As TOKEN_RESPONSE = JsonConvert.DeserializeObject(Of TOKEN_RESPONSE)(myResponse)

            Dim t As New Token(TokenObj.access_token, url, TokenObj.expires_in)
            Return t

        End Function

        Private Function GetDOC360Token(url As String) As Token         '''purge history
            Dim user_id As String = GetMyMSID()
            Dim tokenURI As New Uri(url)
            Dim myToken As New TOKEN_POST_Doc360 With {.appId = ap_id, .domain = dom_id, .userId = user_id}

            'Pass our token object into the json serializer (This will create our JSON string for us)
            Dim myTokenOutput As String = JsonConvert.SerializeObject(myToken)

            'Send the token object as the POST data for the request and retrieve the response (must be UTF8 encoded)
            Dim myResponse As String = sendApiRequest(tokenURI, Encoding.UTF8.GetBytes(myTokenOutput), "application/json", "POST")

            ' MsgBox(tokenURI)

            Dim TokenObj As TOKEN_RESPONSE_Doc360 = JsonConvert.DeserializeObject(Of TOKEN_RESPONSE_Doc360)(myResponse)
            Dim tim As Int32
            tim = 40
            Dim t As New Token(TokenObj.token, url, tim)
            Return t

        End Function
        Private Class TOKEN_POST_Doc360
            Public Property appId As String
            Public Property domain As String
            Public Property userId As String
        End Class
        Public Function GetMyMSID() As String
            GetMyMSID = Environ("USERNAME")
        End Function
        Private Class TOKEN_RESPONSE_Doc360
            Public Property token As String

        End Class


        Private Class TOKEN_POST
            Public Property client_id As String
            Public Property client_secret As String
            Public Property grant_type As String
        End Class

        Private Class TOKEN_RESPONSE
            Public Property token_type As String
            Public Property access_token As String
            Public Property expires_in As String
        End Class

    End Class

    Public Function sendApiRequest(uri As Uri, jsonDataBytes As Byte(), contentType As String, method As String, Optional token As String = "") As String
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Dim response As String
        Dim request As WebRequest
        request = WebRequest.Create(uri)
        If Not jsonDataBytes Is Nothing Then
            request.ContentLength = jsonDataBytes.Length
        End If
        request.ContentType = contentType
        request.Method = method

        'If optional token is passed in, then we will attach it as the auth header
        If token <> "" Then
            request.PreAuthenticate = True
            request.Headers.Add("Authorization", "Bearer " & token)
        End If

        If method = "POST" Then
            Try
                Using requestStream = request.GetRequestStream
                    requestStream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
                    requestStream.Close()
                    Using responseStream = request.GetResponse.GetResponseStream
                        Using reader As New StreamReader(responseStream)
                            response = reader.ReadToEnd
                        End Using
                    End Using
                End Using
                Return response
            Catch ex As System.Net.WebException
                Return Strings.Replace(New StreamReader(ex.Response.GetResponseStream).ReadToEnd(), """error"":", """errorText"":")
            End Try
        ElseIf method = "GET" Then
            Try
                Using responseStream = request.GetResponse.GetResponseStream
                    Using reader As New StreamReader(responseStream)
                        response = reader.ReadToEnd
                    End Using
                End Using
                Return response
            Catch ex As System.Net.WebException
                Return Strings.Replace(New StreamReader(ex.Response.GetResponseStream).ReadToEnd(), """error"":", """errorText"":")
            End Try
        End If

        'EDIT: 1/21/22 JChoui3  
        Return ""
    End Function
    Public Class apiError
        Public Property timestamp As String
        Public Property status As String
        Public Property errorText As String
    End Class

    Public Class Converter(Of t)
        Inherits JsonConverter
        Public Overrides Sub WriteJson(ByVal writer As JsonWriter, ByVal value As Object, ByVal serializer As JsonSerializer)
            Throw New NotImplementedException()
        End Sub
        Public Overrides Function ReadJson(ByVal reader As JsonReader, ByVal objectType As Type, ByVal existingValue As Object, ByVal serializer As JsonSerializer) As Object
            Dim retVal As Object = New Object()

            If reader.TokenType = JsonToken.StartObject Then
                Dim instance As t = CType(serializer.Deserialize(reader, GetType(t)), t)
                retVal = New List(Of t)() From {instance}
            ElseIf reader.TokenType = JsonToken.StartArray Then
                retVal = serializer.Deserialize(reader, objectType)
            End If
            Return retVal
        End Function
        Public Overrides Function CanConvert(ByVal objectType As Type) As Boolean
            Return True
        End Function
    End Class

End Module
