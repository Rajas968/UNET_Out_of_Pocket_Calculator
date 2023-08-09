Imports Newtonsoft.Json

''' <summary>
''' Requires apiCommonFunctions module
''' Documentation at https://data-experience.optum.com/detail/RestAPI/de.explorer%2Fdata-asset-v2%2Fapi-5e2b09b7-6f68-4587-825b-34fc16c84a7d
''' Version 1.0
''' </summary>
Public Class apiPMI

#Region "Declarations"

    Public ReadOnly Property AccessToken As String
        Get
            Return tManager.FetchToken(TokenSource)
        End Get
    End Property

    Private Property Query_URI As Uri

    Private Property TokenSource As String

    Public Sub New(Optional env As Server = Server.Production)
        Select Case env
            Case Server.Production
                TokenSource = TokenURLProduction
                Query_URI = New Uri(ServerURLProduction & QueryOperation)
            Case Server.Release
                TokenSource = TokenURLTest
                Query_URI = New Uri(ServerURLRelease & QueryOperation)
            Case Server.Alpha
                TokenSource = TokenURLTest
                Query_URI = New Uri(ServerURLAlpha & QueryOperation)
            Case Server.Bravo
                TokenSource = TokenURLTest
                Query_URI = New Uri(ServerURLBravo & QueryOperation)
        End Select
    End Sub

    Enum Server
        Production
        Release
        Alpha
        Bravo
    End Enum

    'Token
    Private Const TokenURLProduction As String = "https://gateway-core.optum.com/auth/oauth2/cached/token"
    Private Const TokenURLTest As String = "https://gateway-stage-core.optum.com/auth/oauth2/cached/token"

    'Server URL
    Private Const ServerURLProduction As String = "https://gateway-core.optum.com/api/clm/tops-acura"
    Private Const ServerURLRelease As String = "https://gateway-stage-core.optum.com/api/rlse/clm/tops-acura"
    Private Const ServerURLAlpha As String = "https://gateway-stage-core.optum.com/api/uata/clm/tops-acura"
    Private Const ServerURLBravo As String = "https://gateway-stage-core.optum.com/api/uatb/clm/tops-acura"

    'Operation (append to server URL)
    Private Const QueryOperation As String = "/tops-provider-coverage-reads/v1"

#End Region

#Region "Functions"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="FullTIN">15 digits with prefix and suffix</param>
    ''' <returns></returns>
    Friend Function PerformQuery(FullTIN As String) As pmiData

        Dim pr As New Post_Request With
            {.Request = New Post_Request.PostRequestData With
                {.reqRequiredFlds = New Post_Request.PostRequestData.RequiredFieldData With
                    {.reqPrvTin = FullTIN
                    }
                }
            }

        Dim jsonResult As String = sendApiRequest(Query_URI, System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(pr)), "application/json", "POST", AccessToken)

        If InStr(jsonResult, "error") > 0 Then
            Return New pmiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
        Else
            Return New pmiData With {.Results = JsonConvert.DeserializeObject(Of Post_Response)(jsonResult), .jsonResponse = jsonResult}
        End If
    End Function

#End Region

#Region "Classes"

    Public Class pmiData
        Public Property Results As Post_Response
        Public jsonResponse As String
        Public IsError As Boolean = False
        Private _apiError
        Public Property apiError As apiError
            Get
                Return _apiError
            End Get
            Set(value As apiError)
                _apiError = value
                IsError = True
            End Set
        End Property
    End Class

    Public Class Post_Request
        Public Property Request As PostRequestData

        Public Class PostRequestData
            Public Property ApiConsumer As String = "NIA_Automations"
            Public Property UniqueServiceId As String = "NIADEMOAPP"
            Public Property reqRequiredFlds As RequiredFieldData
            Public Class RequiredFieldData
                Public Property reqRespCode As Integer
                Public Property reqSystem As String
                Public Property reqPrvTin As Long
                Public Property reqViewName As String
                Public Property reqViewVersion As String
            End Class
        End Class
    End Class

    Public Class Post_Response
        <JsonConverter(GetType(Converter(Of ResponseData)))> Property Response As New List(Of ResponseData)
    End Class

    Public Class ResponseData
        Public Property stsReturnStatus As String
        Public Property stsLoc As String
        Public Property stsCode As String
        Public Property stsCodeTypeDesc As String
        Public Property stsAddtlInfo As String
        Public Property rspPrvEpdLogicDel As String
        Public Property rspPrvEpdNumFill As Integer
        Public Property rspPrvType As String
        Public Property rspPrvFlag As String
        Public Property rspPrvStatCd As String
        Public Property rspPrvBulkPay As String
        Public Property rspPrvBulkPaySuff As Integer
        Public Property rspPrvMCode As String
        Public Property rspPrvBillAddr As RspPrvBillAddrData
        Public Property rspPrvEobName As String
        Public Property rspPrvFacId As String
        Public Property rspPrvHospCd As String
        Public Property rspPrvEmpAccInfo As String
        Public Property rspPrvZip As Integer
        Public Property rspPrvDateUpd As Integer
        Public Property rspPrvBankId As Integer
        Public Property rspPrvStatEffDate As Integer
        Public Property rspPrvPcpInd As String
        Public Property rspPrvSpCd As String
        Public Property rspPrvPayToInd As String
        Public Property rspPrv835HipaaInd As String
        Public Property rspPrvSpecCd As String
        Public Property rspPrvFacCd As String
        Public Property rspPrvEmc As String
        Public Property rspPrvEmcEffDt As Integer
        Public Property rspPrvEmcCancDt As Integer
        Public Property rspPrvVerDt As Integer
        Public Property rspPrvDoc As Long
        Public Property rspPrvCoalCd As String
        Public Property rspPrvPpoNo As String
        Public Property rspPrvHmpSsoCd As String
        Public Property rspPrvBatch As String
        Public Property rspPrvComment1 As String
        Public Property rspPrvComment2 As String
        Public Property rspPrvAddrInd As String
        Public Property rspPrvAddrState As Integer
        Public Property rspPrvMxCd As String
        Public Property rspPrvSaAddr As String
        Public Property rspPrvSaCity As String
        Public Property rspPrvSaSt As String
        Public Property rspPrvSaZip As Integer
        Public Property rspPrvCorpOwner As String
        Public Property rspPrvMpin As Integer
        Public Property rspPrvSaAddrReq As Integer
        Public Property rspPrvEftAcctNum As String
        Public Property rspPrvEftRoutingNum As String
        Public Property rspPrvEftCheckDigit As String
        Public Property rspPrvOpid As String
        Public Property rspPrvOfficeCode As String
        Public Property rspPrvTimestamp As Integer
        Public Property rspPrvMedicalReclCd As String

        Public Class RspPrvBillAddrData
            Public Property rspPrvName As String
            Public Property rspPrvName2 As String
            Public Property rspPrvaddr As String
            Public Property rspPrvCity As String
            Public Property rspPrvSt As String
        End Class

    End Class

#End Region

End Class
