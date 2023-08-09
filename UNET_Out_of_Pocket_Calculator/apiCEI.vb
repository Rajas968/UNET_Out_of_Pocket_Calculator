Imports Newtonsoft.Json

''' <summary>
''' Version 1.0
''' Requires apiCommonFunctions module
''' Documentation at https://data-experience.optum.com/detail/RestAPI/de.explorer%2Fdata-asset-v2%2Fapi-9e243bd1-c4cc-4382-8368-5bffb54d23e2
''' </summary>
Public Class apiCEI

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
    Private Const QueryOperation As String = "/tops-member-coverage-reads/v1"

#End Region

#Region "Functions"

    Friend Function PerformQuery(Policy As String, EmployeeID As String, UserID As String, Password As String) As ceiData

        Dim pr As New Post_Request With
            {.Request = New Post_Request.PostRequestData With
                {.reqRequiredFlds = New Post_Request.PostRequestData.ReqRequiredFldsData With
                    {.reqSearchPolicy = Policy, .reqSearchEmpid = EmployeeID, .reqSearchID = UserID, .reqSearchPd = Password}
                }
            }

        Dim jsonResult As String = sendApiRequest(Query_URI, Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(pr)), "application/json", "POST", AccessToken)


        If InStr(jsonResult, "error") > 0 Then
            Return New ceiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
        Else
            Return New ceiData With {.Results = JsonConvert.DeserializeObject(Of Post_Response)(jsonResult), .jsonResponse = jsonResult}
        End If
    End Function

#End Region

#Region "Classes"

    Public Class ceiData
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

    Private Class Post_Request
        Public Property Request As PostRequestData
        Public Class PostRequestData
            Public Property ApiConsumer As String = "NIA_Automations"
            Public Property TypeOfService As String
            Public Property LoggingLevel As String
            Public Property UniqueServiceId As String = "NIADEMOAPP"
            Public Property ServiceInstanceId As String
            Public Property reqRequiredFlds As ReqRequiredFldsData
            Public Class ReqRequiredFldsData
                Public Property reqSearchPolicy As String
                Public Property reqSearchEmpid As String
                Public Property reqSearchID As String
                Public Property reqSearchPd As String
            End Class
        End Class
    End Class

    Public Class Post_Response
        <JsonConverter(GetType(Converter(Of ResponseData)))> Public Property Response As New List(Of ResponseData)
    End Class

    Public Class ResponseData
        Public Property stsReturnStatus As String
        Public Property stsLoc As String
        Public Property stsCode As String
        Public Property stsCodeTypeDesc As String
        Public Property stsAddtlInfo As String
        Public Property rspCustEligHdrData As RspCustEligHdrDataData
        <JsonConverter(GetType(Converter(Of RspErorWarnDataData)))> Public Property rspCustEligCovData As New List(Of RspCustEligCovDataData)
        <JsonConverter(GetType(Converter(Of RspErorWarnDataData)))> Public Property rspErorWarnData As New List(Of RspErorWarnDataData)

        Public Class RspCustEligHdrDataData
            Public Property rspCustEmpFname As String
            Public Property rspCustEmpLname As String
            Public Property rspCustOffName As String
            Public Property rspCustEmpOffNbrSys As RspCustEmpOffNbrSysData
            Public Property rspCustEmpAddr As String
            Public Property rspCustOffAdr1 As String
            Public Property rspCustLogoInd As String
            Public Property rspCustEmpCity As String
            Public Property rspCustEmpSt As String
            Public Property rspCustEmpZip As String
            Public Property rspCustEmpZipExt As String
            Public Property rspCustOffAdr2 As String
            Public Property rspCustCosmosDiv As String
            Public Property rspCustPolName As String
            Public Property rspCustOffCity As String
            Public Property rspCustOffState As String
            Public Property rspCustOffZip As String
            Public Property rspCustOffZipExt As String
            Public Property rspSuffixNbr As String
            Public Property rspPolHolderPhone As String
            Public Property rspGroupNbr As String
            Public Property rspCustNbr As String
            Public Property rspOffPhone As String
            Public Property rspAcctNbr As String
            Public Property rspSrcSysInd As String

            Public Class RspCustEmpOffNbrSysData
                Public Property rspCustEmpOffNbr As String
                Public Property rspCustEmpSys As String
            End Class

        End Class

        Public Class RspCustEligCovDataData
            Public Property rspCustFname As String
            Public Property rspCustRelCd As String
            Public Property rspCustDobMMDDYY As String
            Public Property rspCustSex As String
            Public Property rspCustPcpTinSuf As String
            Public Property rspCustPcpName As String
            Public Property rspCustAsgnVolun As String
            Public Property rspCustNetKey As String
            Public Property rspCustCurPlan As String
            Public Property rspCustCurRptCd As String
            Public Property rspCustCurEffDte As String
            Public Property rspCustCurCanDte As String
            Public Property rspCustCurProduct As String
            Public Property rspCustCurCoverage As String
            Public Property rspCustCurSetPrefix As String
            Public Property rspCustCurSet1 As String
            Public Property rspCustCurSet2 As String
            Public Property rspCustCurSet3 As String
            Public Property rspCustPcpMarket As String
            Public Property rspCustIndCovMktType As String
            Public Property rspCustPcpMpin As String
            Public Property rspCustImcsPlan As String
            Public Property rspCustExclRider As String
            Public Property rspCustPrvPlan As String
            Public Property rspCustPrvRptCd As String
            Public Property rspCustPrvEffDte As String
            Public Property rspCustPrvCanDte As String
            Public Property rspCustPrvProduct As String
            Public Property rspCustPrvCoverage As String
            Public Property rspCustPrvSetPrefix As String
            Public Property rspCustPrvSet1 As String
            Public Property rspCustPrvSet2 As String
            Public Property rspCustPrvSet3 As String
            Public Property rspCustMedCob As String
            Public Property rspCustHistory As String
            Public Property rspCustMediCareCd As String
            Public Property rspCustMedUpdInd As String
            Public Property rspCustPcpIpa As String
            Public Property rspCustCapPay As String
            Public Property rspCustCapModel As String
            Public Property rspCustPrvImcsPlan As String
        End Class

        Public Class RspErorWarnDataData
            Public Property rspErrWarn As String
        End Class

    End Class

#End Region

End Class
