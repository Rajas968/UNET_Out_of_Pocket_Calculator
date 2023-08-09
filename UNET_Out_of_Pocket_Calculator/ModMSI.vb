Imports Newtonsoft.Json

''' <summary>
''' Requires apiCommonFunctions module.
''' Documentation at https://data-experience.optum.com/detail/RestAPI/de.explorer%2Fdata-asset-v2%2Fapi-939f2dd3-240b-4db9-a275-bc02b7566933
''' Version 1.0
''' </summary>
Public Class apiMSI

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
                TokenSource = DomainProduction & TokenServer
                Query_URI = New Uri(DomainProduction & APIServerProduction & OperationQuery)
            Case Server.Alpha
                TokenSource = DomainTest & TokenServer
                Query_URI = New Uri(DomainProduction & APIServerAlpha & OperationQuery)
            Case Server.Bravo
                TokenSource = DomainTest & TokenServer
                Query_URI = New Uri(DomainProduction & APIServerBravo & OperationQuery)
        End Select
    End Sub

    Enum Server
        Production
        Alpha
        Bravo
    End Enum

    'Domain
    Private Const DomainProduction As String = "https://gateway-core.optum.com"
    Private Const DomainTest As String = "https://gateway-stage-core.optum.com"

    'Token Server
    Private Const TokenServer As String = "/auth/oauth2/cached/token"

    'API Servers
    Private Const APIServerProduction As String = "/api/clm/tops-acura"
    Private Const APIServerAlpha As String = "/api/uata/clm/tops-acura"
    Private Const APIServerBravo As String = "/api/uatb/clm/tops-acura"

    'Operation (append to server URL)
    Private Const OperationQuery As String = "/policy-medical-services/v1"

#End Region

#Region "Functions"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Policy">From MXI</param>
    ''' <param name="Plan">From MXI</param>
    ''' <param name="Clss">Railroad only. From MXI.</param>
    ''' <returns></returns>
    Friend Function PerformQuery(Policy As String, Plan As String, Optional Clss As String = "") As msiData

        Dim QueryURL As String = Query_URI.AbsoluteUri & "?pol=" & Policy & "&plan=" & Plan
        'If Plan <> "" Then QueryURL += "&plan=" & Plan
        If Clss <> "" Then QueryURL += "&clss=" & Clss

        Dim jsonResult As String = sendApiRequest(New Uri(QueryURL), Nothing, "application/json", "GET", AccessToken)

        If InStr(jsonResult, "error") > 0 Then
            Return New msiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
        Else

            Return New msiData With {.Results = JsonConvert.DeserializeObject(Of ReturnBlock)(jsonResult), .jsonResponse = jsonResult}

        End If

    End Function

#End Region

#Region "Classes"

    Public Class msiData
        Public Property Results As ReturnBlock
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

    Public Class ReturnBlock
        <JsonConverter(GetType(Converter(Of MsiRow)))> Public Property MsiRows As New List(Of MsiRow)
    End Class

    Public Class MsiRow
        Public Property altTierCalcInd As String
        Public Property autopayInd As String
        Public Property baseCalcCd As String
        Public Property baseCalcDesc As String
        Public Property baseFct As String
        Public Property baseRcSchedNbr As String
        Public Property cancDt As String
        Public Property causCd As String
        Public Property dedCredTypCd As String
        Public Property dedDescBaseCd As String
        Public Property dedDescMmCd As String
        Public Property effDt As String
        Public Property lstUpdtAdjId As String
        Public Property lstUpdtDt As String
        Public Property lvlInd As String
        Public Property mmCalcCd As String
        Public Property mmCalcDesc As String
        Public Property mmFct As String
        Public Property mmRcSchedNbr As String
        Public Property newCoinsApplCd As String
        Public Property parsInd As String
        Public Property pendCd As String
        Public Property posCd As String
        Public Property posTierTypeCd As String
        Public Property protoNmTxt As String
        Public Property rcCd As String
        Public Property rmrkCd As String
        Public Property seqNbr As String
        Public Property spiInd1Cd As String
        Public Property spiInd234Cd As String
        Public Property srvcCatgyBaseCd As String
        Public Property srvcCatgyCd As String
        Public Property srvcCd As String
        Public Property srvcCdNbr As String
        Public Property srvcSetCd As String
        Public Property ssoRelSrvcCd As String
        Public Property suppApplCd As String
    End Class

#End Region

End Class
