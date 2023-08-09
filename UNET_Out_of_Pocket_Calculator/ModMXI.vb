Imports Newtonsoft.Json

Module ModMXI
    'Sub get_mxiDetails() ''gatharing data from note pad if mxi Screen data is available 

    'strInputParameter = "MXI" & "|" & Form1.txt_Policy.Text & "|" & "0001" & "|" & ""

    'Call Input_Parameter(strInputParameter)

    'Call Tracking.Get_MXI()

    'If pmiFile = True Then

    '    For Each Line As String In IO.File.ReadLines(pmiTextFile)
    '        If InStr(Line, "Provider Name") Then
    '            strPROVName = Mid(Line, 18, Len(Line))
    '        End If
    '        If InStr(Line, "Provider SP Code") Then
    '            strSPCode = Mid(Line, 18, Len(Line))
    '        End If
    '    Next
    'End If


    ''' <summary>
    ''' Requires apiCommonFunctions module
    ''' Documentation at https://data-experience.optum.com/detail/RestAPI/de.explorer%2Fdata-asset-v2%2Fapi-56c3f7a1-9731-48b6-b2d1-e956bcc55178
    ''' Version 1.0
    ''' </summary>
    Public Class apiMXI

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
            Alpha
            Bravo
        End Enum

        'Token
        Private Const TokenURLProduction As String = "https://gateway-core.optum.com/auth/oauth2/cached/token"
        Private Const TokenURLTest As String = "https://gateway-stage-core.optum.com/auth/oauth2/cached/token"

        'Server URL
        Private Const ServerURLProduction As String = "https://gateway-core.optum.com/api/clm/tops-acura"
        Private Const ServerURLAlpha As String = "https://gateway-stage-core.optum.com/api/uata/clm/tops-acura"
        Private Const ServerURLBravo As String = "https://gateway-stage-core.optum.com/api/uatb/clm/tops-acura"

        'Operation (append to server URL)
        Private Const QueryOperation As String = "/external-ref-policies/v1"

#End Region

#Region "Functions"

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="Policy">The 6-digit group number, as found on the MRI screen.</param>
        ''' <param name="Plan">The 4-digit plan number, as found on the MRI (07/015) or CEI (10/002) screen. Runs *very* slowly without this.</param>
        ''' <param name="Clss">Railroad only. The 4-digit class number, as found on the MRI (07/020) or CEI (10/007)</param>
        ''' <returns>Returns a wrapper object with Results.MxiRows as the useful list.</returns>
        Friend Function PerformQuery(Policy As String, Optional Plan As String = "", Optional Clss As String = "") As mxiData

            Dim QueryURL As String = Query_URI.AbsoluteUri & "?pol=" & Policy
            If Plan <> "" Then QueryURL += "&plan=" & Plan
            If Clss <> "" Then QueryURL += "&clss=" & Clss

            Dim jsonResult As String = sendApiRequest(New Uri(QueryURL), Nothing, "application/json", "GET", AccessToken)

            If InStr(jsonResult, "error") > 0 Then
                Return New mxiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
            Else
                Return New mxiData With {.Results = JsonConvert.DeserializeObject(Of ReturnBlock)(jsonResult), .jsonResponse = jsonResult}
            End If

        End Function

#End Region

#Region "Classes"
        Public Class mxiData
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
            <JsonConverter(GetType(Converter(Of MxiRow)))> Public Property MxiRows As New List(Of MxiRow)
        End Class

        Public Class MxiRow
            Public Property aaportPct As String
            Public Property acnInd As String
            Public Property alldAnclInd As String
            Public Property appealLangInd As String
            Public Property autoAdjdInd As String
            Public Property baseCovSetNbr As String
            Public Property bnkAcctCd As String
            Public Property busSegPltfm As String
            Public Property cancDt As String
            Public Property cancRsnCd As String
            Public Property capXclsInd As String
            Public Property cchInd As String
            Public Property clssNbr As String
            Public Property clssPrtypNbr As String
            Public Property coreMedPrrAuthCd As String
            Public Property covTypCd As String
            Public Property dfltSrvcRcChrg As String
            Public Property ebdsSet2Id As String
            Public Property ebdsSetId As String
            Public Property effDt As String
            Public Property emergentWrpInd As String
            Public Property enrpDfltPct As String
            Public Property enrpEmrgFaclInd As String
            Public Property enrpErInd As String
            Public Property enrpNonErInd As String
            Public Property enrpNonErPct As String
            Public Property evdBasDialgInd As String
            Public Property franchCd As String
            Public Property freelookInd As String
            Public Property gtdHmoCd As String
            Public Property hcrEhbInd As String
            Public Property hraFamAcssptAmt As String
            Public Property hraIndAcssptAmt As String
            Public Property iplnTypCd As String
            Public Property jqCdReimPct As String
            Public Property lcaInd As String
            Public Property lglEntyCd As String
            Public Property lmtSrvcCd As String
            Public Property lstUpdtDttm As String
            Public Property lstUpdtUserId As String
            Public Property mailCd As String
            Public Property medcrCovSetNbr As String
            Public Property medctEstInd As String
            Public Property mmlCovSetNbr As String
            Public Property mnnrpCd As String
            Public Property mnnrpDmePct As String
            Public Property mnnrpLabPct As String
            Public Property mnnrpPct As String
            Public Property mnrpDfltPct As String
            Public Property mnrpPtPct As String
            Public Property nbSprsInd As String
            Public Property obligId As String
            Public Property optoutUbhtierInd As String
            Public Property payEnrleeCd As String
            Public Property payLoc1Nbr As String
            Public Property payLoc2Nbr As String
            Public Property pcsInd As String
            Public Property plnDedPrortInd As String
            Public Property plnNbr As String
            Public Property plnPrtypNbr As String
            Public Property plnSeqNbr As String
            Public Property polNbr As String
            Public Property polNmAdrInd As String
            Public Property polPrtypNbr As String
            Public Property prdctPlnClssCd As String
            Public Property prefPhrmIdcrdCd As String
            Public Property prefPhrmPrdctCd As String
            Public Property prortEvnt As String
            Public Property prtblCd As String
            Public Property qcareRptSelCd As String
            Public Property reimPolEdtInd As String
            Public Property rgnCd As String
            Public Property rptCdInd As String
            Public Property sfxPrtypCd As String
            Public Property shrArngCd As String
            Public Property srcSysCd As String
            Public Property stdPlnClssNbr As String
            Public Property stdPlnPlnNbr As String
            Public Property stdPlnPolNbr As String
            Public Property tefraApplInd As String
            Public Property uhPremDesgCd As String
        End Class

#End Region

    End Class

    ' End Sub
End Module