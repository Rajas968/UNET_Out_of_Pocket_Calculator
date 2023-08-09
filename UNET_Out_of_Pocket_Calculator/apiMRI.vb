Imports Newtonsoft.Json

''' <summary>
''' Requires apiCommonFunctions module
''' Documentation at https://data-experience.optum.com/detail/RestAPI/de.explorer%2Fdata-asset-v2%2Fapi-89213018-7605-42c5-936e-8e91c0a5d702
''' Version 1.0
''' </summary>
Public Class apiMRI
    'Examples:
    '
    'Dim mriResp = API_MRI.PerformQuery(policy, accountNumber, userName, passWord)

    'For Each item In mriResp.Results.Response
    '   For Each policyLine In item.rspMRIinfo.rspMRICoverageLine
    '       With policyLine                                     'Coverage lines
    '           MsgBox(.rspCovPolicy & "/" & .rspCovPlan & "/" & .rspCovRept & ": " & .rspCovEffDT & " - " & .rspCovCanDT)
    '       End With
    '   Next policyLine

    '   MsgBox(item.rspMRIinfo.rspMRIEmpDsp.rspLastName)        'Employee last name
    '   With item.rspMRIinfo.rspMRIEmployeeCovInfo              'Subscriber info
    '       MsgBox(.rspEEFirstName & "/" & .rspEERelCd & ": " & .rspEEBthDt)
    '   End With

    '   For Each member In item.rspMRIinfo.rspMRIDepCovInfo     'Dep info
    '       With member
    '           MsgBox(.rspDPFirstName & "/" & .rspDPRelCd & ": " & .rspDPBthDt)
    '       End With
    '   Next member
    'Next item

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
    Private Const QueryOperation As String = "/tops-members/v1"

#End Region

#Region "Functions"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Policy">The policy number</param>
    ''' <param name="EmployeeID">10 bytes. The employee ID, preceded by a letter (typically S)</param>
    ''' <param name="unetID">10 bytes. UNET Login Username</param>
    ''' <param name="unetPass">8 bytes. UNET Login Password</param>
    ''' <param name="typeString">1 bytes. No idea. It works with ""</param>
    ''' <returns></returns>
    ''' 
    Friend Function PerformQuery(Policy As String, EmployeeID As String, unetID As String, unetPass As String, Optional typeString As String = "") As mriData

        Dim pr As New Post_Request With
        {.Request = New Post_Request.PostRequestData With
            {.reqRequiredFlds = New Post_Request.PostRequestData.RequiredFieldData With
                {.reqSearchPolicy = Policy, .reqSearchEmpid = EmployeeID, .reqSearchID = unetID, .reqSearchPd = unetPass, .reqtype = typeString}
            }
        }

        Dim jsonResult As String = sendApiRequest(Query_URI, Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(pr)), "application/json", "POST", AccessToken)

        If InStr(jsonResult, "error") > 0 Then
            Return New mriData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
        Else
            Return New mriData With {.Results = JsonConvert.DeserializeObject(Of Post_Response)(jsonResult), .jsonResponse = jsonResult}
        End If
    End Function

#End Region

#Region "Classes"
    Public Class mriData
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
            Public Property UniqueServiceId As String = "NIA_UNASSIGNED"
            Public Property reqRequiredFlds As RequiredFieldData
            Public Class RequiredFieldData

                Public Property reqSearchPolicy As String
                Public Property reqSearchEmpid As String
                Public Property reqSearchID As String
                Public Property reqSearchPd As String
                Public Property reqtype As String

            End Class
        End Class
    End Class

    Public Class Post_Response
        Public Property Response As ResponseData
    End Class

    Public Class ResponseData
        Public Property stsReturnStatus As String
        Public Property stsLoc As String
        Public Property stsCode As String
        Public Property stsCodeTypeDesc As String
        Public Property stsAddtlInfo As String
        Public Property rspMRIinfo As RspMRIinfo
        Public Property rspSFISCreenData As RspSFISCreenData
        Public Property rspESISCreenData As RspESISCreenData
        <JsonConverter(GetType(Converter(Of RspMessageData)))> Public Property rspMessageData As New List(Of RspMessageData)

    End Class
    Public Class RspMRIEmpDsp
        Public Property rspLastName As String
        Public Property rspAddress As String
        Public Property rspMisc As String
        Public Property rspCity As String
        Public Property rspState As String
        Public Property rspZip As String
        Public Property rspSocsecNbr As String
        Public Property rspEmplId As String
        Public Property rspDecDt As String
    End Class

    Public Class RspMRICoverageLine
        Public Property rspCovPolicy As String
        Public Property rspCovPlan As String
        Public Property rspCovRept As String
        Public Property rspCovCV As String
        Public Property rspCovEffDT As String
        Public Property rspCovCanDT As String
        Public Property rspCovNmAddress As String
        Public Property rspCovMO As String
    End Class

    Public Class RspMRIOthCovLine
        Public Property rspOthCovPlan As String
        Public Property rspOthCovRept As String
        Public Property rspOthCovEffDT As String
        Public Property rspOthCovCanDT As String
    End Class

    Public Class RspMRIEmployeeCovInfo
        Public Property rspEEFirstName As String
        Public Property rspEERelCd As String
        Public Property rspEEMedEffDt As String
        Public Property rspEEMedCanDt As String
        Public Property rspEESex As String
        Public Property rspEEBthDt As String
        Public Property rspEEMedicareCode As String
        Public Property rspEEMedicareEffDt As String
        Public Property rspEEMedicareCanDt As String
        Public Property rspEETefra As String
        Public Property rspEERestrictCd As String
        Public Property rspEERetStElig As String
        Public Property rspEECobDt As String
        Public Property rspEEMedUndInd As String
        Public Property rspEECobraInd As String
        Public Property rspEEPortInd As String
        Public Property rspEELateEnt As String
    End Class

    Public Class RspMRIDepCovInfo
        Public Property rspDPFirstName As String
        Public Property rspDPRelCd As String
        Public Property rspDPBthDt As String
        Public Property rspDPMedEffDt As String
        Public Property rspDPMedCanDt As String
        Public Property rspDPSex As String
        Public Property rspDPIndCobCode As String
        Public Property rspDPMedicareCode As String
        Public Property rspDPMedicareEffDt As String
        Public Property rspDPMedicareCanDt As String
        Public Property rspDPTefra As String
        Public Property rspDPRestrictCd As String
        Public Property rspDPRetStElig As String
        Public Property rspDPCobDt As String
        Public Property rspDPMedUndInd As String
        Public Property rspDPCobraInd As String
        Public Property rspDPPortInd As String
        Public Property rspDPLateEnt As String
        Public Property rspDepPhiind As String
        Public Property rspDepCesseqnbr As String
        Public Property rspDepdepnbr As String
        Public Property rspSFIAuditDate As String
        Public Property rspSFIMedLftYear As String
        Public Property rspSFIMedLftDol As String
        Public Property rspSFIMedLftCts As String
        Public Property rspSFIMedAdjDol As String
        Public Property rspSFIMedAdjCts As String
        Public Property rspSFIVsnLftYear As String
        Public Property rspSFIVsnLftDol As String
        Public Property rspSFIVsnLftCts As String
        Public Property rspSFIVsnAdjDol As String
        Public Property rspSFIVsnAdjCts As String
        Public Property rspSFIErrLftYear As String
        Public Property rspSFIErrLftDol As String
        Public Property rspSFIErrLftCts As String
        Public Property rspSFIErrAdjDol As String
        Public Property rspSFIErrAdjCts As String
        Public Property rspSFIRetLftYear As String
        Public Property rspSFIRetLftDol As String
        Public Property rspSFIRetLftCts As String
        Public Property rspSFIRetAdjDol As String
        Public Property rspSFIRetAdjCts As String
        Public Property rspMDILastName As String
        Public Property rspMDIDivIndC As String
        Public Property rspMDIClsIndC As String
        Public Property rspMDIEffDateC As String
        Public Property rspMDICanDateC As String
        Public Property rspMDIDivIndP1 As String
        Public Property rspMDIClsIndP1 As String
        Public Property rspMDIEffDateP1 As String
        Public Property rspMDICanDateP1 As String
        Public Property rspMDIDivIndP2 As String
        Public Property rspMDIClsIndP2 As String
        Public Property rspMDIEffDateP2 As String
        Public Property rspMDICanDateP2 As String
    End Class

    Public Class RspMRIinfo
        Public Property rspMRIEmpDsp As RspMRIEmpDsp
        <JsonConverter(GetType(Converter(Of RspMRICoverageLine)))> Public Property rspMRICoverageLine As New List(Of RspMRICoverageLine)
        <JsonConverter(GetType(Converter(Of RspMRIOthCovLine)))> Public Property rspMRIOthCovLine As New List(Of RspMRIOthCovLine)
        Public Property rspMRIEmployeeCovInfo As RspMRIEmployeeCovInfo
        <JsonConverter(GetType(Converter(Of RspMRIDepCovInfo)))> Public Property rspMRIDepCovInfo As New List(Of RspMRIDepCovInfo)
    End Class

    Public Class RspSFIFamilyData
        Public Property rspSFIFamDedYear As String
        Public Property rspSFIFamNOInd As String
        Public Property rspSFIFamDedAmt As String
        Public Property rspSFIFamDedCts As String
        Public Property rspSFIFamCreAmt As String
        Public Property rspSFIFamCreCts As String
        Public Property rspSFIFamCrAmt As String
        Public Property rspSFIFamCrCts As String
        Public Property rspSFIFamCOind As String
        Public Property rspSFIFamCOAmt As String
        Public Property rspSFIFamCOcts As String
        Public Property rspSFIFamDedYr As String
    End Class

    Public Class RspSFIFamily
        <JsonConverter(GetType(Converter(Of RspSFIFamilyData)))> Public Property rspSFIFamilyData As New List(Of RspSFIFamilyData)
    End Class

    Public Class RspSFISCreenData
        Public Property rspSFIFamily As RspSFIFamily
    End Class

    Public Class RspESISalandDed
        Public Property rspESIYear As String
        Public Property rspESIInnDedAmt As String
        Public Property rspESIInnNCAmt As String
        Public Property rspESIOonDedAmt As String
        Public Property rspESIOonNCAmt As String
    End Class

    Public Class RspESIEligCovData
        Public Property rspESIMbrCovCd As String
        Public Property rspESIMbrStartDt As String
    End Class

    Public Class RspESISCreenData
        <JsonConverter(GetType(Converter(Of RspESISalandDed)))> Public Property rspESISalandDed As New List(Of RspESISalandDed)
        <JsonConverter(GetType(Converter(Of RspESIEligCovData)))> Public Property rspESIEligCovData As New List(Of RspESIEligCovData)
    End Class

    Public Class RspMessageData
        Public Property rspMessage As String
    End Class


#End Region

End Class
