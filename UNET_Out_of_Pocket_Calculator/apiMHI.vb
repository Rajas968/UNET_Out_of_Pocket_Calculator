Imports Newtonsoft.Json

''' <summary>
''' Version 1.0
''' Documentation at https://data-experience.optum.com/detail/RestAPI/de.explorer%2Fdata-asset-v2%2Fapi-6490b6b1-fc21-43b3-91f0-024f70e772c5
''' Requires apiCommonFunctions module
''' </summary>
Public Class apiMHI

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
            Case Server.Release
                TokenSource = DomainTest & TokenServer
                Query_URI = New Uri(DomainProduction & APIServerRelease & OperationQuery)
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
        Release
        Alpha
        Bravo
    End Enum

    'Domain
    Private Const DomainProduction As String = "https://gateway-core.optum.com"
    Private Const DomainTest As String = "https://gateway-stage-core.optum.com"

    'Token Servers
    Private Const TokenServer As String = "/auth/oauth2/cached/token"

    'API Servers
    Private Const APIServerProduction As String = "/api/clm/tops-acura"
    Private Const APIServerRelease As String = "/api/rlse/clm/tops-acura"
    Private Const APIServerAlpha As String = "/api/uata/clm/tops-acura"
    Private Const APIServerBravo As String = "/api/uatb/clm/tops-acura"

    'Operation
    Private Const OperationQuery As String = "/tops-history-claims/v1"

#End Region

#Region "Functions"

    Friend Function PerformQuery(Policy As String, EmployeeID As String, PatientFirstName As String, RelationshipCode As String, ICN As String, Optional DraftNumber As String = "") As mhiData

        Dim pr As New Post_Request With
            {.Request = New Post_Request.PostRequestData With
                {.reqRequiredFlds = New Post_Request.PostRequestData.RequiredFieldData With
                    {.reqSearchPolicy = Policy, .reqSearchEmpid = EmployeeID, .reqSearchEmpname = PatientFirstName, .reqSearchEmprelation = RelationshipCode, .reqSearchIcn = ICN, .reqSearchDraftNbr = DraftNumber
                    }
                }
            }

        Dim jsonResult As String = sendApiRequest(Query_URI, Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(pr)), "application/json", "POST", AccessToken)

        If InStr(jsonResult, "error") > 0 Then
            Return New mhiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
        Else
            Return New mhiData With {.Results = JsonConvert.DeserializeObject(Of Post_Response)(jsonResult), .jsonResponse = jsonResult}
        End If

    End Function

#End Region

#Region "Classes"

    Public Class mhiData
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
            Public Property UniqueServiceId As String = "NIADEMOAPP"
            Public Property reqRequiredFlds As RequiredFieldData
            Public Class RequiredFieldData
                Public Property reqSearchPolicy As String = String.Empty
                Public Property reqSearchEmpid As String = String.Empty
                Public Property reqSearchEmpname As String = String.Empty
                Public Property reqSearchEmprelation As String = String.Empty
                Public Property reqSearchDraftNbr As String = String.Empty
                Public Property reqSearchIcn As String = String.Empty
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
        Public Property rspComments As String
        Public Property rspPolicy As String
        Public Property rspEEID As String
        Public Property rspRelCD As String
        Public Property rspRegionInd As String
        Public Property rspEmpFirstName As String
        Public Property rspEmpLastName As String
        Public Property rspEmpAddress As String
        Public Property rspEmpCity As String
        Public Property rspEmpState As String
        Public Property rspEmpZip As String
        Public Property rspLifetimePaid As String
        Public Property rspLifetimePaid_string As String
        Public Property rspLifetimeRem As String
        Public Property rspLifetimeRem_string As String
        <JsonConverter(GetType(Converter(Of ClaimHeaderData)))> Public Property rspClaimHeader As New List(Of ClaimHeaderData)

        Public Class ClaimHeaderData
            Public Property rspPatientNbr As String
            Public Property rspWavInd As String
            Public Property rspSpr As String
            Public Property rspSpr_String As String
            Public Property rspHospDiscCd As String
            Public Property rspAoTheJob As String
            Public Property rspVendorInfo As String
            Public Property rspPolicyDeductible As String
            Public Property rspPolicyDeductible_string As String
            Public Property rspFamilyDeductible As String
            Public Property rspFamilyDeductible_string As String
            Public Property rspFamilyDeductibleDate As String
            Public Property rspDeductiblePeriodStartDt As String
            Public Property rspDeductiblePeriodEndDt As String
            Public Property rspDiag1 As String
            Public Property rspDiag2 As String
            Public Property rspIcd9Com1 As String
            Public Property rspDrg As String
            Public Property rspCreditReserveYear As String
            Public Property rspCob As String
            Public Property rspCOBCredResAmt As Long
            Public Property rspCOBCredResAmt_string As String
            Public Property rspCOBRemaining As String
            Public Property rspCOBRemaining_string As String
            Public Property rspClaimNbr As String
            Public Property rspDiagDescription As String
            Public Property rspCauseCode As String
            Public Property rspMediCredRsv As String
            Public Property rspMediCredRsv_string As String
            Public Property rspMediRemaining As String
            Public Property rspMediRemaining_string As String
            <JsonConverter(GetType(Converter(Of rsp6ModData)))> Public Property rsp6Mod As New List(Of rsp6ModData)
            <JsonConverter(GetType(Converter(Of rspFModData)))> Public Property rspFMod As New List(Of rspFModData)
            <JsonConverter(GetType(Converter(Of rspSModData)))> Public Property rspSMod As New List(Of rspSModData)
            <JsonConverter(GetType(Converter(Of rspNModData)))> Public Property rspNMod As New List(Of rspNModData)
            <JsonConverter(GetType(Converter(Of rspRModData)))> Public Property rspRMod As New List(Of rspRModData)
            <JsonConverter(GetType(Converter(Of rsp2ModData)))> Public Property rsp2Mod As New List(Of rsp2ModData)
            <JsonConverter(GetType(Converter(Of rspPoundModData)))> Public Property rspPoundMod As New List(Of rspPoundModData)
            <JsonConverter(GetType(Converter(Of rspDollarModData)))> Public Property rspDollarMod As New List(Of rspDollarModData)
            <JsonConverter(GetType(Converter(Of rspBModData)))> Public Property rspBMod As New List(Of rspBModData)

            <JsonConverter(GetType(Converter(Of rspWModData)))> Public Property rspWMod As New List(Of rspWModData)

            <JsonConverter(GetType(Converter(Of rspAtModData)))> Public Property rspAtMod As New List(Of rspAtModData)
            <JsonConverter(GetType(Converter(Of rspDModData)))> Public Property rspDMod As New List(Of rspDModData)
            <JsonConverter(GetType(Converter(Of rspIModData)))> Public Property rspIMod As New List(Of rspIModData)
            <JsonConverter(GetType(Converter(Of rsp9ModData)))> Public Property rsp9Mod As New List(Of rsp9ModData)
            <JsonConverter(GetType(Converter(Of rsp8ModData)))> Public Property rsp8Mod As New List(Of rsp8ModData)
            <JsonConverter(GetType(Converter(Of rspLineDataData)))> Public Property rspLineData As New List(Of rspLineDataData)
            <JsonConverter(GetType(Converter(Of rspTotalLine1Data)))> Public Property rspTotalLine1 As New List(Of rspTotalLine1Data)
            <JsonConverter(GetType(Converter(Of rspTotLine2Data)))> Public Property rspTotLine2 As New List(Of rspTotLine2Data)
            Public Class rsp6ModData
                Public Property rspPrefInd As String
            End Class
            Public Class rspFModData
                Public Property rspMedicareId As String
            End Class
            Public Class rspSModData
                Public Property rspClaimSourceCode As String
                Public Property rspBillType As String
                Public Property rspNmfPercent As String
                Public Property rspNmfPercent_string As String
                Public Property rspNmfAmount As String
                Public Property rspNmfAmount_string As String
            End Class
            Public Class rspNModData
                Public Property rspNonCleanInd As String
                Public Property rspNonCleanElements As New List(Of String)
                Public Property rspMoreNonCleanElements As String
            End Class
            Public Class rspRModData
                Public Property rspResubInd As String
                Public Property rspResubDate As String
            End Class
            Public Class rsp2ModData
                Public Property rspCheckInfoRecpt As String
                Public Property rspCheckInfoAge As String
                Public Property rspCheckInfoPmtTyp As String
                Public Property rspCheckInfoPpSt As String
                Public Property rspCheckInfoProduct As String
                Public Property rspCheckInfoVarId As String
                Public Property rspCheckInfoImdRls As String
                Public Property rspCheckInfoRsnCd As String
            End Class
            Public Class rspPoundModData
                Public Property rspProvTpsmCd As String
            End Class
            Public Class rspDollarModData
                Public Property rspProvBenInd As String
            End Class
            Public Class rspBModData
                Public Property rspBillNpiId As String
                Public Property rspRendNpild As String
                Public Property rspRefNpild As String
                Public Property rspAttnNpild As String
            End Class
            Public Class rspWModData
                Public Property rspPatNbrMsg As String
                Public Property rspPatNbrNum As String
            End Class
            Public Class rspAtModData
                Public Property rspProvSpecNum As String
                Public Property rspProvSpecCd As String
            End Class
            Public Class rspDModData
                Public Property rspSuffTotChg As String
                Public Property rspSuffTotChg_String As String
            End Class
            Public Class rspIModData
                Public Property rspDolResubStatus As String
                Public Property rspDolXrefIcn As String
            End Class
            Public Class rsp9ModData
                Public Property rsp835OrigIcn As String
                Public Property rsp835OrgDrft As String
                Public Property rsp835UpInd As String
            End Class
            Public Class rsp8ModData
                Public Property rsp835CxIcn As String
                Public Property rsp835CxSufx As String
                Public Property rsp835CxDate As String
                Public Property rsp835CxTime As String
                Public Property rsp835CxTrans As String
            End Class
            Public Class rspLineDataData
                Public Property rspPlaceOfServ As String
                Public Property rspServiceCd As String
                Public Property rspFirstDate As String
                Public Property rspLastDate As String
                Public Property rspNumber As String
                Public Property rspNumber_string As String
                Public Property rspOverrideCd As String
                Public Property rspPayeeCd As String
                Public Property rspProvPosNbr As String
                Public Property rspRemarkCd As String
                Public Property rspSanctionInd As String
                Public Property rspCharge As String
                Public Property rspCharge_string As String
                Public Property rspNotCovered As String
                Public Property rspNotCovered_string As String
                Public Property rspBaseCovered As String
                Public Property rspBaseCovered_string As String
                Public Property rspBaseDedAmt As String
                Public Property rspBaseDedAmt_string As String
                Public Property rspBaseDedDesc As String
                Public Property rspBasePct As String
                Public Property rspBaseAmt As String
                Public Property rspBaseAmt_string As String
                Public Property rspSuppAmt As String
                Public Property rspSuppAmt_string As String
                Public Property rspMmCoveredAmt As String
                Public Property rspMmCoveredAmt_string As String
                Public Property rspMmDedAmt As String
                Public Property rspMmDedAmt_string As String
                Public Property rspMmDedDesc As String
                Public Property rspMmPct As String
                Public Property rspMmAmt As String
                Public Property rspMmAmt_string As String
                Public Property rspCrAmt As String
                Public Property rspCrAmt_string As String
                <JsonConverter(GetType(Converter(Of rspDetailLine1Data)))> Public Property rspDetailLine1 As New List(Of rspDetailLine1Data)
                Public Class rspDetailLine1Data
                    Public Property rspSourceId As String
                    Public Property rspAuthNbr As String
                    Public Property rspProc As String
                    Public Property rspRemarkCode As String
                    Public Property rspProcCd As String
                    Public Property rspEbdRdlgyProcInd As String
                    Public Property rspUpdInd As String
                    Public Property rspEbdProcLvlCd As String
                    Public Property rspProcMatchCdU As String
                    Public Property rspSrvcProcCdU As String
                    Public Property rspSiteCareProcCd As String
                    Public Property rspLnOncChrgInd As String
                    Public Property rspLnGenTstChrgInd As String
                    Public Property rspSecOpinVendInd As String
                    Public Property rspSpclRxProcCd As String
                    Public Property rspPmSourceCd As String
                    Public Property rspPhyMedCatg As String
                    Public Property rspPhyMedThrpy As String
                    Public Property rspPhyMedCount As String
                    Public Property rspPhyMedMskInd As String
                    Public Property rspPhyMedMskCount As String
                    Public Property rspLineFinalCauseCd As String
                End Class
                <JsonConverter(GetType(Converter(Of rspDetailLine2Data)))> Public Property rspDetailLine2 As New List(Of rspDetailLine2Data)
                Public Class rspDetailLine2Data
                    Public Property rspClaimTypCd As String
                    Public Property rspServiceLnCd As String
                    Public Property rspServiceProcCd As String
                    Public Property rspProcMtchCdL As String
                    Public Property rspSrvcProcCdL As String
                End Class
                <JsonConverter(GetType(Converter(Of rsp1ModData)))> Public Property rsp1Mod As New List(Of rsp1ModData)
                Public Class rsp1ModData
                    Public Property rspHcrCatType As String
                    Public Property rspHcrCatName As String
                End Class
                <JsonConverter(GetType(Converter(Of rspHModData)))> Public Property rspHMod As New List(Of rspHModData)
                Public Class rspHModData
                    Public Property rspMhiEtReferDeny As String
                    Public Property rspMhiEtOlSpiTblId As String
                    Public Property rspMhiEtXriType As String
                End Class
                <JsonConverter(GetType(Converter(Of rspXModData)))> Public Property rspXMod As New List(Of rspXModData)
                Public Class rspXModData
                    Public Property rspPrmptPayStMsg As String
                    Public Property rspPrmptPayStNum As String
                End Class
                <JsonConverter(GetType(Converter(Of rspVModData)))> Public Property rspVMod As New List(Of rspVModData)
                Public Class rspVModData
                    Public Property rsOrigPOS As String
                End Class
                <JsonConverter(GetType(Converter(Of rspCModData)))> Public Property rspCMod As New List(Of rspCModData)
                Public Class rspCModData
                    Public Property rspCopayFam As String
                    Public Property rspCopayAmt As String
                    Public Property rspCopayAmt_string As String
                    Public Property rspCopayType As String
                End Class
                <JsonConverter(GetType(Converter(Of rspJModData)))> Public Property rspJMod As New List(Of rspJModData)
                Public Class rspJModData
                    Public Property rspPpmcCode As String
                    Public Property rspPmCode As String
                End Class
                <JsonConverter(GetType(Converter(Of rspPModData)))> Public Property rspPMod As New List(Of rspPModData)
                Public Class rspPModData
                    Public Property rspPnltyNtfyTypCd1 As String
                    Public Property rspPnltyAmt1 As String
                    Public Property rspPnltyAmt1_string As String
                    Public Property rspPnltyRemarkCd1 As String
                    Public Property rspPnltyBypassCd1 As String
                    Public Property rspAdjAmt1 As String
                    Public Property rspAdjAmt1_string As String
                    Public Property rspPnltyDays1 As String
                    Public Property rspPnltyMthd1 As String
                    Public Property rspPnltyNtfyTypCd2 As String
                    Public Property rsPnltyAmt2 As String
                    Public Property rsPnltyAmt2_string As String
                    Public Property rspPnltyRemarkCd2 As String
                    Public Property rspPnltyBypassCd2 As String
                    Public Property rspAdjAmt2 As String
                    Public Property rspAdjAmt2_string As String
                    Public Property rspPnltyDays2 As String
                    Public Property rspPnltyMthd2 As String
                End Class
                <JsonConverter(GetType(Converter(Of rspKModData)))> Public Property rspKMod As New List(Of rspKModData)
                Public Class rspKModData
                    Public Property rspSpltIndMsg As String
                End Class
                <JsonConverter(GetType(Converter(Of rspOModData)))> Public Property rspOMod As New List(Of rspOModData)
                Public Class rspOModData
                    Public Property rspSurgDiagInd As String
                End Class
                <JsonConverter(GetType(Converter(Of rspQModData)))> Public Property rspQMod As New List(Of rspQModData)
                Public Class rspQModData
                    Public Property rspNewOldSvcCdNum As String
                    Public Property rspNewOldSvcCdLvl As String
                    Public Property rspNewOldPrdCd As String
                End Class
                <JsonConverter(GetType(Converter(Of rsp4ModData)))> Public Property rsp4Mod As New List(Of rsp4ModData)
                Public Class rsp4ModData
                    Public Property rspDedCreditAmt As String
                    Public Property rspDedCreditAmt_string As String
                End Class
                <JsonConverter(GetType(Converter(Of rsp5ModData)))> Public Property rsp5Mod As New List(Of rsp5ModData)
                Public Class rsp5ModData
                    Public Property rsMednecRelSvcId As String
                    Public Property rsMednecRuleNbr As String
                    Public Property rsMednecRuleEffDt As String
                    Public Property rsMednecAuthNbr As String
                    Public Property rspMednecAuthSrcId As String
                End Class

                <JsonConverter(GetType(Converter(Of rsp6ModData)))> Public Property rsp6Mod As New List(Of rsp6ModData)
                Public Class rsp6ModData
                    Public Property rspNavigtTierCd As String
                End Class
                <JsonConverter(GetType(Converter(Of rsp7ModData)))> Public Property rsp7Mod As New List(Of rsp7ModData)
                Public Class rsp7ModData
                    Public Property rspEocIdentifier As String
                    Public Property rspEocTrgrIcn As String
                    Public Property rspEocTrgrSfx As String
                    Public Property rspEocTrgrLnNbr As String
                End Class
                <JsonConverter(GetType(Converter(Of rspZModData)))> Public Property rspZMod As New List(Of rspZModData)
                Public Class rspZModData
                    Public Property rspProvPrsRemarkCd As String
                End Class
                <JsonConverter(GetType(Converter(Of rspExclamationModData)))> Public Property rspExclamationMod As New List(Of rspExclamationModData)
                Public Class rspExclamationModData
                    Public Property rspPhycMedDetails As String
                End Class
                <JsonConverter(GetType(Converter(Of rspAModData)))> Public Property rspAMod As New List(Of rspAModData)
                Public Class rspAModData
                    Public Property rspProvMsgEobName As String
                End Class
                <JsonConverter(GetType(Converter(Of rspMMod_Bag)))> Public Property rspMMod As New List(Of rspMMod_Bag)
                Public Class rspMMod_Bag
                    <JsonConverter(GetType(Converter(Of rspMModData)))> Public Property rspModifierCdTbl As New List(Of rspMModData)
                    Public Class rspMModData
                        Public Property rspOriginalMod As String = String.Empty
                        Public Property rspTranslatedMod As String = String.Empty
                    End Class
                End Class
                Public Class rspLModData
                    Public Property rspClosureCdTbl As String
                End Class
                <JsonConverter(GetType(Converter(Of rsp3ModData)))> Public Property rsp3Mod As New List(Of rsp3ModData)
                Public Class rsp3ModData
                    Public Property rspOopTbl As String
                    Public Property rspOopAccumAmt As String
                    Public Property rspOopAccumAmt_string As String
                    Public Property rspOopAccumFamIndvCd As String
                End Class
            End Class
            Public Class rspTotalLine1Data
                Public Property rspProviderNbr As String
                Public Property rspMedNecInd As String
                Public Property rspDraftNbr As String
                Public Property rspDateProc As String
                Public Property rspAdjNbr As String
                Public Property rspRapInd As String
                Public Property rspRapRcl As String
                Public Property rspDcSt As String
                Public Property rspGtTyp As String
                Public Property rspTotCharge As String
                Public Property rspTotCharge_string As String
                Public Property rspTotPaid As String
                Public Property rspTotPaid_string As String
                Public Property rspTotProvAdj As String
                Public Property rspTotProvAdj_string As String
                Public Property rspTotPatResp As String
                Public Property rspTotPatResp_string As String
                Public Property rspFacContrMeth As String
                Public Property rspNewCobInd As String
                Public Property rsp835RelInd As String

            End Class
            Public Class rspTotLine2Data
                Public Property rspIcn As String
                Public Property rspIcnSuffix As String
                Public Property rspActPaidDt As String
                Public Property rspInterestInd As String
                Public Property rspFlimLocNbr As String
                Public Property rspFlnOff As String
                Public Property rspRptSuf As String
                Public Property rspRptAccnt As String
                Public Property rspPrsInd As String
                Public Property rspSIInd As String
            End Class
        End Class

    End Class

#End Region

End Class


