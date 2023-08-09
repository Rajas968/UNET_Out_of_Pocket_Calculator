Imports Newtonsoft.Json

''' <summary>
''' Requires apiCommonFunctions module
''' Documentation at https://data-experience.optum.com/detail/RestAPI/de.explorer%2Fdata-asset-v2%2Fapi-6ca0d0fc-677d-4865-97f0-d2c823d092b4
''' Version 1.0
''' </summary>
Public Class apiMMI

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
    Private Const QueryOperation As String = "/master-policies/v1"

#End Region

#Region "Functions"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Policy">NOT the group number. 6-digit stdPlnPolNbr from MXI (coordinates 06/002) </param>
    ''' <param name="Plan">4-digit stdPlnPlnNbr from MXI (coordinates 06/010)</param>
    ''' <param name="Clss">4-digit stdPlnClssNbr from MXI (coordinates 06/016)</param>
    ''' <param name="Page">1 or 2 digit MMI page. MMI1, MMI2, and so on.</param>
    ''' <returns></returns>
    Friend Function PerformQuery(Policy As String, Plan As String, Optional Clss As String = "", Optional Page As String = "") As mmiData

        Dim QueryURL As String = Query_URI.AbsoluteUri & "?"

        If Policy <> "" Then QueryURL += "&pol=" & Policy
        If Plan <> "" Then QueryURL += "&plan=" & Plan
        If Clss <> "" Then QueryURL += "&clss=" & Clss
        If Page <> "" Then QueryURL += "&page=" & Page

        Dim jsonResult As String = sendApiRequest(New Uri(QueryURL), Nothing, "application/json", "GET", AccessToken)

        If InStr(jsonResult, "error") > 0 Then
            Return New mmiData With {.apiError = JsonConvert.DeserializeObject(Of apiError)(jsonResult), .jsonResponse = jsonResult}
        Else
            Return New mmiData With {.Results = JsonConvert.DeserializeObject(Of MmiReturn)(jsonResult), .jsonResponse = jsonResult}
        End If

    End Function

#End Region

#Region "Classes"
    Public Class mmiData
        Public Property Results As MmiReturn
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

    Public Class MmiReturn
        Public Property mmi1Return As Mmi1Return
        Public Property mmi2Return As Mmi2Return
        Public Property mmi3Return As Mmi3Return
        Public Property mmi4Return As Mmi4Return
        Public Property mmi5Return As Mmi5Return
        Public Property mmi6Return As Mmi6Return
        Public Property mmi7Return As Mmi7Return
        Public Property mmi8Return As Mmi8Return
        Public Property mmi9Return As Mmi9Return
        Public Property mmi10Return As Mmi10Return
        Public Property mmi11Return As Mmi11Return
        Public Property mmi12Return As Mmi12Return
        Public Property mmi13Return As Mmi13Return
    End Class

#Region "MMI 1"

    Public Class Mmi1Return
        <JsonConverter(GetType(Converter(Of Mmi1ARows)))> Public Property mmi1ARows As New List(Of Mmi1ARows)
        <JsonConverter(GetType(Converter(Of Mmi1BRows)))> Public Property mmi1BRows As New List(Of Mmi1BRows)
        <JsonConverter(GetType(Converter(Of Mmi1CRows)))> Public Property mmi1CRows As New List(Of Mmi1CRows)
    End Class

    Public Class Mmi1ARows
        Public Property allncCd As String
        Public Property altClsCd As String
        Public Property altPrdCd As String
        Public Property appealLangInd As String
        Public Property autopayInd As String
        Public Property bnkAcctCd As String
        Public Property bulkEligCd As String
        Public Property bulkMailCd As String
        Public Property busSegPltfm As String
        Public Property cancDt As String
        Public Property caseIndCd As String
        Public Property cliSzCd As String
        Public Property clmPicTypCd As String
        Public Property cmMedcMirInd As String
        Public Property cobMcCd As String
        Public Property cobMoNbr As String
        Public Property cobPayPursCd As String
        Public Property cobStdCd As String
        Public Property contrStCd As String
        Public Property covTypCd As String
        Public Property crsEligInd As String
        Public Property cxiInd As String
        Public Property diagVsSurgInd As String
        Public Property dsesStTblNbr As String
        Public Property ebdsBasCovSetId As String
        Public Property ebdsCovBaseCd As String
        Public Property ebdsMcrCovSetVal As String
        Public Property ebdsMmeCovSetVal As String
        Public Property ebdsSet2Id As String
        Public Property effDt As String
        Public Property eligCd As String
        Public Property eligFmCd As String
        Public Property erInd As String
        Public Property erisaInd As String
        Public Property faclShrSvCd As String
        Public Property flTypCd As String
        Public Property fundTypCd As String
        Public Property gdrRuleCd As String
        Public Property hiDedPlnCd As String
        Public Property intgrCd As String
        Public Property lftmMaxAmt As String
        Public Property lftmMaxXapplInd As String
        Public Property lineBusCd As String
        Public Property lmtSrvcCd As String
        Public Property lscd As String
        Public Property mbrPrdctCd As String
        Public Property mbrPrdctType As String
        Public Property mcrCd As String
        Public Property mdlPolNm As String
        Public Property medMailCd As String
        Public Property medcrPartdCustElecCd As String
        Public Property newOldSrvcInd As String
        Public Property nonEmbdDedCd As String
        Public Property nonEmrgInd As String
        Public Property nptRulePkgId As String
        Public Property obligId As String
        Public Property offOfRgst As String
        Public Property oldNewEligInd As String
        Public Property onLineAddCd As String
        Public Property oonAccumCd As String
        Public Property oonLftmMaxAmt As String
        Public Property oonReinstAmt As String
        Public Property oonReinstFreqCd As String
        Public Property oopMaxCovCd As String
        Public Property payLoc2Nbr As String
        Public Property payLocCd As String
        Public Property pcpBenLvlRule As String
        Public Property plnFturInd As String
        Public Property plnNbr As String
        Public Property pnaCd As String
        Public Property polIdCd As String
        Public Property polNbr As String
        Public Property polTypCd As String
        Public Property polYrPlnDt As String
        Public Property ppoRaplInd As String
        Public Property prdctCdId As String
        Public Property prdctKeyCd As String
        Public Property prdctMcSelCd As String
        Public Property prefLabNtwkInd As String
        Public Property prsInd As String
        Public Property prvntBenCd As String
        Public Property psychDefXcptInd As String
        Public Property qaClmSelPct As String
        Public Property raplNoblxCasCd As String
        Public Property rgnCd As String
        Public Property rhaOptCd As String
        Public Property riskClsCd As String
        Public Property rptAmtNbr As String
        Public Property rptCdInd As String
        Public Property rxSbscrIdTypCd As String
        Public Property shrArngCd As String
        Public Property signOnOffNbr As String
        Public Property slotTblId As String
        Public Property spclProc10Cd As String
        Public Property spclProc1Cd As String
        Public Property spclProc2Cd As String
        Public Property spclProc3Cd As String
        Public Property spclProc4Cd As String
        Public Property spclProc5Cd As String
        Public Property spclProc6Cd As String
        Public Property spclProc7Cd As String
        Public Property spclProc8Cd As String
        Public Property spclProc9Cd As String
        Public Property srvcCatgyTblId As String
        Public Property srvcCdASetInd As String
        Public Property srvcCdBSetInd As String
        Public Property srvcCdCSetInd As String
        Public Property srvcCdDSetInd As String
        Public Property srvcCdESetInd As String
        Public Property srvcCdFSetInd As String
        Public Property srvcCdGSetInd As String
        Public Property srvcCdHSetInd As String
        Public Property srvcCdISetInd As String
        Public Property srvcCdJSetInd As String
        Public Property srvcCdKSetInd As String
        Public Property srvcCdLSetInd As String
        Public Property srvcCdMSetInd As String
        Public Property srvcCdNSetInd As String
        Public Property srvcCdOSetInd As String
        Public Property srvcCdPSetInd As String
        Public Property srvcCdQSetInd As String
        Public Property srvcCdRSetInd As String
        Public Property srvcCdSSetInd As String
        Public Property srvcCdTSetInd As String
        Public Property srvcCdUSetInd As String
        Public Property srvcCdVSetInd As String
        Public Property srvcCdWSetInd As String
        Public Property srvcCdXSetInd As String
        Public Property srvcCdYSetInd As String
        Public Property srvcCdZSetInd As String
        Public Property strctUpdtDt As String
        Public Property sviTblPriNbr As String
        Public Property sviTblTerNbr As String
        Public Property tefraApplInd As String
        Public Property termAgeNbr As String
        Public Property termCd As String
        Public Property tier1LftmAccumCd As String
        Public Property tier1LftmMaxAmt As String
        Public Property tier1ReinstAmt As String
        Public Property tier1ReinstFreqCd As String
        Public Property tier1UrgntCareOopCd As String
        Public Property uhpCd As String
        Public Property varPrdFct As String
        Public Property xrefPortDt As String
        Public Property xtraTerrStMandInd As String
    End Class
    Public Class Mmi1BRows
        Public Property combExchgTypCd As String
        Public Property combRmrkCd As String
        Public Property combVendCatgyCd As String
        Public Property combVendId As String
        Public Property covTypCd As String
    End Class

    Public Class Mmi1CRows
        Public Property covTypCd As String
        Public Property lftAccumCd As String
        Public Property ovragLftAmt As String
        Public Property ovragReinstAmt As String
        Public Property reinstAgeNbr As String
        Public Property reinstAmt As String
        Public Property reinstFreqCd As String
    End Class

#End Region
#Region "MMI 2"

    Public Class Mmi2Return
        <JsonConverter(GetType(Converter(Of Mmi2ARows)))> Public Property mmi2ARows As New List(Of Mmi2ARows)
        <JsonConverter(GetType(Converter(Of Mmi2BRows)))> Public Property mmi2BRows As New List(Of Mmi2BRows)
        <JsonConverter(GetType(Converter(Of Mmi2CRows)))> Public Property mmi2CRows As New List(Of Mmi2CRows)
    End Class

    Public Class Mmi2ARows
        Public Property covTypCd As String
        Public Property futEffPrxstCovAmt As String
        Public Property initEffPrxstCovAmt As String
        Public Property phrmCpnInd As String
        Public Property prxstXclsMoVal As String
        Public Property psyLftmCnfmMaxAmt As String
        Public Property psyLftmCombPrscMaxAmt As String
        Public Property psyLftmNonCnfmMaxAmt As String
        Public Property rxSbscrIdTypCd As String
        Public Property secOpinVendCd As String
        Public Property sovCharMtchNbr As String
        Public Property sovDisDedCd As String
        Public Property sovDisincntNcoinsInd As String
        Public Property sovDisincntPct As String
        Public Property sovIncntPct As String
        Public Property sovTblId As String
    End Class

    Public Class Mmi2BRows
        Public Property combExchgTypCd As String
        Public Property combRmrkCd As String
        Public Property combVendCatgyCd As String
        Public Property combVendId As String
        Public Property covTypCd As String
    End Class

    Public Class Mmi2CRows
        Public Property covTypCd As String
        Public Property errLftAccumCd As String
        Public Property errLftMaxAmt As String
        Public Property errReinstAmt As String
        Public Property errReinstFreqCd As String
        Public Property mmCausAmt As String
        Public Property mmCausPrdCd As String
        Public Property mmMaxAccumCd As String
        Public Property mmMaxAmt As String
        Public Property mmMaxPrdCd As String
        Public Property psyLftCnfnAccumCd As String
        Public Property psyLftCnfnReinstAmt As String
        Public Property psyLftCombAccumCd As String
        Public Property psyLftNcnfnAccumCd As String
        Public Property psyLftNcnfnReinstAmt As String
        Public Property psyReinstCombAmt As String
        Public Property psyReinstFreqCnfnCd As String
        Public Property psyReinstFreqCombCd As String
        Public Property psyReinstFreqNcnfnCd As String
        Public Property rrLftAccumCd As String
        Public Property rrLftMaxAmt As String
        Public Property rrReinstAmt As String
        Public Property rrReinstFreqCd As String
    End Class

#End Region
#Region "MMI 3"

    Public Class Mmi3Return
        <JsonConverter(GetType(Converter(Of Mmi3ARows)))> Public Property mmi3ARows As New List(Of Mmi3ARows)
        <JsonConverter(GetType(Converter(Of Mmi3BRows)))> Public Property mmi3BRows As New List(Of Mmi3BRows)
        <JsonConverter(GetType(Converter(Of Mmi3CRows)))> Public Property mmi3CRows As New List(Of Mmi3CRows)
    End Class

    Public Class Mmi3ARows
        Public Property alcCyCombMaxCovAmt As String
        Public Property alcCyIptntMaxCovAmt As String
        Public Property alcCyOptntMaxCovAmt As String
        Public Property alcLftmCombMaxCovAmt As String
        Public Property alcLftmIptntMaxCovAmt As String
        Public Property alcLftmOptntMaxCovAmt As String
        Public Property benAutsm2AgeYrFrVal As String
        Public Property benAutsm2AgeYrToVal As String
        Public Property benAutsm2AllCd As String
        Public Property benAutsm2PostCd As String
        Public Property benAutsmAgeYrFrVal As String
        Public Property benAutsmAgeYrToVal As String
        Public Property benAutsmAllCd As String
        Public Property benAutsmPostCd As String
        Public Property benMaxEhbAuralHabCnt As String
        Public Property benMaxEhbAuralRhabCnt As String
        Public Property benMaxEhbCogHabCnt As String
        Public Property benMaxEhbCogRhabCnt As String
        Public Property benMaxEhbCrdcRhabCnt As String
        Public Property benMaxEhbHabRmrkCd As String
        Public Property benMaxEhbOtHabCnt As String
        Public Property benMaxEhbOtRhabCnt As String
        Public Property benMaxEhbPlmryRhabCnt As String
        Public Property benMaxEhbPosmHabCnt As String
        Public Property benMaxEhbPosmRhabCnt As String
        Public Property benMaxEhbPostHabCnt As String
        Public Property benMaxEhbPostRhabCnt As String
        Public Property benMaxEhbPtHabCnt As String
        Public Property benMaxEhbPtRhabCnt As String
        Public Property benMaxEhbPtotHabCnt As String
        Public Property benMaxEhbPtotRhabCnt As String
        Public Property benMaxEhbRhabRmrkCd As String
        Public Property benMaxEhbSpchHabCnt As String
        Public Property benMaxEhbSpchRhabCnt As String
        Public Property benMaxEhbSpneHabCnt As String
        Public Property benMaxEhbSpneRhabCnt As String
        Public Property benPhysMedcnPrdCd As String
        Public Property chMinDaysNbr As String
        Public Property chMmYrsNbr As String
        Public Property chldCovCeaseCd As String
        Public Property chldCovCeasePrdCd As String
        Public Property chldMaxCovAgeVal As String
        Public Property covTypCd As String
        Public Property famCyrMaxAmt As String
        Public Property goldenRuleInd As String
        Public Property psyCyrCnfmMaxAmt As String
        Public Property psyCyrCombPrscMaxAmt As String
        Public Property psyCyrNonCnfmMaxAmt As String
        Public Property sdCovCessCd As String
        Public Property spsrDepMaxAgeVal As String
        Public Property stdntCovCeaseCd As String
        Public Property stdntCovCeasePrdCd As String
        Public Property stdntMaxAgeVal As String
        Public Property sviCanc1Dt As String
        Public Property sviCanc2Dt As String
        Public Property sviCanc3Dt As String
        Public Property sviCanc4Dt As String
        Public Property sviCanc5Dt As String
        Public Property sviDtTbl1Nbr As String
        Public Property sviDtTbl2Nbr As String
        Public Property sviDtTbl3Nbr As String
        Public Property sviDtTbl4Nbr As String
        Public Property sviDtTbl5Nbr As String
        Public Property sviEff1Dt As String
        Public Property sviEff2Dt As String
        Public Property sviEff3Dt As String
        Public Property sviEff4Dt As String
        Public Property sviEff5Dt As String
        Public Property sviTbl4Nbr As String
        Public Property sviTbl5Nbr As String
        Public Property sviTbl6Nbr As String
    End Class

    Public Class Mmi3BRows
        Public Property alcCyrCnfnAccumCd As String
        Public Property alcCyrCombAccumCd As String
        Public Property alcCyrNcnfnAccumCd As String
        Public Property alcLftCnfnAccumCd As String
        Public Property alcLftCombAccumCd As String
        Public Property alcLftNcnfnAccumCd As String
        Public Property alcReinstCnfnAmt As String
        Public Property alcReinstCombAmt As String
        Public Property alcReinstCombFreqCd As String
        Public Property alcReinstFreqCnfnCd As String
        Public Property alcReinstFreqNcnfnCd As String
        Public Property alcReinstNcnfnAmt As String
        Public Property covTypCd As String
        Public Property psyCyrCnfnAccumCd As String
        Public Property psyCyrCombAccumCd As String
        Public Property psyCyrNcnfnAccumCd As String
    End Class

    Public Class Mmi3CRows
        Public Property sviCancDt As String
        Public Property sviEffDt As String
        Public Property sviTblNbr As String
        Public Property sviTblOrdrNbr As String
    End Class

#End Region
#Region "MMI 4"

    Public Class Mmi4Return
        <JsonConverter(GetType(Converter(Of Mmi4ARows)))> Public Property mmi4ARows As New List(Of Mmi4ARows)
        <JsonConverter(GetType(Converter(Of Mmi4BRows)))> Public Property mmi4BRows As New List(Of Mmi4BRows)
        <JsonConverter(GetType(Converter(Of Mmi4CRows)))> Public Property mmi4CRows As New List(Of Mmi4CRows)
        <JsonConverter(GetType(Converter(Of Mmi4DRows)))> Public Property mmi4DRows As New List(Of Mmi4DRows)
    End Class

    Public Class Mmi4ARows
        Public Property combPrscDedPriCd As String
        Public Property covTypCd As String
        Public Property famMbrCnt As String
        Public Property famTxtSwapCd As String
        Public Property nonEmbdCoreDedCd As String
        Public Property prortEvntTypCd As String
        Public Property prortIntrvlFreqCd As String
        Public Property tierLblInd As String
    End Class

    Public Class Mmi4BRows
        Public Property combDed1Ind As String
        Public Property combDed2Ind As String
        Public Property combVendCatgyCd As String
        Public Property covTypCd As String
    End Class

    Public Class Mmi4CRows
        Public Property covTypCd As String
        Public Property dedEeChrgAmt As String
        Public Property dedEePls1Amt As String
        Public Property dedEeSpoAmt As String
        Public Property dedFreqPrdCd As String
        Public Property dedMbrCnt As String
        Public Property dedMbrDesc As String
        Public Property dedMultFct As String
        Public Property dedMultSalryPct As String
        Public Property famDedAmt As String
        Public Property famDedCc As String
        Public Property famDedCd As String
        Public Property famDedCo As String
        Public Property oopMultFct As String
        Public Property seqNbr As String
    End Class

    Public Class Mmi4DRows
        Public Property covTypCd As String
        Public Property dedAccumCd As String
        Public Property dedAccumPdAmt As String
        Public Property dedBenPdAmt As String
        Public Property dedCobCd As String
        Public Property dedEndDt As String
        Public Property dedFreqCd As String
        Public Property dedMntAmt As String
        Public Property dedMntCd As String
        Public Property dedMntPdAmt As String
        Public Property dedNtwkTypCd As String
        Public Property dedSemiPvtRtCd As String
        Public Property dedSrvcDesc As String
        Public Property indDedAmt As String
        Public Property indDedCc As String
        Public Property indDedCd As String
        Public Property indDedCo As String
        Public Property seqNbr As String
    End Class

#End Region
#Region "MMI 5"

    Public Class Mmi5Return
        <JsonConverter(GetType(Converter(Of Mmi5ARows)))> Public Property mmi5ARows As New List(Of Mmi5ARows)
        <JsonConverter(GetType(Converter(Of Mmi5BRows)))> Public Property mmi5BRows As New List(Of Mmi5BRows)
        <JsonConverter(GetType(Converter(Of Mmi5CRows)))> Public Property mmi5CRows As New List(Of Mmi5CRows)
        <JsonConverter(GetType(Converter(Of Mmi5DRows)))> Public Property mmi5DRows As New List(Of Mmi5DRows)
    End Class

    Public Class Mmi5ARows
        Public Property combPrscDedQualCd As String
        Public Property covTypCd As String
        Public Property hospDfnCd As String
    End Class

    Public Class Mmi5BRows
        Public Property combDed3Ind As String
        Public Property combDed4Ind As String
        Public Property combVendCatgyCd As String
        Public Property covTypCd As String
    End Class

    Public Class Mmi5CRows
        Public Property covTypCd As String
        Public Property dfdeDedAmt As String
        Public Property dfdeDedCaroCd As String
        Public Property dfdeDedCstCntnCd As String
        Public Property dfdeDedEeChrgAmt As String
        Public Property dfdeDedEePls1Amt As String
        Public Property dfdeDedEeSpoAmt As String
        Public Property dfdeDedFreqPrdCd As String
        Public Property dfdeDedMbrCnt As String
        Public Property dfdeDedMbrDesc As String
        Public Property dfdeDedMultFct As String
        Public Property seqNbr As String
    End Class

    Public Class Mmi5DRows
        Public Property covTypCd As String
        Public Property dideDedAccumCd As String
        Public Property dideDedAccumPdAmt As String
        Public Property dideDedAmt As String
        Public Property dideDedBenPdAmt As String
        Public Property dideDedCaroCd As String
        Public Property dideDedCd As String
        Public Property dideDedCobCd As String
        Public Property dideDedCstCntnCd As String
        Public Property dideDedFreqCd As String
        Public Property dideDedMntAmt As String
        Public Property dideDedMntCd As String
        Public Property dideDedMntPdAmt As String
        Public Property dideDedNtwkTypCd As String
        Public Property dideDedSemiPvtRtCd As String
        Public Property dideDedSrvcDesc As String
        Public Property seqNbr As String
    End Class

#End Region
#Region "MMI 6"

    Public Class Mmi6Return
        <JsonConverter(GetType(Converter(Of Mmi6ARows)))> Public Property mmi6ARows As New List(Of Mmi6ARows)
        <JsonConverter(GetType(Converter(Of Mmi6BRows)))> Public Property mmi6BRows As New List(Of Mmi6BRows)
        <JsonConverter(GetType(Converter(Of Mmi6CRows)))> Public Property mmi6CRows As New List(Of Mmi6CRows)
    End Class

    Public Class Mmi6ARows
        Public Property combPrscDedTirCd As String
        Public Property covTypCd As String
        Public Property dedCred1Amt As String
        Public Property dedCred2Amt As String
        Public Property dedCred3Amt As String
        Public Property dedCred4Amt As String
        Public Property dedCredTyp1Cd As String
        Public Property dedCredTyp2Cd As String
        Public Property dedCredTyp3Cd As String
        Public Property dedCredTyp4Cd As String
        Public Property famHraApLmtAmt As String
        Public Property indvHraApLmtAmt As String
        Public Property plnNbr As String
    End Class

    Public Class Mmi6BRows
        Public Property covTypCd As String
        Public Property dedEeChrgAmt As String
        Public Property dedEePls1Amt As String
        Public Property dedEeSpoAmt As String
        Public Property dedFreqPrdCd As String
        Public Property dedMbrCnt As String
        Public Property dedMbrDesc As String
        Public Property dedMultFct As String
        Public Property famDedAmt As String
        Public Property famDedCc As String
        Public Property famDedCo As String
        Public Property seqNbr As String
    End Class

    Public Class Mmi6CRows
        Public Property covTypCd As String
        Public Property dedAccumCd As String
        Public Property dedAccumPdAmt As String
        Public Property dedBenPdAmt As String
        Public Property dedCd As String
        Public Property dedCobCd As String
        Public Property dedFreqCd As String
        Public Property dedMntAmt As String
        Public Property dedMntCd As String
        Public Property dedMntPdAmt As String
        Public Property dedNtwkTypCd As String
        Public Property dedSemiPvtRtCd As String
        Public Property dedSrvcDesc As String
        Public Property indDedAmt As String
        Public Property indDedCc As String
        Public Property indDedCo As String
        Public Property seqNbr As String
    End Class

#End Region
#Region "MMI 7"

    Public Class Mmi7Return
        <JsonConverter(GetType(Converter(Of Mmi7ARows)))> Public Property mmi7ARows As New List(Of Mmi7ARows)
        <JsonConverter(GetType(Converter(Of Mmi7BRows)))> Public Property mmi7BRows As New List(Of Mmi7BRows)
        <JsonConverter(GetType(Converter(Of Mmi7CRows)))> Public Property mmi7CRows As New List(Of Mmi7CRows)
    End Class

    Public Class Mmi7ARows
        Public Property covTypCd As String
        Public Property dfltChrgPct As String
        Public Property matEligCd As String
        Public Property matTypCd As String
        Public Property mmMultSurgInd As String
        Public Property mtrnCovCd As String
        Public Property mtrnDepCovCd As String
        Public Property pcntRateJqCodesFound As String
        Public Property rsnCustyPrdCd As String
        Public Property sdayMultSurgCovCd As String
        Public Property somePrdCd As String
        Public Property suppExtCd As String
        Public Property suppRncCd As String
    End Class

    Public Class Mmi7BRows
        Public Property covTypCd As String
        Public Property prvsnMaxCapAmt As String
        Public Property psyPaBaseCnCd As String
        Public Property psyPaBasePct As String
        Public Property psyPaBcalcCd As String
        Public Property psyPaDedDescCd As String
        Public Property psyPaMcalcCd As String
        Public Property psyPaMmCnCd As String
        Public Property psyPaMmNcoindCd As String
        Public Property psyPaMmPct As String
        Public Property psyPaPsiInd As String
        Public Property sso2BaseCnCd As String
        Public Property sso2BasePct As String
        Public Property sso2CnfmNcnfmCd As String
        Public Property sso2CnfmNcnfmSchNbr As String
        Public Property sso2MmCnCd As String
        Public Property sso2MmNcoindCd As String
        Public Property sso2MmPct As String
        Public Property sso2NoObtnBcalcCd As String
        Public Property sso2NoObtnMmcalcCd As String
        Public Property sso2ObtnBcalcCd As String
        Public Property sso2ObtnMmcalcCd As String
        Public Property sso3BaseCnCd As String
        Public Property sso3BasePct As String
        Public Property sso3CfmnNcfmnCd As String
        Public Property sso3CfmnNcfmnSchNbr As String
        Public Property sso3DedDescCd As String
        Public Property sso3MmCnCd As String
        Public Property sso3MmPct As String
        Public Property sso3NcoindCd As String
        Public Property sso3NoObtnBcalcCd As String
        Public Property sso3NoObtnMmcalcCd As String
        Public Property sso3ObtnBcalcCd As String
        Public Property sso3ObtnMmcalcCd As String
        Public Property ssoAgrInd As String
        Public Property ssoApplInd As String
        Public Property ssoBaseCnCd As String
        Public Property ssoBasePct As String
        Public Property ssoCnfmNcnfmDedDescCd As String
        Public Property ssoCobXclsCd As String
        Public Property ssoDedDescCd As String
        Public Property ssoMaxCapAmt As String
        Public Property ssoMmCnCd As String
        Public Property ssoMmNcoindCd As String
        Public Property ssoMmPct As String
        Public Property ssoNoObtnBcalcCd As String
        Public Property ssoNoObtnMmcalcCd As String
        Public Property ssoObtnBcalcCd As String
        Public Property ssoObtnMmcalcCd As String
        Public Property ssoPrvsnOrdrCd As String
        Public Property ssoSchNbr As String
    End Class

    Public Class Mmi7CRows
        Public Property covTypCd As String
        Public Property dedAmt As String
        Public Property dedCaroCd As String
        Public Property dedCstCntnCd As String
        Public Property dedFreqPrdCd As String
        Public Property dedMbrCnt As String
        Public Property dedMbrDesc As String
        Public Property dedMultFct As String
    End Class

#End Region
#Region "MMI 8"

    Public Class Mmi8Return
        <JsonConverter(GetType(Converter(Of Mmi8ARows)))> Public Property mmi8ARows As New List(Of Mmi8ARows)
        <JsonConverter(GetType(Converter(Of Mmi8BRows)))> Public Property mmi8BRows As New List(Of Mmi8BRows)
        <JsonConverter(GetType(Converter(Of Mmi8CRows)))> Public Property mmi8CRows As New List(Of Mmi8CRows)
    End Class

    Public Class Mmi8ARows
        Public Property altMktTypCd As String
        Public Property bhCcrCd As String
        Public Property bhMktNbr As String
        Public Property bhParsMktNbr As String
        Public Property bhParsPauthTblNbr As String
        Public Property bhPauthTblNbr As String
        Public Property bhvHlthVendCd As String
        Public Property caiApplInd As String
        Public Property careMgtInd As String
        Public Property chrpNtwkInd As String
        Public Property cntgsMktOvrdInd As String
        Public Property coinsCopayCd As String
        Public Property copayAmt As String
        Public Property copayMaxAnnlAmt As String
        Public Property copayVarId As String
        Public Property copayWaivTblId As String
        Public Property coreMedMktNum As String
        Public Property coreMedPrrAuthCd As String
        Public Property coreMedTblNum As String
        Public Property covTypCd As String
        Public Property cptnXclsInd As String
        Public Property crsTrtDaysCd As String
        Public Property dolTlrInd As String
        Public Property eapVendCd As String
        Public Property eligXrefCd As String
        Public Property emergentWrpInd As String
        Public Property enrpDfltPct As String
        Public Property enrpEmrgFaclInd As String
        Public Property enrpErInd As String
        Public Property enrpNonErInd As String
        Public Property enrpNonErPct As String
        Public Property evdBasDialgInd As String
        Public Property facilityMnnrpPct As String
        Public Property faclShrSvCd As String
        Public Property fertCtrcptvCd As String
        Public Property genTstPolPrtcpCd As String
        Public Property iplanTypCd As String
        Public Property liabOptCd As String
        Public Property mbrNtwkKeyMtchCd As String
        Public Property medcnMgtInd As String
        Public Property mngPsychInd As String
        Public Property mnnrpDmePct As String
        Public Property mnnrpLabPct As String
        Public Property mnrpCd As String
        Public Property mnrpDfltPct As String
        Public Property mnrpPct As String
        Public Property mnrpPtPct As String
        Public Property mntlUbhCd As String
        Public Property nhpNtwkFlexInd As String
        Public Property nonEmrgInd As String
        Public Property nonPpoCd As String
        Public Property ntfyCrdcEpInd As String
        Public Property ntwkPcpCopayAmt As String
        Public Property obgynPcpCopayInd As String
        Public Property ofcVstMaxRmrkCd As String
        Public Property oncPolPrtcpCd As String
        Public Property oopUrgntCareCd As String
        Public Property optoutUbhtierInd As String
        Public Property othMntlNvsXcldCd As String
        Public Property parsMntlNvsXcldCd As String
        Public Property pcpCopayCd As String
        Public Property pcpSpecCoinsInd As String
        Public Property physnShrSvCd As String
        Public Property plnNbr As String
        Public Property ppoInd As String
        Public Property ppoMaxCapAmt As String
        Public Property ppoMinEmrgInd As String
        Public Property ppoMntlNvsXcldCd As String
        Public Property ppoOoaProvCd As String
        Public Property ppoPmntCd As String
        Public Property provCapMinPct As String
        Public Property provOrdrCd As String
        Public Property radcardMktNbr As String
        Public Property radcardTblNbr As String
        Public Property rcprctyTblId As String
        Public Property relSrvcInd As String
        Public Property rhapsodyCopayCd As String
        Public Property rhapsodyCopayDayCnt As String
        Public Property siteCareProcCd As String
        Public Property siteSrvcPrtcpCd As String
        Public Property spclRxInd As String
        Public Property tier1CopayAmt As String
        Public Property tier1UrgntCareAmt As String
        Public Property tier1UrgntCareOopCd As String
        Public Property travBenMktNbr As String
        Public Property travBenTblNbr As String
        Public Property uciRuleCd As String
        Public Property urgntCareAmt As String
        Public Property visnCd As String
        Public Property xcldProvCd As String
    End Class

    Public Class Mmi8BRows
        Public Property basPctOvrlayCd As String
        Public Property benLvlTypCd As String
        Public Property covTypCd As String
        Public Property dedDescCd As String
        Public Property incntPntlyBasCd As String
        Public Property incntPntlyBasPct As String
        Public Property incntPntlyMedCd As String
        Public Property incntPntlyMedPct As String
        Public Property newCoinsCd As String
        Public Property overlayPctCd As String
    End Class

    Public Class Mmi8CRows
        Public Property covTypCd As String
        Public Property prvsnMaxCapAmt As String
    End Class

#End Region
#Region "MMI 9"

    Public Class Mmi9Return
        <JsonConverter(GetType(Converter(Of Mmi9ARows)))> Public Property mmi9ARows As New List(Of Mmi9ARows)
        <JsonConverter(GetType(Converter(Of Mmi9BRows)))> Public Property mmi9BRows As New List(Of Mmi9BRows)
    End Class

    Public Class Mmi9ARows
        Public Property authPrdNbr As String
        Public Property benMaxAuralCiPriCd As String
        Public Property benMaxAuralCiPriCnt As String
        Public Property benMaxAuralCiSecCd As String
        Public Property benMaxAuralCiSecCnt As String
        Public Property benMaxCogTrpyExclCd As String
        Public Property benMaxCogTrpyPriCd As String
        Public Property benMaxCogTrpyPriCnt As String
        Public Property benMaxCogTrpySecCd As String
        Public Property benMaxCogTrpySecCnt As String
        Public Property benMaxCombPriCd As String
        Public Property benMaxCombPriCnt As String
        Public Property benMaxCombSecCd As String
        Public Property benMaxCombSecCnt As String
        Public Property benMaxCrdcRehabPriCd As String
        Public Property benMaxCrdcRehabPriCnt As String
        Public Property benMaxCrdcRehabSecCd As String
        Public Property benMaxCrdcRehabSecCnt As String
        Public Property benMaxOcpTrpyPriCd As String
        Public Property benMaxOcpTrpyPriCnt As String
        Public Property benMaxOcpTrpySecCd As String
        Public Property benMaxOcpTrpySecCnt As String
        Public Property benMaxPhOcSpchPriCd As String
        Public Property benMaxPhOcSpchPriCnt As String
        Public Property benMaxPhOcSpchSecCd As String
        Public Property benMaxPhOcSpchSecCnt As String
        Public Property benMaxPhOcTrpyPriCd As String
        Public Property benMaxPhOcTrpyPriCnt As String
        Public Property benMaxPhOcTrpySecCd As String
        Public Property benMaxPhOcTrpySecCnt As String
        Public Property benMaxPhysTrpyPriCd As String
        Public Property benMaxPhysTrpyPriCnt As String
        Public Property benMaxPhysTrpySecCd As String
        Public Property benMaxPhysTrpySecCnt As String
        Public Property benMaxPlmryRehbPriCd As String
        Public Property benMaxPlmryRehbPriCnt As String
        Public Property benMaxPlmryRehbSecCd As String
        Public Property benMaxPlmryRehbSecCnt As String
        Public Property benMaxSpchTrpyPriCd As String
        Public Property benMaxSpchTrpyPriCnt As String
        Public Property benMaxSpchTrpySecCd As String
        Public Property benMaxSpchTrpySecCnt As String
        Public Property benMaxSpneMnipPriCd As String
        Public Property benMaxSpneMnipPriCnt As String
        Public Property benMaxSpneMnipSecCd As String
        Public Property benMaxSpneMnipSecCnt As String
        Public Property clmAutoDenyInd As String
        Public Property coreMedPrrAuthCd As String
        Public Property covTypCd As String
        Public Property diagVsSurgInd As String
        Public Property eciTblId As String
        Public Property emrgParsNtfyInd As String
        Public Property emrgParsNtfyPrdCd As String
        Public Property eocInd As String
        Public Property evdBasDialgInd As String
        Public Property eviTblId As String
        Public Property faclClmEdtInd As String
        Public Property ncnfnCd As String
        Public Property noblxLabInd As String
        Public Property ntfyCrdcEpInd As String
        Public Property othrRmrkCd As String
        Public Property padmisPrdNbr As String
        Public Property pars3NobtnBcalcCd As String
        Public Property pars3NobtnMmcalcCd As String
        Public Property parsAgrCd As String
        Public Property parsList As String
        Public Property parsMnlProcInd As String
        Public Property parsOopLmtAmt As String
        Public Property parsSurgSchedNbr As String
        Public Property pauthRvwPrdCd As String
        Public Property physMedcnPrdCd As String
        Public Property rapl3TierInd As String
        Public Property relSrvcInd As String
        Public Property rmrkCd As String
        Public Property rmrkSpineManipCd As String
        Public Property tinMtchCd As String
        Public Property uciRuleCd As String
        Public Property varWndwPrdNbr As String
    End Class

    Public Class Mmi9BRows
        Public Property basPctOvrlayCd As String
        Public Property benLvlTypCd As String
        Public Property covTypCd As String
        Public Property dedDescCd As String
        Public Property incntPntlyBasCd As String
        Public Property incntPntlyBasPct As String
        Public Property incntPntlyMedCd As String
        Public Property incntPntlyMedPct As String
        Public Property newCoinsCd As String
        Public Property overlayPctCd As String
    End Class

#End Region
#Region "MMI 10"

    Public Class Mmi10Return
        <JsonConverter(GetType(Converter(Of Mmi10ARows)))> Public Property mmi10ARows As New List(Of Mmi10ARows)
        <JsonConverter(GetType(Converter(Of Mmi10BRows)))> Public Property mmi10BRows As New List(Of Mmi10BRows)
        <JsonConverter(GetType(Converter(Of Mmi10CRows)))> Public Property mmi10CRows As New List(Of Mmi10CRows)
        <JsonConverter(GetType(Converter(Of Mmi10DRows)))> Public Property mmi10DRows As New List(Of Mmi10DRows)
    End Class

    Public Class Mmi10ARows
        Public Property covTypCd As String
        Public Property dualOopIndvPct As String
        Public Property famNewCoinsAmt As String
        Public Property nbSprsInd As String
        Public Property newCoinsAccumCd As String
        Public Property newCoinsAmt As String
        Public Property newCoinsChgAmt As String
        Public Property newCoinsCobXcldInd As String
        Public Property newCoinsCombPrscCd As String
        Public Property newCoinsDedTypCd As String
        Public Property newCoinsEndDt As String
        Public Property newCoinsIndvMaxPct As String
        Public Property newCoinsIndvMinPct As String
        Public Property newCoinsIndvNbr As String
        Public Property newCoinsPrdCd As String
        Public Property newCoinsSalFamMultFct As String
        Public Property newCoinsSalFamTypCd As String
        Public Property newCoinsSalIndvTypCd As String
        Public Property nonEmbdNewCoinsCd As String
        Public Property oopCombEeChrgAmt As String
        Public Property oopCombEePls1Amt As String
        Public Property oopCombEeSpoAmt As String
        Public Property oopCombFamAmt As String
        Public Property oopCombIndvAmt As String
        Public Property oopCombNbrCd As String
        Public Property oopInNtwkEeChrgAmt As String
        Public Property oopInNtwkEePls1Amt As String
        Public Property oopInNtwkEeSpoAmt As String
        Public Property plnNbr As String
        Public Property tier1FamNewCoinsAmt As String
        Public Property tier1NewCoinsAmt As String
    End Class

    Public Class Mmi10BRows
        Public Property benLvlTypCd As String
        Public Property covTypCd As String
        Public Property newCoinsCd As String
    End Class

    Public Class Mmi10CRows
        Public Property combDualOopInd As String
        Public Property combNewCoinsInd As String
        Public Property combTier1Ind As String
        Public Property combVendCatgyCd As String
        Public Property covTypCd As String
    End Class

    Public Class Mmi10DRows
        Public Property covTypCd As String
        Public Property poolAllAmt As String
        Public Property poolAllCd As String
        Public Property poolAllClssNbr As String
        Public Property poolAllDivNbr As String
        Public Property poolAllDt As String
        Public Property poolAllPlcyNbr As String
        Public Property poolMmAmt As String
        Public Property poolMmClssNbr As String
        Public Property poolMmDivNbr As String
        Public Property poolMmPlcyNbr As String
    End Class

#End Region
#Region "MMI 11"

    Public Class Mmi11Return
        <JsonConverter(GetType(Converter(Of Mmi11ARows)))> Public Property mmi11ARows As New List(Of Mmi11ARows)
        <JsonConverter(GetType(Converter(Of Mmi11BRows)))> Public Property mmi11BRows As New List(Of Mmi11BRows)
        <JsonConverter(GetType(Converter(Of Mmi11CRows)))> Public Property mmi11CRows As New List(Of Mmi11CRows)
    End Class

    Public Class Mmi11ARows
        Public Property coreOopNcapInd As String
        Public Property covTypCd As String
        Public Property famMultPct As String
        Public Property famSalryTypCd As String
        Public Property famVal As String
        Public Property indvCopayCaroCd As String
        Public Property indvOopCd As String
        Public Property indvPrdCd As String
        Public Property indvSalryTypCd As String
        Public Property nonEmbdCopayCd As String
        Public Property nonEmbdCoreOopCd As String
        Public Property xapplyCopayCd As String
        Public Property xapplyOopCd As String
    End Class

    Public Class Mmi11BRows
        Public Property accumBenTypCd As String
        Public Property accumRuleTypCd As String
        Public Property covTypCd As String
        Public Property eePls1Amt As String
        Public Property eePlsChAmt As String
        Public Property eePlsSpAmt As String
        Public Property famMaxAmt As String
        Public Property famOopAmt As String
        Public Property indvMaxAmt As String
        Public Property indvOopAmt As String
    End Class

    Public Class Mmi11CRows
        Public Property combCoreInnCpyInd As String
        Public Property combCoreInnOopInd As String
        Public Property combCoreOonCpyInd As String
        Public Property combCoreOonOopInd As String
        Public Property combCoreTier1CpyInd As String
        Public Property combCoreTier1OopInd As String
        Public Property combVendCatgyCd As String
        Public Property covTypCd As String
    End Class

#End Region
#Region "MMI 12"

    Public Class Mmi12Return
        <JsonConverter(GetType(Converter(Of Mmi12ARows)))> Public Property mmi12ARows As New List(Of Mmi12ARows)
    End Class

    Public Class Mmi12ARows
        Public Property covTypCd As String
        Public Property custVarTablId As String
        Public Property cxiSeqNbr As String
        Public Property plnNbr As String
    End Class

#End Region
#Region "MMI 13"

    Public Class Mmi13Return
        <JsonConverter(GetType(Converter(Of Mmi13ARows)))> Public Property mmi13ARows As New List(Of Mmi13ARows)
        <JsonConverter(GetType(Converter(Of Mmi13BRows)))> Public Property mmi13BRows As New List(Of Mmi13BRows)
        <JsonConverter(GetType(Converter(Of Mmi13CRows)))> Public Property mmi13CRows As New List(Of Mmi13CRows)
    End Class

    Public Class Mmi13ARows
        Public Property benAutsmAgeYrFrVal As String
        Public Property benAutsmAgeYrToVal As String
        Public Property benAutsmAllCd As String
        Public Property benAutsmPostCd As String
        Public Property benMaxCogTrpyExclCd As String
        Public Property benMaxEhbHabRmrkCd As String
        Public Property benMaxEhbRhabRmrkCd As String
        Public Property covTypCd As String
        Public Property othrRmrkCd As String
        Public Property physMedcnPrdCd As String
        Public Property rmrkSpineManipCd As String
    End Class

    Public Class Mmi13BRows
        Public Property covTypCd As String
        Public Property ntwkStsNparCd As String
        Public Property ntwkStsParCd As String
        Public Property trpyBenLmtNparCnt As String
        Public Property trpyBenLmtParCnt As String
        Public Property trpyCtgyCd As String
        Public Property trpyTypCd As String
    End Class

    Public Class Mmi13CRows
        Public Property covTypCd As String
        Public Property habCombCiLmtCd As String
        Public Property habCombCtLmtCd As String
        Public Property habCombMtLmtCd As String
        Public Property habCombOtLmtCd As String
        Public Property habCombPtLmtCd As String
        Public Property habCombSmLmtCd As String
        Public Property habCombStLmtCd As String
        Public Property rhabCombCiLmtCd As String
        Public Property rhabCombCrLmtCd As String
        Public Property rhabCombCtLmtCd As String
        Public Property rhabCombMtLmtCd As String
        Public Property rhabCombOtLmtCd As String
        Public Property rhabCombPrLmtCd As String
        Public Property rhabCombPtLmtCd As String
        Public Property rhabCombSmLmtCd As String
        Public Property rhabCombStLmtCd As String
    End Class

#End Region

#End Region

End Class
