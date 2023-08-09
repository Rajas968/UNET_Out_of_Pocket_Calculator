Imports System.ComponentModel
Imports AutConnMgrTypeLibrary
Imports AutSessTypeLibrary
Imports Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Net
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop.Word
Imports System.Web
Imports System.Diagnostics.Eventing.Reader
Imports System.Data.SqlClient
Imports DocumentFormat.OpenXml.Presentation

'Imports DocumentFormat.OpenXml.Spreadsheet
Public Class Form1
    Public strUnetID
    Public strUnetPass
    Public strICN
    Public strENG
    Public strMemID
    Public strPTLNM
    Public fromDT As Date
    Public I As Integer
    Public mydt As Date
    Public rewYear As String
    Public frmDT As Date
    Public toDT As Date
    Public ChoiceType As String
    Public myYear As String
    Public CurrentYear As String
    Public myYear2 As String, Member, strPrevChoice, CarryOvr
    Public cryptOutput, sysValue, sysVariable, Enc_COM_Exists
    Public strUserId, strPsw, returnValue
    Public strDoc360Data As String, strMXIdata As String

    Private api_MRI As New apiMRI
    Private api_MXI As New apiMXI
    Private api_MMI As New apiMMI
    Private api_MHI As New apiMHI
    Private api_AHI As New apiAHI
    Private api_CEI As New apiCEI
    Private api_PMI As New apiPMI
    Private api_MSI As New apiMSI
    Private api_MCI As New apiMCI
    'Dim apiMHIobj As apiMHI = New apiMHI
    Dim apiMHIobj As apiMHI = New apiMHI
    Private mmiDetails As New List(Of apiMMI.MmiReturn)
    Private claimList As New List(Of apiMHI.ResponseData)
    Private ceiList As New List(Of apiCEI.ResponseData)
    Public rw As Integer
    Public ptFName As String
    Public ptRelation As String
    Public MHIPTName() As String
    Public MMI1 As Worksheet
    Public MMI4 As Worksheet
    Public MMI5 As Worksheet
    Public MMI6 As Worksheet
    Public MMI8 As Worksheet
    Public MMI10 As Worksheet
    Public MMI11 As Worksheet
    Public Copays As Worksheet
    Public MHISheet As Worksheet
    Public OONPercent()
    Public InnIndDed() As Double
    Public InnDedCross() As Boolean
    Public OONIndDed() As Double
    Public OONDedCross() As Boolean
    Public InnIndOOP() As Double
    Public DedToOOP() As Boolean
    Public OONIndOOP() As Double
    Public OOPCross() As String
    Public InnFamDed() As Double
    Public InnFamDedCross() As Boolean
    Public InnFamOOP() As Double
    Public OONFamDed() As Double
    Public OONFamDedCross() As Boolean
    Public DxCauseCode() As String  'Captures cause codes from Copays tab
    Public CopayPOS() As String     'Captures place of service codes from Copays tab
    Public OONFamOOP() As Double
    Public INNStatus() As Boolean  'Used when IndDed is listed as a 'U'
    Public TieredCross() As Boolean
    Public EffDte() As Date
    Public CancelDate() As Date
    Public DateOfSvc As Date
    Public MMICount As Integer
    Public NonEmbed() As String
    Public TotalRows As Integer
    Public blnBeforeEff() As Boolean
    ' new declaration 
    Dim colCntPG1 As Int32
    Dim rowCnt As Int32
    Public strfullTin As String, strProv As String, proType, strPlnnbr, msiResp, strcause, mciResp
    Public PlaceSvc As String, MCalcNum, spccode As String, spccode1 As String, strCopaySet As String, dtSt_ToDate, strPTName As String, nrow As Integer
    Public OOPSheet As Worksheet, blnTiered As Boolean, intChoice As Integer, q As Integer, Cnt As Integer
    Public Const conOvCode As String = "H,D,T"
    Public strCOVPlan As String
    Public mmiFlag As Boolean = False
    Public pName As String
    Public pRel As String, strDedCode
    Dim strInputDeductible As String
    Dim Rowchk As Integer
    Public strProvName As String
    Public blnFacility As Boolean, OIOIMClaim As Boolean, OtherIns As Boolean = 3000
    Public starttime, FComments, MHIstarttime, MHIendtime, MHIHistoryCnt, MMIstarttime, MMIendtime, strCopay, strDed, strDedToOOP, blnCopays
    Public respLogin, loginStatus
    Public index As Integer
    Dim mycolor As String
    Public sArray As String()
    Public BrLDate As String
    Public intRow As Integer
    Public oldPName()
    Dim ptN()
    Dim xt As Integer
    Dim row_yellow As Integer
    Dim strSSN As String
    Private rowColors As New Dictionary(Of Integer, Color)


    ''Added - 05/08/2023 Sanjeet


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        'Fills the year selection with the last 5 years
        '(assumes the user will be gathering for the current year)
        For yearNumber = CInt(Format(CDate(Now), "yyyy")) To CInt(CInt(Format(CDate(Now), "yyyy")) - 5) Step -1
            yearList.Items.Add(yearNumber.ToString)
        Next yearNumber
        yearList.SelectedIndex = 0

        loginStatus = False
        ''calling to API Exe File to launch api .
        Call Get_UserID()
        Call Get_UserPass()

        If loginStatus = True Then
            Call Get_UserID()
            Call Get_UserPass()
        End If


        Call GetTable()

        TabControl1.SelectedIndex = 0

        Dim dtMonth As Integer = Format(Now, "MM")
        Dim dtDay As Integer = Format(Now, "dd")

        Dim yearStart As String = "01/01/" & yearList.Text
        Dim yearEnd As String = "12/31/" & yearList.Text
        'Dim yearEnd As String = dtMonth & "/" & dtDay & "/" & yearList.Text
        startSelect.Text = yearStart
        endSelect.Text = yearEnd


        mmiPag1RowsAdd()
        mmiPag4RowsAdd()
        mmiPag5RowsAdd()

        MMIPage10()

        MMIOverviw()


    End Sub


    Sub mmiPag1RowsAdd()
        DGrid_PG1.Rows.Add(19)
    End Sub

    Sub mmiPag4RowsAdd()
        DGrid_PG4.Rows.Add(64)
    End Sub
    Sub mmiPag5RowsAdd()
        DGrid_PG5.Rows.Add(34)
    End Sub

    Sub MMIPage10()
        DGrid_PG10.Rows.Add(46)
    End Sub

    Sub MMIOverviw()
        DGridOverview.Rows.Add(32)
    End Sub


    Public Sub Get_UserID()
        I = 10
        'Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Volatile Environment", "USERNAME", Nothing)
        Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Volatile Environment", "Eso_i", Nothing)
        'Decrypt password, if needed
        returnValue = readValue
        If returnValue = Nothing Then
            ''Launching Unet E_signon Macro
            MsgBox("You are not logged in Unet application. Please click OK to proceed to Login through E_Signon macro")
            Call Launch_vbs()

            respLogin = MsgBox("Please click on Yes when E_signon Macro completed? ", vbYesNo, "Out of Pocket Calculator")

            If respLogin = vbYes Then
                loginStatus = True
            End If
            'Application.Restart()

            '' need to check if user click on login on Unet Sanjeet           

        ElseIf returnValue.Length = 64 Then
            Try
                Dim encryptionMethod As SignonSecurity.HashingCredential = New SignonSecurity.HashingCredential
                returnValue = encryptionMethod.Decrypt(returnValue)
                Me.txtUserID.Text = returnValue.ToString()
            Catch ex As Exception
                returnValue = ""
            End Try
        End If
    End Sub
    Public Sub Get_UserPass()


        Dim readValue = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Volatile Environment", "Eso_p", Nothing)
        'Decrypt password, if needed
        returnValue = readValue

        If returnValue = Nothing Then


        ElseIf returnValue.Length = 64 Then
            Try
                Dim encryptionMethod As SignonSecurity.HashingCredential = New SignonSecurity.HashingCredential
                returnValue = encryptionMethod.Decrypt(returnValue)
                Me.txtUserPass.Text = returnValue.ToString()
            Catch ex As Exception
                returnValue = ""
            End Try
        End If

    End Sub

    Private Function PrepareTokenRequest(ByVal clientId As String, ByVal clientSecret As String) As String
        Dim tokenRequest As String = String.Empty
        tokenRequest = $"grant_type=client_credentials&client_id={clientId}&client_secret={clientSecret}"
        Return tokenRequest
    End Function
    Sub History_detail()

        strMemID = txt_Policy.Text
        strPTLNM = UCase(txt_SSN.Text)

        MsgBox("Calling to gather data from history")

        TabControl1.SelectedIndex = 0

        Dim i As Integer
        Dim n As Long
        n = 1
        i = 0
        '''adding to to MHI Grid
        For i = 0 To 10
            tblMHI.Rows.Add(n)
            n = n + 1
        Next

        '''Updating data from MHI API

        For q = 0 To 10
            tblMHI.Rows(q).Cells(0S).Value = "11/01/2022"
        Next

        ' TryCast(tblMHI.DataSource, DataTable).DefaultView.RowFilter = "From >= '" + (D1.Value, "MM/dd/yyyy") + "'"

        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        RichTextBox1.SelectionBullet = True
        RichTextBox1.AppendText("Completed to get the data from MHI History" & vbCrLf)

        ' Application.Exit()

    End Sub

    Private Sub btnCEIExport_Click(sender As Object, e As EventArgs) Handles btnCEIExport.Click
        Call Initialize()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        RichTextBox1.SelectionBullet = True
        RichTextBox1.AppendText("Exporting data into Excel........." & vbCrLf)
        Call Export_MHI()
        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        RichTextBox1.SelectionBullet = True
        RichTextBox1.AppendText("Exported History data in Excel" & vbCrLf)
    End Sub

    Private Sub btnClaimInfo_Click(sender As Object, e As EventArgs) Handles btnClaimInfo.Click
        '''Fetching hisotry data for all Family member
        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        RichTextBox1.SelectionBullet = True
        RichTextBox1.AppendText("Gathering data from MHI screen ...." & vbCrLf)
        ''------------------------------------------------------------------------------
        MHIstarttime = Now()
        Dim AHI_Response

        strSSN = "S" & UCase(Trim(txt_SSN.Text))
        If InStr(strSSN, "SS") > 0 Then
            strSSN = Replace(strSSN, "SS", "S")
        End If
        frmDT = startSelect.Text

        Dim dcnList As New List(Of String)
        Dim intLenght As Integer
        ' tblMHI.Rows.Clear()
        'Loops through checked members
        For memberNumber = 0 To memberList.CheckedItems.Count - 1
            'Splits the member name and rel so it's usable in the AHI Query
            Dim splitMbr = Split(memberList.CheckedItems(memberNumber), "/")
            Dim icnList As New List(Of String)

            'AHI API Query using the name and rel from the checked list along with
            'the year range gathered from the previous modules
            If InStr(splitMbr(0), " ") > 0 Then
                intLenght = Trim(InStr(splitMbr(0), " "))
                splitMbr(0) = Mid(splitMbr(0), 1, intLenght)
            End If

            Try
                AHI_Response = api_AHI.QueryAllResults(
                    Trim(txt_Policy.Text), strSSN,
                    Trim(splitMbr(0)), Trim(splitMbr(1)), ,
                    Format(CDate(startSelect.Text), "MM/dd/yyyy"),
                    Format(CDate(endSelect.Text), "MM/dd/yyyy")).Results.Response(0)
            Catch ex As Exception

                MsgBox("AHI Detail is not available to seleted Year, Please select correct Year and re-start")
                Exit For
            End Try
            'AHI_Response = api_AHI.QueryAllResults(
            '    Trim(txt_Policy.Text), Trim(txt_SSN.Text),
            '    Trim(splitMbr(0)), Trim(splitMbr(1)), ,
            '    Format(CDate(startSelect.Text), "MM/dd/yyyy"),
            '    Format(CDate(endSelect.Text), "MM/dd/yyyy")).Results.Response(0)

            If (AHI_Response Is Nothing) OrElse (AHI_Response.rspLineData.Count = 0) Then Continue For

            'Loops through all the AHI claims and checks if the temporary ICN List contains 
            'the ICN from the AHI Query, If Not it adds it to my ICN List
            For Each ahiClaim In AHI_Response.rspLineData
                If Not icnList.Contains(ahiClaim.rspIcn) Then icnList.Add(ahiClaim.rspIcn)
            Next ahiClaim

            If icnList.Count = 0 Then Continue For
            'Loops through all the tempoarary ICNs gathered previously and gathers the MHI data for those ICNs
            For Each icnNo In icnList
                'MHI API query that uses the previous gathered ICN
                '''Sanjeet updating in new way 

                Dim mhiResp As apiMHI.ResponseData = api_MHI.PerformQuery(Trim(txt_Policy.Text), strSSN, Trim(splitMbr(0)), Trim(splitMbr(1)), Trim(icnNo)).Results.Response(0)
                '  Dim mhiResp As apiMHI.ResponseData = api_MHI.PerformQuery("", "", "", "", Trim("DH19338319")).Results.Response(0)

                If (mhiResp Is Nothing) Then Continue For

                'Loops through all the MHI claims and checks if the DCN List contains 
                'the DCN from the MHI Query, If Not it adds it to my DCN List

                'For Each mhiClaim In mhiResp.rspClaimHeader
                'For Each totalLine In mhiClaim.rspTotalLine1
                'Dim strdraftnn As String = totalLine.rspDraftNbr

                '        If Not dcnList.Contains(totalLine.rspDraftNbr) Or totalLine.rspDraftNbr = "0000000000" Then  ''added by Sanjeet 
                'dcnList.Add(totalLine.rspDraftNbr)  'Adds to a list of all gathered DCNs so we can skip duplicates
                claimList.Add(mhiResp)              'Adds the entire claim to an outside list so it can be used later
                '        End If
                'Next totalLine
                'Next mhiClaim
            Next icnNo

            ''adding paitent name
            'DGrid_PG1.Rows(3).Cells(memberNumber + 1).Value = Trim(splitMbr(0)) & "/" & Trim(splitMbr(1))
            'DGrid_PG4.Rows(3).Cells(memberNumber + 1).Value = Trim(splitMbr(0)) & "/" & Trim(splitMbr(1))
            'DGrid_PG5.Rows(3).Cells(memberNumber + 1).Value = Trim(splitMbr(0)) & "/" & Trim(splitMbr(1))
            'DGrid_PG10.Rows(3).Cells(memberNumber + 1).Value = Trim(splitMbr(0)) & "/" & Trim(splitMbr(1))
            'DGridOverview.Rows(3).Cells(memberNumber + 1).Value = Trim(splitMbr(0)) & "/" & Trim(splitMbr(1))

        Next memberNumber

        'Loops through all claims in the claim list
        'Loops through all claims in the claim list

        rw = 0
        Dim cntNum As Integer = 0
        Dim myicnNo As String

        If claimList.Count < 0 Then
            MsgBox("There is no History data found selected year")
            Exit Sub
        End If

        For Each claim In claimList

            ptFName = claim.rspEmpFirstName
            ptRelation = claim.rspRelCD
            Dim sl = 0
            cntNum = 0
            'Loops through all DCNs in each claim
            For Each mhiHeader In claim.rspClaimHeader
                'Loops through all claim lines within each DCN
                Dim clmNumber As String = mhiHeader.rspClaimNbr
                Dim causeCode As String = mhiHeader.rspCauseCode
                cntNum = cntNum + 1
                For Each claimLine In mhiHeader.rspLineData

                    DGridMHI.Rows.Add(claimLine.rspFirstDate,
                        claimLine.rspLastDate,
                        claimLine.rspServiceCd,
                        claimLine.rspPlaceOfServ,
                        claimLine.rspNumber_string,
                        claimLine.rspOverrideCd,
                        claimLine.rspPayeeCd,
                        claimLine.rspProvPosNbr,
                        claimLine.rspRemarkCd,
                        claimLine.rspCharge,
                        claimLine.rspNotCovered,
                        claimLine.rspMmDedDesc,
                        claimLine.rspMmCoveredAmt,
                        claimLine.rspMmDedAmt,
                        claimLine.rspMmDedDesc,
                        claimLine.rspMmPct,
                        claimLine.rspMmAmt,
                        claimLine.rspCrAmt,
                        claimLine.rspCrAmt_string,
                        claimLine.rspSanctionInd)

                    For Each footer In mhiHeader.rspTotalLine1
                        DGridMHI.Rows(rw).Cells(20).Value = Mid(causeCode, 1, 1)
                        DGridMHI.Rows(rw).Cells(21).Value = Mid((footer.rspProviderNbr), 1, 1)
                        DGridMHI.Rows(rw).Cells(22).Value = Mid((footer.rspProviderNbr), 2, 9)
                        DGridMHI.Rows(rw).Cells(23).Value = Mid((footer.rspProviderNbr), 11, 5)
                        DGridMHI.Rows(rw).Cells(24).Value = clmNumber
                        DGridMHI.Rows(rw).Cells(25).Value = (footer.rspDraftNbr)
                        DGridMHI.Rows(rw).Cells(26).Value = (footer.rspDateProc)
                        DGridMHI.Rows(rw).Cells(28).Value = (footer.rspTotCharge)
                        DGridMHI.Rows(rw).Cells(29).Value = (footer.rspTotPaid)
                        If DGridMHI.Rows(rw).Cells(2).Value = "RX" Then
                            Threading.Thread.Sleep(100)
                            DGridMHI.Rows(rw).Cells(22).Value = "Pharmacy"
                            DGridMHI.Rows(rw).Cells(51).Value = "Pharmacy"
                            DGridMHI.Rows(rw).Cells(52).Value = "RX"
                        End If
                    Next footer
                    'The same footer loop can exist here as well as in the claimLine because
                    'it is nested under the mhiHeader class
                    For Each footer In mhiHeader.rspTotLine2
                        DGridMHI.Rows(rw).Cells(30).Value = footer.rspIcn
                        DGridMHI.Rows(rw).Cells(31).Value = footer.rspIcnSuffix
                        DGridMHI.Rows(rw).Cells(32).Value = footer.rspFlnOff
                        DGridMHI.Rows(rw).Cells(33).Value = footer.rspPrsInd
                        DGridMHI.Rows(rw).Cells(34).Value = footer.rspSIInd
                        DGridMHI.Rows(rw).Cells(35).Value = ptFName
                        DGridMHI.Rows(rw).Cells(36).Value = ""
                        DGridMHI.Rows(rw).Cells(37).Value = ptRelation
                        DGridMHI.Rows(rw).Cells(38).Value = ptFName & "/" & ptRelation
                        DGridMHI.Rows(rw).Cells(49).Value = footer.rspIcn & " " & footer.rspIcnSuffix
                        myicnNo = footer.rspIcn
                    Next footer
                    rw = rw + 1
                    sl = sl + 1
                Next claimLine
                'Exit For                
            Next mhiHeader

        Next claim

        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        RichTextBox1.SelectionBullet = True
        RichTextBox1.AppendText("Data fetched from MHI screen" & vbCrLf)
        'Adds place of service code from previous line to Copay line.
        'Used during calculations for out of pocket
        Dim PSCode As String
        Dim icnt As Integer

        For icnt = 0 To DGridMHI.Rows.Count - 1
            If InStr(DGridMHI.Rows(icnt).Cells(2).Value, "NCOPAY") > 0 Or InStr(DGridMHI.Rows(icnt).Cells(2).Value, "COPAY") > 0 Then
                PSCode = DGridMHI.Rows(icnt - 1).Cells(3).Value
                DGridMHI.Rows(icnt).Cells(3).Value = PSCode
            End If
        Next
        '**********************inserting blank row            

        For rnxt = 0 To DGridMHI.Rows.Count - 2

            If DGridMHI.Rows(rnxt).Cells(38).Value <> DGridMHI.Rows(rnxt + 1).Cells(38).Value And Not IsDBNull(DGridMHI.Rows(rnxt + 1).Cells(38).Value) And DGridMHI.Rows(rnxt + 1).Cells(38).Value <> Nothing Then
                DGridMHI.Rows.Insert(rnxt + 1)
                rnxt = rnxt + 1
            End If
        Next


        Dim dt_n As New System.Data.DataTable

        dt_n = dgridmhi_()


        'For rnxt = 0 To DGridMHI.Rows.Count - 2

        '    If Format(CDate(DGridMHI.Rows(rnxt).Cells(0).Value), "yyyy") <> Format(CDate(DGridMHI.Rows(rnxt + 1).Cells(0).Value), "yyyy") And Not IsDBNull(DGridMHI.Rows(rnxt + 1).Cells(38).Value) And DGridMHI.Rows(rnxt + 1).Cells(38).Value <> Nothing Then
        '        DGridMHI.Rows.Insert(rnxt + 1)
        '        rnxt = rnxt + 1
        '    End If
        'Next

        ''Calling to create column and insert data for histor
        'Dim resp = MsgBox("Do you need to access claims from Purged History? ", vbYesNo, "****Purged History option not available at this time****")

        Dim resp = MsgBox("Do you need to access claims from Purged History? ", vbYesNo, "Out of Pocket Calculator")
        Rowchk = 1
        If resp = vbYes Then
            Purgemain.Show()
            'MsgBox("Purged History option not available at this time, Data will be pulled from MHI History only")
            'Call getMHI(dt_n)
        ElseIf resp = vbNo Then
            Call getMHI(dt_n)
        End If

        ''remove previous year data 
        Dim dtStartYear As Date = Format(CDate(startSelect.Text), "MM/dd/yyyy")
        Dim dtCompareYear As Integer
        dtCompareYear = dtStartYear.Year


        Dim intLastrow As Integer = tblMHI.Rows.Count - 1
        For I = 0 To intLastrow
            If I > intLastrow Then Exit For
            Try
                Dim rvalue = Format(CDate(tblMHI.Rows(I).Cells(0).Value), "yyyy")

                'If Int(rvalue) < Int(yearList.Text) And rvalue <> "0001" Then

                If Int(rvalue) < dtCompareYear And rvalue <> "0001" Then
                    tblMHI.Refresh()
                    tblMHI.Rows.RemoveAt(I)
                    tblMHI.Refresh()
                    intLastrow = intLastrow - 1
                    I = I - 1
                End If
            Catch ex As Exception

            End Try
        Next I
        ''-----------------------------------------------Added code 04/27/2023
        Dim trow As Integer
        Try

            For trow = 0 To tblMHI.Rows.Count - 1

                If IsDBNull(tblMHI.Rows(trow).Cells(0).Value) And IsDBNull(tblMHI.Rows(trow + 1).Cells(0).Value) Then
                    tblMHI.Rows.RemoveAt(trow + 1)

                    If IsDBNull(tblMHI.Rows(trow + 2).Cells(0).Value) Then
                        tblMHI.Rows.RemoveAt(trow + 2)
                    End If
                End If

            Next

        Catch ex As Exception

        End Try

        Try

            For trow = 0 To tblMHI.Rows.Count - 1

                If IsDBNull(tblMHI.Rows(trow).Cells(0).Value) And IsDBNull(tblMHI.Rows(trow + 1).Cells(0).Value) Then
                    tblMHI.Rows.RemoveAt(trow + 1)

                    If IsDBNull(tblMHI.Rows(trow + 2).Cells(0).Value) Then
                        tblMHI.Rows.RemoveAt(trow + 2)
                    End If
                End If

            Next

        Catch ex As Exception

        End Try

        ''-----------------------------------------------Added code 04/27/2023

        Dim strMbr As String                                                                '''NEED TO CHECK 

        For crow = 0 To tblCopay.Rows.Count - 2
            For memberNumber = 0 To memberList.CheckedItems.Count - 1
                'Splits the member name and rel so it's usable in the AHI Query
                strMbr = memberList.CheckedItems(memberNumber)
                tblCopay.Rows(crow).Cells(3).Value = strMbr
            Next
        Next

        Try
            If IsDBNull(tblMHI.Rows(0).Cells(0).Value) Then
                tblMHI.Rows.RemoveAt(0)
            End If
        Catch ex As Exception

        End Try

        MHIendtime = Now()
        MHIHistoryCnt = tblMHI.Rows.Count - 1



        ''''UPDATING OIM                DONE - Sanjeet 05/16/2023

        For trow = 0 To tblMHI.Rows.Count - 1

            If Not IsDBNull(tblMHI.Rows(trow).Cells(2).Value) Then
                Dim strOIM As String = Mid(tblMHI.Rows(trow).Cells(2).Value, 1, 2)
                If strOIM = "OI" Then
                    Dim strICN As String = Trim(tblMHI.Rows(trow).Cells(30).Value)

                    For crow = 0 To tblMHI.Rows.Count - 1

                        If Trim(isNullOrEmpty(tblMHI.Rows(crow).Cells(30).Value)) = strICN Then

                            tblMHI.Rows(crow).Cells(47).Value = True
                        End If
                    Next
                End If
            End If
        Next


        For crow = 0 To tblMHI.Rows.Count - 1

            If isNullOrEmpty(tblMHI.Rows(crow).Cells(47).Value) Then
                tblMHI.Rows(crow).Cells(47).Value = False
            End If
        Next

        If isNullOrEmpty(DGridCEI.Rows(2).Cells(2).Value) Then
            '   Call Delete_Rows()
        End If

        'MsgBox("Data fetched successfully")

    End Sub

    Sub Delete_Rows()
        Try

            For trow = 0 To tblMHI.Rows.Count - 1

                If isNullOrEmpty(tblMHI.Rows(trow).Cells(0).Value) Then
                    tblMHI.Rows.RemoveAt(trow)

                    If IsDBNull(tblMHI.Rows(trow + 2).Cells(0).Value) Then
                        tblMHI.Rows.RemoveAt(trow + 2)
                    End If
                End If

            Next

            Call Build_Accural_Column_Formulas()  'Rebuilds formulas in columns AU - AT since deleting a

        Catch ex As Exception

        End Try
    End Sub
    Function dgridmhi_() As System.Data.DataTable
        Dim row_ince As Integer = 2
        Dim dt As New System.Data.DataTable()
        dt = creating_dt()

        For i As Int32 = 0 To DGridMHI.Rows.Count - 2
            dt.Rows.Add(New Object() {DGridMHI.Rows(i).Cells(0).Value,
 DGridMHI.Rows(i).Cells(1).Value,
 DGridMHI.Rows(i).Cells(2).Value,
 DGridMHI.Rows(i).Cells(3).Value,
 DGridMHI.Rows(i).Cells(4).Value,
 DGridMHI.Rows(i).Cells(5).Value,
 DGridMHI.Rows(i).Cells(6).Value,
 DGridMHI.Rows(i).Cells(7).Value,
 DGridMHI.Rows(i).Cells(8).Value,
 DGridMHI.Rows(i).Cells(9).Value,
 DGridMHI.Rows(i).Cells(10).Value,
 DGridMHI.Rows(i).Cells(11).Value,
 DGridMHI.Rows(i).Cells(12).Value,
 DGridMHI.Rows(i).Cells(13).Value,
 DGridMHI.Rows(i).Cells(14).Value,
 DGridMHI.Rows(i).Cells(15).Value,
 DGridMHI.Rows(i).Cells(16).Value,
 DGridMHI.Rows(i).Cells(17).Value,
 DGridMHI.Rows(i).Cells(18).Value,
 DGridMHI.Rows(i).Cells(19).Value,
 DGridMHI.Rows(i).Cells(20).Value,
 DGridMHI.Rows(i).Cells(21).Value,
 DGridMHI.Rows(i).Cells(22).Value,
 DGridMHI.Rows(i).Cells(23).Value,
 DGridMHI.Rows(i).Cells(24).Value,
 DGridMHI.Rows(i).Cells(25).Value,
 DGridMHI.Rows(i).Cells(26).Value,
 DGridMHI.Rows(i).Cells(27).Value,
 DGridMHI.Rows(i).Cells(28).Value,
 DGridMHI.Rows(i).Cells(29).Value,
 DGridMHI.Rows(i).Cells(30).Value,
 DGridMHI.Rows(i).Cells(31).Value,
 DGridMHI.Rows(i).Cells(32).Value,
 DGridMHI.Rows(i).Cells(33).Value,
 DGridMHI.Rows(i).Cells(34).Value,
 DGridMHI.Rows(i).Cells(35).Value,
 DGridMHI.Rows(i).Cells(36).Value,
 DGridMHI.Rows(i).Cells(37).Value,
 DGridMHI.Rows(i).Cells(38).Value,
 DGridMHI.Rows(i).Cells(39).Value,
 DGridMHI.Rows(i).Cells(40).Value,
 DGridMHI.Rows(i).Cells(41).Value,
 DGridMHI.Rows(i).Cells(42).Value,
 DGridMHI.Rows(i).Cells(43).Value,
 DGridMHI.Rows(i).Cells(44).Value,
 DGridMHI.Rows(i).Cells(45).Value,
 DGridMHI.Rows(i).Cells(46).Value,
 DGridMHI.Rows(i).Cells(47).Value,
 DGridMHI.Rows(i).Cells(48).Value,
 DGridMHI.Rows(i).Cells(49).Value,
 DGridMHI.Rows(i).Cells(50).Value,
 DGridMHI.Rows(i).Cells(51).Value})
        Next


        Dim dttables_temp As List(Of System.Data.DataTable)
        dttables_temp = list(dt, dt.Rows.Count)
        Dim chkblank As Boolean = False
        Dim chkblank_else As Boolean = False
        Dim dt_count_ As Integer = 0
        'Form1.tblMHI.DataSource = dttables_
        Dim Final_dt_temp As New System.Data.DataTable
        Dim teamptable As New System.Data.DataTable



        teamptable = dt.Clone



        Final_dt_temp = dt.Clone
        Dim cnt_ As Integer
        Dim dt_count_temp As Integer = dttables_temp.Count
        For Each datatable As System.Data.DataTable In dttables_temp
            Dim dtview As New DataView(datatable)
            dtview.Sort = "From DESC"

            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt_ = 0
            dt_count_ += 1

            'For Each rw As DataRow In dt_dt.Rows

            '    If dt_dt.Rows.Count - 1 = cnt_ And dt_count_ <> dt_count_temp Then
            '        Final_dt_temp.ImportRow(rw)
            '        rw = Final_dt_temp.NewRow()
            '        Final_dt_temp.Rows.Add(rw)
            '    Else
            '        Final_dt_temp.ImportRow(rw)
            '    End If

            '    cnt_ += 1
            'Next
            'Dim newrow As DataRow = Final_dt.NewRow()
            'newrow = Final_dt.NewRow()

            ''------------------------------------------ New logic 



            If Not isNullOrEmpty(DGridCEI.Rows(row_ince).Cells(3).Value) Then
                BrLDate = Format(CDate(DGridCEI.Rows(row_ince).Cells(3).Value), "MM/dd/yyyy")
                chkblank = False
                For q = 0 To dt_dt.Rows.Count - 1
                    If Not isNullOrEmpty(dt_dt.Rows(q).Item(0)) Then
                        If Format(CDate(dt_dt.Rows(q).Item(0)), "MM/dd/yyyy") <= CDate(BrLDate) Then
                            teamptable.ImportRow(dt_dt.Rows(q))
                            'Exit For
                            chkblank = True
                        End If
                    End If
                Next

                If chkblank Then
                    teamptable.Rows.Add(teamptable.NewRow())
                End If
                chkblank = False
                For q = 0 To dt_dt.Rows.Count - 1
                    If Not isNullOrEmpty(dt_dt.Rows(q).Item(0)) Then
                        If Format(CDate(dt_dt.Rows(q).Item(0)), "MM/dd/yyyy") > CDate(BrLDate) Then
                            teamptable.ImportRow(dt_dt.Rows(q))
                            'Exit For
                            chkblank = True
                        End If
                    End If
                Next
                If chkblank Then
                    teamptable.Rows.Add(teamptable.NewRow())
                End If

            Else

                chkblank_else = False
                For Each rw As DataRow In dt_dt.Rows
                    If dt_dt.Rows.Count - 1 = cnt_ And dt_count_ <> dt_count_temp Then
                        teamptable.ImportRow(rw)
                        rw = teamptable.NewRow()
                        teamptable.Rows.Add(rw)
                        chkblank_else = True
                    Else
                        chkblank_else = True
                        teamptable.ImportRow(rw)
                    End If
                    cnt_ += 1
                Next


            End If


            row_ince += 3

            ''------------------------------------
        Next

        '       Dim tblcnt As Integer = teamptable.Rows.Count - 1
        Try
            If chkblank_else = False Then
                Dim tblcnt As Integer = teamptable.Rows.Count - 1
                teamptable.Rows.RemoveAt(tblcnt)
            End If
        Catch ex As Exception

        End Try





        'DGridMHI_II.Refresh()

        '        Return Final_dt_
        Return teamptable

    End Function

    'Call GetTable()
    Private Sub ResetViewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetViewToolStripMenuItem.Click
        tblMHI.Sort(tblMHI.Columns(37), ListSortDirection.Ascending)
        tblMHI.Sort(tblMHI.Columns(0), ListSortDirection.Descending)
    End Sub

    Sub Format_Sort()

        tblMHI.Sort(tblMHI.Columns(26), ListSortDirection.Ascending)
        tblMHI.Refresh()

        'Range("A5").Select
        'Selection.Sort Key1:=Range("AA5"), Order1:=xlAscending, Key2:=Range("AW5") _
        '            , Order2:=xlAscending, Key3:=Range("Y5"), Order3:=xlAscending, Header _
        '            :=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom

        Call Build_Accural_Column_Formulas()  'Rebuilds formulas in columns AU - AT since deleting a


    End Sub

    'Sort by Member name then DOS
    Sub MbrSort()

        Dim tblCount As Integer = 0

        Dim dttables_SORT As New List(Of System.Data.DataTable)
        dttables_SORT = list(Final_dt_main, Final_dt_main.Rows.Count)
        Dim dt_count As Integer = dttables_SORT.Count
        'Form1.tblMHI.DataSource = dttables_


        Final_dt_main_SORT = New Data.DataTable

        Final_dt_main_SORT = Final_dt_main.Clone
        Dim cnt As Integer
        For Each datatable As System.Data.DataTable In dttables_SORT
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "PT_Name Asc,ICNandSuffix Asc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows


                If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_main_SORT.ImportRow(rw)
                    rw = Final_dt_main_SORT.NewRow()
                    Final_dt_main_SORT.Rows.Add(rw)
                Else
                    Final_dt_main_SORT.ImportRow(rw)
                End If
                cnt += 1
            Next
        Next
        tblMHI.DataSource = Nothing
        tblMHI.DataSource = Final_dt_main_SORT

        'Dim dtview As New DataView(Final_dt_main)
        'dtview.Sort = "PT_Name Asc,ICNandSuffix Asc,From Asc"
        'Dim dt_dt As System.Data.DataTable = dtview.ToTable()
        'tblMHI.DataSource = Nothing
        'tblMHI.DataSource = dt_dt


        Call Build_Accural_Column_Formulas()
    End Sub
    'Sort by Member name, Processed date and ICN/Suffix
    Sub MbrSort_Second()


        Dim tblCount As Integer = 0



        Dim dttables_SORT As New List(Of System.Data.DataTable)
        dttables_SORT = list(Final_dt_main, Final_dt_main.Rows.Count)
        Dim dt_count As Integer = dttables_SORT.Count
        'Form1.tblMHI.DataSource = dttables_


        Final_dt_main_SORT = New Data.DataTable

        Final_dt_main_SORT = Final_dt_main.Clone
        Dim cnt As Integer
        For Each datatable As System.Data.DataTable In dttables_SORT
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "ProcDate Asc,ICNandSuffix Asc,ProcDate Asc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows


                If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_main_SORT.ImportRow(rw)
                    rw = Final_dt_main_SORT.NewRow()
                    Final_dt_main_SORT.Rows.Add(rw)
                Else
                    Final_dt_main_SORT.ImportRow(rw)
                End If
                cnt += 1
            Next
        Next
        tblMHI.DataSource = Nothing
        tblMHI.DataSource = Final_dt_main_SORT


        'Dim dtview As New DataView(Final_dt_main)
        'dtview.Sort = "ProcDate Asc,ICNandSuffix Asc,ProcDate Asc"
        'Dim dt_dt As System.Data.DataTable = dtview.ToTable()
        'tblMHI.DataSource = Nothing
        'tblMHI.DataSource = dt_dt

        Call Build_Accural_Column_Formulas()
    End Sub
    Sub ICNSort()

        Dim tblCount As Integer = 0

        Dim dttables_SORT As New List(Of System.Data.DataTable)
        dttables_SORT = list(Final_dt_main, Final_dt_main.Rows.Count)
        Dim dt_count As Integer = dttables_SORT.Count
        'Form1.tblMHI.DataSource = dttables_


        Final_dt_main_SORT = New Data.DataTable

        Final_dt_main_SORT = Final_dt_main.Clone
        Dim cnt As Integer
        For Each datatable As System.Data.DataTable In dttables_SORT
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "PT_Name Asc,Suf DESC"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows


                If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_main_SORT.ImportRow(rw)
                    rw = Final_dt_main_SORT.NewRow()
                    Final_dt_main_SORT.Rows.Add(rw)
                Else
                    Final_dt_main_SORT.ImportRow(rw)
                End If
                cnt += 1
            Next
        Next
        tblMHI.DataSource = Nothing
        tblMHI.DataSource = Final_dt_main_SORT

        'Dim dtview As New DataView(Final_dt_main)
        'dtview.Sort = "PT_Name Asc,Suf DESC"
        'Dim dt_dt As System.Data.DataTable = dtview.ToTable()
        'tblMHI.DataSource = Nothing
        'tblMHI.DataSource = dt_dt

        Call Build_Accural_Column_Formulas()
    End Sub
    Sub ProvSort()

        Dim tblCount As Integer = 0

        Dim dttables_SORT As New List(Of System.Data.DataTable)
        dttables_SORT = list(Final_dt_main, Final_dt_main.Rows.Count)
        Dim dt_count As Integer = dttables_SORT.Count
        'Form1.tblMHI.DataSource = dttables_


        Final_dt_main_SORT = New Data.DataTable

        Final_dt_main_SORT = Final_dt_main.Clone
        Dim cnt As Integer
        For Each datatable As System.Data.DataTable In dttables_SORT
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "PT_Name DESC,TIN Asc,Suffix Asc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows


                If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_main_SORT.ImportRow(rw)
                    rw = Final_dt_main_SORT.NewRow()
                    Final_dt_main_SORT.Rows.Add(rw)
                Else
                    Final_dt_main_SORT.ImportRow(rw)
                End If
                cnt += 1
            Next
        Next
        tblMHI.DataSource = Nothing
        tblMHI.DataSource = Final_dt_main_SORT

        'Dim dtview As New DataView(Final_dt_main)
        'dtview.Sort = "PT_Name DESC,TIN Asc,Suffix Asc"
        'Dim dt_dt As System.Data.DataTable = dtview.ToTable()
        'tblMHI.DataSource = Nothing
        'tblMHI.DataSource = dt_dt

        Call Build_Accural_Column_Formulas()

    End Sub
    'Sort by TIN
    Sub TINSort()
        tblMHI.Sort(tblMHI.Columns(38), ListSortDirection.Ascending)
        tblMHI.Sort(tblMHI.Columns(22), ListSortDirection.Descending)
        tblMHI.Sort(tblMHI.Columns(0), ListSortDirection.Ascending)

        Call Build_Accural_Column_Formulas()
    End Sub
    'Sort by Ded Indicator
    Sub DedIndSort()
        tblMHI.Sort(tblMHI.Columns(14), ListSortDirection.Descending)
        Call Build_Accural_Column_Formulas()
    End Sub
    'Sort by %
    Sub PercentSort()
        tblMHI.Sort(tblMHI.Columns(15), ListSortDirection.Descending)
        Call Build_Accural_Column_Formulas()
    End Sub
    Sub ICN_Draft_Sort()

        Dim dt As New System.Data.DataTable()
        dt = creating_dt()

        For i As Int32 = 0 To tblMHI.Rows.Count - 2
            dt.Rows.Add(New Object() {tblMHI.Rows(i).Cells(0).Value,
                        tblMHI.Rows(i).Cells(1).Value,
                        tblMHI.Rows(i).Cells(2).Value,
                        tblMHI.Rows(i).Cells(3).Value,
                        tblMHI.Rows(i).Cells(4).Value,
                        tblMHI.Rows(i).Cells(5).Value,
                        tblMHI.Rows(i).Cells(6).Value,
                        tblMHI.Rows(i).Cells(7).Value,
                        tblMHI.Rows(i).Cells(8).Value,
                        tblMHI.Rows(i).Cells(9).Value,
                        tblMHI.Rows(i).Cells(10).Value,
                        tblMHI.Rows(i).Cells(11).Value,
                        tblMHI.Rows(i).Cells(12).Value,
                        tblMHI.Rows(i).Cells(13).Value,
                        tblMHI.Rows(i).Cells(14).Value,
                        tblMHI.Rows(i).Cells(15).Value,
                        tblMHI.Rows(i).Cells(16).Value,
                        tblMHI.Rows(i).Cells(17).Value,
                        tblMHI.Rows(i).Cells(18).Value,
                        tblMHI.Rows(i).Cells(19).Value,
                        tblMHI.Rows(i).Cells(20).Value,
                        tblMHI.Rows(i).Cells(21).Value,
                        tblMHI.Rows(i).Cells(22).Value,
                        tblMHI.Rows(i).Cells(23).Value,
                        tblMHI.Rows(i).Cells(24).Value,
                        tblMHI.Rows(i).Cells(25).Value,
                        tblMHI.Rows(i).Cells(26).Value,
                        tblMHI.Rows(i).Cells(27).Value,
                        tblMHI.Rows(i).Cells(28).Value,
                        tblMHI.Rows(i).Cells(29).Value,
                        tblMHI.Rows(i).Cells(30).Value,
                        tblMHI.Rows(i).Cells(31).Value,
                        tblMHI.Rows(i).Cells(32).Value,
                        tblMHI.Rows(i).Cells(33).Value,
                        tblMHI.Rows(i).Cells(34).Value,
                        tblMHI.Rows(i).Cells(35).Value,
                        tblMHI.Rows(i).Cells(36).Value,
                        tblMHI.Rows(i).Cells(37).Value,
                        tblMHI.Rows(i).Cells(38).Value,
                        tblMHI.Rows(i).Cells(39).Value,
                        tblMHI.Rows(i).Cells(40).Value,
                        tblMHI.Rows(i).Cells(41).Value,
                        tblMHI.Rows(i).Cells(42).Value,
                        tblMHI.Rows(i).Cells(43).Value,
                        tblMHI.Rows(i).Cells(44).Value,
                        tblMHI.Rows(i).Cells(45).Value,
                        tblMHI.Rows(i).Cells(46).Value,
                        tblMHI.Rows(i).Cells(47).Value,
                        tblMHI.Rows(i).Cells(48).Value,
                        tblMHI.Rows(i).Cells(49).Value,
                        tblMHI.Rows(i).Cells(50).Value,
                        tblMHI.Rows(i).Cells(51).Value})
        Next

        Dim tblCount As Integer = 0

        Dim dttables_ As New List(Of System.Data.DataTable)
        dttables_ = list(dt, dt.Rows.Count)
        Dim dt_count As Integer = dttables_.Count
        'Form1.tblMHI.DataSource = dttables_

        Dim Final_dt As New System.Data.DataTable
        Final_dt = dt.Clone
        Dim cnt As Integer
        For Each datatable As System.Data.DataTable In dttables_
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "From DESC"
            dtview.Sort = "PT_Rel Asc,ClaimNumber Asc,TotalPaid Asc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows
                Final_dt.ImportRow(rw)
                cnt += 1
            Next
        Next

        tblMHI.DataSource = Nothing
        tblMHI.Refresh()
        tblMHI.DataSource = Final_dt


        Call Build_Accural_Column_Formulas()
    End Sub

    'Sort by Patient then Processed date
    Sub Sort_ProcDate()
        Dim SortOrder, OOPCalcRun

        SortOrder = MsgBox("Do you want the process date sorted in Ascending order? Clicking " &
                            "No will sort in Descending order.", vbYesNo + vbDefaultButton2)
        If SortOrder = vbYes Then
            SortOrder = "xlAscending"
        Else
            SortOrder = "xlDescending"
        End If
        OOPCalcRun = tblMHI.Rows(0).Cells(48)

        If SortOrder = vbYes And OOPCalcRun = "" Then

            Dim tblCount As Integer = 0

            Dim dttables_SORT As New List(Of System.Data.DataTable)
            dttables_SORT = list(Final_dt_main, Final_dt_main.Rows.Count)
            Dim dt_count As Integer = dttables_SORT.Count
            'Form1.tblMHI.DataSource = dttables_


            Final_dt_main_SORT = New Data.DataTable

            Final_dt_main_SORT = Final_dt_main.Clone
            Dim cnt As Integer
            For Each datatable As System.Data.DataTable In dttables_SORT
                tblCount += 1
                Dim dtview As New DataView(datatable)
                dtview.Sort = "PT_Name Desc,From ASC"
                Dim dt_dt As System.Data.DataTable = dtview.ToTable()
                cnt = 0
                For Each rw As DataRow In dt_dt.Rows


                    If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                        Final_dt_main_SORT.ImportRow(rw)
                        rw = Final_dt_main_SORT.NewRow()
                        Final_dt_main_SORT.Rows.Add(rw)
                    Else
                        Final_dt_main_SORT.ImportRow(rw)
                    End If
                    cnt += 1
                Next
            Next
            tblMHI.DataSource = Nothing
            tblMHI.DataSource = Final_dt_main_SORT
        Else
            Call Custom_Sort(SortOrder)
        End If
        Call Build_Accural_Column_Formulas()
    End Sub
    'Sort by processed date only
    Sub ProcDate_Only(ByVal blnFormat As Boolean)
        Dim LastRow, SortOrder, OOPCalcRun

        OOPCalcRun = DGridMHI.Rows(0).Cells(48).Value

        'If OOPCalcRun = "" Then
        '    'MsgBox("Please run the OOP calculation macro prior to sorting by processed date.")
        '    Exit Sub
        'End If
        'If blnFormat = False Or blnFormat = "" Then
        '    SortOrder = MsgBox("Do you want the process date sorted in Ascending order? Clicking " &
        '                   "No will sort in Descending order.", vbYesNo + vbDefaultButton2)
        'Else
        '    SortOrder = vbYes
        'End If

        ''need to clear value 
        'Range(Cells(LastRow, 44), Cells(5, 47)).Select
        'Selection.Clear

        'If SortOrder = vbYes Then
        '    tblMHI.Sort(tblMHI.Columns(0), ListSortDirection.Ascending)

        'Else
        '    tblMHI.Sort(tblMHI.Columns(0), ListSortDirection.Descending)
        'End If

        Sort_ProcesssedDateOnly()
        Sort_ProcesssedDateOnly()
        'tblMHI.Sort(tblMHI.Columns(26), ListSortDirection.Ascending)


    End Sub
    Sub Custom_Sort(SortOrder)
        Dim tblCount As Integer = 0

        Dim dttables_SORT As New List(Of System.Data.DataTable)
        dttables_SORT = list(Final_dt_main, Final_dt_main.Rows.Count)
        Dim dt_count As Integer = dttables_SORT.Count
        'Form1.tblMHI.DataSource = dttables_


        Final_dt_main_SORT = New Data.DataTable

        Final_dt_main_SORT = Final_dt_main.Clone
        Dim cnt As Integer
        For Each datatable As System.Data.DataTable In dttables_SORT
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "PT_Name Desc,From Desc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows


                If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_main_SORT.ImportRow(rw)
                    rw = Final_dt_main_SORT.NewRow()
                    Final_dt_main_SORT.Rows.Add(rw)
                Else
                    Final_dt_main_SORT.ImportRow(rw)
                End If
                cnt += 1
            Next
        Next
        tblMHI.DataSource = Nothing
        tblMHI.DataSource = Final_dt_main_SORT
    End Sub

    Private Sub SortByICNToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SortByICNToolStripMenuItem.Click
        ICNSort()
        cell_colorback()
    End Sub

    Private Sub PatientAndProcessedDateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PatientAndProcessedDateToolStripMenuItem.Click
        Call MbrSort_Second()
        cell_colorback()
    End Sub

    Private Sub ProviderTinAndSuffixToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProviderTinAndSuffixToolStripMenuItem.Click
        ProvSort()
        cell_colorback()
    End Sub

    Private Sub DeductibleIndicatorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeductibleIndicatorToolStripMenuItem.Click
        Call DedIndSort()
        cell_colorback()
    End Sub

    Private Sub SortByPercentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SortByPercentToolStripMenuItem.Click
        PercentSort()
        cell_colorback()
    End Sub
    ''deleting duplicate row 

    Sub Del_DuplicateRow()
        Try

            Dim r As Integer
            Dim q As Integer

            Dim uniqueRow As String
            Dim duprow As String

            For r = 0 To tblMHI.Rows.Count - 1

                If Trim(tblMHI.Rows(r).Cells(0).Value) = "" Then
                    Exit For
                End If

                uniqueRow = (Trim(tblMHI.Rows(r).Cells(0).Value) &
                        Trim(tblMHI.Rows(r).Cells(1).Value) &
                        Trim(tblMHI.Rows(r).Cells(2).Value) &
                        Trim(tblMHI.Rows(r).Cells(9).Value) &
                        Trim(tblMHI.Rows(r).Cells(10).Value) &
                        Trim(tblMHI.Rows(r).Cells(12).Value) &
                        Trim(tblMHI.Rows(r).Cells(13).Value) &
                        Trim(tblMHI.Rows(r).Cells(22).Value) &
                        Trim(tblMHI.Rows(r).Cells(28).Value) &
                        Trim(tblMHI.Rows(r).Cells(29).Value) &
                        Trim(tblMHI.Rows(r).Cells(30).Value))
                q = 0

                For q = (r + 1) To tblMHI.Rows.Count - 2

                    If Trim(tblMHI.Rows(q).Cells(0).Value) = "" Then
                        Exit For
                    End If
                    duprow = (Trim(tblMHI.Rows(q).Cells(0).Value) &
                        Trim(tblMHI.Rows(q).Cells(1).Value) &
                        Trim(tblMHI.Rows(q).Cells(2).Value) &
                        Trim(tblMHI.Rows(q).Cells(9).Value) &
                        Trim(tblMHI.Rows(q).Cells(10).Value) &
                        Trim(tblMHI.Rows(q).Cells(12).Value) &
                        Trim(tblMHI.Rows(q).Cells(13).Value) &
                        Trim(tblMHI.Rows(q).Cells(22).Value) &
                        Trim(tblMHI.Rows(q).Cells(28).Value) &
                        Trim(tblMHI.Rows(q).Cells(29).Value) &
                        Trim(tblMHI.Rows(q).Cells(30).Value))

                    'MsgBox(uniqueRow)
                    'MsgBox(duprow)

                    If uniqueRow = duprow Then
                        tblMHI.Rows.RemoveAt(q)
                        tblMHI.Refresh()
                    End If

                Next q
            Next r
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnMain_Click(sender As Object, e As EventArgs) Handles btnMain.Click
        memberList.Items.Clear()
        'COSMOS_Window_Selection(txt_Policy.Text, txt_SSN.Text)
        starttime = Now()
        Call Get_CIE()
    End Sub
    Sub Get_CIE()
        strSSN = "S" & UCase(txt_SSN.Text)
        If InStr(strSSN, "SS") > 0 Then
            strSSN = Replace(strSSN, "SS", "S")
        End If
        Dim CEI_Response = api_CEI.PerformQuery(txt_Policy.Text, strSSN, txtUserID.Text, txtUserPass.Text)
        Dim colcnt As Int32 = 0
        Dim rcount As Integer = 1


        Try

            Dim strName As String
            Dim strdob As String
            For Each ceiRsp In CEI_Response.Results.Response

                strName = ceiRsp.rspCustEligHdrData.rspCustEmpFname
                strdob = ceiRsp.rspCustEligHdrData.rspCustEmpLname
                strdob = ceiRsp.rspCustEligHdrData.rspCustOffCity
                ''updating member in MMIOverview
                '     DGridOverview.Rows(0).Cells(1).Value = ceiRsp.rspCustEligHdrData.rspCustEmpFname & " " & ceiRsp.rspCustEligHdrData.rspCustEmpLname
                'DGridOverview.Rows(1).Cells(1).Value = ceiRsp.rspCustEligHdrData.rspCustEmpAddr
                'DGridOverview.Rows(2).Cells(1).Value = ceiRsp.rspCustEligHdrData.rspCustEmpCity & " " & ceiRsp.rspCustEligHdrData.rspCustEmpSt & " " & ceiRsp.rspCustEligHdrData.rspCustEmpZip


                DGridMInfo.Rows.Add(ceiRsp.rspCustEligHdrData.rspCustEmpLname,
                                      ceiRsp.rspCustEligHdrData.rspCustEmpAddr,
                                      ceiRsp.rspCustEligHdrData.rspCustEmpOffNbrSys.rspCustEmpOffNbr & ceiRsp.rspCustEligHdrData.rspCustEmpOffNbrSys.rspCustEmpSys)
                DGridMInfo.Rows.Add(rcount)
                DGridMInfo.Rows(1).Cells(1).Value = ceiRsp.rspCustEligHdrData.rspCustEmpCity & " " & ceiRsp.rspCustEligHdrData.rspCustEmpSt & " " & ceiRsp.rspCustEligHdrData.rspCustEmpZip

                ''''line

                Dim rcnt As Integer = 0
                Dim j = ceiRsp.rspCustEligCovData.Count - 1
                For memCnt = 0 To j
                    DGridCEI.Rows.Insert(rcnt)

                    DGridCEI.Rows(rcnt).Cells(0).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustFname
                    DGridCEI.Rows(rcnt).Cells(1).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustRelCd
                    DGridCEI.Rows(rcnt).Cells(2).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustDobMMDDYY
                    rcnt = rcnt + 1
                    DGridCEI.Rows.Insert(rcnt)

                    DGridCEI.Rows(rcnt).Cells(0).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustCurPlan
                    DGridCEI.Rows(rcnt).Cells(1).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustCurRptCd
                    DGridCEI.Rows(rcnt).Cells(2).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustCurEffDte
                    DGridCEI.Rows(rcnt).Cells(3).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustCurCanDte
                    rcnt = rcnt + 1
                    DGridCEI.Rows.Insert(rcnt)

                    DGridCEI.Rows(rcnt).Cells(0).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustPrvPlan
                    DGridCEI.Rows(rcnt).Cells(1).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustPrvRptCd
                    DGridCEI.Rows(rcnt).Cells(2).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustPrvEffDte
                    DGridCEI.Rows(rcnt).Cells(3).Value = ceiRsp.rspCustEligCovData(memCnt).rspCustPrvCanDte
                    rcnt = rcnt + 1
                    DGridCEI.Rows.Insert(rcnt)

                    memberList.Items.Add(ceiRsp.rspCustEligCovData(memCnt).rspCustFname & "/" & ceiRsp.rspCustEligCovData(memCnt).rspCustRelCd)

                    'DGridCEI.Rows.Add(ceiRsp.rspCustEligCovData(memCnt).rspCustFname,
                    '                  ceiRsp.rspCustEligCovData(memCnt).rspCustRelCd,
                    '                  ceiRsp.rspCustEligCovData(memCnt).rspCustDobMMDDYY,
                    '                  ceiRsp.rspCustEligCovData(memCnt).rspCustCurEffDte,
                    '                  ceiRsp.rspCustEligCovData(memCnt).rspCustCurCanDte)
                Next

            Next

        Catch ex As Exception
            MsgBox(ex.Message.ToString())
            '            COSMOS_Window_Selection(txt_Policy.Text, txt_SSN.Text)
        End Try
    End Sub

    Public Sub CopayTablist(pol As String, plancode As String, strYear As String, strname As String, causcd As String, strpos As String, mmccode As String, spcode1 As String, spcode2 As String, strcopayset As String)
        tblCopay.Rows.Add(pol, plancode, strYear, strPTName, causcd, strpos, mmccode, spcode1, spcode2, strcopayset)
    End Sub

    '        Me.memberList.Items.Add("STEVEN/SP")

    'Dim apiDoc360obj As apiDoc360 = New apiDoc360
    ' strDoc360Data = apiDoc360obj.PerformQuery("S535192852", "", "").jsonResponse

    'Call get_ceiDetails() '''calling to get data from CEI Api
    'Call get_pmiDetails() '''calling to get data from PMI Api

    'Call Get_mxiDetails() '''calling to get data from PMI Api

    'Call History_detail() '''calling to get data from History Api
    'CHECK: If date ranges overlap
    Public Shared Function compareRange(s1Start As String, s1End As String, s2Start As String, s2End As String) As Boolean
        If IsDate(s1Start) And IsDate(s1End) And IsDate(s2Start) And IsDate(s2End) Then
            If ((CDate(Trim(s1Start)) >= CDate(Trim(s2Start))) And  'SET 1: Start   >=  SET 2: Start
            (CDate(Trim(s1Start)) <= CDate(Trim(s2End)))) Or        'SET 1: Start   <=  SET 2:End
            ((CDate(Trim(s1End)) >= CDate(Trim(s2Start))) And       'SET 1: End     >=  SET 2: Start
            (CDate(Trim(s1End)) <= CDate(Trim(s2End)))) Or          'SET 1: End     <=  SET 2:End
            ((CDate(Trim(s2Start)) >= CDate(Trim(s1Start))) And     'SET 2: Start   >=  SET 1: Start   
            (CDate(Trim(s2Start))) <= CDate(Trim(s1End))) Or        'SET 2: Start   <=  SET 1:End
            ((CDate(Trim(s2End)) >= CDate(Trim(s1Start))) And       'SET 2: End     >=  Set 1: Start
            (CDate(Trim(s2End))) <= CDate(Trim(s1End))) Then        'SET 2: End     <=  Set 1:End
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Private Sub GETHistoryToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim apiDoc360obj As apiDoc360 = New apiDoc360
        strDoc360Data = apiDoc360obj.PerformQuery("", "", "").jsonResponse
    End Sub

    Private Sub tblCopay_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles tblCopay.CellContentClick

    End Sub

    Private Sub ClearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
        Try
            MHI.Enabled = True
            MHI.Visible = True
            'Chkselect_memlist.Checked = False
            memberList.Items.Clear()



            DGridCEI.Rows.Clear() : DGridCEI.Refresh()
            DGridMInfo.Rows.Clear() : DGridMInfo.Refresh()
            DGridMHI.Rows.Clear() : DGridMHI.Refresh()



            DGridOverview.Rows.Clear() : DGridOverview.Refresh()
            DGridADJ.Rows.Clear() : DGridADJ.Refresh()
            DGrid_PG1.Rows.Clear() : DGrid_PG1.Refresh()
            DGrid_PG10.Rows.Clear() : DGrid_PG10.Refresh()



            DGrid_PG4.Rows.Clear() : DGrid_PG4.Refresh()



            DGrid_PG5.Rows.Clear() : DGrid_PG5.Refresh()



            tblOOP.Rows.Clear() : tblOOP.Refresh()
            tblCopay.Rows.Clear() : tblCopay.Refresh()
            DGridADJ.Rows.Clear() : DGridADJ.Refresh()
            DGrid_PG10.DataSource = Nothing
            DGridMInfo.DataSource = Nothing
            DGrid_PG1.DataSource = Nothing
            DGridOverview.DataSource = Nothing
            DGridCEI.DataSource = Nothing
            DGridMHI.DataSource = Nothing
            DGrid_PG4.DataSource = Nothing
            DGrid_PG5.DataSource = Nothing
            tblOOP.DataSource = Nothing
            DGridADJ.DataSource = Nothing
            tblCopay.DataSource = Nothing
            tblMHI.DataSource = Nothing
            For memberNumber = 0 To memberList.CheckedItems.Count - 1
                memberList.Items.Clear()
                memberList.Refresh()
            Next
            txt_Policy.Clear()
            txt_SSN.Clear()
            'DgridCopay.Rows.Clear()            
            tblMHI.DataSource = Nothing
            tblMHI.Refresh()



            mmiPag1RowsAdd()
            mmiPag4RowsAdd()
            mmiPag5RowsAdd()
            MMIPage10()
            MMIOverviw()
            mmiFlag = False
            mmiDetails.Clear()
            RichTextBox1.Clear()
            tblMHI.Columns.Clear()
            tblOOP.Columns.Clear()
            claimList.Clear()
            ceiList.Clear()
            Erase sArray

        Catch ex As Exception

        End Try
    End Sub



    Private Sub CalculateOOPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculateOOPToolStripMenuItem.Click


        'Call MbrSort()    ''''''''''''''calling for Sort first


        Erase EffDte

        rowCnt = DGrid_PG1.RowCount
        colCntPG1 = DGrid_PG1.Columns.Count

        q = 1
        Do
            If DGrid_PG1.Rows(0).Cells(q).Value = "" Then
                'q = q - 1
                Exit Do
            Else
                q = q + 1
            End If
        Loop

        For Cnt = 1 To q

            ReDim Preserve EffDte(MMICount)
            ReDim Preserve CancelDate(MMICount)
            ReDim Preserve MHIPTName(MMICount)

            'Format(CDate(startSelect.Text), "MM/dd/yyyy"),
            Dim dt = Split(DGrid_PG1.Rows(2).Cells(Cnt).Value, "-")

            If InStr(Mid(DGrid_PG1.Rows(2).Cells(Cnt).Value, 1, 10), "-") > 0 Then
                'EffDte(MMICount) = Format(CDate(Trim(dt(0))), "MM/dd/yyyy")
            Else

                'EffDte(MMICount) = Format(CDate(Trim(Mid(dt(0), 1, 10))), "MM/dd/yyyy")
            End If

            If InStr(DGrid_PG1.Rows(2).Cells(Cnt).Value, "- 99999999") > 0 Then
                'CancelDate(MMICount) = "12/31/" & Format(Now, "yyyy")
                ' CancelDate(MMICount) = Format(Now, "MM/dd/yyyy")
            Else
                'Dim dtvalue As String = dt(1).ToString
                'Dim dateTime As String = dtvalue
                'CancelDate(MMICount) = Convert.ToDateTime(dtvalue)

                'CancelDate(MMICount) = CDate(dtvalue.ToString)
                '    MsgBox(CancelDate(MMICount))
            End If

        Next

        Call SeparateData()
        Call GatherData()
        'Sorts spreadsheet by Member name, Processed date and ICN/Suffix.
        'Processed date allows macro to closely match the order in which the system applied
        'claims to the different buckets.
        'ICN/Suffix keeps all claim info for that specific ICN together on spreadsheet.
        'Call MbrSort_Second()               ''need to check 05/11/2023 Sanjeet
        Call PerfCalcs()                ''main Calculation procedure

    End Sub

    Sub SeparateData()

        MMICount = 0

        'MMI page 1

        rowCnt = DGrid_PG1.RowCount

        colCntPG1 = DGrid_PG1.Columns.Count
        rowCnt = DGrid_PG1.RowCount
        q = 1
        Do
            If DGrid_PG1.Rows(0).Cells(q).Value = "" Then
                q = q - 1
                Exit Do
            Else
                q = q + 1
            End If
        Loop


        For Cnt = 1 To q


            If DGridOverview.Rows(0).Cells(Cnt).Value = "" Then
                Exit For
            End If

            ReDim Preserve EffDte(MMICount)
            ReDim Preserve CancelDate(MMICount)
            ReDim Preserve MHIPTName(MMICount)

            Dim dt = Split(DGridOverview.Rows(6).Cells(Cnt).Value, "-")

            If InStr(Mid(DGridOverview.Rows(6).Cells(Cnt).Value, 1, 10), "-") > 0 Then
                EffDte(MMICount) = Format(CDate(Trim(dt(0))), "MM/dd/yyyy")
            Else

                EffDte(MMICount) = CDate(Trim(Mid(DGridOverview.Rows(6).Cells(Cnt).Value, 1, 10)))
            End If

            If InStr(DGridOverview.Rows(6).Cells(Cnt).Value, "- 99999999") > 0 Then
                CancelDate(MMICount) = "12/31/" & Format(Now, "yyyy")
            Else

                Dim dtvalue As String = dt(1).ToString
                'Dim dateTime As String = dtvalue
                CancelDate(MMICount) = Convert.ToDateTime(dtvalue)

                CancelDate(MMICount) = CDate(dtvalue.ToString)
                'MsgBox(CancelDate(MMICount))

            End If
            MHIPTName(MMICount) = DGridOverview.Rows(2).Cells(Cnt).Value
            'MsgBox(MHIPTName(0))

            If DGridOverview.Rows(0).Cells(Cnt).Value = "" Then Exit For

            Dim MmiEffdte
            MMICount = MMICount + 1
        Next

        Dim PTRows()
        MMICount = 0
        Cnt = 0

        For rnxt = 0 To tblMHI.Rows.Count - 2
            If Not IsDBNull(tblMHI.Rows(rnxt + 1).Cells(38).Value) Then
                If Trim(tblMHI.Rows(rnxt).Cells(38).Value.ToString()) = "" Then Continue For
                If tblMHI.Rows(rnxt).Cells(38).Value <> tblMHI.Rows(rnxt + 1).Cells(38).Value Then
                    ReDim Preserve PTRows(MMICount)
                    PTRows(MMICount) = rnxt
                    rnxt = rnxt + 1
                    MMICount = MMICount + 1
                End If
            Else
                rnxt = rnxt + 1
                MMICount = MMICount + 1
            End If
        Next






        Dim PTName
        Dim NbrAddedRows
        Dim nxtrow As Integer = 0
        For MMICount = UBound(EffDte) To 0 Step -1
            Try

                If tblMHI.Rows(MMICount).Cells(0).Value = "" Then
                    Exit For
                End If

            Catch ex As Exception
                'MessageBox.Show("Before click on OOP Calculator, you should click first on Get History button. Pelase clear the tool and run again OOP Calculator", "OOP Calculator", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End Try
            PTName = tblMHI.Rows(MMICount).Cells(38).Value
            DateOfSvc = Format(CDate(tblMHI.Rows(MMICount).Cells(0).Value), "MM/dd/yyyy")
            If PTName = MHIPTName(MMICount) Then
                If DateOfSvc >= EffDte(MMICount) And DateOfSvc <= CancelDate(MMICount) Then
                    Do
                        nxtrow = nxtrow + 1
                        If Trim(tblMHI.Rows(MMICount).Cells(0).Value) = "" Then
                            'MMICount = MMICount + 1
                            MMICount = UBound(EffDte) + 1
                            Exit Do
                        End If
                        If tblMHI.Rows.Count - 2 = nxtrow Then
                            MMICount = UBound(EffDte) + 1
                            Exit Do
                        End If

                        If IsDBNull(tblMHI.Rows(nxtrow).Cells(0).Value) Then            ''Added on 04/28/2023 Sanjeet to skip the error blank cell Sanjeet
                            Exit For
                        End If

                        DateOfSvc = Format(CDate(tblMHI.Rows(nxtrow).Cells(0).Value), "MM/dd/yyyy")

                        Exit For
                        If DateOfSvc < Format(EffDte(MMICount), "MM/dd/yyyy") Or DateOfSvc > Format(CancelDate(MMICount), "MM/dd/yyyy") Then
                            'Rows(ActiveCell.Row & ":" & ActiveCell.Row).Select                          '''''need to check this line sanjeet
                            'Selection.Insert Shift:=xlDown
                            'MMICount = MMICount + 1
                            NbrAddedRows = NbrAddedRows + 1
                            MMICount = UBound(EffDte) + 1
                            Exit Do
                        End If
                    Loop

                End If
            End If
        Next

    End Sub

    Private Sub MHIHistoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MHIHistoryToolStripMenuItem.Click

    End Sub

    Private Sub FormatMHISheetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FormatMHISheetToolStripMenuItem.Click

        Call Format_Sheet()

    End Sub

    Sub GatherData()

        Dim IndDed As String
        Dim FamDed As String
        Dim NewCoinsCode
        Dim IndNewCoins
        Dim EffCanclDate As String

        Erase EffDte
        Erase CancelDate
        Erase blnBeforeEff

        MMICount = 0

        'MMI page 1

        rowCnt = DGrid_PG1.RowCount
        q = 1
        Do
            If DGrid_PG1.Rows(0).Cells(q).Value = "" Then
                'q = q - 1
                Exit Do
            Else
                q = q + 1
            End If
        Loop
        colCntPG1 = DGrid_PG1.Columns.Count




        For Cnt = 1 To q - 1

            ReDim Preserve INNStatus(MMICount)
            INNStatus(MMICount) = True
            ReDim Preserve MHIPTName(MMICount)
            ReDim Preserve NonEmbed(MMICount)
            ReDim Preserve blnBeforeEff(MMICount)



            MHIPTName(MMICount) = DGridOverview.Rows(2).Cells(Cnt).Value
            NonEmbed(MMICount) = DGrid_PG1.Rows(19).Cells(Cnt).Value 'ActiveCell.Offset(19, 0).Value , Need to check
            EffCanclDate = DGridOverview.Rows(6).Cells(Cnt).Value

            ReDim Preserve EffDte(MMICount)
            ReDim Preserve CancelDate(MMICount)
            ReDim Preserve MHIPTName(MMICount)



            Dim dt = Split(DGridOverview.Rows(6).Cells(Cnt).Value, "-")

            If InStr(Mid(DGridOverview.Rows(6).Cells(Cnt).Value, 1, 10), "-") > 0 Then
                EffDte(MMICount) = Format(CDate(Trim(dt(0))), "MM/dd/yyyy")
            Else

                EffDte(MMICount) = CDate(Trim(Mid(DGridOverview.Rows(6).Cells(Cnt).Value, 1, 10)))
            End If
            'MsgBox(UBound(EffDte))
            If InStr(DGridOverview.Rows(6).Cells(Cnt).Value, "- 99999999") > 0 Then
                CancelDate(MMICount) = "12/31/" & Format(Now, "yyyy")
            Else

                Dim dtvalue As String = dt(1).ToString
                'Dim dateTime As String = dtvalue
                CancelDate(MMICount) = Convert.ToDateTime(dtvalue)

                CancelDate(MMICount) = CDate(dtvalue.ToString)
            End If
            'MsgBox(CancelDate(MMICount))

            ''  ActiveCell.Offset(0, 1).Select

            'MMI page 4, Individual and Family deductibles
            '       MMI4.Activate()
            ReDim Preserve InnIndDed(MMICount)
            ReDim Preserve OONIndDed(MMICount)
            ReDim Preserve InnFamDed(MMICount)
            ReDim Preserve InnDedCross(MMICount)  'Individual INN ded applies to OON ded
            ReDim Preserve OONDedCross(MMICount)  'Individual OON ded applies to INN ded
            ReDim Preserve OONFamDedCross(MMICount)  'Family OON ded applies to INN ded
            ReDim Preserve TieredCross(MMICount)    'Used for tiered plans


            IndDed = Trim(DGrid_PG4.Rows(7).Cells(Cnt).Value)
            If Trim(IndDed) = "MG" Then
                OONDedCross(MMICount) = True
            ElseIf IndDed = "U" Then
                INNStatus(MMICount) = False
                OONDedCross(MMICount) = False
                InnDedCross(MMICount) = False
            Else
                OONDedCross(MMICount) = False
            End If

            If IndDed = "MZ" Then
                TieredCross(MMICount) = True
            ElseIf InStr(IndDed, "M") > 0 And InStr(IndDed, "Z") > 0 Then
                TieredCross(MMICount) = True
            Else
                TieredCross(MMICount) = False
            End If

            'Row 21 on MMI page 4
            'IndDed = ActiveCell.Offset(20, 0).Value
            IndDed = Trim(DGrid_PG4.Rows(30).Cells(Cnt).Value)
            If Trim(IndDed) = "GM" Then
                InnDedCross(MMICount) = True
            Else
                InnDedCross(MMICount) = False
            End If
            FamDed = Trim(DGrid_PG4.Rows(28).Cells(Cnt).Value)
            If Trim(FamDed) = "MG" Then
                OONFamDedCross(MMICount) = True
            Else
                OONFamDedCross(MMICount) = False
            End If

            InnIndDed(MMICount) = Int(DGrid_PG4.Rows(6).Cells(Cnt).Value)        ''' need to check data type sanjeet
            OONIndDed(MMICount) = Int(DGrid_PG4.Rows(29).Cells(Cnt).Value)
            InnFamDed(MMICount) = Int(DGrid_PG4.Rows(48).Cells(Cnt).Value)
            'ActiveCell.Offset(0, 1).Select  need to check for next column value 

            'MMI page 5, Continuation of Family deductibles from page 4

            ReDim Preserve OONFamDed(MMICount)
            ReDim Preserve InnFamDedCross(MMICount)  'Family INN ded applies to OON ded

            'MMI page 5, Continuation of Family deductibles from page 4
            ReDim Preserve OONFamDed(MMICount)
            ReDim Preserve InnFamDedCross(MMICount)  'Family INN ded applies to OON ded

            FamDed = Trim(DGrid_PG5.Rows(25).Cells(Cnt).Value)

            If Trim(FamDed) = "GM" Then
                InnFamDedCross(MMICount) = True
            Else
                InnFamDedCross(MMICount) = True
            End If

            OONFamDed(MMICount) = DGrid_PG5.Rows(24).Cells(Cnt).Value

            'MMI page 10, Individual/Family out of pocket
            'MMI10.Activate()
            ReDim Preserve InnIndOOP(MMICount)
            ReDim Preserve DedToOOP(MMICount)  'Ind INN/OON Ded applies to OOP
            ReDim Preserve OONIndOOP(MMICount)
            ReDim Preserve OOPCross(MMICount)  'OOP cross applies
            ReDim Preserve InnFamOOP(MMICount)
            ReDim Preserve OONFamOOP(MMICount)
            ReDim Preserve OONPercent(MMICount)


            InnIndOOP(MMICount) = Int(DGrid_PG10.Rows(33).Cells(Cnt).Value)
            OONIndOOP(MMICount) = Int(DGrid_PG10.Rows(35).Cells(Cnt).Value)
            OONPercent(MMICount) = Convert.ToDouble(DGrid_PG10.Rows(7).Cells(Cnt).Value) * 100        '''need to check Sanjeet 07/25/2023
            InnFamOOP(MMICount) = Int(DGrid_PG10.Rows(36).Cells(Cnt).Value)
            OONFamOOP(MMICount) = Int(DGrid_PG10.Rows(37).Cells(Cnt).Value)

            NewCoinsCode = DGrid_PG10.Rows(46).Cells(Cnt).Value     'NEED TO CHECK Sanjeet

            IndNewCoins = DGrid_PG10.Rows(4).Cells(Cnt).Value
            'Indicates if Deductible cross applies to OOP.
            If IndNewCoins = "2" Then
                DedToOOP(MMICount) = True
            Else
                DedToOOP(MMICount) = False
            End If

            Select Case NewCoinsCode
                Case "0"
                    'No cross application
                    OOPCross(MMICount) = "None"

                Case "1"
                    'Dual - INN to OON and OON to INN
                    OOPCross(MMICount) = "1"

                Case "2"
                    'Dual to Normal
                    OOPCross(MMICount) = "2"

                Case "3"
                    'Normal to Dual
                    OOPCross(MMICount) = "3"

                Case "4"
                    'Dual/Tiered
                    OOPCross(MMICount) = "1"

                Case "5"
                    OOPCross(MMICount) = "None"

                Case "6"
                    OOPCross(MMICount) = "None"

                Case "7"
                    OOPCross(MMICount) = "None"

                Case Else
                    OOPCross(MMICount) = "None"

            End Select
            'Determines if copay applies towards the out of pocket. Only applies to special
            'processing codes "C" and "D".
            Dim SpclProcCode

            '''validating form Capay Datagrid
            ReDim Preserve DxCauseCode(MMICount)
            ReDim Preserve CopayPOS(MMICount)
            DxCauseCode(MMICount) = ""

            For crow = 0 To tblCopay.Rows.Count - 1

                'If tblCopay.Rows(crow).Cells(0).Value = 2 And tblCopay.Rows(crow).Cells(0).Value = "" Then
                If crow = 0 And tblCopay.Rows(crow).Cells(0).Value = "" Then
                    DxCauseCode(MMICount) = "N/A"
                    CopayPOS(MMICount) = "N/A"
                    Exit For
                End If

                If MHIPTName(MMICount) = Trim(tblCopay.Rows(crow).Cells(3).Value) Then
                    If Format(tblCopay.Rows(crow).Cells(2).Value, "MM/dd/yyyy") = Format(EffCanclDate, "MM/dd/yyyy") Then

                        SpclProcCode = Trim(tblCopay.Rows(crow).Cells(7).Value)
                        If SpclProcCode = "C" Or SpclProcCode = "D" Then
                            If Trim(tblCopay.Rows(crow).Cells(4).Value) = "" Then
                                If DxCauseCode(MMICount) = "Empty" Or DxCauseCode(MMICount) = "" Then
                                    DxCauseCode(MMICount) = "All"
                                End If
                            End If
                            If Trim(tblCopay.Rows(crow).Cells(5).Value) = "" Then
                                If CopayPOS(MMICount) = "Empty" Or CopayPOS(MMICount) = "" Then
                                    CopayPOS(MMICount) = "All"
                                End If
                            End If
                            If DxCauseCode(MMICount) = "Empty" Or DxCauseCode(MMICount) = "" Then
                                DxCauseCode(MMICount) = Trim(tblCopay.Rows(crow).Cells(4).Value)
                                CopayPOS(MMICount) = Trim(tblCopay.Rows(crow).Cells(5).Value)
                            Else
                                If DxCauseCode(MMICount) <> "All" Then
                                    DxCauseCode(MMICount) = DxCauseCode(MMICount) & "/" & Trim(tblCopay.Rows(crow).Cells(4).Value)

                                End If
                                If CopayPOS(MMICount) <> "All" Then
                                    CopayPOS(MMICount) = CopayPOS(MMICount) & "/" & Trim(tblCopay.Rows(crow).Cells(5).Value)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            MMICount = MMICount + 1
        Next
        ''do shorting top to Button For mhi sheet

    End Sub

    Sub PerfCalcs()

        'RichTextBox1.SelectionIndent = 5
        'RichTextBox1.BulletIndent = 4
        'RichTextBox1.SelectionBullet = True
        'RichTextBox1.AppendText("OOP Calculating in progress....." & vbCrLf)

        Dim SvcCode, OVRide, Payee, Percent, Remark, DedCode
        Dim Covered As Double    'currency 
        Dim NotCovered As Double 'currency 
        Dim Deduct As Double     'currency
        Dim ClmPaid As Double    'currency 
        Dim Charge As Double 'currency 
        Dim Oop As Double    'currency 
        Dim FirstRun As Boolean
        Dim blnRegClaim As Boolean
        Dim CauseCode As String, PotPaid As Double 'currency   
        Dim HxAdj, CurrentRow, PlanFound, OldPtName, PTNameChk, Sanction, InpHosp, PTName

        HxAdj = "H,D,T"     'Indicates claim is a history backload
        FirstRun = True
        'MMICount = 1
        Covered = FormatCurrency(Covered)
        Dim loopCounter As Integer = 0

        'For MMICount = UBound(EffDte) To 0 Step -1        
        'For nrow = tblMHI.Rows.Count To 0 Step -1
        Try

            'Dim nrow As Integer
            nrow = 0
            'For MMICount = UBound(EffDte) To 0 Step -1
            For MMICount = 0 To UBound(EffDte)
                'nrow = 0
                PTNameChk = MHIPTName(MMICount)
                '''''''added 05/01/2023

                'For nrow = 0 To tblMHI.Rows.Count
                Do

                    PlanFound = False
                    blnRegClaim = False
                    CurrentRow = nrow

                    DateOfSvc = Format(CDate(tblMHI.Rows(nrow).Cells(0).Value), "MM/dd/yyyy")

                    '''updating code for OIM and other  insurance and Facilty 
                    'Added for use with OI/OIM calculations with Inpatient
                    'facility bills

                    Dim strSVC As String = Mid(tblMHI.Rows(nrow).Cells(2).Value, 1, 2)
                    If InStr(strSVC, "PR") > 0 Or InStr(strSVC, "SP") > 0 Or InStr(strSVC, "IC") > 0 Then
                        blnFacility = True
                    End If

                    'Distinguishes an OI/OIM claim from a normal claim.
                    'OIOIMClaim = False
                    'If InStr(Mid(tblMHI.Rows(nrow).Cells(0).Value, 1, 2), "OI") > 0 Then
                    '    OIOIMClaim = True
                    'End If
                    'If OIOIMClaim = True Then
                    '    OtherIns = False
                    'Else
                    '    OtherIns = False
                    'End If

                    'tblMHI.Rows(nrow).Cells(47).Value = OtherIns

                    '''Adding new code to pupulate 0.00
                    'tblMHI.Rows(nrow).Cells(39).Value.NumberFormat = "0.00_);[Red](0.00)"
                    'tblMHI.Rows(nrow).Cells(40).Value.NumberFormat = "0.00_);[Red](0.00)"
                    'tblMHI.Rows(nrow).Cells(41).Value.NumberFormat = "0.00_);[Red](0.00)"
                    'tblMHI.Rows(nrow).Cells(42).Value.NumberFormat = "0.00_);[Red](0.00)"

                    tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                    '''
                    PTNameChk = MHIPTName(MMICount)

                    PTName = tblMHI.Rows(nrow).Cells(38).Value
                    If CurrentRow = 0 Then OldPtName = PTName

                    If PTName <> OldPtName Then
                        'MMICount = UBound(EffDte) + 1
                        OldPtName = PTName
                        Exit Do
                    End If                                                           '''need to check line no 1304 also need to uncomment 
                    If PTName <> PTNameChk Then Exit Do
                    Remark = tblMHI.Rows(nrow).Cells(8).Value

                    'If DateOfSvc >= EffDte(MMICount) And DateOfSvc <= CancelDate(MMICount) Then
                    PlanFound = True
                        If blnBeforeEff(MMICount) = True Then
                            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                            tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                            tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                            OldPtName = PTName
                            tblMHI.Rows(nrow).Cells(48).Value = FirstRun
                        Else
                            SvcCode = Trim(tblMHI.Rows(nrow).Cells(2).Value)
                            OVRide = Trim(tblMHI.Rows(nrow).Cells(5).Value)
                            Payee = tblMHI.Rows(nrow).Cells(6).Value
                            Percent = Replace(tblMHI.Rows(nrow).Cells(15).Value, "%", "")
                            Charge = tblMHI.Rows(nrow).Cells(9).Value
                            Charge = Math.Round(Charge, 2)
                            NotCovered = tblMHI.Rows(nrow).Cells(10).Value
                            NotCovered = Math.Round(NotCovered, 2)
                            Covered = tblMHI.Rows(nrow).Cells(12).Value
                            Covered = Math.Round(Covered, 2)
                            Deduct = tblMHI.Rows(nrow).Cells(13).Value
                            Deduct = Math.Round(Deduct, 2)
                            DedCode = tblMHI.Rows(nrow).Cells(14).Value
                            If DedCode = "Z" And TieredCross(MMICount) = True Then DedCode = "M"
                            ClmPaid = tblMHI.Rows(nrow).Cells(16).Value
                            ClmPaid = Math.Round(ClmPaid, 2)
                            Remark = tblMHI.Rows(nrow).Cells(8).Value
                            PlaceSvc = tblMHI.Rows(nrow).Cells(3).Value
                        CauseCode = tblMHI.Rows(nrow).Cells(20).Value
                        Try
                            OtherIns = tblMHI.Rows(nrow).Cells(47).Value
                        Catch ex As Exception
                        End Try

                        Sanction = UCase(tblMHI.Rows(nrow).Cells(19).Value)
                            tblMHI.Rows(nrow).Cells(50).Value = False
                            InpHosp = tblMHI.Rows(nrow).Cells(50).Value
                            tblMHI.Rows(nrow).Cells(48).Value = FirstRun

                            '*******************************************************
                            '*** Check what scenerio the service line applies to
                            '*******************************************************

                            If InStr(HxAdj, Payee) > 0 And InStr(SvcCode, "COPAY") = 0 Then
                                'Charge, Covered and Paid are equal go to Lifetime max
                                If Covered = ClmPaid And Charge = ClmPaid And Deduct = "0" Then
                                    tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                    tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                ElseIf Payee = "H" Or Payee = "T" Then
                                    If OVRide = "20" Or OVRide = "30" Then
                                        'Decrease deductible by amount in deductible field
                                        'Increase OOP by amount in covered field
                                        If Covered <> "0.00" And Covered <> "0" Then
                                            If INNStatus(MMICount) = True Then
                                                tblMHI.Rows(nrow).Cells(40).Value = Covered
                                                tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                                tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                                If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                                                    tblMHI.Rows(nrow).Cells(42).Value = Covered
                                                Else
                                                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                                End If

                                            Else
                                                tblMHI.Rows(nrow).Cells(42).Value = Covered
                                                tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                                tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                            End If
                                        Else
                                            If INNStatus(MMICount) = True Then
                                                Call Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)

                                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                                tblMHI.Rows(nrow).Cells(42).Value = "0.00"

                                            Else
                                                tblMHI.Rows(nrow).Cells(41).Value = Deduct
                                            End If
                                        End If
                                    ElseIf OVRide = "01" Or OVRide = "02" Then
                                        If Covered <> "0.00" And Covered <> "0" Then
                                            tblMHI.Rows(nrow).Cells(40).Value = Covered
                                            If INNStatus(MMICount) = True And Deduct <> "0.00" And Deduct <> "0" Then
                                                Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
                                            Else
                                                tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                            End If
                                            If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                                                tblMHI.Rows(nrow).Cells(42).Value = Covered
                                            Else
                                                tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                            End If
                                        Else
                                            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                        End If
                                    End If
                                ElseIf Payee = "D" Then
                                    If OVRide = "20" Then
                                        'Decrease deductible by amount in deductible field
                                        'Increase OOP by amount in covered field
                                        If Covered <> "0.00" And Covered <> "0" Then
                                            tblMHI.Rows(nrow).Cells(42).Value = Covered
                                            If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                                                tblMHI.Rows(nrow).Cells(40).Value = Covered
                                            Else
                                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                            End If
                                            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                        Else
                                            If INNStatus(MMICount) = True Then
                                                Call Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
                                            Else
                                                tblMHI.Rows(nrow).Cells(41).Value = Deduct
                                            End If

                                            tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                        End If
                                    ElseIf OVRide = "30" Then
                                        'Increase deductible by amount in deductible field
                                        'Decrease OOP by amount in covered field
                                        If Covered <> "0.00" And Covered <> "0" Then
                                            tblMHI.Rows(nrow).Cells(42).Value = Covered
                                            If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                                                tblMHI.Rows(nrow).Cells(40).Value = Covered
                                            Else
                                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                            End If
                                        Else
                                            If INNStatus(MMICount) = True Then
                                                Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
                                            Else
                                                tblMHI.Rows(nrow).Cells(41).Value = Deduct
                                            End If
                                            tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                        End If
                                    ElseIf OVRide = "01" Or OVRide = "02" Then
                                        If Covered <> "0.00" And Covered <> "0" Then
                                            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(42).Value = Covered
                                            If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                                                tblMHI.Rows(nrow).Cells(40).Value = Covered
                                            Else
                                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                            End If

                                        End If
                                    End If
                                End If
                            ElseIf OVRide = "20" And Payee = "1" Then
                                'Reimbursement of deductible/OOP to member
                                If INNStatus(MMICount) = True Then
                                    If Deduct <> "0.00" And Deduct <> "0" Then
                                        Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
                                    End If
                                    If Covered = "0" Or Covered = "0.00" Then
                                        If Remark = "74" Or Remark = "71" Then
                                            tblMHI.Rows(nrow).Cells(40).Value = (ClmPaid)   ''Abs(ClmPaid) need to check sanjeet
                                        Else
                                            If InStr(ClmPaid, "-") = 0 Then ClmPaid = "-" & ClmPaid
                                            tblMHI.Rows(nrow).Cells(40).Value = ClmPaid
                                        End If
                                        If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                                            If Remark = "74" Or Remark = "71" Then
                                            tblMHI.Rows(nrow).Cells(42).Value = ClmPaid   'Abs(ClmPaid)
                                        Else
                                                If InStr(ClmPaid, "-") = 0 Then ClmPaid = "-" & ClmPaid
                                                tblMHI.Rows(nrow).Cells(42).Value = ClmPaid
                                            End If

                                        End If
                                    End If
                                Else
                                    If Remark = "74" Or Remark = "71" Then
                                        tblMHI.Rows(nrow).Cells(41).Value = Deduct
                                    tblMHI.Rows(nrow).Cells(42).Value = ClmPaid  'Abs(ClmPaid)
                                Else
                                        If InStr(ClmPaid, "-") = 0 Then ClmPaid = "-" & ClmPaid
                                        tblMHI.Rows(nrow).Cells(41).Value = Deduct
                                        tblMHI.Rows(nrow).Cells(42).Value = ClmPaid
                                    End If
                                End If

                            ElseIf OVRide = "30" Or Remark = "70" Or OVRide = "Y0" Then
                                tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                            ElseIf OVRide = "P" And Remark = "71" Then
                                tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                            ElseIf OVRide = "R" And Remark = "71" Then
                                tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                'OI/OIM claims

                            ElseIf OtherIns = True Then
                                If InStr(SvcCode, "OI") > 0 And InpHosp = False Then
                                    tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                    tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                Else
                                    OI_OIM_Calcs(Percent, Charge, Covered, Deduct, DedCode, ClmPaid, NotCovered, PlaceSvc, CauseCode, SvcCode, InpHosp, Remark, OVRide)
                                End If
                                'Copay/Ncopay line
                            ElseIf InStr(SvcCode, "COPAY") > 0 Then
                                Copay_Calcs(NotCovered, PlaceSvc, CauseCode, nrow)
                            ElseIf Sanction = "Y" Then
                                Sactioned_Claim(Percent, Covered, Deduct, DedCode, ClmPaid, blnRegClaim, nrow)
                            Else
                                'All other scenerios ("regular claims")
                                blnRegClaim = True
                                If Percent = "100" Then
                                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                    If Deduct <> "0.00" And Deduct <> "0" Then
                                        If InnDedCross(MMICount) = True Or OONDedCross(MMICount) = True Then
                                            'Individual INN ded to OON ded or
                                            'Individual OON ded to INN ded applies
                                            Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
                                            If DedCode = "M" Then
                                                If DedToOOP(MMICount) = True Then
                                                    If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                                                        tblMHI.Rows(nrow).Cells(42).Value = Deduct
                                                    Else
                                                        tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                                    End If
                                                Else
                                                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                                End If
                                            ElseIf DedCode = "G" Then
                                                If DedToOOP(MMICount) = True Then
                                                    If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                                                        tblMHI.Rows(nrow).Cells(40).Value = Deduct
                                                    Else
                                                        tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                                    End If
                                                Else
                                                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                                End If
                                            End If
                                        Else
                                            Call Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
                                            If DedCode = "M" Then
                                                If DedToOOP(MMICount) = True Then

                                                    If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                                                        tblMHI.Rows(nrow).Cells(42).Value = Deduct
                                                    Else
                                                        tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                                    End If
                                                Else
                                                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                                End If

                                            ElseIf DedCode = "G" Then
                                                If DedToOOP(MMICount) = True Then
                                                    If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                                                        tblMHI.Rows(nrow).Cells(40).Value = Deduct
                                                    Else
                                                        tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                                    End If
                                                Else
                                                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                                End If
                                            End If
                                        End If
                                    End If

                                Else
                                    If INNStatus(MMICount) = True Then
                                        If Deduct <> "0.00" And Deduct <> "0" Then
                                            Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
                                            If DedCode = "M" Then
                                                'INN Deductible
                                                If DedToOOP(MMICount) = True Then
                                                    'Ded applies to OOP
                                                    Oop = Covered - ClmPaid
                                                Else
                                                    Oop = Covered - Deduct - ClmPaid
                                                End If
                                                Oop = Math.Round(Oop, 2)
                                                tblMHI.Rows(nrow).Cells(40).Value = Oop
                                                If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                                                    tblMHI.Rows(nrow).Cells(42).Value = Oop
                                                Else
                                                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                                End If
                                            ElseIf DedCode = "G" Then
                                                If DedToOOP(MMICount) = True Then
                                                    'Ded applies to OOP
                                                    Oop = Covered - ClmPaid
                                                Else
                                                    Oop = Covered - Deduct - ClmPaid
                                                End If
                                                Oop = Math.Round(Oop, 2)
                                                tblMHI.Rows(nrow).Cells(42).Value = Oop
                                                If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                                                    tblMHI.Rows(nrow).Cells(40).Value = Oop
                                                Else
                                                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                                End If
                                            End If
                                        Else
                                            If ClmPaid = "0" Then
                                                If Covered <> "0.01" Then
                                                    'PotPaid = ((Covered * (Percent / 100))   ''CCur(Covered * (Percent / 100))  sanjeet
                                                    Oop = Covered - PotPaid
                                                Else
                                                    Oop = Covered
                                                End If
                                            Else
                                                Oop = Covered - ClmPaid
                                            End If
                                            Oop = Math.Round(Oop, 2)
                                            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                                            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                                            If (Percent > OONPercent(MMICount) Or DedCode = "M") And
                                                                  DedCode <> "G" Then
                                                tblMHI.Rows(nrow).Cells(40).Value = Oop
                                                If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                                                    tblMHI.Rows(nrow).Cells(42).Value = Oop
                                                Else
                                                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                                                End If
                                            Else
                                                tblMHI.Rows(nrow).Cells(42).Value = Oop
                                                If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                                                    tblMHI.Rows(nrow).Cells(40).Value = Oop
                                                Else
                                                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                                                End If
                                            End If
                                        End If
                                    Else

                                    End If
                                End If
                            End If
                        End If
                        'End If
                        'Next nrow
                        nrow = nrow + 1

                    If IsDBNull(tblMHI.Rows(nrow).Cells(38).Value) Then
                        nrow = nrow + 1
                    End If
                    If IsDBNull(tblMHI.Rows(nrow).Cells(38).Value) Then
                        Exit Do
                    End If
                Loop 'Until tblMHI.Rows(nrow).Cells(0).Value = ""
            Next MMICount
        Catch ex As Exception
            'MessageBox.Show("Please check if you have entered manual data. Is there any blank cell you entered?  If Not Please connect with the NAT team", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

        'Format columns AN-AQ to numbers
        'ActiveCell.Offset(0, 39).NumberFormat = "0.00_);[Red](0.00)"  '' need to check 
        'ActiveCell.Offset(0, 40).NumberFormat = "0.00_);[Red](0.00)"
        'ActiveCell.Offset(0, 41).NumberFormat = "0.00_);[Red](0.00)"
        'ActiveCell.Offset(0, 42).NumberFormat = "0.00_);[Red](0.00)"

        'Next
        'Next        

        Dim rRow As Integer

        'For rRow = 0 To tblMHI.Rows.Count - 1
        '    'hello
        '    'tblMHI.Rows(nrow).Cells(39).Value = Int(tblMHI.Rows(nrow).Cells(39).Value)
        '    'tblMHI.Rows(nrow).Cells(40).Value = Int(tblMHI.Rows(nrow).Cells(40).Value)
        '    'tblMHI.Rows(nrow).Cells(41).Value = Int(tblMHI.Rows(nrow).Cells(41).Value)
        '    'tblMHI.Rows(nrow).Cells(42).Value = Int(tblMHI.Rows(nrow).Cells(42).Value)
        '    'Sumit singh
        '    If Not IsDBNull(tblMHI.Rows(nrow).Cells(39).Value) Then
        '        tblMHI.Rows(nrow).Cells(39).Value = Int(tblMHI.Rows(nrow).Cells(39).Value)
        '    End If

        '    If Not IsDBNull(tblMHI.Rows(nrow).Cells(40).Value) Then
        '        tblMHI.Rows(nrow).Cells(40).Value = Int(tblMHI.Rows(nrow).Cells(40).Value)
        '    End If

        '    If Not IsDBNull(tblMHI.Rows(nrow).Cells(41).Value) Then
        '        tblMHI.Rows(nrow).Cells(41).Value = Int(tblMHI.Rows(nrow).Cells(41).Value)
        '    End If

        '    If Not IsDBNull(tblMHI.Rows(nrow).Cells(42).Value) Then
        '        tblMHI.Rows(nrow).Cells(42).Value = Int(tblMHI.Rows(nrow).Cells(42).Value)
        '    End If

        'Next

        Build_Accural_Column_Formulas()

        'Call Sum_Data()

        TabControl1.SelectedIndex = 0

        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        RichTextBox1.SelectionBullet = True
        RichTextBox1.AppendText("OOP Calculation - Completed" & vbCrLf)




    End Sub

    Private Sub InstructionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InstructionsToolStripMenuItem.Click
        'Instructions.Show()
        Dim word_app As Word._Application = New Word.Application
        word_app.Visible = True
        Dim word_doc As Word._Document
        Dim para As Word.Paragraph
        Dim Mydocpath As String
        Mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\" + "Insight Software\Macro Express\Macro Files\NET\OOPCalculator" + "\" + "Instruction - OOP Calculator.docx"
        word_doc = word_app.Documents.Add(Mydocpath)
        word_doc.Activate()
    End Sub

    Private Sub btnOOPExport_Click(sender As Object, e As EventArgs) Handles btnOOPExport.Click

        'Call OOPSpreadSheet_Export()
        Call ExportToExcel(tblOOP)

    End Sub

    Private Sub ExportToExcel(ByVal dgv As DataGridView)
        Dim xlApp As New Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Add()
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(xlWorkBook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        Dim misValue As Object = System.Reflection.Missing.Value

        ' Write headers to Excel
        Dim i As Integer = 0
        Dim j As Integer = 0
        For Each column As DataGridViewColumn In dgv.Columns
            xlWorkSheet.Cells(1, j + 1) = column.HeaderText
            j += 1
        Next

        ' Write data to Excel
        For i = 0 To dgv.Rows.Count - 1
            For j = 0 To dgv.Columns.Count - 1
                xlWorkSheet.Cells(i + 2, j + 1) = dgv.Rows(i).Cells(j).Value
            Next
        Next

        ' Save the Excel file
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel 97-2003 Workbook (*.xls)|*.xls"
        saveFileDialog1.FilterIndex = 1
        saveFileDialog1.FileName = "MyExcelFile"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            xlWorkBook.SaveAs(saveFileDialog1.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        End If

        ' Cleanup
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)

        xlWorkSheet = Nothing
        xlWorkBook = Nothing
        xlApp = Nothing

        GC.Collect()
    End Sub


    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim msg As String = "Are you want to close the Tool?"
        Dim title As String = "Tool clossing"
        Dim result = MessageBox.Show(msg, title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
        If result = DialogResult.Cancel Then
            e.Cancel = True
        End If

    End Sub

    Sub Sum_Data()              ''SUM and color code
        'tblMHI.Rows.Add()
        Dim sum As Integer = 0
        Dim icnt As Integer
        For j As Integer = 39 To 42
            icnt = 0
            For icnt = 0 To tblMHI.Rows.Count() - 2
                Try
                    sum = sum + Int(tblMHI.Rows(icnt).Cells(j).Value)
                Catch ex As Exception

                End Try

            Next
            tblMHI.Rows(icnt).Cells(j).Value = sum
            tblMHI.Rows(icnt).Cells(j).Style.BackColor = Color.Gray
            tblMHI.Rows(icnt).Cells(38).Value = "Total"
            sum = 0.00
        Next



    End Sub
    Sub OI_OIM_Calcs(Percent, Charge, Covered, Deduct, DedCode, ClmPaid, NotCovered, PlaceSvc, CauseCode, SvcCode, InpHosp, Remark, OVRide)
        '    Dim OIoop As Currency
        '    Dim DPaid As Currency
        '    Dim OICovered As Currency
        '    Dim OIPaid As Currency
        '    Dim Draft, ClmOop As Currency
        '    Dim UhgCovered As Currency
        '    Dim UhgPaid As Currency
        '    Dim UhgNotCov As Currency
        '    Dim UhgCharge As Currency
        '    Dim UhgDeduct As Currency
        '    Dim intRowCnt, OICount

        Dim OIoop As Double
        Dim DPaid As Double
        Dim OICovered As Double
        Dim OIPaid As Double
        Dim Draft, ClmOop As Double
        Dim UhgCovered As Double
        Dim UhgPaid As Double
        Dim UhgNotCov As Double
        Dim UhgCharge As Double
        Dim UhgDeduct As Double
        Dim intRowCnt, OICount

        If InStr(SvcCode, "COPAY") > 0 Then

            Copay_Calcs(NotCovered, PlaceSvc, CauseCode, nrow)
        ElseIf InpHosp = True Then
            'Inpatient Facility bill calculates OOP differently for OI/OIM
            ClmOop = 0
            intRowCnt = nrow
            UhgNotCov = 0
            UhgCharge = 0
            UhgDeduct = 0
            Do
                'Draft = ActiveCell.Offset(intRowCnt, 24).Value
                Draft = tblMHI.Rows(intRowCnt).Cells(24).Value
                If InStr(SvcCode, "OI") > 0 Then
                    OIoop = (Charge - ClmPaid)            'CCur(Charge) - CCur(ClmPaid) Need to check sanjeet
                    OICovered = Charge              'CCur(Charge)
                    OIPaid = ClmPaid                    'CCur(ClmPaid)
                Else
                    If OVRide <> "30" And OVRide <> "Y0" Then
                        Covered = tblMHI.Rows(intRowCnt).Cells(12).Value
                        DPaid = tblMHI.Rows(intRowCnt).Cells(18).Value
                        UhgNotCov = UhgNotCov + tblMHI.Rows(intRowCnt).Cells(10).Value  'CCur(ActiveCell.Offset(intRowCnt, 10).Value) need to check
                        UhgCharge = UhgCharge + tblMHI.Rows(intRowCnt).Cells(9).Value   'CCur(ActiveCell.Offset(intRowCnt, 9).Value)
                        UhgCovered = UhgCovered + Covered
                        UhgCovered = Math.Round([UhgCovered], 2)
                        UhgPaid = UhgPaid + DPaid
                        UhgPaid = Math.Round([UhgPaid], 2)
                        DedCode = tblMHI.Rows(intRowCnt).Cells(14).Value
                        Percent = tblMHI.Rows(intRowCnt).Cells(15).Value
                        UhgDeduct = UhgDeduct + tblMHI.Rows(intRowCnt).Cells(13).Value   'CCur(ActiveCell.Offset(intRowCnt, 13).Value)
                        UhgDeduct = Math.Round([UhgDeduct], 2)
                    End If
                End If
                intRowCnt = intRowCnt + 1
                If tblMHI.Rows(intRowCnt).Cells(14).Value <> Draft Then Exit Do
                SvcCode = tblMHI.Rows(intRowCnt).Cells(2).Value
            Loop

            If UhgCharge = UhgNotCov Then
                For OICount = 0 To (intRowCnt - 1)
                    tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                    If OICount = intRowCnt - 1 Then Exit For
                Next
                '            ActiveCell.Offset(intRowCnt - 1, 0).Select
                Exit Sub

            ElseIf OIPaid <= UhgPaid Then
                For OICount = 0 To (intRowCnt - 1)                     ''updating sanjeet
                    tblMHI.Rows(nrow).Cells(49).Value = False
                Next
                tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                tblMHI.Rows(nrow).Cells(42).Value = "0.00"
            Else
                If DedCode = "M" Then
                    tblMHI.Rows(nrow).Cells(39).Value = UhgDeduct
                    If InnDedCross(MMICount) = True Then
                        tblMHI.Rows(nrow).Cells(41).Value = UhgDeduct
                    Else
                        tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                    End If
                    tblMHI.Rows(nrow).Cells(40).Value = OIoop
                    If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                        tblMHI.Rows(nrow).Cells(42).Value = OIoop
                    Else
                        tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                    End If
                ElseIf DedCode = "G" Then
                    tblMHI.Rows(nrow).Cells(41).Value = UhgDeduct
                    If OONDedCross(MMICount) = True Then
                        tblMHI.Rows(nrow).Cells(39).Value = UhgDeduct
                    Else
                        tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                    End If
                    tblMHI.Rows(nrow).Cells(42).Value = OIoop
                    If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                        tblMHI.Rows(nrow).Cells(40).Value = OIoop
                    Else
                        tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                    End If
                End If
                For OICount = 1 To (intRowCnt - 1)
                    tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                    If OICount = intRowCnt - 1 Then Exit For
                Next
                '            ActiveCell.Offset(intRowCnt - 1, 0).Select
            End If

        Else
            If Deduct = "0" Or Deduct = "0.00" Then
                tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                tblMHI.Rows(nrow).Cells(41).Value = "0.00"
            Else
                If DedCode = "M" Then
                    tblMHI.Rows(nrow).Cells(39).Value = Deduct
                    If InnDedCross(MMICount) = True Then
                        tblMHI.Rows(nrow).Cells(41).Value = Deduct
                    Else
                        tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                    End If
                ElseIf DedCode = "G" Then
                    tblMHI.Rows(nrow).Cells(41).Value = Deduct
                    If OONDedCross(MMICount) = True Then
                        tblMHI.Rows(nrow).Cells(39).Value = Deduct
                    Else
                        tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                    End If
                End If
            End If

            'If Percent = "100" Then
            If Percent = "100" And DedToOOP(MMICount) = False Then
                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                tblMHI.Rows(nrow).Cells(42).Value = "0.00"
            Else
                DPaid = tblMHI.Rows(nrow).Cells(18).Value
                If DedToOOP(MMICount) = True Then
                    OIoop = Covered - ClmPaid - DPaid
                Else
                    OIoop = Covered - Deduct - ClmPaid - DPaid
                End If
                If Percent > OONPercent(MMICount) Or DedCode = "M" Then
                    tblMHI.Rows(nrow).Cells(40).Value = OIoop
                    If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then            '''need to check  06/27/2023
                        tblMHI.Rows(nrow).Cells(42).Value = OIoop
                    End If
                ElseIf DedCode = "G7" Or DedCode = "M7" Then
                    tblMHI.Rows(nrow).Cells(39).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(41).Value = "0.00"
                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                Else
                    tblMHI.Rows(nrow).Cells(42).Value = OIoop
                    If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                        tblMHI.Rows(nrow).Cells(40).Value = OIoop
                    End If
                End If

            End If

        End If

        Call Build_Accural_Column_Formulas()

    End Sub

    Private Sub ProcessedDateOnlyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcessedDateOnlyToolStripMenuItem.Click
        Sort_ProcesssedDateOnly()
        cell_colorback()
    End Sub

    Sub Sort_ProcesssedDateOnly()

        ''---------------

        Dim intLastrow As Integer = tblMHI.Rows.Count - 1
        For I = 0 To intLastrow
            If I > intLastrow Then Exit For
            Try
                If IsDBNull(tblMHI.Rows(I).Cells(0).Value) Then
                    tblMHI.Refresh()
                    tblMHI.Rows.RemoveAt(I)
                    tblMHI.Refresh()
                    intLastrow = intLastrow - 1
                    I = I - 1
                End If
            Catch ex As Exception

            End Try
        Next I



        Dim tblCount As Integer = 0

        Dim dttables_SORT As New List(Of System.Data.DataTable)
        dttables_SORT = list(Final_dt_main, Final_dt_main.Rows.Count)
        Dim dt_count As Integer = dttables_SORT.Count
        'Form1.tblMHI.DataSource = dttables_

        Final_dt_main_SORT = New Data.DataTable

        Final_dt_main_SORT = Final_dt_main.Clone
        Dim cnt As Integer
        For Each datatable As System.Data.DataTable In dttables_SORT
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "ProcDate Asc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows


                If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_main_SORT.ImportRow(rw)
                    rw = Final_dt_main_SORT.NewRow()
                    Final_dt_main_SORT.Rows.Add(rw)
                Else
                    Final_dt_main_SORT.ImportRow(rw)
                End If
                cnt += 1
            Next
        Next
        tblMHI.DataSource = Nothing
        tblMHI.DataSource = Final_dt_main_SORT

        Call Build_Accural_Column_Formulas()
    End Sub

    ''Process date and draft
    Private Sub ProcessedDateAndDraftToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcessedDateAndDraftToolStripMenuItem.Click


        Dim intLastrow As Integer = tblMHI.Rows.Count - 1
        For I = 0 To intLastrow
            If I > intLastrow Then Exit For
            Try
                If IsDBNull(tblMHI.Rows(I).Cells(0).Value) Then
                    tblMHI.Refresh()
                    tblMHI.Rows.RemoveAt(I)
                    tblMHI.Refresh()
                    intLastrow = intLastrow - 1
                    I = I - 1
                End If
            Catch ex As Exception

            End Try
        Next I


        Dim SortOrder
        SortOrder = MsgBox("Do you want the process date sorted in Ascending order? Clicking " &
                       "No will sort in Descending order.", vbYesNo + vbDefaultButton2)

        If SortOrder = vbYes Then
            Dim dtview As New DataView(Final_dt_main)
            dtview.Sort = "ProcDate Desc ,Draft Asc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            tblMHI.DataSource = Nothing
            tblMHI.DataSource = dt_dt
        Else
            Dim dtview As New DataView(Final_dt_main)
            dtview.Sort = "ProcDate Desc,Draft Desc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            tblMHI.DataSource = Nothing
            tblMHI.DataSource = dt_dt
        End If
        Call Build_Accural_Column_Formulas()
        cell_colorback()
    End Sub

    Private Sub ProviderTinToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProviderTinToolStripMenuItem.Click

        Dim tblCount As Integer = 0

        Dim dttables_SORT As New List(Of System.Data.DataTable)
        dttables_SORT = list(Final_dt_main, Final_dt_main.Rows.Count)
        Dim dt_count As Integer = dttables_SORT.Count
        'Form1.tblMHI.DataSource = dttables_


        Final_dt_main_SORT = New Data.DataTable

        Final_dt_main_SORT = Final_dt_main.Clone
        Dim cnt As Integer
        For Each datatable As System.Data.DataTable In dttables_SORT
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "PT_Name Asc,TIN Desc,From Asc"
            Dim dt_dt As System.Data.DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows


                If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_main_SORT.ImportRow(rw)
                    rw = Final_dt_main_SORT.NewRow()
                    Final_dt_main_SORT.Rows.Add(rw)
                Else
                    Final_dt_main_SORT.ImportRow(rw)
                End If
                cnt += 1
            Next
        Next
        tblMHI.DataSource = Nothing
        tblMHI.DataSource = Final_dt_main_SORT

        'Dim dtview As New DataView(Final_dt_main)
        'dtview.Sort = "PT_Name Asc,TIN Desc,From Asc"
        'Dim dt_dt As System.Data.DataTable = dtview.ToTable()
        'tblMHI.DataSource = Nothing
        'tblMHI.DataSource = dt_dt
        Call Build_Accural_Column_Formulas()
        cell_colorback()
    End Sub

    Private Sub DateOfServiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DateOfServiceToolStripMenuItem.Click
        tblMHI.Sort(tblMHI.Columns(0), ListSortDirection.Descending)
        Call Build_Accural_Column_Formulas()
        cell_colorback()
    End Sub

    Private Sub PaitentNameAndDOSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PaitentNameAndDOSToolStripMenuItem.Click
        Call MbrSort()
        cell_colorback()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Export_MMI_OVERVIEW()
    End Sub

    Private Sub ClearAllFilterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllFilterToolStripMenuItem.Click
        tblMHI.CleanFilter()
    End Sub

    Private Sub tblMHI_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles tblMHI.CellClick
        index = e.RowIndex
        Dim nrow As Integer = 0

        'Try
        '    If Not isNullOrEmpty(sArray) Then
        '        For q = 0 To UBound(sArray)
        '            If isNullOrEmpty(sArray(q)) Then Continue For
        '            For t = 0 To tblMHI.Rows.Count - 1
        '                If isNullOrEmpty(tblMHI.Rows(t).Cells(57).Value) Then Exit For
        '                If tblMHI.Rows(t).Cells(57).Value > sArray(q) Then Exit For



        '                If tblMHI.Rows(t).Cells(57).Value = sArray(q) Then 'And tblMHI.Rows(t).Cells(57).Value > sArray(q) Then
        '                    For I = 0 To 56
        '                        tblMHI.Rows(t).Cells(I).Style.BackColor = Color.Yellow
        '                    Next
        '                    Exit For
        '                End If
        '            Next
        '        Next
        '    End If
        'Catch ex As Exception
        '    'MsgBox("got it")
        'End Try



    End Sub

    Private Sub ApplyColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ApplyColorToolStripMenuItem.Click
        Dim color_check As Boolean

        Try
            If tblMHI.Rows(index).Cells(0).Style.BackColor.Name = "Yellow" Then
                For I = 0 To 56
                    tblMHI.Rows(index).Cells(I).Style.BackColor = Color.White
                Next



                For i As Integer = 0 To UBound(sArray)
                    If isNullOrEmpty(sArray(i)) Then Continue For
                    If sArray(i) = tblMHI.Rows(index).Cells(57).Value Then
                        sArray(i) = ""
                    End If
                Next



                color_check = True
            End If



            If color_check = False Then
                For I = 0 To 56
                    tblMHI.Rows(index).Cells(I).Style.BackColor = Color.Yellow
                Next

                ReDim Preserve sArray(row_yellow)
                sArray(row_yellow) = tblMHI.Rows(index).Cells(57).Value
                row_yellow += 1
            End If
        Catch ex As Exception



        End Try



    End Sub

    Private Sub tblMHI_SelectionChanged(sender As Object, e As EventArgs) Handles tblMHI.SelectionChanged
        'Try
        '    If Not isNullOrEmpty(sArray) Then
        '        For q = 0 To UBound(sArray)
        '            If isNullOrEmpty(sArray(q)) Then Continue For
        '            For t = 0 To tblMHI.Rows.Count - 2
        '                If tblMHI.Rows(t).Cells(57).Value = sArray(q) Then
        '                    For I = 0 To 56
        '                        tblMHI.Rows(sArray(q)).Cells(I).Style.BackColor = Color.Yellow
        '                    Next
        '                End If
        '            Next
        '            'tblMHI.Rows(sArray(q)).Cells(0).Style.BackColor = Color.Yellow
        '        Next
        '    End If



        'Catch ex As Exception



        'End Try
    End Sub

    Private Sub tblMHI_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles tblMHI.CellFormatting
        If rowColors.ContainsKey(e.RowIndex) Then
            e.CellStyle.BackColor = rowColors(e.RowIndex)
        End If
        'MsgBox("cell formating")
        Try
            If Not isNullOrEmpty(sArray) Then
                For q = 0 To UBound(sArray)
                    If isNullOrEmpty(sArray(q)) Then Continue For
                    For t = 0 To tblMHI.Rows.Count - 1
                        If isNullOrEmpty(tblMHI.Rows(t).Cells(57).Value) Then Exit For
                        If tblMHI.Rows(t).Cells(57).Value > sArray(q) Then Exit For





                        If tblMHI.Rows(t).Cells(57).Value = sArray(q) Then 'And tblMHI.Rows(t).Cells(57).Value > sArray(q) Then
                            For I = 0 To 56
                                tblMHI.Rows(t).Cells(I).Style.BackColor = Color.Yellow
                            Next
                            Exit For
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            'MsgBox("got it")
        End Try
    End Sub

    Sub cell_colorback()



        Try
            If Not isNullOrEmpty(sArray) Then
                For q = 0 To UBound(sArray)
                    If isNullOrEmpty(sArray(q)) Then Continue For
                    For t = 0 To tblMHI.Rows.Count - 1
                        If isNullOrEmpty(tblMHI.Rows(t).Cells(57).Value) Then Exit For
                        'If tblMHI.Rows(t).Cells(57).Value > sArray(q) Then Exit For





                        If tblMHI.Rows(t).Cells(57).Value = sArray(q) Then 'And tblMHI.Rows(t).Cells(57).Value > sArray(q) Then
                            For I = 0 To 56
                                tblMHI.Rows(t).Cells(I).Style.BackColor = Color.Yellow
                            Next
                            Exit For
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            'MsgBox("got it")
        End Try
    End Sub

    Sub Build_Accural_Column_Formulas()

        'Try

        Dim BoolNxtMemberRow As Boolean
        BoolNxtMemberRow = False

        Dim i As Integer
        Dim int39Value As Double
        Dim int40Value As Double
        Dim int41Value As Double
        Dim int42Value As Double

        Dim int43Value As Double
        Dim int44Value As Double
        Dim int45Value As Double
        Dim int46Value As Double

        Dim int43Sum As Double
        Dim int44Sum As Double
        Dim int45Sum As Double
        Dim int46Sum As Double

        For i = 0 To tblMHI.Rows.Count - 2                  ''added as per date 04/07/2023 before it was '-2'

            If i = 0 Or BoolNxtMemberRow = True Then
                Try
                    BoolNxtMemberRow = False
                    int39Value = tblMHI.Rows(i).Cells(39).Value : int39Value = Math.Round([int39Value], 2)
                    int40Value = tblMHI.Rows(i).Cells(40).Value : int40Value = Math.Round([int40Value], 2)
                    int41Value = tblMHI.Rows(i).Cells(41).Value : int41Value = Math.Round([int41Value], 2)
                    int42Value = tblMHI.Rows(i).Cells(42).Value : int42Value = Math.Round([int42Value], 2)

                Catch ex As Exception

                End Try

                tblMHI.Rows(i).Cells(43).Value = int39Value
                tblMHI.Rows(i).Cells(44).Value = int40Value
                tblMHI.Rows(i).Cells(45).Value = int41Value
                tblMHI.Rows(i).Cells(46).Value = int42Value
            Else
                Try
                    int43Value = tblMHI.Rows(i - 1).Cells(43).Value : int43Value = Math.Round([int43Value], 2)
                    int44Value = tblMHI.Rows(i - 1).Cells(44).Value : int44Value = Math.Round([int44Value], 2)
                    int45Value = tblMHI.Rows(i - 1).Cells(45).Value : int45Value = Math.Round([int45Value], 2)
                    int46Value = tblMHI.Rows(i - 1).Cells(46).Value : int46Value = Math.Round([int46Value], 2)

                    int39Value = tblMHI.Rows(i).Cells(39).Value : int39Value = Math.Round([int39Value], 2)
                    int40Value = tblMHI.Rows(i).Cells(40).Value : int40Value = Math.Round([int40Value], 2)
                    int41Value = tblMHI.Rows(i).Cells(41).Value : int41Value = Math.Round([int41Value], 2)
                    int42Value = tblMHI.Rows(i).Cells(42).Value : int42Value = Math.Round([int42Value], 2)
                Catch ex As Exception

                End Try
                int43Sum = int43Value + int39Value
                int44Sum = int44Value + int40Value
                int45Sum = int45Value + int41Value
                int46Sum = int46Value + int42Value

                int43Sum = Math.Round([int43Sum], 2)
                int44Sum = Math.Round([int44Sum], 2)
                int45Sum = Math.Round([int45Sum], 2)
                int46Sum = Math.Round([int46Sum], 2)


                'int43Value = Int(tblMHI.Rows(i - 1).Cells(43).Value) + Int(tblMHI.Rows(i).Cells(39).Value)
                'int44Value = Int(tblMHI.Rows(i - 1).Cells(44).Value) + Int(tblMHI.Rows(i).Cells(40).Value)
                'int45Value = Int(tblMHI.Rows(i - 1).Cells(45).Value) + Int(tblMHI.Rows(i).Cells(41).Value)
                'int46Value = Int(tblMHI.Rows(i - 1).Cells(46).Value) + Int(tblMHI.Rows(i).Cells(42).Value)

                tblMHI.Rows(i).Cells(43).Value = int43Sum
                tblMHI.Rows(i).Cells(44).Value = int44Sum
                tblMHI.Rows(i).Cells(45).Value = int45Sum
                tblMHI.Rows(i).Cells(46).Value = int46Sum
            End If

            If IsDBNull(tblMHI.Rows(i + 1).Cells(38).Value) Then
                i = i + 1
                BoolNxtMemberRow = True
            End If
            Try
                If IsDBNull(tblMHI.Rows(i + 1).Cells(38).Value) Then
                    Exit Sub
                End If
            Catch ex As Exception

            End Try


        Next i

        'Catch ex As Exception

        'End Try

    End Sub
    Sub Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
        Dim InNetwork As Boolean
        Dim DedCodeSkips
        DedCodeSkips = "1,2,3,4,5,6,7"

        If DedCode = "M" Then
            'INN Deductible
            tblMHI.Rows(nrow).Cells(39).Value = Deduct
            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
            InNetwork = True
            If InnDedCross(MMICount) = True Then
                tblMHI.Rows(nrow).Cells(41).Value = Deduct
            End If
            If blnRegClaim = True Then
                If DedToOOP(MMICount) = True Then
                    tblMHI.Rows(nrow).Cells(40).Value = Deduct
                End If
            End If
        ElseIf DedCode = "G" Then
            'OON Deductible
            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
            tblMHI.Rows(nrow).Cells(41).Value = Deduct
            If OONDedCross(MMICount) = True Or OONFamDedCross(MMICount) = True Then
                tblMHI.Rows(nrow).Cells(39).Value = Deduct
            End If
            If blnRegClaim = True Then
                If DedToOOP(MMICount) = True Then
                    tblMHI.Rows(nrow).Cells(42).Value = Deduct
                End If
            End If
        ElseIf DedCode = "G7" Or DedCode = "M7" Then
            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
        ElseIf InStr(1, DedCodeSkips, DedCode) > 0 Then
            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
        End If
        ''need to check sanjeet
    End Sub
    Sub Copay_Calcs(NotCovered, PlaceSvc, CauseCode, nrow)
        If DxCauseCode(MMICount) = "" Or DxCauseCode(MMICount) = "N/A" Then
            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
            tblMHI.Rows(nrow).Cells(40).Value = "0.00"
            tblMHI.Rows(nrow).Cells(41).Value = "0.00"
            tblMHI.Rows(nrow).Cells(42).Value = "0.00"
        Else
            tblMHI.Rows(nrow).Cells(39).Value = "0.00"
            tblMHI.Rows(nrow).Cells(41).Value = "0.00"


            If DxCauseCode(MMICount) = "All" Or CopayPOS(MMICount) = "All" Or
                (InStr(DxCauseCode(MMICount), CauseCode) > 0 And InStr(CopayPOS(MMICount), PlaceSvc) > 0) Then
                tblMHI.Rows(nrow).Cells(40).Value = NotCovered
                If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then                ''NEED TO CHECK SANJEET
                    tblMHI.Rows(nrow).Cells(42).Value = NotCovered
                Else
                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                End If
            Else
                tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                tblMHI.Rows(nrow).Cells(42).Value = "0.00"
            End If
        End If
    End Sub
    Sub Sactioned_Claim(Percent, Covered, Deduct, DedCode, ClmPaid, blnRegClaim, nrow)
        'Dim Oop As Currency
        Dim Oop As Double

        If Deduct <> "0.00" And Deduct <> "0" Then
            Deduct_Calcs(DedCode, Deduct, blnRegClaim, nrow)
        End If

        If DedCode = "G7" Or DedCode = "M7" Then
            tblMHI.Rows(nrow).Cells(40).Value = "0.00"
            tblMHI.Rows(nrow).Cells(42).Value = "0.00"
        Else
            If Deduct < Covered And Percent <> "100" Then
                If DedToOOP(MMICount) = True Then
                    Oop = (Covered - Deduct) * ((100 - Percent) / 100) + Deduct
                Else
                    Oop = (Covered - Deduct) * ((100 - Percent) / 100)
                End If
            Else
                If DedToOOP(MMICount) = True Then
                    Oop = Deduct
                Else
                    Oop = "0.00"
                End If
            End If
            Oop = Math.Round(Oop, 2)
            If Percent > OONPercent(MMICount) Or DedCode = "M" Then
                tblMHI.Rows(nrow).Cells(40).Value = Oop
                If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "3" Then
                    tblMHI.Rows(nrow).Cells(42).Value = Oop
                Else
                    tblMHI.Rows(nrow).Cells(42).Value = "0.00"
                End If
            Else
                tblMHI.Rows(nrow).Cells(42).Value = Oop
                If OOPCross(MMICount) = "1" Or OOPCross(MMICount) = "2" Then
                    tblMHI.Rows(nrow).Cells(40).Value = Oop
                Else
                    tblMHI.Rows(nrow).Cells(40).Value = "0.00"
                End If
            End If
        End If

    End Sub
    Private Sub GatherProvInfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GatherProvInfoToolStripMenuItem.Click

        GetProvInfo()
    End Sub

    Private Sub GetMMIToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetMMIToolStripMenuItem.Click
        'The following code will take the account number, policy number, and year and do the following:
        '1: It will gather the MRI data which will include the plans under that policy
        '2: It will determine which MRI plans fall into the given year and gather MXI data for those plans
        '3: It will determine which MMI plans fall into the given year under the previously given MXI plans
        '4: It will store all applicable MMI plans into a separate list so they can be accessed outside the loop
        '5: It will determine the proper start-end dates based on information from MMI
        '6: It will gather all possible MMI pages given the final time frame
        '7: It will find the final date range given any possible MMI pages
        '8: It will use that date range to add all members in the policy to a list and check them off based on
        '   their current effective status


        'Assigns a temporary date of the end of the year of the given year
        'This can be used to automatically determine the year ranges by automatically gathering
        'mmi data for that year
        blnCopays = False
        MMIstarttime = Now()

        RichTextBox1.AppendText("Gathering data from MMI screen..." & vbCrLf)
        RichTextBox1.SelectionBullet = True
        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4


        Dim useDate As String = "12/31/" & yearList.Text   '''commented by Sanjeet

        '       Dim useDate As String = "1/1/" & yearList.Text


        If IsDate(useDate) = False Then
            Exit Sub
        Else
            useDate = Format(CDate(useDate), "MM/dd/yyyy")
        End If

        Dim tempYearStart As String = ""
        Dim tempYearEnd As String = ""

        Dim stSdt = Format(CDate(startSelect.Text), "MM/dd/yyyy")
        Dim stEdt = Format(CDate(endSelect.Text), "MM/dd/yyyy")

        If stEdt = Format(Now, "MM/dd/yyyy") Then
            stEdt = "12/31/" & yearList.Text
        End If


        Dim yearStart As String = "12/31/" & yearList.Text
        Dim yearEnd As String = "01/01/1901"

        memberList.SelectedIndex = -1

        Dim colcnt As Int32 = 0
        Dim rcount As Integer = 1
        Dim dtMonth
        Dim dtYear
        Dim lastrow As Integer = tblMHI.Rows.Count - 2


        dtMonth = Format(CDate(tblMHI.Rows(lastrow).Cells(0).Value), "MM").ToString()
        dtYear = Format(CDate(tblMHI.Rows(lastrow).Cells(0).Value), "yyyy").ToString()
        dtMonth = dtMonth & "/" & "01" & "/" & dtYear




        'Dim mriResp = api_MRI.PerformQuery(Trim(txt_Policy.Text), Trim(txt_SSN.Text), Trim(txtUserID.Text), Trim(txtUserPass.Text))
        Dim CEIPalnCode As String
        Dim CEIClassCode As String
        Dim MMICheck As Boolean
        Dim MMICheck_Second As Boolean
        Dim MMICnt As Integer = 0
        Dim X As Integer = 0
        Dim dtcheck As Boolean

        For memberNumber = 0 To memberList.CheckedItems.Count - 1
            dtcheck = False
            MMICnt = 0
            'X = 0

            '   Dim mriResp = api_MRI.PerformQuery(Trim(txt_Policy.Text), Trim(txt_SSN.Text), Trim(txtUserID.Text), Trim(txtUserPass.Text))
            '   For Each mriPolicy In mriResp.Results.Response.rspMRIinfo.rspMRICoverageLine


            'Dim mriResp = api_MRI.PerformQuery(Trim(txt_Policy.Text), Trim(txt_SSN.Text), Trim(txtUserID.Text), Trim(txtUserPass.Text))

            'For Each mriPolicy In mriResp.Results.Response.rspMRIinfo.rspMRICoverageLine
            '    Dim mriStart As String = policyRange(mriPolicy.rspCovEffDT)                        'Coverage Line: Start
            '    Dim mriEnd As String = policyRange(mriPolicy.rspCovCanDT, mriPolicy.rspCovEffDT)   'Coverage Line: End
            'Next


            Dim splitMbr = Split(memberList.CheckedItems(memberNumber), "/")

            CEIPalnCode = Get_CEIPlan(splitMbr(0), intRow)

            CEIClassCode = Get_CEIClass(splitMbr(0))
            Dim pvrcMember As Boolean

            'For nxt = xt To UBound(Final_Date)
            '    ptN = Split(Final_Date(nxt), ",")
            '    If ptN(1) = memberList.CheckedItems(memberNumber) Then
            '        xt += 1
            '        pvrcMember = True
            '        Exit For
            '    End If

            'Next

            'If pvrcMember = False Then Continue For

            'Checks if coverage info has a valid date. 
            'If useDate isn't within that range it will continue to the next line

            'If ((IsDate(mriStart) = False) Or (IsDate(mriEnd) = False)) OrElse
            '((CDate(useDate) < CDate(mriStart)) Or (CDate(useDate) > CDate(mriEnd))) Then Continue For

            'This is the object that will be filled using the apiMXI module (it isn't called yet because
            'we need to check if the policy is railroad [railroad uses a larger MXI control line including rpt])            

            Dim mxiResp = Nothing
                If checkRailroad(Trim(txt_Policy.Text)) = True Then
                    mxiResp = api_MXI.PerformQuery(Trim(txt_Policy.Text), CEIPalnCode, CEIClassCode).Results
                Else
                    mxiResp = api_MXI.PerformQuery(Trim(txt_Policy.Text), CEIPalnCode).Results
                End If

                'This will loop through all the coverage screens found in MXI for that particular MRI policy/plan
                'UNET screen equivelant: mxi,policy,plan

                For Each mxiPolicy In mxiResp.mxirows

                    Dim mxiStart As String = policyRange(mxiPolicy.effDt)                   'Coverage Line: Start   sanjeet
                    Dim mxiEnd As String = policyRange(mxiPolicy.cancDt, mxiPolicy.effDt)   'Coverage Line: End
                    MMICnt = MMICnt + 1


                If MMICnt > 12 Then Exit For
                pvrcMember = False
                'Checks if coverage info has a valid date. 
                'If useDate isn't within that range it will continue to the next line

                'If ((IsDate(mxiStart) = False) Or (IsDate(mxiEnd) = False)) OrElse
                '((CDate(useDate) < CDate(mxiStart)) Or (CDate(useDate) > CDate(mxiEnd))) Then Continue For

                'If ((IsDate(mxiStart) = False) Or (IsDate(mxiEnd) = False)) OrElse
                '((CDate(useDate) < CDate(mxiStart)) Or (CDate(useDate) > CDate(mxiEnd))) Then Continue For



                'If CDate(Final_Date(X - 1)) > CDate(mxiStart) And CDate(Final_Date(X - 1)) <= CDate(mxiEnd) Then

                'If CDate(tblMHI.Rows(0).Cells(0).Value) >= CDate(mxiStart) Or dtMonth >= CDate(mxiStart) Then                    ''Added on 05/01/2023 in orderto fix the MMI Extra plans issue 
                'If CDate(stSdt) <= CDate(mxiStart) Or (CDate(mxiEnd) <= CDate(stEdt)) Then

                'Format(CDate(tblMHI.Rows(lastrow).Cells(0).Value), "yyyy").ToString()

                'MsgBox(Final_Date(X))

                Try
                    If X > 0 Then
                        oldPName = Split(Final_Date(X - 1), ",")
                        If InStr(Final_Date(X), oldPName(1)) > 0 Then
                            If memberNumber <> 0 Then
                                memberNumber = memberNumber - 1
                            End If

                        End If
                    End If
                Catch ex As Exception

                End Try


                Try
                        Format(CDate(tblMHI.Rows(lastrow).Cells(0).Value), "yyyy").ToString()

                        If Format(CDate(stEdt), "yyyy") = Format(CDate(mxiEnd), "yyyy") And dtcheck = False Then

                            If CDate(stEdt) > CDate(Mid(Final_Date(X), 1, 10)) Then stEdt = Mid(Final_Date(X), 1, 10)
                            X += 1
                            dtcheck = True
                        End If
                    Catch ex As Exception

                    End Try


                    If CDate(mxiStart) <= CDate(stSdt) And (CDate(mxiEnd) >= CDate(stEdt)) Then

                        MMICheck = True
                        'stEdt
                        '                    Dim mmiFlag As Boolean = False
                        If mmiDetails Is Nothing Then   'Checks if mmiDetails list has been populated/initiated
                            mmiDetails = New List(Of apiMMI.MmiReturn)
                        Else
                            'Checks if this MMI page has been added to the list previously (don't need to add duplicates)
                            If mmiDetails.Count > 0 Then
                                For Each mmiPolicy In mmiDetails
                                    If (mmiPolicy.mmi1Return.mmi1ARows(0).polNbr = mxiPolicy.stdPlnPolNbr) And
                                    (mmiPolicy.mmi1Return.mmi1ARows(0).plnNbr = mxiPolicy.stdPlnPlnNbr) Then
                                        mmiFlag = True
                                        Exit For
                                    End If
                                Next mmiPolicy
                            Else
                                mmiDetails = New List(Of apiMMI.MmiReturn)
                            End If
                        End If

                        'If this set of mmi pages has already been added to the list previously it will skip this page
                        If mmiFlag = True Then Continue For

                        'Adds the mmiAPI response (the set of MMI Pages) object to a list to be used later
                        If checkRailroad(Trim(txt_Policy.Text)) = True Then
                            mmiDetails.Add(api_MMI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr, mxiPolicy.stdPlnClssNbr, "").Results)
                            strPlnnbr = Trim(CEIPalnCode) & "/" & mxiPolicy.stdPlnPlnNbr

                            colcnt = colcnt + 1
                            If mxiStart < stSdt Then mxiStart = stSdt
                            Call mmiItems(colcnt, strPlnnbr, mxiStart, mxiEnd)

                            DGridOverview.Rows(2).Cells(0).Value = "Patient Name"
                            DGridOverview.Rows(2).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            DGrid_PG1.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            DGrid_PG4.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            DGrid_PG5.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            DGrid_PG10.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            strPTName = memberList.CheckedItems(memberNumber)
                            ''Consuming MSI Api
                            Dim plnCode As String = CEIPalnCode

                        Else
                            mmiDetails.Add(api_MMI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr, "", "").Results)
                            ''''added by sanjeet
                            strPlnnbr = Trim(CEIPalnCode) & "/" & mxiPolicy.stdPlnPlnNbr

                            colcnt = colcnt + 1
                            If CDate(mxiStart) < CDate(stSdt) Then mxiStart = stSdt
                            Call mmiItems(colcnt, strPlnnbr, mxiStart, mxiEnd)

                            DGridOverview.Rows(2).Cells(0).Value = "Patient Name"
                            DGridOverview.Rows(2).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            DGrid_PG1.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            DGrid_PG4.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            DGrid_PG5.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            DGrid_PG10.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                            strPTName = memberList.CheckedItems(memberNumber)
                            ''Consuming MSI Api
                            Dim plnCode As String = CEIPalnCode
                            'Dim msiResp = api_MSI.PerformQuery(Trim(txt_Policy.Text), mxiPolicy.stdPlnPlnNbr, "90000").Results
                            If checkRailroad(Trim(txt_Policy.Text)) = True Then
                                msiResp = api_MSI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr, mxiPolicy.stdPlnClssNbr)
                            Else
                                msiResp = api_MSI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr)
                            End If

                            For Each msiRow In msiResp.Results.MsiRows
                                Dim msiStart As String = policyRange(msiRow.effDt)                  'Coverage Line: Start
                                Dim msiEnd As String = policyRange(msiRow.cancDt, msiRow.effDt)     'Coverage Line: End 
                                'Not sure if the calculator needs to compare a date range so I pulled the properties and logic for it
                                'Double check the end date - all the date ranges I've found seem to end 9999 so they'll all show as effective
                                ' - you might need to add in new logic for this
                                '''sanjeet
                                If (IsDate(msiStart) = False) OrElse
                                        (compareRange(yearStart, yearEnd, msiStart, msiEnd) = False) Then Continue For

                                If ((UCase(msiRow.srvcCd) = "EMERG") Or (UCase(msiRow.srvcCd) = "90000")) And ((UCase(msiRow.srvcSetCd) = "N") Or (UCase(msiRow.srvcSetCd) = "P") Or (UCase(msiRow.srvcSetCd) = "T")) Then
                                    'MsgBox(msiStart & " - " & msiEnd & vbNewLine &
                                    'msiRow.srvcCd & " / " & msiRow.srvcSetCd)
                                    If msiRow.causCd <> "[ ]" And msiRow.causCd <> "[]" Then
                                        strcause = msiRow.causCd
                                        PlaceSvc = msiRow.posCd
                                        MCalcNum = Mid(msiRow.mmCalcCd, 3, Len(msiRow.mmCalcCd))

                                        spccode = ""
                                        spccode1 = ""
                                        dtSt_ToDate = mxiStart & "-" & mxiEnd
                                        strCopaySet = msiRow.srvcSetCd

                                        'applying API MCI
                                        If checkRailroad(Trim(txt_Policy.Text)) = True Then         ''Need to check Sanjeet
                                            mciResp = api_MCI.PerformQuery(MCalcNum)
                                        Else
                                            mciResp = api_MCI.PerformQuery("121147")
                                        End If

                                        For Each mciRow In mciResp.Results.MciRows
                                            spccode = mciRow.specialProcCode
                                            Exit For
                                        Next

                                        If (spccode = "C" Or spccode = "D") And blnCopays = False Then          ''05/08/2023
                                            blnCopays = True
                                            DGridOverview.Rows(32).Cells(1).Value = blnCopays
                                        End If


                                        'COSMOS_Window_Selection(MCalcNum, "MCI")
                                        '''------------------------
                                        'spccode = GetText(COSMOS1, 7, 72, 1)

                                        Call CopayTablist(txt_Policy.Text, mxiPolicy.stdPlnPlnNbr, dtSt_ToDate, strPTName, strcause, PlaceSvc, MCalcNum, spccode, strCopaySet, spccode1)

                                    End If

                                End If
                            Next msiRow
                        End If
                    End If
                    '  End If
                Next mxiPolicy

                'Next mriPolicy

                '''''''''''''''''''''''''''if Tool did not found the MMI             

                Dim mmicntNxt As Integer = 0

                If MMICheck = False Then


                    For Each mxiPolicy In mxiResp.mxirows

                        Dim mxiStart As String = policyRange(mxiPolicy.effDt)                   'Coverage Line: Start   sanjeet
                        Dim mxiEnd As String = policyRange(mxiPolicy.cancDt, mxiPolicy.effDt)   'Coverage Line: End

                        'Checks if coverage info has a valid date. 
                        'If useDate isn't within that range it will continue to the next line
                        'If ((IsDate(mxiStart) = False) Or (IsDate(mxiEnd) = False)) OrElse
                        '((CDate(useDate) < CDate(mxiStart)) Or (CDate(useDate) > CDate(mxiEnd))) Then Continue For

                        ' If (CDate(mxiStart) < CDate(stSdt)) Or (CDate(stEdt) > CDate(mxiEnd)) Then Continue For
                        If (CDate(stSdt) <= CDate(mxiStart) Or CDate(mxiStart) <= CDate(stSdt)) And Strings.Right(CDate(mxiEnd), 2) = Strings.Right(CDate(stEdt), 2) Then
                            mmicntNxt = mmicntNxt + 1
                            MMICheck_Second = True
                            If mmicntNxt > 1 Then Exit For
                            'MMICheck = True
                            Dim mmiFlag As Boolean = False
                            If mmiDetails Is Nothing Then   'Checks if mmiDetails list has been populated/initiated
                                mmiDetails = New List(Of apiMMI.MmiReturn)
                            Else
                                'Checks if this MMI page has been added to the list previously (don't need to add duplicates)
                                If mmiDetails.Count > 0 Then
                                    For Each mmiPolicy In mmiDetails
                                        If (mmiPolicy.mmi1Return.mmi1ARows(0).polNbr = mxiPolicy.stdPlnPolNbr) And
                                            (mmiPolicy.mmi1Return.mmi1ARows(0).plnNbr = mxiPolicy.stdPlnPlnNbr) Then
                                            mmiFlag = True
                                            Exit For
                                        End If
                                    Next mmiPolicy
                                Else
                                    mmiDetails = New List(Of apiMMI.MmiReturn)
                                End If
                            End If

                            'If this set of mmi pages has already been added to the list previously it will skip this page
                            If mmiFlag = True Then Continue For

                            'Adds the mmiAPI response (the set of MMI Pages) object to a list to be used later
                            If checkRailroad(Trim(txt_Policy.Text)) = True Then
                                mmiDetails.Add(api_MMI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr, mxiPolicy.stdPlnClssNbr, "").Results)
                                strPlnnbr = Trim(CEIPalnCode) & "/" & mxiPolicy.stdPlnPlnNbr

                                colcnt = colcnt + 1
                                Call mmiItems(colcnt, strPlnnbr, mxiStart, mxiEnd)

                                DGridOverview.Rows(2).Cells(0).Value = "Patient Name"
                                DGridOverview.Rows(2).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                DGrid_PG1.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                DGrid_PG4.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                DGrid_PG5.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                DGrid_PG10.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                strPTName = memberList.CheckedItems(memberNumber)
                                ''Consuming MSI Api
                                Dim plnCode As String = CEIPalnCode
                            Else
                                mmiDetails.Add(api_MMI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr, "", "").Results)
                                ''''added by sanjeet
                                'strPlnnbr = Trim(mriPolicy.rspCovPlan) & "/" & mxiPolicy.stdPlnPlnNbr   ''NEED TO UNCOMMENT                                                        
                                strPlnnbr = CEIPalnCode & "/" & mxiPolicy.stdPlnPlnNbr   ''NEED TO UNCOMMENT

                                colcnt = colcnt + 1
                                Call mmiItems(colcnt, strPlnnbr, mxiStart, mxiEnd)

                                DGridOverview.Rows(2).Cells(0).Value = "Patient Name"
                                DGridOverview.Rows(2).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                DGrid_PG1.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                DGrid_PG4.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                DGrid_PG5.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                DGrid_PG10.Rows(3).Cells(colcnt).Value = memberList.CheckedItems(memberNumber)
                                strPTName = memberList.CheckedItems(memberNumber)
                                MMICheck = False
                                ''Consuming MSI Api
                                Dim plnCode As String = CEIPalnCode
                                'Dim msiResp = api_MSI.PerformQuery(Trim(txt_Policy.Text), mxiPolicy.stdPlnPlnNbr, "90000").Results
                                If checkRailroad(Trim(txt_Policy.Text)) = True Then
                                    msiResp = api_MSI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr, mxiPolicy.stdPlnClssNbr)
                                Else
                                    msiResp = api_MSI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr)
                                End If

                                For Each msiRow In msiResp.Results.MsiRows
                                    Dim msiStart As String = policyRange(msiRow.effDt)                  'Coverage Line: Start
                                    Dim msiEnd As String = policyRange(msiRow.cancDt, msiRow.effDt)     'Coverage Line: End 
                                    'Not sure if the calculator needs to compare a date range so I pulled the properties and logic for it
                                    'Double check the end date - all the date ranges I've found seem to end 9999 so they'll all show as effective
                                    ' - you might need to add in new logic for this
                                    '''sanjeet
                                    If (IsDate(msiStart) = False) OrElse
                                            (compareRange(yearStart, yearEnd, msiStart, msiEnd) = False) Then Continue For

                                    If ((UCase(msiRow.srvcCd) = "EMERG") Or (UCase(msiRow.srvcCd) = "90000")) And ((UCase(msiRow.srvcSetCd) = "N") Or (UCase(msiRow.srvcSetCd) = "P") Or (UCase(msiRow.srvcSetCd) = "T")) Then
                                        'MsgBox(msiStart & " - " & msiEnd & vbNewLine &
                                        'msiRow.srvcCd & " / " & msiRow.srvcSetCd)
                                        If msiRow.causCd <> "[ ]" And msiRow.causCd <> "[]" Then
                                            strcause = msiRow.causCd
                                            PlaceSvc = msiRow.posCd
                                            MCalcNum = Mid(msiRow.mmCalcCd, 3, Len(msiRow.mmCalcCd))

                                            spccode = ""
                                            spccode1 = ""
                                            dtSt_ToDate = mxiStart & "-" & mxiEnd
                                            strCopaySet = msiRow.srvcSetCd

                                            'applying API MCI
                                            If checkRailroad(Trim(txt_Policy.Text)) = True Then         ''Need to check Sanjeet
                                                mciResp = api_MCI.PerformQuery(MCalcNum)
                                            Else
                                                mciResp = api_MCI.PerformQuery("121147")
                                            End If

                                            For Each mciRow In mciResp.Results.MciRows
                                                spccode = mciRow.specialProcCode
                                                Exit For
                                            Next

                                            If (spccode = "C" Or spccode = "D") And blnCopays = False Then          ''05/08/2023
                                                blnCopays = True
                                                DGridOverview.Rows(32).Cells(1).Value = blnCopays
                                            End If

                                            'COSMOS_Window_Selection(MCalcNum, "MCI")
                                            '''------------------------
                                            'spccode = GetText(COSMOS1, 7, 72, 1)

                                            Call CopayTablist(txt_Policy.Text, mxiPolicy.stdPlnPlnNbr, dtSt_ToDate, strPTName, strcause, PlaceSvc, MCalcNum, spccode, strCopaySet, spccode1)

                                        End If

                                    End If
                                Next msiRow
                            End If
                        End If

                    Next mxiPolicy

                End If

            Next memberNumber

            Dim conYear As String = ""
        Dim coFlag As Boolean = False

        'For Each mmiPolicy In mmiDetails
        '    'Checks if any of the gathered MMI policy sets had carry over
        '    If (mmiPolicy.mmi4Return.mmi4DRows(0).indDedCo = "1") Or
        '        (mmiPolicy.mmi4Return.mmi4CRows(0).famDedCo = "1") Or
        '        ((mmiPolicy.mmi10Return.mmi10ARows(0).newCoinsPrdCd = "3") Or
        '        (mmiPolicy.mmi10Return.mmi10ARows(0).newCoinsPrdCd = "6") Or
        '        (mmiPolicy.mmi10Return.mmi10ARows(0).newCoinsPrdCd = "8")) Then
        '        coFlag = True
        '    End If

        '    'Checks if any of the gathered MMI policy sets were contract year
        '    Select Case mmiPolicy.mmi1Return.mmi1ARows(0).polTypCd
        '        Case "33", "35", "37", "38", "39"
        '            conYear = Microsoft.VisualBasic.Right(mmiPolicy.mmi1Return.mmi1ARows(0).polYrPlnDt.Replace(" ", "/"), 5) & "-" & yearList.Text
        '            If IsDate(conYear) Then conYear = Format(CDate(conYear), "MM/dd/yyyy")
        '    End Select
        'Next mmiPolicy

        ''Adjusts the start of the year for contract/calendar
        'tempYearStart = "01/01/" & yearList.Text
        'tempYearEnd = ""

        'If conYear <> "" Then tempYearStart = conYear

        ''Readjusts the end of the year based on the beginning of the year
        'tempYearEnd = DateAdd("yyyy", 1, Format(CDate(tempYearStart), "MM/dd/yyyy"))
        'tempYearEnd = DateAdd("d", -1, Format(CDate(tempYearEnd), "MM/dd/yyyy"))

        ''Checks If range is still within year (can be shifted by contract year plans)
        'If DatePart("yyyy", CDate(tempYearEnd)).ToString <> yearList.Text Then
        '    tempYearStart = DateAdd("yyyy", -1, CDate(tempYearStart))
        '    tempYearEnd = DateAdd("yyyy", -1, CDate(tempYearEnd))
        'End If

        ''Readjusts the beginning of the year based on carry over
        'If coFlag = True Then tempYearStart = Format(DateAdd("m", -3, CDate(tempYearStart)), "MM/dd/yyyy")

        ''Assigns the starting and end date if those dates are the earliest/latest in the loop
        'If CDate(tempYearStart) < CDate(yearStart) Then yearStart = tempYearStart
        'If CDate(tempYearEnd) > CDate(yearEnd) Then yearEnd = tempYearEnd

        ''It will now repeat the process given the full range of the year
        ''(This ensures that all split plans are gathered and accounted for)
        'For Each mriPolicy In mriResp.Results.Response.rspMRIinfo.rspMRICoverageLine
        '    Dim mriStart As String = policyRange(mriPolicy.rspCovEffDT)                        'Coverage Line: Start
        '    Dim mriEnd As String = policyRange(mriPolicy.rspCovCanDT, mriPolicy.rspCovEffDT)   'Coverage Line: End

        '    'Checks if coverage info has a valid date. 
        '    'If the year range isn't within or overlapping that range it will continue to the next line
        '    If ((IsDate(mriStart) = False) Or (IsDate(mriEnd) = False)) OrElse
        '    (compareRange(yearStart, yearEnd, mriStart, mriEnd) = False) Then Continue For

        '    'This is the object that will be filled using the apiMXI module (it isn't called yet because
        '    'we need to check if the policy is railroad [railroad uses a larger MXI control line including rpt])
        '    Dim mxiResp = Nothing
        '    If checkRailroad(Trim(txt_Policy.Text)) = True Then
        '        mxiResp = api_MXI.PerformQuery(Trim(mriPolicy.rspCovPolicy), Trim(mriPolicy.rspCovPlan), Trim(mriPolicy.rspCovRept)).Results
        '    Else
        '        mxiResp = api_MXI.PerformQuery(Trim(mriPolicy.rspCovPolicy), Trim(mriPolicy.rspCovPlan)).Results
        '    End If

        '    'This will loop through all the coverage screens found in MXI for that particular MRI policy/plan
        '    'UNET screen equivelant: mxi,policy,plan
        '    For Each mxiPolicy In mxiResp.mxirows
        '        Dim mxiStart As String = policyRange(mxiPolicy.effDt)                   'Coverage Line: Start
        '        Dim mxiEnd As String = policyRange(mxiPolicy.cancDt, mxiPolicy.effDt)   'Coverage Line: End

        '        'Checks if coverage info has a valid date. 
        '        'If useDate isn't within that range it will continue to the next line
        '        If ((IsDate(mxiStart) = False) Or (IsDate(mxiEnd) = False)) OrElse
        '        (compareRange(yearStart, yearEnd, mxiStart, mxiEnd) = False) Then Continue For

        '        Dim mmiFlag As Boolean = False
        '        If mmiDetails Is Nothing Then   'Checks if mmiDetails list has been populated/initiated
        '            mmiDetails = New List(Of apiMMI.MmiReturn)
        '        Else
        '            'Checks if this MMI page has been added to the list previously (don't need to add duplicates)
        '            If mmiDetails.Count > 0 Then
        '                For Each mmiPolicy In mmiDetails
        '                    If (mmiPolicy.mmi1Return.mmi1ARows(0).polNbr = mxiPolicy.stdPlnPolNbr) And
        '                    (mmiPolicy.mmi1Return.mmi1ARows(0).plnNbr = mxiPolicy.stdPlnPlnNbr) Then
        '                        mmiFlag = True
        '                        Exit For
        '                    End If
        '                Next mmiPolicy
        '            Else
        '                mmiDetails = New List(Of apiMMI.MmiReturn)
        '            End If
        '        End If

        '        'If this set of mmi pages has already been added to the list previously it will skip this page
        '        If mmiFlag = True Then Continue For

        '        'Adds the mmiAPI response (the set of MMI Pages) object to a list to be used later
        '        If checkRailroad(Trim(txt_Policy.Text)) = True Then
        '            mmiDetails.Add(api_MMI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr, mxiPolicy.stdPlnClssNbr, "").Results)
        '        Else
        '            mmiDetails.Add(api_MMI.PerformQuery(mxiPolicy.stdPlnPolNbr, mxiPolicy.stdPlnPlnNbr, "", "").Results)
        '        End If

        '    Next mxiPolicy

        'Next mriPolicy

        ''We now have every possible MMI policy within that year so we can do our final adjustments to the date
        ''range. We also have all the MMI pages needed for the OOP Calculator.
        'For Each mmiPolicy In mmiDetails
        '    'Checks if any of the gathered MMI policy sets had carry over
        '    'I = mmiPolicy.mmi4Return.mmi4DRows(0).indDedCo
        '    'I = mmiPolicy.mmi4Return.mmi4DRows(1).indDedAmt

        '    If (mmiPolicy.mmi4Return.mmi4DRows(0).indDedCo = "1") Or
        '        (mmiPolicy.mmi4Return.mmi4CRows(0).famDedCo = "1") Or
        '        ((mmiPolicy.mmi10Return.mmi10ARows(0).newCoinsPrdCd = "3") Or
        '        (mmiPolicy.mmi10Return.mmi10ARows(0).newCoinsPrdCd = "6") Or
        '        (mmiPolicy.mmi10Return.mmi10ARows(0).newCoinsPrdCd = "8")) Then
        '        coFlag = True
        '    End If

        '    'Checks if any of the gathered MMI policy sets were contract year
        '    Select Case mmiPolicy.mmi1Return.mmi1ARows(0).polTypCd
        '        Case "33", "35", "37", "38", "39"
        '            conYear = Microsoft.VisualBasic.Right(mmiPolicy.mmi1Return.mmi1ARows(0).polYrPlnDt.Replace(" ", "/"), 5) & "-" & yearList.Text
        '            If IsDate(conYear) Then conYear = Format(CDate(conYear), "MM/dd/yyyy")
        '    End Select
        'Next mmiPolicy

        ''Adjusts the start of the year for contract/calendar
        'tempYearStart = "01/01/" & yearList.Text
        'tempYearEnd = ""

        'If conYear <> "" Then tempYearStart = conYear

        ''Readjusts the end of the year based on the beginning of the year
        'tempYearEnd = DateAdd("yyyy", 1, Format(CDate(tempYearStart), "MM/dd/yyyy"))
        'tempYearEnd = DateAdd("d", -1, Format(CDate(tempYearEnd), "MM/dd/yyyy"))

        ''Checks If range is still within year (can be shifted by contract year plans)
        'If DatePart("yyyy", CDate(tempYearEnd)).ToString <> yearList.Text Then
        '    tempYearStart = DateAdd("yyyy", -1, CDate(tempYearStart))
        '    tempYearEnd = DateAdd("yyyy", -1, CDate(tempYearEnd))
        'End If

        ''Readjusts the beginning of the year based on carry over
        'If coFlag = True Then tempYearStart = Format(DateAdd("m", -3, CDate(tempYearStart)), "MM/dd/yyyy")

        ''Assigns the starting and end date if those dates are the earliest/latest in the loop
        'If CDate(tempYearStart) < CDate(yearStart) Then yearStart = tempYearStart
        'If CDate(tempYearEnd) > CDate(yearEnd) Then yearEnd = tempYearEnd

        ''We will iterate through the MRI policies one final time to add a list of all possible members
        ''found in the MRI API object and check their eligibility to make sure they are active within the
        ''date range we've found
        'For Each mriPolicy In mriResp.Results.Response.rspMRIinfo.rspMRICoverageLine
        '    Dim mriStart As String = policyRange(mriPolicy.rspCovEffDT)                        'Coverage Line: Start
        '    Dim mriEnd As String = policyRange(mriPolicy.rspCovCanDT, mriPolicy.rspCovEffDT)   'Coverage Line: End

        '    'Checks if coverage info has a valid date. 
        '    'If the year range isn't within or overlapping that range it will continue to the next line
        '    If ((IsDate(mriStart) = False) Or (IsDate(mriEnd) = False)) OrElse
        '    (compareRange(yearStart, yearEnd, mriStart, mriEnd) = False) Then Continue For

        '    'Adds the subscriber to the member list
        '    Dim eeStart As String = policyRange(mriResp.Results.Response.rspMRIinfo.rspMRIEmployeeCovInfo.rspEEMedEffDt)
        '    Dim eeEnd As String = policyRange(mriResp.Results.Response.rspMRIinfo.rspMRIEmployeeCovInfo.rspEEMedCanDt,
        '        mriResp.Results.Response.rspMRIinfo.rspMRIEmployeeCovInfo.rspEEMedEffDt)

        '    'memberList.Items.Add(mriResp.Results.Response.rspMRIinfo.rspMRIEmployeeCovInfo.rspEEFirstName & "/" &
        '    '    mriResp.Results.Response.rspMRIinfo.rspMRIEmployeeCovInfo.rspEERelCd,
        '    '    compareRange(yearStart, yearEnd, eeStart, eeEnd))

        'Next mriPolicy

        'Assigns the final 
        Try
            'startSelect.Text = yearStart
            'endSelect.Text = yearEnd
        Catch
        End Try
        mmiDetails.Clear()

        If mmiFlag = True Then
            MsgBox("Looks like Plan is not falling in date range, Please select valid date range and re-run")
        End If

        'For crow = 0 To tblCopay.Rows.Count - 2
        '    tblCopay.Rows(crow).Cells(3).Value = Trim(DGridCEI.Rows(0).Cells(0).Value) & "/" & Trim(DGridCEI.Rows(0).Cells(1).Value) 'ptFName & "/" & ptRelation
        'Next

        For memberNumber = 0 To memberList.CheckedItems.Count - 1
            'Splits the member name and rel so it's usable in the AHI Query
            Dim splitMbr = memberList.CheckedItems(memberNumber)
        Next


        RichTextBox1.AppendText("Data fetched from MMI screen   " & vbCrLf)
        RichTextBox1.SelectionBullet = True
        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        MMIendtime = Now()

        TabControl1.SelectedIndex = 2

        'MsgBox("Please check if policy is spliting in same year for example 04/01/2023-05/31/2023 and 01/01/2023-03/31/2023 then update the date as below for that Member " & vbCrLf & "01/01/2023-12/31/2023")

    End Sub


    Private Sub Chk_select_memlist_CheckedChanged(sender As Object, e As EventArgs) Handles Chk_select_memlist.CheckedChanged
        If Chk_select_memlist.Checked Then
            For i As Integer = 0 To memberList.Items.Count - 1
                memberList.SetItemChecked(i, True)
            Next
        Else
            For i As Integer = 0 To memberList.Items.Count - 1
                memberList.SetItemChecked(i, False)
            Next
        End If
    End Sub

    Sub GetProvInfo()
        Dim strStartRow As Integer
        Dim strEndRow As Integer
        Dim NextPref As String
        Dim NextTin As String
        Dim NextSffx As String

        Dim ProvPref As String
        Dim ProvTin As String
        Dim ProvSffx As String


        For q = 0 To tblMHI.Rows.Count - 2
            strStartRow = q
            ProvPref = tblMHI.Rows(q).Cells(21).Value
            ProvTin = tblMHI.Rows(q).Cells(22).Value
            ProvSffx = tblMHI.Rows(q).Cells(23).Value
            strfullTin = tblMHI.Rows(q).Cells(21).Value & tblMHI.Rows(q).Cells(22).Value & tblMHI.Rows(q).Cells(23).Value

            Do
                If Not IsDBNull(tblMHI.Rows(q + 1).Cells(21).Value) Then
                    NextPref = tblMHI.Rows(q + 1).Cells(21).Value
                    NextTin = tblMHI.Rows(q + 1).Cells(22).Value
                    NextSffx = tblMHI.Rows(q + 1).Cells(23).Value
                End If
                If ProvPref <> NextPref Or ProvTin <> NextTin Or ProvSffx <> NextSffx Then
                    If ProvTin = "" And NextTin = "" Then
                        Exit Sub
                    Else
                        Exit Do
                    End If
                Else
                    q = q + 1
                End If

            Loop

            strEndRow = q
            If ProvTin <> "Pharmacy" Then


                Dim pmi_response = api_PMI.PerformQuery(strfullTin)
                If strfullTin = "000000000000000" Then Continue For
                For Each pmiRsp In pmi_response.Results.Response
                    strProv = pmiRsp.rspPrvBillAddr.rspPrvName
                    strProv = strProv.Replace(";", "")
                    proType = pmiRsp.rspPrvType
                    Exit For
                Next

                For intRowCnt = strStartRow To strEndRow
                    tblMHI.Rows(intRowCnt).Cells(51).Value = strProv
                    tblMHI.Rows(intRowCnt).Cells(52).Value = proType
                    If intRowCnt = strEndRow Then Exit For
                Next

                'tblMHI.Rows(q).Cells(51).Value = strProv
                'tblMHI.Rows(q).Cells(52).Value = proType

            Else
                tblMHI.Rows(q).Cells(51).Value = "Pharmacy"

                tblMHI.Rows(q).Cells(52).Value = "Rx"
            End If

        Next q

        Dim t As Integer
        'For t = 0 To tblMHI.Rows.Count - 1
        '    Try
        '        strProvName = tblMHI.Rows(t).Cells(51).Value
        '        strfullTin = tblMHI.Rows(t).Cells(49).Value
        '        Call update_Provider_Name(strfullTin, strProvName)
        '    Catch ex As Exception

        '    End Try

        'Next


        'MsgBox("Gathered Prov Info")
        RichTextBox1.AppendText("Data fetched for Provider name and Type...   " & vbCrLf)
        RichTextBox1.SelectionBullet = True
        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        tblMHI.Refresh()
    End Sub

    Sub update_Provider_Name(strfullTin, ProviderName)
        For c = 0 To tblMHI.Rows.Count - 1
            Try
                If tblMHI.Rows(c).Cells(49).Value = strfullTin Then
                    tblMHI.Rows(c).Cells(51).Value = strProvName
                End If
            Catch ex As Exception

            End Try

        Next
    End Sub

    Private Sub yearList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles yearList.SelectedIndexChanged

        Dim dtMonth As Integer = 12
        Dim dtDay As Integer = 31

        Dim yearStart As String = "01/01/" & yearList.Text
        'Dim yearEnd As String = "12/31/" & yearList.Text
        Dim yearEnd As String = dtMonth & "/" & dtDay & "/" & yearList.Text
        startSelect.Text = yearStart
        endSelect.Text = yearEnd

    End Sub

    Public Shared Function policyRange(reqDate As String, Optional optDate As String = "") As String
        Select Case optDate
            Case ""
                Return checkDate(reqDate)
                Exit Function
            Case Else
                Dim tempReq As String = checkDate(reqDate)
                Dim tempOpt As String = checkDate(optDate)

                If (IsDate(tempReq)) And (IsDate(tempOpt)) Then
                    If CDate(tempReq) < CDate(tempOpt) Then
                        For yearAdd = 1 To 50
                            tempReq = Format(CDate(DateAdd("yyyy", +1, CDate(tempReq))), "MM/dd/yyyy")
                            If CDate(tempReq) >= CDate(tempOpt) Then Exit For
                        Next yearAdd
                    End If

                    Return checkDate(tempReq)
                    Exit Function
                ElseIf (IsDate(tempReq) = False) And (IsDate(tempOpt)) Then
                    If IsDate(tempOpt) Then
                        'Dim dateStart As String = Format(CDate(DateAdd("yyyy", +1, CDate(tempOpt))), "MM/dd/yyyy")
                        Dim dateStart As String = Format(CDate(DateAdd("yyyy", +0, CDate(tempOpt))), "MM/dd/yyyy")
                        'dateStart = Format(CDate(DateAdd("d", -1, CDate(dateStart))), "MM/dd/yyyy")
                        dateStart = Format(Now, "MM/dd/yyyy")

                        'For yearAdd = 1 To 50
                        '    'dateStart = Format(CDate(DateAdd("yyyy", +1, CDate(dateStart))), "MM/dd/yyyy")
                        '    dateStart = Format(CDate(DateAdd("yyyy", +0, CDate(dateStart))), "MM/dd/yyyy")
                        '    If CDate(dateStart) >= CDate(Now) Then Exit For
                        'Next yearAdd

                        Return Format(CDate(dateStart), "MM/dd/yyyy")
                        Exit Function
                    End If
                End If

                Return ""
        End Select
    End Function

    'CHECK: Dates (Adjusts dates for unet out - 9999 years)
    Public Shared Function checkDate(dateCheck As String)
        If IsDate(dateCheck) Then
            If DatePart("yyyy", CDate(dateCheck)) = "9999" Then
                Dim dateStart As String = DatePart("m", dateCheck) & "/" & DatePart("d", dateCheck) & "/" & DatePart("yyyy", Now)

                If CDate(dateStart) < CDate(Now) Then
                    For yearAdd = 1 To 50
                        dateStart = Format(CDate(DateAdd("yyyy", +1, CDate(dateStart))), "MM/dd/yyyy")
                        If CDate(dateStart) >= CDate(Now) Then Exit For
                    Next yearAdd
                End If

                Return Format(CDate(dateStart), "MM/dd/yyyy")
                Exit Function
            Else
                Return Format(CDate(dateCheck), "MM/dd/yyyy")
                Exit Function
            End If
        Else
            If InStr(dateCheck, "-") Then
                Dim tempDate As String = dateCheck.Replace("-", "/")
                Return checkDate(tempDate)
                Exit Function
            Else
                Return ""
                Exit Function
            End If
        End If
    End Function

    'CHECK: If policy is RAILROAD
    Public Shared Function checkRailroad(policyNumber As String) As Boolean
        If (Trim(policyNumber) = "000001") Or (Trim(policyNumber) = "023000") Or
        (Trim(policyNumber) = "690100") Or (Trim(policyNumber) = "046000") Or
        (Trim(policyNumber) = "023111") Or (Trim(policyNumber) = "107300") Then
            Return True
        Else
            Return False
        End If
    End Function

    Sub Get_mxiDetails()
        Dim apiMXIobj As apiMXI = New apiMXI
        strMXIdata = apiMXIobj.PerformQuery("182019", "1236", "").jsonResponse
        Call Parse_MXI(strMXIdata)
    End Sub
    Private Sub ELGSLetterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ELGSLetterToolStripMenuItem.Click
        Form2.Show()
    End Sub

    'Private Sub InstructionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InstructionToolStripMenuItem.Click
    '    Login.Show()
    'End Sub        ''need to check sanjeet

    'Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
    '    About.Show()
    'End Sub

    'These are all the MMI items I use for the Workspace - all you have to do is substitute the
    'mmiSelect for whichever mmiObject you're using. 
    'Note: Not all of them have been confirmed And tested - but it's a good example of how to index 
    'the different MMI pages.

    Public Sub mmiItems(ByVal colno As Int32, ByVal Plnnbr As String, ByVal Startdt As Date, ByVal strEnddt As Date)

        'Try


        'strPTName = DGridCEI.Rows(0).Cells(0).Value & "/" & DGridCEI.Rows(0).Cells(1).Value
        'Dim colno As Int32 = 1
        For Each mmiSelect In mmiDetails

            '    'MMI PAGE 1x
            Threading.Thread.Sleep(500)

            DGrid_PG1.Rows(0).Cells(0).Value = "Policy" : DGrid_PG1.Rows(0).Cells(colno).Value = txt_Policy.Text
            DGrid_PG1.Rows(1).Cells(0).Value = "Plan Code/Reporting Code/Plan Var" : DGrid_PG1.Rows(1).Cells(colno).Value = Plnnbr
            DGrid_PG1.Rows(2).Cells(0).Value = "Year" : DGrid_PG1.Rows(2).Cells(colno).Value = Startdt & "-" & strEnddt
            DGrid_PG1.Rows(3).Cells(0).Value = "Patient Name" : DGrid_PG1.Rows(3).Cells(colno).Value = strPTName
            DGrid_PG1.Rows(4).Cells(0).Value = "Effective Date" : DGrid_PG1.Rows(4).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).effDt
            DGrid_PG1.Rows(5).Cells(0).Value = "Cancel Date" : DGrid_PG1.Rows(5).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).cancDt
            DGrid_PG1.Rows(6).Cells(0).Value = "Eligibility Code" : DGrid_PG1.Rows(6).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).eligCd
            DGrid_PG1.Rows(7).Cells(0).Value = "ASO Type" : DGrid_PG1.Rows(7).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).fundTypCd
            DGrid_PG1.Rows(8).Cells(0).Value = "Policy Type" : DGrid_PG1.Rows(8).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).polTypCd
            DGrid_PG1.Rows(9).Cells(0).Value = "COB Code" : DGrid_PG1.Rows(9).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).cobStdCd
            DGrid_PG1.Rows(10).Cells(0).Value = "COB Medicare Code" : DGrid_PG1.Rows(10).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).cobMcCd
            DGrid_PG1.Rows(11).Cells(0).Value = "COB Pay and Pursue" : DGrid_PG1.Rows(11).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).cobPayPursCd
            DGrid_PG1.Rows(12).Cells(0).Value = "Risk" : DGrid_PG1.Rows(12).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).riskClsCd
            DGrid_PG1.Rows(13).Cells(0).Value = "Lifetime Max Amount" : DGrid_PG1.Rows(13).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).lftmMaxAmt
            DGrid_PG1.Rows(14).Cells(0).Value = "Lifetime Accum Code" : DGrid_PG1.Rows(14).Cells(colno).Value = mmiSelect.mmi1Return.mmi1CRows(0).lftAccumCd
            DGrid_PG1.Rows(15).Cells(0).Value = "Lifetime Max Cross Apply" : DGrid_PG1.Rows(15).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).lftmMaxXapplInd
            DGrid_PG1.Rows(16).Cells(0).Value = "Contract State" : DGrid_PG1.Rows(16).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).contrStCd
            DGrid_PG1.Rows(17).Cells(0).Value = "Medicare Part D Election" : DGrid_PG1.Rows(17).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).medcrPartdCustElecCd
            DGrid_PG1.Rows(18).Cells(0).Value = "Plan/Vendor Attribute" : DGrid_PG1.Rows(18).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).hiDedPlnCd
            DGrid_PG1.Rows(19).Cells(0).Value = "Non-Embedded Ded" : DGrid_PG1.Rows(19).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).nonEmbdDedCd

            '' commented by sanjeet for page 1
            '    'MMI PAGE 2
            '    plnDet_MMIList.Rows.Add("2", "Vendor: RX ID", mmiSelect.mmi2Return.mmi2BRows(3).combVendId,
            '        "This field identifies the pharmacy vendor applicable to the combined RX feed.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: RX Feed", mmiSelect.mmi2Return.mmi2BRows(3).combExchgTypCd,
            '        "This field will identify if the pharmacy vendor is real time, batch, or highway.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: RX RC", mmiSelect.mmi2Return.mmi2BRows(3).combRmrkCd,
            '        "A one-shot populated the RX RC field with remark code B3 on any TOPS policymaster with the vendor code MI (Med Impact) coded in the RX VDR field.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: Vision ID", mmiSelect.mmi2Return.mmi2BRows(4).combVendId,
            '        "This field will identify the vision vendor applicable to the combined medical/vision feed.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: Vision Feed", mmiSelect.mmi2Return.mmi2BRows(4).combExchgTypCd,
            '        "This field will identify if the vision vendor is Real Time, Batch, or Highway.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: Vision RC", mmiSelect.mmi2Return.mmi2BRows(4).combRmrkCd,
            '        "")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: Dental ID", mmiSelect.mmi2Return.mmi2BRows(1).combVendId,
            '      "This field identifies the dental vendor applicable to the combined Dental feed.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: Dental Feed", mmiSelect.mmi2Return.mmi2BRows(1).combExchgTypCd,
            '        "This field will identify if the dental vendor is real time, batch, or highway.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: Dental RC", mmiSelect.mmi2Return.mmi2BRows(1).combRmrkCd,
            '        "")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: BH ID", mmiSelect.mmi2Return.mmi2BRows(0).combVendId,
            '        "This field identifies the BH vendor applicable to the combined BH feed.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: BH Feed", mmiSelect.mmi2Return.mmi2BRows(0).combExchgTypCd,
            '        "This field will identify if the BH vendor is real time, batch, or highway.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: BH RC", mmiSelect.mmi2Return.mmi2BRows(0).combRmrkCd,
            '        "")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: MD ID", mmiSelect.mmi2Return.mmi2BRows(2).combVendId,
            '        "This field identifies the MD vendor applicable to the combined MD feed.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: MD Feed", mmiSelect.mmi2Return.mmi2BRows(2).combExchgTypCd,
            '        "This field will identify if the MD vendor is real time, batch, or highway.")
            '    plnDet_MMIList.Rows.Add("2", "Vendor: MD RC", mmiSelect.mmi2Return.mmi2BRows(2).combRmrkCd,
            '        "")


            'MMI PAGE 4
            '[
            'MsgBox(mmiSelect.mmi4Return.mmi4BRows.Count)

            Try

                'mmiSelect.mmi4Return.mmi4DRows(0).dedEndDt

                DGrid_PG4.Rows(0).Cells(0).Value = "Policy" : DGrid_PG4.Rows(0).Cells(colno).Value = txt_Policy.Text
                DGrid_PG4.Rows(1).Cells(0).Value = "Plan Code/Reporting Code/Plan Var" : DGrid_PG4.Rows(1).Cells(colno).Value = Plnnbr
                DGrid_PG4.Rows(2).Cells(0).Value = "Year" : DGrid_PG4.Rows(2).Cells(colno).Value = Startdt & "-" & strEnddt
                DGrid_PG4.Rows(3).Cells(0).Value = "Patient Name" : DGrid_PG4.Rows(3).Cells(colno).Value = strPTName
                DGrid_PG4.Rows(4).Cells(0).Value = "IND DED 1: End Date" : DGrid_PG4.Rows(4).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedEndDt
                DGrid_PG4.Rows(5).Cells(0).Value = "IND DED 1: Network Type" : DGrid_PG4.Rows(5).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedNtwkTypCd
                DGrid_PG4.Rows(6).Cells(0).Value = "IND DED 1: Amount" : DGrid_PG4.Rows(6).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).indDedAmt
                DGrid_PG4.Rows(7).Cells(0).Value = "IND DED 1: Description" : DGrid_PG4.Rows(7).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedSrvcDesc
                DGrid_PG4.Rows(8).Cells(0).Value = "IND DED 1: Cost Containment" : DGrid_PG4.Rows(8).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).indDedCc
                DGrid_PG4.Rows(9).Cells(0).Value = "IND DED 1: Type Code" : DGrid_PG4.Rows(9).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).covTypCd
                DGrid_PG4.Rows(10).Cells(0).Value = "IND DED 1: Frequency Code" : DGrid_PG4.Rows(10).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedFreqCd
                DGrid_PG4.Rows(11).Cells(0).Value = "IND DED 1: Benefit Period" : DGrid_PG4.Rows(11).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedMntPdAmt
                DGrid_PG4.Rows(12).Cells(0).Value = "IND DED 1: Carry-Over Code" : DGrid_PG4.Rows(12).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).indDedCo
                DGrid_PG4.Rows(13).Cells(0).Value = "IND DED 1: COB Exclusion" : DGrid_PG4.Rows(13).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedCobCd
                DGrid_PG4.Rows(14).Cells(0).Value = "IND DED 1: X-Semi Private Rate" : DGrid_PG4.Rows(14).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedSemiPvtRtCd
                DGrid_PG4.Rows(15).Cells(0).Value = "IND DED 1: Accum Code" : DGrid_PG4.Rows(15).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedAccumCd
                DGrid_PG4.Rows(16).Cells(0).Value = "IND DED 1: Accum Period" : DGrid_PG4.Rows(16).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedAccumPdAmt
                DGrid_PG4.Rows(17).Cells(0).Value = "IND DED 1: Maint Amount" : DGrid_PG4.Rows(17).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedMntAmt
                DGrid_PG4.Rows(18).Cells(0).Value = "IND DED 1: Maint Code" : DGrid_PG4.Rows(18).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedMntCd
                DGrid_PG4.Rows(10).Cells(0).Value = "IND DED 1: Maint Period" : DGrid_PG4.Rows(19).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedMntPdAmt
                DGrid_PG4.Rows(20).Cells(0).Value = "IND DED 1: Combined RX" : DGrid_PG4.Rows(20).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(3).combDed1Ind
                DGrid_PG4.Rows(21).Cells(0).Value = "IND DED 1: Combined Vision" : DGrid_PG4.Rows(21).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(4).combDed1Ind
                DGrid_PG4.Rows(22).Cells(0).Value = "IND DED 1: Combined Dental" : DGrid_PG4.Rows(22).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(1).combDed1Ind
                DGrid_PG4.Rows(23).Cells(0).Value = "IND DED 1: Combined BH" : DGrid_PG4.Rows(23).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(0).combDed1Ind
                DGrid_PG4.Rows(24).Cells(0).Value = "IND DED 1: Combined MD" : DGrid_PG4.Rows(24).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(2).combDed1Ind
                DGrid_PG4.Rows(25).Cells(0).Value = "IND DED 2: End Date" : DGrid_PG4.Rows(25).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedEndDt
                DGrid_PG4.Rows(26).Cells(0).Value = "IND DED 2: End Date" : DGrid_PG4.Rows(26).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedEndDt
                DGrid_PG4.Rows(27).Cells(0).Value = "IND DED 2: End Date" : DGrid_PG4.Rows(27).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedEndDt
                DGrid_PG4.Rows(28).Cells(0).Value = "IND DED 2: Network Type" : DGrid_PG4.Rows(28).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedNtwkTypCd
                DGrid_PG4.Rows(29).Cells(0).Value = "IND DED 2: Amount" : DGrid_PG4.Rows(29).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).indDedAmt
                DGrid_PG4.Rows(30).Cells(0).Value = "IND DED 2: Description" : DGrid_PG4.Rows(30).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedSrvcDesc
                DGrid_PG4.Rows(31).Cells(0).Value = "IND DED 2: Cost Containment" : DGrid_PG4.Rows(31).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).indDedCc
                DGrid_PG4.Rows(32).Cells(0).Value = "IND DED 2: Type Code" : DGrid_PG4.Rows(32).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).covTypCd
                DGrid_PG4.Rows(33).Cells(0).Value = "IND DED 2: Frequency Code" : DGrid_PG4.Rows(33).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedFreqCd
                DGrid_PG4.Rows(34).Cells(0).Value = "IND DED 2: Benefit Period" : DGrid_PG4.Rows(34).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedMntPdAmt
                DGrid_PG4.Rows(35).Cells(0).Value = "IND DED 2: Carry-Over Code" : DGrid_PG4.Rows(35).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).indDedCo
                DGrid_PG4.Rows(36).Cells(0).Value = "IND DED 2: COB Exclusion" : DGrid_PG4.Rows(36).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedCobCd
                DGrid_PG4.Rows(37).Cells(0).Value = "IND DED 2: X-Semi Private Rate" : DGrid_PG4.Rows(37).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedSemiPvtRtCd
                DGrid_PG4.Rows(38).Cells(0).Value = "IND DED 2: Accum Code" : DGrid_PG4.Rows(38).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedAccumCd
                DGrid_PG4.Rows(39).Cells(0).Value = "IND DED 2: Accum Period" : DGrid_PG4.Rows(39).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedAccumPdAmt
                DGrid_PG4.Rows(40).Cells(0).Value = "IND DED 2: Maint Amount" : DGrid_PG4.Rows(40).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedMntAmt
                DGrid_PG4.Rows(41).Cells(0).Value = "IND DED 2: Maint Code" : DGrid_PG4.Rows(41).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedMntCd
                DGrid_PG4.Rows(42).Cells(0).Value = "IND DED 2: Maint Period" : DGrid_PG4.Rows(42).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).dedMntPdAmt
                DGrid_PG4.Rows(43).Cells(0).Value = "IND DED 2: Combined RX" : DGrid_PG4.Rows(43).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(3).combDed2Ind
                DGrid_PG4.Rows(44).Cells(0).Value = "IND DED 2: Combined Vision" : DGrid_PG4.Rows(44).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(4).combDed2Ind
                DGrid_PG4.Rows(45).Cells(0).Value = "IND DED 2: Combined Dental" : DGrid_PG4.Rows(45).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(1).combDed2Ind
                DGrid_PG4.Rows(46).Cells(0).Value = "IND DED 2: Combined BH" : DGrid_PG4.Rows(46).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(0).combDed2Ind
                DGrid_PG4.Rows(47).Cells(0).Value = "IND DED 2: Combined MD" : DGrid_PG4.Rows(47).Cells(colno).Value = mmiSelect.mmi4Return.mmi4BRows(2).combDed2Ind
                DGrid_PG4.Rows(48).Cells(0).Value = "FAM DED 1: Amount" : DGrid_PG4.Rows(48).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).famDedAmt
                DGrid_PG4.Rows(49).Cells(0).Value = "FAM DED 1: Description" : DGrid_PG4.Rows(49).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).dedMbrDesc
                DGrid_PG4.Rows(50).Cells(0).Value = "FAM DED 1: Cost Containment" : DGrid_PG4.Rows(50).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).famDedCc
                DGrid_PG4.Rows(51).Cells(0).Value = "FAM DED 1: Type Code" : DGrid_PG4.Rows(51).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).covTypCd
                DGrid_PG4.Rows(52).Cells(0).Value = "FAM DED 1: Frequency Code" : DGrid_PG4.Rows(52).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).dedFreqPrdCd
                DGrid_PG4.Rows(53).Cells(0).Value = "FAM DED 1: Carry-Over Code" : DGrid_PG4.Rows(53).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).famDedCo
                DGrid_PG4.Rows(54).Cells(0).Value = "FAM DED 1: Multiplier" : DGrid_PG4.Rows(54).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).dedMultSalryPct
                DGrid_PG4.Rows(55).Cells(0).Value = "FAM DED 1: Multiplier (DED)" : DGrid_PG4.Rows(55).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).dedMultFct
                DGrid_PG4.Rows(56).Cells(0).Value = "FAM DED 1: Multiplier (OOP)" : DGrid_PG4.Rows(56).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).oopMultFct
                DGrid_PG4.Rows(57).Cells(0).Value = "FAM DED 1: Text" : DGrid_PG4.Rows(57).Cells(colno).Value = mmiSelect.mmi4Return.mmi4ARows(0).famTxtSwapCd
                DGrid_PG4.Rows(58).Cells(0).Value = "IND TIER 1: Label" : DGrid_PG4.Rows(58).Cells(colno).Value = mmiSelect.mmi4Return.mmi4ARows(0).tierLblInd
                DGrid_PG4.Rows(59).Cells(0).Value = "ALT DED 1: EE+1 Amount" : DGrid_PG4.Rows(59).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).dedEePls1Amt
                DGrid_PG4.Rows(60).Cells(0).Value = "ALT DED 1: EE+SP Amount" : DGrid_PG4.Rows(60).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).dedEeSpoAmt
                DGrid_PG4.Rows(61).Cells(0).Value = "ALT DED 1: EE+CH Amount" : DGrid_PG4.Rows(61).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).dedEeChrgAmt
                DGrid_PG4.Rows(62).Cells(0).Value = "FAM DED 1: Prorate Interval" : DGrid_PG4.Rows(62).Cells(colno).Value = mmiSelect.mmi4Return.mmi4ARows(0).prortIntrvlFreqCd
                DGrid_PG4.Rows(63).Cells(0).Value = "FAM DED 1: Prorate Interval" : DGrid_PG4.Rows(63).Cells(colno).Value = mmiSelect.mmi4Return.mmi4ARows(0).prortEvntTypCd

            Catch ex As Exception
                DGrid_PG4.Rows(0).Cells(0).Value = "Policy" : DGrid_PG4.Rows(0).Cells(colno).Value = txt_Policy.Text
                DGrid_PG4.Rows(1).Cells(0).Value = "Plan Code/Reporting Code/Plan Var" : DGrid_PG4.Rows(1).Cells(colno).Value = Plnnbr
                DGrid_PG4.Rows(2).Cells(0).Value = "Year" : DGrid_PG4.Rows(2).Cells(colno).Value = Startdt & "-" & strEnddt
                DGrid_PG4.Rows(3).Cells(0).Value = "Patient Name" : DGrid_PG4.Rows(3).Cells(colno).Value = strPTName
                DGrid_PG4.Rows(4).Cells(0).Value = "IND DED 1: End Date" : DGrid_PG4.Rows(4).Cells(colno).Value = 0
                DGrid_PG4.Rows(5).Cells(0).Value = "IND DED 1: Network Type" : DGrid_PG4.Rows(5).Cells(colno).Value = 0
                DGrid_PG4.Rows(6).Cells(0).Value = "IND DED 1: Amount" : DGrid_PG4.Rows(6).Cells(colno).Value = 0
                DGrid_PG4.Rows(7).Cells(0).Value = "IND DED 1: Description" : DGrid_PG4.Rows(7).Cells(colno).Value = 0
                DGrid_PG4.Rows(8).Cells(0).Value = "IND DED 1: Cost Containment" : DGrid_PG4.Rows(8).Cells(colno).Value = 0
                DGrid_PG4.Rows(9).Cells(0).Value = "IND DED 1: Type Code" : DGrid_PG4.Rows(9).Cells(colno).Value = 0
                DGrid_PG4.Rows(10).Cells(0).Value = "IND DED 1: Frequency Code" : DGrid_PG4.Rows(10).Cells(colno).Value = 0
                DGrid_PG4.Rows(11).Cells(0).Value = "IND DED 1: Benefit Period" : DGrid_PG4.Rows(11).Cells(colno).Value = 0
                DGrid_PG4.Rows(12).Cells(0).Value = "IND DED 1: Carry-Over Code" : DGrid_PG4.Rows(12).Cells(colno).Value = 0
                DGrid_PG4.Rows(13).Cells(0).Value = "IND DED 1: COB Exclusion" : DGrid_PG4.Rows(13).Cells(colno).Value = 0
                DGrid_PG4.Rows(14).Cells(0).Value = "IND DED 1: X-Semi Private Rate" : DGrid_PG4.Rows(14).Cells(colno).Value = 0
                DGrid_PG4.Rows(15).Cells(0).Value = "IND DED 1: Accum Code" : DGrid_PG4.Rows(15).Cells(colno).Value = 0
                DGrid_PG4.Rows(16).Cells(0).Value = "IND DED 1: Accum Period" : DGrid_PG4.Rows(16).Cells(colno).Value = 0
                DGrid_PG4.Rows(17).Cells(0).Value = "IND DED 1: Maint Amount" : DGrid_PG4.Rows(17).Cells(colno).Value = 0
                DGrid_PG4.Rows(18).Cells(0).Value = "IND DED 1: Maint Code" : DGrid_PG4.Rows(18).Cells(colno).Value = 0
                DGrid_PG4.Rows(10).Cells(0).Value = "IND DED 1: Maint Period" : DGrid_PG4.Rows(19).Cells(colno).Value = 0
                DGrid_PG4.Rows(20).Cells(0).Value = "IND DED 1: Combined RX" : DGrid_PG4.Rows(20).Cells(colno).Value = 0
                DGrid_PG4.Rows(21).Cells(0).Value = "IND DED 1: Combined Vision" : DGrid_PG4.Rows(21).Cells(colno).Value = 0
                DGrid_PG4.Rows(22).Cells(0).Value = "IND DED 1: Combined Dental" : DGrid_PG4.Rows(22).Cells(colno).Value = 0
                DGrid_PG4.Rows(23).Cells(0).Value = "IND DED 1: Combined BH" : DGrid_PG4.Rows(23).Cells(colno).Value = 0
                DGrid_PG4.Rows(24).Cells(0).Value = "IND DED 1: Combined MD" : DGrid_PG4.Rows(24).Cells(colno).Value = 0
                DGrid_PG4.Rows(25).Cells(0).Value = "IND DED 2: End Date" : DGrid_PG4.Rows(25).Cells(colno).Value = 0
                DGrid_PG4.Rows(26).Cells(0).Value = "IND DED 2: End Date" : DGrid_PG4.Rows(26).Cells(colno).Value = 0
                DGrid_PG4.Rows(27).Cells(0).Value = "IND DED 2: End Date" : DGrid_PG4.Rows(27).Cells(colno).Value = 0
                DGrid_PG4.Rows(28).Cells(0).Value = "IND DED 2: Network Type" : DGrid_PG4.Rows(28).Cells(colno).Value = 0
                DGrid_PG4.Rows(29).Cells(0).Value = "IND DED 2: Amount" : DGrid_PG4.Rows(29).Cells(colno).Value = 0
                DGrid_PG4.Rows(30).Cells(0).Value = "IND DED 2: Description" : DGrid_PG4.Rows(30).Cells(colno).Value = 0
                DGrid_PG4.Rows(31).Cells(0).Value = "IND DED 2: Cost Containment" : DGrid_PG4.Rows(31).Cells(colno).Value = 0
                DGrid_PG4.Rows(32).Cells(0).Value = "IND DED 2: Type Code" : DGrid_PG4.Rows(32).Cells(colno).Value = 0
                DGrid_PG4.Rows(33).Cells(0).Value = "IND DED 2: Frequency Code" : DGrid_PG4.Rows(33).Cells(colno).Value = 0
                DGrid_PG4.Rows(34).Cells(0).Value = "IND DED 2: Benefit Period" : DGrid_PG4.Rows(34).Cells(colno).Value = 0
                DGrid_PG4.Rows(35).Cells(0).Value = "IND DED 2: Carry-Over Code" : DGrid_PG4.Rows(35).Cells(colno).Value = 0
                DGrid_PG4.Rows(36).Cells(0).Value = "IND DED 2: COB Exclusion" : DGrid_PG4.Rows(36).Cells(colno).Value = 0
                DGrid_PG4.Rows(37).Cells(0).Value = "IND DED 2: X-Semi Private Rate" : DGrid_PG4.Rows(37).Cells(colno).Value = 0
                DGrid_PG4.Rows(38).Cells(0).Value = "IND DED 2: Accum Code" : DGrid_PG4.Rows(38).Cells(colno).Value = 0
                DGrid_PG4.Rows(39).Cells(0).Value = "IND DED 2: Accum Period" : DGrid_PG4.Rows(39).Cells(colno).Value = 0
                DGrid_PG4.Rows(40).Cells(0).Value = "IND DED 2: Maint Amount" : DGrid_PG4.Rows(40).Cells(colno).Value = 0
                DGrid_PG4.Rows(41).Cells(0).Value = "IND DED 2: Maint Code" : DGrid_PG4.Rows(41).Cells(colno).Value = 0
                DGrid_PG4.Rows(42).Cells(0).Value = "IND DED 2: Maint Period" : DGrid_PG4.Rows(42).Cells(colno).Value = 0
                DGrid_PG4.Rows(43).Cells(0).Value = "IND DED 2: Combined RX" : DGrid_PG4.Rows(43).Cells(colno).Value = 0
                DGrid_PG4.Rows(44).Cells(0).Value = "IND DED 2: Combined Vision" : DGrid_PG4.Rows(44).Cells(colno).Value = 0
                DGrid_PG4.Rows(45).Cells(0).Value = "IND DED 2: Combined Dental" : DGrid_PG4.Rows(45).Cells(colno).Value = 0
                DGrid_PG4.Rows(46).Cells(0).Value = "IND DED 2: Combined BH" : DGrid_PG4.Rows(46).Cells(colno).Value = 0
                DGrid_PG4.Rows(47).Cells(0).Value = "IND DED 2: Combined MD" : DGrid_PG4.Rows(47).Cells(colno).Value = 0
                DGrid_PG4.Rows(48).Cells(0).Value = "FAM DED 1: Amount" : DGrid_PG4.Rows(48).Cells(colno).Value = 0
                DGrid_PG4.Rows(49).Cells(0).Value = "FAM DED 1: Description" : DGrid_PG4.Rows(49).Cells(colno).Value = 0
                DGrid_PG4.Rows(50).Cells(0).Value = "FAM DED 1: Cost Containment" : DGrid_PG4.Rows(50).Cells(colno).Value = 0
                DGrid_PG4.Rows(51).Cells(0).Value = "FAM DED 1: Type Code" : DGrid_PG4.Rows(51).Cells(colno).Value = 0
                DGrid_PG4.Rows(52).Cells(0).Value = "FAM DED 1: Frequency Code" : DGrid_PG4.Rows(52).Cells(colno).Value = 0
                DGrid_PG4.Rows(53).Cells(0).Value = "FAM DED 1: Carry-Over Code" : DGrid_PG4.Rows(53).Cells(colno).Value = 0
                DGrid_PG4.Rows(54).Cells(0).Value = "FAM DED 1: Multiplier" : DGrid_PG4.Rows(54).Cells(colno).Value = 0
                DGrid_PG4.Rows(55).Cells(0).Value = "FAM DED 1: Multiplier (DED)" : DGrid_PG4.Rows(55).Cells(colno).Value = 0
                DGrid_PG4.Rows(56).Cells(0).Value = "FAM DED 1: Multiplier (OOP)" : DGrid_PG4.Rows(56).Cells(colno).Value = 0
                DGrid_PG4.Rows(57).Cells(0).Value = "FAM DED 1: Text" : DGrid_PG4.Rows(57).Cells(colno).Value = 0
                DGrid_PG4.Rows(58).Cells(0).Value = "IND TIER 1: Label" : DGrid_PG4.Rows(58).Cells(colno).Value = 0
                DGrid_PG4.Rows(59).Cells(0).Value = "ALT DED 1: EE+1 Amount" : DGrid_PG4.Rows(59).Cells(colno).Value = 0
                DGrid_PG4.Rows(60).Cells(0).Value = "ALT DED 1: EE+SP Amount" : DGrid_PG4.Rows(60).Cells(colno).Value = 0
                DGrid_PG4.Rows(61).Cells(0).Value = "ALT DED 1: EE+CH Amount" : DGrid_PG4.Rows(61).Cells(colno).Value = 0
                DGrid_PG4.Rows(62).Cells(0).Value = "FAM DED 1: Prorate Interval" : DGrid_PG4.Rows(62).Cells(colno).Value = 0
                DGrid_PG4.Rows(63).Cells(0).Value = "FAM DED 1: Prorate Interval" : DGrid_PG4.Rows(63).Cells(colno).Value = 0
            End Try
            'MMI PAGE 5

            DGrid_PG5.Rows(0).Cells(0).Value = "Policy" : DGrid_PG5.Rows(0).Cells(colno).Value = txt_Policy.Text
            DGrid_PG5.Rows(1).Cells(0).Value = "Plan Code/Reporting Code/Plan Var" : DGrid_PG5.Rows(1).Cells(colno).Value = Plnnbr
            DGrid_PG5.Rows(2).Cells(0).Value = "Year" : DGrid_PG5.Rows(2).Cells(colno).Value = Startdt & "-" & strEnddt
            DGrid_PG5.Rows(3).Cells(0).Value = "Patient Name" : DGrid_PG5.Rows(3).Cells(colno).Value = strPTName


            If mmiSelect.mmi5Return.mmi5DRows.Count = 1 Then

                'DGrid_PG5.Rows(4).Cells(0).Value = "IND DED 3: Network Type" : DGrid_PG5.Rows(4).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedNtwkTypCd
                'DGrid_PG5.Rows(5).Cells(0).Value = "IND DED 3: Amount" : DGrid_PG5.Rows(5).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedAmt
                'DGrid_PG5.Rows(6).Cells(0).Value = "IND DED 3: Description" : DGrid_PG5.Rows(6).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedSrvcDesc
                'DGrid_PG5.Rows(7).Cells(0).Value = "IND DED 3: Cost Containment" : DGrid_PG5.Rows(7).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedCstCntnCd
                'DGrid_PG5.Rows(8).Cells(0).Value = "IND DED 3: Type Code" : DGrid_PG5.Rows(8).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).covTypCd
                'DGrid_PG5.Rows(9).Cells(0).Value = "IND DED 3: Frequency Code" : DGrid_PG5.Rows(9).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedFreqCd
                'DGrid_PG5.Rows(10).Cells(0).Value = "IND DED 3: Benefit Period" : DGrid_PG5.Rows(10).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedMntPdAmt
                'DGrid_PG5.Rows(11).Cells(0).Value = "IND DED 3: Carry-Over Code" : DGrid_PG5.Rows(11).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedCaroCd
                'DGrid_PG5.Rows(12).Cells(0).Value = "IND DED 3: COB Exclusion" : DGrid_PG5.Rows(12).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedCobCd
                'DGrid_PG5.Rows(13).Cells(0).Value = "IND DED 3: X-Semi Private Rate" : DGrid_PG5.Rows(13).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedSemiPvtRtCd
                'DGrid_PG5.Rows(14).Cells(0).Value = "IND DED 3: Accum Code" : DGrid_PG5.Rows(14).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedAccumCd
                'DGrid_PG5.Rows(15).Cells(0).Value = "IND DED 3: Accum Period" : DGrid_PG5.Rows(15).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedAccumPdAmt
                'DGrid_PG5.Rows(16).Cells(0).Value = "IND DED 3: Maint Amount" : DGrid_PG5.Rows(16).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedMntAmt
                'DGrid_PG5.Rows(17).Cells(0).Value = "IND DED 3: Maint Code" : DGrid_PG5.Rows(17).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedMntCd
                'DGrid_PG5.Rows(18).Cells(0).Value = "IND DED 3: Maint Period" : DGrid_PG5.Rows(18).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedMntPdAmt
                'DGrid_PG5.Rows(19).Cells(0).Value = "IND DED 1: Combined RX" : DGrid_PG5.Rows(19).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(1).combVendCatgyCd
                'DGrid_PG5.Rows(20).Cells(0).Value = "IND DED 1: Combined Vision" : DGrid_PG5.Rows(20).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(3).combVendCatgyCd
                'DGrid_PG5.Rows(21).Cells(0).Value = "IND DED 1: Combined Dental" : DGrid_PG5.Rows(21).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                'DGrid_PG5.Rows(22).Cells(0).Value = "IND DED 1: Combined BH" : DGrid_PG5.Rows(22).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                'DGrid_PG5.Rows(23).Cells(0).Value = "IND DED 1: Combined MD" : DGrid_PG5.Rows(23).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(1).combVendCatgyCd
                DGrid_PG5.Rows(4).Cells(0).Value = "IND DED 3: Network Type" : DGrid_PG5.Rows(4).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedNtwkTypCd
                DGrid_PG5.Rows(5).Cells(0).Value = "IND DED 3: Amount" : DGrid_PG5.Rows(5).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedAmt
                DGrid_PG5.Rows(6).Cells(0).Value = "IND DED 3: Description" : DGrid_PG5.Rows(6).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedSrvcDesc
                DGrid_PG5.Rows(7).Cells(0).Value = "IND DED 3: Cost Containment" : DGrid_PG5.Rows(7).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedCstCntnCd
                DGrid_PG5.Rows(8).Cells(0).Value = "IND DED 3: Type Code" : DGrid_PG5.Rows(8).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).covTypCd
                DGrid_PG5.Rows(9).Cells(0).Value = "IND DED 3: Frequency Code" : DGrid_PG5.Rows(9).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedFreqCd
                DGrid_PG5.Rows(10).Cells(0).Value = "IND DED 3: Benefit Period" : DGrid_PG5.Rows(10).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedMntPdAmt
                DGrid_PG5.Rows(11).Cells(0).Value = "IND DED 3: Carry-Over Code" : DGrid_PG5.Rows(11).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedCaroCd
                DGrid_PG5.Rows(12).Cells(0).Value = "IND DED 3: COB Exclusion" : DGrid_PG5.Rows(12).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedCobCd
                DGrid_PG5.Rows(13).Cells(0).Value = "IND DED 3: X-Semi Private Rate" : DGrid_PG5.Rows(13).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedSemiPvtRtCd
                DGrid_PG5.Rows(14).Cells(0).Value = "IND DED 3: Accum Code" : DGrid_PG5.Rows(14).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedAccumCd
                DGrid_PG5.Rows(15).Cells(0).Value = "IND DED 3: Accum Period" : DGrid_PG5.Rows(15).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedAccumPdAmt
                DGrid_PG5.Rows(16).Cells(0).Value = "IND DED 3: Maint Amount" : DGrid_PG5.Rows(16).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedMntAmt
                DGrid_PG5.Rows(17).Cells(0).Value = "IND DED 3: Maint Code" : DGrid_PG5.Rows(17).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedMntCd
                DGrid_PG5.Rows(18).Cells(0).Value = "IND DED 3: Maint Period" : DGrid_PG5.Rows(18).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(0).dideDedMntPdAmt
                DGrid_PG5.Rows(19).Cells(0).Value = "IND DED 1: Combined RX" : DGrid_PG5.Rows(19).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                DGrid_PG5.Rows(20).Cells(0).Value = "IND DED 1: Combined Vision" : DGrid_PG5.Rows(20).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(2).combVendCatgyCd
                DGrid_PG5.Rows(21).Cells(0).Value = "IND DED 1: Combined Dental" : DGrid_PG5.Rows(21).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                DGrid_PG5.Rows(22).Cells(0).Value = "IND DED 1: Combined BH" : DGrid_PG5.Rows(22).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                DGrid_PG5.Rows(23).Cells(0).Value = "IND DED 1: Combined MD" : DGrid_PG5.Rows(23).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedAmt
                Catch ex As Exception
                End Try
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(1).dfdeDedAmt
                Catch ex As Exception
                End Try
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(2).dfdeDedAmt
                Catch ex As Exception

                End Try
                Try
                    DGrid_PG5.Rows(25).Cells(0).Value = "FAM DED 2: Description" : DGrid_PG5.Rows(25).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMbrDesc
                    DGrid_PG5.Rows(26).Cells(0).Value = "FAM DED 2: Cost Containment" : DGrid_PG5.Rows(26).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedCstCntnCd
                    DGrid_PG5.Rows(27).Cells(0).Value = "FAM DED 2: Type Code" : DGrid_PG5.Rows(27).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).covTypCd
                    DGrid_PG5.Rows(28).Cells(0).Value = "FAM DED 2: Frequency Code" : DGrid_PG5.Rows(28).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedFreqPrdCd
                    DGrid_PG5.Rows(29).Cells(0).Value = "FAM DED 2: Carry-Over Code" : DGrid_PG5.Rows(29).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedCaroCd
                    DGrid_PG5.Rows(30).Cells(0).Value = "FAM DED 2: Multiplier" : DGrid_PG5.Rows(30).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMultFct
                    DGrid_PG5.Rows(31).Cells(0).Value = "FAM DED 2: Individuals" : DGrid_PG5.Rows(31).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMbrCnt
                    DGrid_PG5.Rows(32).Cells(0).Value = "ALT DED 1: EE+1 Amount" : DGrid_PG5.Rows(32).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEePls1Amt
                    DGrid_PG5.Rows(33).Cells(0).Value = "ALT DED 1: EE+SP Amount" : DGrid_PG5.Rows(33).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEeSpoAmt
                    DGrid_PG5.Rows(34).Cells(0).Value = "ALT DED 1: EE+CH Amount" : DGrid_PG5.Rows(34).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEeChrgAmt
                Catch ex As Exception

                End Try

            End If

            '''-----------------------------------------            
            If mmiSelect.mmi5Return.mmi5DRows.Count = 2 Then

                DGrid_PG5.Rows(4).Cells(0).Value = "IND DED 3: Network Type" : DGrid_PG5.Rows(4).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedNtwkTypCd
                DGrid_PG5.Rows(5).Cells(0).Value = "IND DED 3: Amount" : DGrid_PG5.Rows(5).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedAmt
                DGrid_PG5.Rows(6).Cells(0).Value = "IND DED 3: Description" : DGrid_PG5.Rows(6).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedSrvcDesc
                DGrid_PG5.Rows(7).Cells(0).Value = "IND DED 3: Cost Containment" : DGrid_PG5.Rows(7).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedCstCntnCd
                DGrid_PG5.Rows(8).Cells(0).Value = "IND DED 3: Type Code" : DGrid_PG5.Rows(8).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).covTypCd
                DGrid_PG5.Rows(9).Cells(0).Value = "IND DED 3: Frequency Code" : DGrid_PG5.Rows(9).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedFreqCd
                DGrid_PG5.Rows(10).Cells(0).Value = "IND DED 3: Benefit Period" : DGrid_PG5.Rows(10).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedMntPdAmt
                DGrid_PG5.Rows(11).Cells(0).Value = "IND DED 3: Carry-Over Code" : DGrid_PG5.Rows(11).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedCaroCd
                DGrid_PG5.Rows(12).Cells(0).Value = "IND DED 3: COB Exclusion" : DGrid_PG5.Rows(12).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedCobCd
                DGrid_PG5.Rows(13).Cells(0).Value = "IND DED 3: X-Semi Private Rate" : DGrid_PG5.Rows(13).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedSemiPvtRtCd
                DGrid_PG5.Rows(14).Cells(0).Value = "IND DED 3: Accum Code" : DGrid_PG5.Rows(14).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedAccumCd
                DGrid_PG5.Rows(15).Cells(0).Value = "IND DED 3: Accum Period" : DGrid_PG5.Rows(15).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedAccumPdAmt
                DGrid_PG5.Rows(16).Cells(0).Value = "IND DED 3: Maint Amount" : DGrid_PG5.Rows(16).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedMntAmt
                DGrid_PG5.Rows(17).Cells(0).Value = "IND DED 3: Maint Code" : DGrid_PG5.Rows(17).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedMntCd
                DGrid_PG5.Rows(18).Cells(0).Value = "IND DED 3: Maint Period" : DGrid_PG5.Rows(18).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(1).dideDedMntPdAmt
                DGrid_PG5.Rows(19).Cells(0).Value = "IND DED 1: Combined RX" : DGrid_PG5.Rows(19).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(1).combVendCatgyCd
                DGrid_PG5.Rows(20).Cells(0).Value = "IND DED 1: Combined Vision" : DGrid_PG5.Rows(20).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(3).combVendCatgyCd
                DGrid_PG5.Rows(21).Cells(0).Value = "IND DED 1: Combined Dental" : DGrid_PG5.Rows(21).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                DGrid_PG5.Rows(22).Cells(0).Value = "IND DED 1: Combined BH" : DGrid_PG5.Rows(22).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                DGrid_PG5.Rows(23).Cells(0).Value = "IND DED 1: Combined MD" : DGrid_PG5.Rows(23).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(1).combVendCatgyCd
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedAmt
                Catch ex As Exception
                End Try
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(1).dfdeDedAmt
                Catch ex As Exception
                End Try
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(2).dfdeDedAmt
                Catch ex As Exception

                End Try
                Try
                    DGrid_PG5.Rows(25).Cells(0).Value = "FAM DED 2: Description" : DGrid_PG5.Rows(25).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMbrDesc
                    DGrid_PG5.Rows(26).Cells(0).Value = "FAM DED 2: Cost Containment" : DGrid_PG5.Rows(26).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedCstCntnCd
                    DGrid_PG5.Rows(27).Cells(0).Value = "FAM DED 2: Type Code" : DGrid_PG5.Rows(27).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).covTypCd
                    DGrid_PG5.Rows(28).Cells(0).Value = "FAM DED 2: Frequency Code" : DGrid_PG5.Rows(28).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedFreqPrdCd
                    DGrid_PG5.Rows(29).Cells(0).Value = "FAM DED 2: Carry-Over Code" : DGrid_PG5.Rows(29).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedCaroCd
                    DGrid_PG5.Rows(30).Cells(0).Value = "FAM DED 2: Multiplier" : DGrid_PG5.Rows(30).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMultFct
                    DGrid_PG5.Rows(31).Cells(0).Value = "FAM DED 2: Individuals" : DGrid_PG5.Rows(31).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMbrCnt
                    DGrid_PG5.Rows(32).Cells(0).Value = "ALT DED 1: EE+1 Amount" : DGrid_PG5.Rows(32).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEePls1Amt
                    DGrid_PG5.Rows(33).Cells(0).Value = "ALT DED 1: EE+SP Amount" : DGrid_PG5.Rows(33).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEeSpoAmt
                    DGrid_PG5.Rows(34).Cells(0).Value = "ALT DED 1: EE+CH Amount" : DGrid_PG5.Rows(34).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEeChrgAmt
                Catch ex As Exception

                End Try
            End If
            '''-----------------------------------------            

            If mmiSelect.mmi5Return.mmi5DRows.Count = 3 Then               ''''This is original code for MMI Page5

                DGrid_PG5.Rows(4).Cells(0).Value = "IND DED 3: Network Type" : DGrid_PG5.Rows(4).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedNtwkTypCd
                DGrid_PG5.Rows(5).Cells(0).Value = "IND DED 3: Amount" : DGrid_PG5.Rows(5).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedAmt
                DGrid_PG5.Rows(6).Cells(0).Value = "IND DED 3: Description" : DGrid_PG5.Rows(6).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedSrvcDesc
                DGrid_PG5.Rows(7).Cells(0).Value = "IND DED 3: Cost Containment" : DGrid_PG5.Rows(7).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedCstCntnCd
                DGrid_PG5.Rows(8).Cells(0).Value = "IND DED 3: Type Code" : DGrid_PG5.Rows(8).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).covTypCd
                DGrid_PG5.Rows(9).Cells(0).Value = "IND DED 3: Frequency Code" : DGrid_PG5.Rows(9).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedFreqCd
                DGrid_PG5.Rows(10).Cells(0).Value = "IND DED 3: Benefit Period" : DGrid_PG5.Rows(10).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedMntPdAmt
                DGrid_PG5.Rows(11).Cells(0).Value = "IND DED 3: Carry-Over Code" : DGrid_PG5.Rows(11).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedCaroCd
                DGrid_PG5.Rows(12).Cells(0).Value = "IND DED 3: COB Exclusion" : DGrid_PG5.Rows(12).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedCobCd
                DGrid_PG5.Rows(13).Cells(0).Value = "IND DED 3: X-Semi Private Rate" : DGrid_PG5.Rows(13).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedSemiPvtRtCd
                DGrid_PG5.Rows(14).Cells(0).Value = "IND DED 3: Accum Code" : DGrid_PG5.Rows(14).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedAccumCd
                DGrid_PG5.Rows(15).Cells(0).Value = "IND DED 3: Accum Period" : DGrid_PG5.Rows(15).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedAccumPdAmt
                DGrid_PG5.Rows(16).Cells(0).Value = "IND DED 3: Maint Amount" : DGrid_PG5.Rows(16).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedMntAmt
                DGrid_PG5.Rows(17).Cells(0).Value = "IND DED 3: Maint Code" : DGrid_PG5.Rows(17).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedMntCd
                DGrid_PG5.Rows(18).Cells(0).Value = "IND DED 3: Maint Period" : DGrid_PG5.Rows(18).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(2).dideDedMntPdAmt
                DGrid_PG5.Rows(19).Cells(0).Value = "IND DED 1: Combined RX" : DGrid_PG5.Rows(19).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(2).combVendCatgyCd
                DGrid_PG5.Rows(20).Cells(0).Value = "IND DED 1: Combined Vision" : DGrid_PG5.Rows(20).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(4).combVendCatgyCd
                DGrid_PG5.Rows(21).Cells(0).Value = "IND DED 1: Combined Dental" : DGrid_PG5.Rows(21).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(1).combVendCatgyCd
                DGrid_PG5.Rows(22).Cells(0).Value = "IND DED 1: Combined BH" : DGrid_PG5.Rows(22).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(0).combVendCatgyCd
                DGrid_PG5.Rows(23).Cells(0).Value = "IND DED 1: Combined MD" : DGrid_PG5.Rows(23).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(2).combVendCatgyCd
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedAmt
                Catch ex As Exception
                    Try
                        DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(1).dfdeDedAmt
                    Catch ex1 As Exception
                        Try
                            DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(2).dfdeDedAmt
                        Catch ex2 As Exception

                        End Try
                    End Try
                End Try

                Try
                    DGrid_PG5.Rows(25).Cells(0).Value = "FAM DED 2: Description" : DGrid_PG5.Rows(25).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMbrDesc
                    DGrid_PG5.Rows(26).Cells(0).Value = "FAM DED 2: Cost Containment" : DGrid_PG5.Rows(26).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedCstCntnCd
                    DGrid_PG5.Rows(27).Cells(0).Value = "FAM DED 2: Type Code" : DGrid_PG5.Rows(27).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).covTypCd
                    DGrid_PG5.Rows(28).Cells(0).Value = "FAM DED 2: Frequency Code" : DGrid_PG5.Rows(28).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedFreqPrdCd
                    DGrid_PG5.Rows(29).Cells(0).Value = "FAM DED 2: Carry-Over Code" : DGrid_PG5.Rows(29).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedCaroCd
                    DGrid_PG5.Rows(30).Cells(0).Value = "FAM DED 2: Multiplier" : DGrid_PG5.Rows(30).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMultFct
                    DGrid_PG5.Rows(31).Cells(0).Value = "FAM DED 2: Individuals" : DGrid_PG5.Rows(31).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMbrCnt
                    DGrid_PG5.Rows(32).Cells(0).Value = "ALT DED 1: EE+1 Amount" : DGrid_PG5.Rows(32).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEePls1Amt
                    DGrid_PG5.Rows(33).Cells(0).Value = "ALT DED 1: EE+SP Amount" : DGrid_PG5.Rows(33).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEeSpoAmt
                    DGrid_PG5.Rows(34).Cells(0).Value = "ALT DED 1: EE+CH Amount" : DGrid_PG5.Rows(34).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEeChrgAmt
                Catch ex As Exception

                End Try


            End If

            If mmiSelect.mmi5Return.mmi5DRows.Count = 4 Or mmiSelect.mmi5Return.mmi5DRows.Count = 5 Then

                DGrid_PG5.Rows(4).Cells(0).Value = "IND DED 3: Network Type" : DGrid_PG5.Rows(4).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedNtwkTypCd
                DGrid_PG5.Rows(5).Cells(0).Value = "IND DED 3: Amount" : DGrid_PG5.Rows(5).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedAmt
                DGrid_PG5.Rows(6).Cells(0).Value = "IND DED 3: Description" : DGrid_PG5.Rows(6).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedSrvcDesc
                DGrid_PG5.Rows(7).Cells(0).Value = "IND DED 3: Cost Containment" : DGrid_PG5.Rows(7).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedCstCntnCd
                DGrid_PG5.Rows(8).Cells(0).Value = "IND DED 3: Type Code" : DGrid_PG5.Rows(8).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).covTypCd
                DGrid_PG5.Rows(9).Cells(0).Value = "IND DED 3: Frequency Code" : DGrid_PG5.Rows(9).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedFreqCd
                DGrid_PG5.Rows(10).Cells(0).Value = "IND DED 3: Benefit Period" : DGrid_PG5.Rows(10).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedMntPdAmt
                DGrid_PG5.Rows(11).Cells(0).Value = "IND DED 3: Carry-Over Code" : DGrid_PG5.Rows(11).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedCaroCd
                DGrid_PG5.Rows(12).Cells(0).Value = "IND DED 3: COB Exclusion" : DGrid_PG5.Rows(12).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedCobCd
                DGrid_PG5.Rows(13).Cells(0).Value = "IND DED 3: X-Semi Private Rate" : DGrid_PG5.Rows(13).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedSemiPvtRtCd
                DGrid_PG5.Rows(14).Cells(0).Value = "IND DED 3: Accum Code" : DGrid_PG5.Rows(14).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedAccumCd
                DGrid_PG5.Rows(15).Cells(0).Value = "IND DED 3: Accum Period" : DGrid_PG5.Rows(15).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedAccumPdAmt
                DGrid_PG5.Rows(16).Cells(0).Value = "IND DED 3: Maint Amount" : DGrid_PG5.Rows(16).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedMntAmt
                DGrid_PG5.Rows(17).Cells(0).Value = "IND DED 3: Maint Code" : DGrid_PG5.Rows(17).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedMntCd
                DGrid_PG5.Rows(18).Cells(0).Value = "IND DED 3: Maint Period" : DGrid_PG5.Rows(18).Cells(colno).Value = mmiSelect.mmi5Return.mmi5DRows(3).dideDedMntPdAmt
                DGrid_PG5.Rows(19).Cells(0).Value = "IND DED 1: Combined RX" : DGrid_PG5.Rows(19).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(3).combVendCatgyCd
                DGrid_PG5.Rows(20).Cells(0).Value = "IND DED 1: Combined Vision" : DGrid_PG5.Rows(20).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(4).combVendCatgyCd
                DGrid_PG5.Rows(21).Cells(0).Value = "IND DED 1: Combined Dental" : DGrid_PG5.Rows(21).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(2).combVendCatgyCd
                DGrid_PG5.Rows(22).Cells(0).Value = "IND DED 1: Combined BH" : DGrid_PG5.Rows(22).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(1).combVendCatgyCd
                DGrid_PG5.Rows(23).Cells(0).Value = "IND DED 1: Combined MD" : DGrid_PG5.Rows(23).Cells(colno).Value = mmiSelect.mmi5Return.mmi5BRows(3).combVendCatgyCd
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedAmt
                Catch ex As Exception
                End Try

                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(1).dfdeDedAmt
                Catch ex As Exception

                End Try
                Try
                    DGrid_PG5.Rows(24).Cells(0).Value = "FAM DED 2: Amount" : DGrid_PG5.Rows(24).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(2).dfdeDedAmt
                Catch ex As Exception

                End Try
                DGrid_PG5.Rows(25).Cells(0).Value = "FAM DED 2: Description" : DGrid_PG5.Rows(25).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMbrDesc
                DGrid_PG5.Rows(26).Cells(0).Value = "FAM DED 2: Cost Containment" : DGrid_PG5.Rows(26).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedCstCntnCd
                DGrid_PG5.Rows(27).Cells(0).Value = "FAM DED 2: Type Code" : DGrid_PG5.Rows(27).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).covTypCd
                DGrid_PG5.Rows(28).Cells(0).Value = "FAM DED 2: Frequency Code" : DGrid_PG5.Rows(28).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedFreqPrdCd
                DGrid_PG5.Rows(29).Cells(0).Value = "FAM DED 2: Carry-Over Code" : DGrid_PG5.Rows(29).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedCaroCd
                DGrid_PG5.Rows(30).Cells(0).Value = "FAM DED 2: Multiplier" : DGrid_PG5.Rows(30).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMultFct
                DGrid_PG5.Rows(31).Cells(0).Value = "FAM DED 2: Individuals" : DGrid_PG5.Rows(31).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedMbrCnt
                DGrid_PG5.Rows(32).Cells(0).Value = "ALT DED 1: EE+1 Amount" : DGrid_PG5.Rows(32).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEePls1Amt
                DGrid_PG5.Rows(33).Cells(0).Value = "ALT DED 1: EE+SP Amount" : DGrid_PG5.Rows(33).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEeSpoAmt
                DGrid_PG5.Rows(34).Cells(0).Value = "ALT DED 1: EE+CH Amount" : DGrid_PG5.Rows(34).Cells(colno).Value = mmiSelect.mmi5Return.mmi5CRows(0).dfdeDedEeChrgAmt

            End If

            'MMI PAGE 10

            DGrid_PG10.Rows(0).Cells(0).Value = "Policy" : DGrid_PG10.Rows(0).Cells(colno).Value = txt_Policy.Text
            DGrid_PG10.Rows(1).Cells(0).Value = "Plan Code/Reporting Code/Plan Var" : DGrid_PG10.Rows(1).Cells(colno).Value = Plnnbr
            DGrid_PG10.Rows(2).Cells(0).Value = "Year" : DGrid_PG10.Rows(2).Cells(colno).Value = Startdt & "-" & strEnddt
            DGrid_PG10.Rows(3).Cells(0).Value = "Patient Name" : DGrid_PG10.Rows(3).Cells(colno).Value = strPTName
            DGrid_PG10.Rows(4).Cells(0).Value = "IND COIN: Code" : DGrid_PG10.Rows(4).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsAccumCd
            DGrid_PG10.Rows(5).Cells(0).Value = "IND COIN: Carry-Over Code" : DGrid_PG10.Rows(5).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsPrdCd
            DGrid_PG10.Rows(6).Cells(0).Value = "IND COIN: Salary Type" : DGrid_PG10.Rows(5).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsSalIndvTypCd
            DGrid_PG10.Rows(7).Cells(0).Value = "IND COIN: Min Percent" : DGrid_PG10.Rows(7).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsIndvMinPct
            DGrid_PG10.Rows(8).Cells(0).Value = "IND COIN: DED Ind" : DGrid_PG10.Rows(8).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsDedTypCd
            DGrid_PG10.Rows(9).Cells(0).Value = "IND COIN: COB Exclusion" : DGrid_PG10.Rows(9).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsCobXcldInd
            DGrid_PG10.Rows(10).Cells(0).Value = "IND COIN: Mid-Year Change Amount" : DGrid_PG10.Rows(10).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsChgAmt
            DGrid_PG10.Rows(11).Cells(0).Value = "Carryover" : DGrid_PG10.Rows(11).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsPrdCd
            DGrid_PG10.Rows(12).Cells(0).Value = "IND COIN: Mid-Year Change End Date" : DGrid_PG10.Rows(12).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsEndDt
            DGrid_PG10.Rows(13).Cells(0).Value = "FAM COIN: Salary Type" : DGrid_PG10.Rows(13).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsSalFamTypCd
            DGrid_PG10.Rows(14).Cells(0).Value = "FAM COIN: Salary Multiplier" : DGrid_PG10.Rows(14).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsSalFamMultFct
            DGrid_PG10.Rows(15).Cells(0).Value = "FAM COIN: Number of Members" : DGrid_PG10.Rows(15).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsIndvNbr
            DGrid_PG10.Rows(16).Cells(0).Value = "INN COIN: Combined RX" : DGrid_PG10.Rows(16).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(3).combNewCoinsInd
            DGrid_PG10.Rows(17).Cells(0).Value = "OON COIN: Combined RX" : DGrid_PG10.Rows(17).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(3).combDualOopInd
            DGrid_PG10.Rows(18).Cells(0).Value = "T1 COIN: Combined RX" : DGrid_PG10.Rows(18).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(3).combTier1Ind
            DGrid_PG10.Rows(19).Cells(0).Value = "INN COIN: Combined Vision" : DGrid_PG10.Rows(19).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(4).combNewCoinsInd
            DGrid_PG10.Rows(20).Cells(0).Value = "OON COIN: Combined Vision" : DGrid_PG10.Rows(20).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(4).combDualOopInd
            DGrid_PG10.Rows(21).Cells(0).Value = "T1 COIN: Combined Vision" : DGrid_PG10.Rows(21).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(4).combTier1Ind
            DGrid_PG10.Rows(22).Cells(0).Value = "INN COIN: Combined Dental" : DGrid_PG10.Rows(22).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(1).combNewCoinsInd
            DGrid_PG10.Rows(23).Cells(0).Value = "OON COIN: Combined Dental" : DGrid_PG10.Rows(23).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(1).combDualOopInd
            DGrid_PG10.Rows(24).Cells(0).Value = "T1 COIN: Combined Dental" : DGrid_PG10.Rows(24).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(1).combTier1Ind
            DGrid_PG10.Rows(25).Cells(0).Value = "INN COIN: Combined BH" : DGrid_PG10.Rows(25).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(0).combNewCoinsInd
            DGrid_PG10.Rows(26).Cells(0).Value = "OON COIN: Combined BH" : DGrid_PG10.Rows(26).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(0).combDualOopInd
            DGrid_PG10.Rows(27).Cells(0).Value = "T1 COIN: Combined BH" : DGrid_PG10.Rows(27).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(0).combTier1Ind
            DGrid_PG10.Rows(28).Cells(0).Value = "INN COIN: Combined MD" : DGrid_PG10.Rows(28).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(2).combNewCoinsInd
            DGrid_PG10.Rows(29).Cells(0).Value = "OON COIN: Combined MD" : DGrid_PG10.Rows(29).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(2).combDualOopInd
            DGrid_PG10.Rows(30).Cells(0).Value = "T1 COIN: Combined MD" : DGrid_PG10.Rows(30).Cells(colno).Value = mmiSelect.mmi10Return.mmi10CRows(2).combTier1Ind
            DGrid_PG10.Rows(31).Cells(0).Value = "NB Suppression" : DGrid_PG10.Rows(31).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).nbSprsInd
            DGrid_PG10.Rows(32).Cells(0).Value = "IND COIN: INN Percent" : DGrid_PG10.Rows(32).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsIndvMaxPct
            DGrid_PG10.Rows(33).Cells(0).Value = "IND COIN: INN Amount" : DGrid_PG10.Rows(33).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsAmt
            DGrid_PG10.Rows(34).Cells(0).Value = "IND COIN: OON Percent" : DGrid_PG10.Rows(34).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).dualOopIndvPct
            DGrid_PG10.Rows(35).Cells(0).Value = "IND COIN: OON Amount" : DGrid_PG10.Rows(35).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombIndvAmt
            DGrid_PG10.Rows(36).Cells(0).Value = "FAM COIN: INN Amount" : DGrid_PG10.Rows(36).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).famNewCoinsAmt
            DGrid_PG10.Rows(37).Cells(0).Value = "FAM COIN: OON Amount" : DGrid_PG10.Rows(37).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombFamAmt
            DGrid_PG10.Rows(38).Cells(0).Value = "IND COIN: T1 Amount" : DGrid_PG10.Rows(38).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).tier1NewCoinsAmt
            DGrid_PG10.Rows(39).Cells(0).Value = "FAM COIN: T1 Amount" : DGrid_PG10.Rows(39).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).tier1FamNewCoinsAmt
            DGrid_PG10.Rows(40).Cells(0).Value = "ALT INN COIN: EE+1 Amount" : DGrid_PG10.Rows(40).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopInNtwkEePls1Amt
            DGrid_PG10.Rows(41).Cells(0).Value = "ALT OON COIN: EE+1 Amount" : DGrid_PG10.Rows(41).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombEePls1Amt
            DGrid_PG10.Rows(42).Cells(0).Value = "ALT INN COIN: EE+SP Amount" : DGrid_PG10.Rows(42).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopInNtwkEeSpoAmt
            DGrid_PG10.Rows(43).Cells(0).Value = "ALT OON COIN: EE+SP Amount" : DGrid_PG10.Rows(43).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombEeSpoAmt
            DGrid_PG10.Rows(44).Cells(0).Value = "ALT INN COIN: EE+CH Amount" : DGrid_PG10.Rows(44).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopInNtwkEeChrgAmt
            DGrid_PG10.Rows(45).Cells(0).Value = "ALT OON COIN: EE+CH Amount" : DGrid_PG10.Rows(45).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombEeChrgAmt
            DGrid_PG10.Rows(46).Cells(0).Value = "CROSS-APPLY OOP IND" : DGrid_PG10.Rows(46).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombNbrCd

            'MMI PAGE 11          
            'DGrid_PG11.Rows.Add("Core INN EE + 1", mmiSelect.mmi11Return.mmi11BRows(1).eePls1Amt)
            'DGrid_PG11.Rows.Add("Core INN EE + SP", mmiSelect.mmi11Return.mmi11BRows(1).eePlsSpAmt)
            'DGrid_PG11.Rows.Add("Core INN EE + CH", mmiSelect.mmi11Return.mmi11BRows(1).eePlsChAmt)
            'DGrid_PG11.Rows.Add("Core INN EE + CH", mmiSelect.mmi11Return.mmi11BRows(1))

            '--------MMI OVERVIEW------            
            Dim PTAge As String = DGridCEI.Rows(0).Cells(2).Value
            PTAge = PTAge.Replace("-", "")
            Dim dtNum As Integer
            If PTAge = "" Then

                Exit Sub
            End If
            If Mid(PTAge, 5, 2) > 30 And Mid(PTAge, 5, 2) <= 99 Then
                dtNum = 19
            Else
                dtNum = 20
            End If

            Dim dtPTAge As Date = Mid(PTAge, 1, 2) & "/" & Mid(PTAge, 3, 2) & "/" & dtNum & Mid(PTAge, 5, 2)
            dtPTAge = Format(dtPTAge, "MM/dd/yyyy")

            Dim dtAge = Now.Year - dtPTAge.Year

            Dim intAge = (Int(dtAge) / 365)

            DGridOverview.Rows(0).Cells(0).Value = "Member Name" : DGridOverview.Rows(0).Cells(colno).Value = DGridCEI.Rows(0).Cells(0).Value & " " & DGridMInfo.Rows(0).Cells(0).Value
            DGridOverview.Rows(1).Cells(0).Value = "Patient address" : DGridOverview.Rows(1).Cells(colno).Value = DGridMInfo.Rows(0).Cells(1).Value & " " & DGridMInfo.Rows(1).Cells(1).Value
            'DGridOverview.Rows(2).Cells(0).Value = "Patient Name" : DGridOverview.Rows(2).Cells(colno).Value = DGridCEI.Rows(0).Cells(0).Value & " " & DGridMInfo.Rows(0).Cells(0).Value
            DGridOverview.Rows(3).Cells(0).Value = "Patient Age" : DGridOverview.Rows(3).Cells(colno).Value = dtAge
            DGridOverview.Rows(4).Cells(0).Value = "Policy" : DGridOverview.Rows(4).Cells(colno).Value = txt_Policy.Text
            DGridOverview.Rows(5).Cells(0).Value = "Plan Code/Reporting Code/Plan Var" : DGridOverview.Rows(5).Cells(colno).Value = Plnnbr
            DGridOverview.Rows(6).Cells(0).Value = "Year" : DGridOverview.Rows(6).Cells(colno).Value = Startdt & "-" & strEnddt
            DGridOverview.Rows(7).Cells(0).Value = "Patient Name/Relationship" : DGridOverview.Rows(7).Cells(colno).Value = ""
            DGridOverview.Rows(8).Cells(0).Value = "Medicare Indicator" : DGridOverview.Rows(8).Cells(colno).Value = ""
            'PG1
            DGridOverview.Rows(9).Cells(0).Value = "MMI Page1"
            DGridOverview.Rows(10).Cells(0).Value = "Comb RX Vendor ID" : DGridOverview.Rows(10).Cells(colno).Value = mmiSelect.mmi2Return.mmi2BRows(3).combVendId
            DGridOverview.Rows(11).Cells(0).Value = "Embedded/Non-Embedded" : DGridOverview.Rows(11).Cells(colno).Value = mmiSelect.mmi1Return.mmi1ARows(0).nonEmbdDedCd
            'PG4
            DGridOverview.Rows(12).Cells(0).Value = "MMI Page4"
            Try
                DGridOverview.Rows(13).Cells(0).Value = "(DED1) Individual Deductible Amount" : DGridOverview.Rows(13).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).indDedAmt
                DGridOverview.Rows(14).Cells(0).Value = "(DED2) Individual Deductible Amount" : DGridOverview.Rows(14).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(1).indDedAmt
                DGridOverview.Rows(15).Cells(0).Value = "(DED3) Family Deductible Amount" : DGridOverview.Rows(15).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).famDedAmt
                DGridOverview.Rows(16).Cells(0).Value = "EE + 1 Ded 1" : DGridOverview.Rows(16).Cells(colno).Value = mmiSelect.mmi4Return.mmi4CRows(0).dedEePls1Amt
                DGridOverview.Rows(17).Cells(0).Value = "IND DED 1: Description" : DGridOverview.Rows(17).Cells(colno).Value = mmiSelect.mmi4Return.mmi4DRows(0).dedSrvcDesc
            Catch ex As Exception

            End Try

            'mmiSelect.mmi4Return.mmi4DRows(0).dedSrvcDesc
            'PG5
            DGridOverview.Rows(18).Cells(0).Value = "MMI Page5"
            DGridOverview.Rows(19).Cells(0).Value = "(DED2) Family Deductible 2 Description" : DGridOverview.Rows(19).Cells(colno).Value = DGrid_PG5.Rows(25).Cells(colno).Value
            DGridOverview.Rows(20).Cells(0).Value = "(DED3) Family Deductible Amount" : DGridOverview.Rows(20).Cells(colno).Value = DGrid_PG5.Rows(24).Cells(colno).Value
            DGridOverview.Rows(21).Cells(0).Value = "EE + 1 Ded 2" : DGridOverview.Rows(21).Cells(colno).Value = DGrid_PG5.Rows(32).Cells(colno).Value
            'pg10
            DGridOverview.Rows(22).Cells(0).Value = "MMI Page10"
            DGridOverview.Rows(23).Cells(0).Value = "Individual New Coinsurance Amount" : DGridOverview.Rows(23).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).newCoinsAmt
            DGridOverview.Rows(24).Cells(0).Value = "Individual New Coinsurance Dual Out-of-Pocket Amount" : DGridOverview.Rows(24).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombIndvAmt
            DGridOverview.Rows(25).Cells(0).Value = "Individual New Minimum Coinsurance Percent" : DGridOverview.Rows(25).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).dualOopIndvPct
            DGridOverview.Rows(26).Cells(0).Value = "Family New Coinsurance Amount" : DGridOverview.Rows(26).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).famNewCoinsAmt
            DGridOverview.Rows(27).Cells(0).Value = "FAM NEW COINS DUAL OOP AMT" : DGridOverview.Rows(27).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombFamAmt
            DGridOverview.Rows(28).Cells(0).Value = "EE + 1 INN OOP AMT" : DGridOverview.Rows(28).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopInNtwkEePls1Amt
            DGridOverview.Rows(29).Cells(0).Value = "EE + 1 DUAL OOP AMT" : DGridOverview.Rows(29).Cells(colno).Value = mmiSelect.mmi10Return.mmi10ARows(0).oopCombEePls1Amt    '' need to check 
            DGridOverview.Rows(31).Cells(0).Value = "Deductible to OOP"
            DGridOverview.Rows(32).Cells(0).Value = "Copay to OOP"
            '''Closing mmi overview 
            Exit For
        Next


        strDedToOOP = DGrid_PG10.Rows(4).Cells(1).Value
        If strDedToOOP = "2" Then
            strDedToOOP = "Y"
        Else
            strDedToOOP = "N"
        End If
        DGridOverview.Rows(31).Cells(1).Value = strDedToOOP


        ' Form1.RichTextBox1.AppendText("Completed to gather data for MMI Page 5")

        mmiDetails.Clear()

        'Catch ex As Exception

        'End Try

    End Sub
    Function Setup_Headers(ByVal intSheetType As Integer, ByVal strPtLast As String,
                    ByVal strPayloc As String, ByVal strMyChoice As String, intChoice As Integer)

        tblOOP.Columns.Add("A", "A")
        tblOOP.Columns.Add("B", "B")
        tblOOP.Columns.Add("C", "C")
        tblOOP.Columns.Add("D", "D")
        tblOOP.Columns.Add("E", "E")
        tblOOP.Rows.Add(5)
        Select Case intChoice
            Case 1
                OOP_Headers()
            Case 2
                Medicare_Headers()
            Case 3
                MedicarePPO_Headers()
            Case 4
                OI_Headers()
            Case 5
                OIMEDI_Headers()
        End Select

        'tblOOP.Rows.Add() : tblOOP.Rows.Add() : tblOOP.Rows.Add() : tblOOP.Rows.Add() : tblOOP.Rows.Add()
        tblOOP.Rows(0).Cells(1).Value = "First Name:" : tblOOP.Rows(0).Cells(1).Style.ForeColor = Color.Black
        tblOOP.Rows(0).Cells(2).Value = pName
        tblOOP.Rows(1).Cells(1).Value = "Last Name:" : tblOOP.Rows(1).Cells(1).Style.ForeColor = Color.Black
        tblOOP.Rows(1).Cells(2).Value = DGridMInfo.Rows(0).Cells(0).Value
        tblOOP.Rows(0).Cells(3).Value = "Member ID:" : tblOOP.Rows(0).Cells(3).Style.ForeColor = Color.Black
        tblOOP.Rows(1).Cells(3).Value = "XXX-XX-" & Mid(txt_SSN.Text, 7, 4)
        tblOOP.Rows(0).Cells(11).Value = "Year" : tblOOP.Rows(1).Cells(11).Style.ForeColor = Color.Black
        tblOOP.Rows(0).Cells(12).Value = yearList.Text
        tblOOP.Rows(1).Cells(12).Value = DGridMInfo.Rows(0).Cells(2).Value

        strDed = strDedToOOP

        If blnCopays = True Then
            strCopay = "Y"
        Else
            strCopay = "N"
        End If

        If intChoice <> 2 Then
            tblOOP.Rows(0).Cells(6).Value = "Copay applies To OOP?"
            tblOOP.Rows(0).Cells(9).Value = strCopay                                ''pending to apply by Sanjeet 
            tblOOP.Rows(1).Cells(6).Value = "INN ded applies To OOP?"
            tblOOP.Rows(1).Cells(9).Value = strDed
            tblOOP.Rows(2).Cells(6).Value = "OON ded applies To OOP?"
            tblOOP.Rows(2).Cells(9).Value = strDed
        Else
            tblOOP.Rows(0).Cells(6).Value = "Deductible applies To OOP?"
            tblOOP.Rows(0).Cells(9).Value = "strDed"
        End If
        For qrow = 0 To tblOOP.Columns.Count - 1
            tblOOP.Rows(4).Cells(qrow).Style.BackColor = Color.Gray
        Next
        tblOOP.Rows(4).Cells(0).Value = "From"
        tblOOP.Rows(4).Cells(1).Value = "Thru"
        tblOOP.Rows(4).Cells(2).Value = "Provider"
        tblOOP.Rows(4).Cells(3).Value = "INN?"
        tblOOP.Rows(4).Cells(4).Value = "Total Charge"


        'tblMHI.Rows(icnt).Cells(j).Style.BackColor = Color.Gray

        'Range("E6", Range("M6").End(xlDown)).NumberFormat = "$#,###.00;[Red]-$#,###.00;$0.00"
        'Range("H6", Range("H6").End(xlDown)).Interior.Color = RGB(151, 255, 255)
        'Range("I6", Range("I6").End(xlDown)).Interior.Color = RGB(151, 255, 255)
        'Range("K6", Range("K6").End(xlDown)).Interior.Color = RGB(151, 255, 255)
        'Range("L6", Range("L6").End(xlDown)).Interior.Color = RGB(151, 255, 255)

    End Function

    Sub OOP_Headers()
        tblOOP.Columns.Add("F", "F")
        tblOOP.Columns.Add("G", "G")
        tblOOP.Columns.Add("H", "H")
        tblOOP.Columns.Add("I", "I")
        tblOOP.Columns.Add("J", "J")
        tblOOP.Columns.Add("K", "K")
        tblOOP.Columns.Add("L", "L")
        tblOOP.Columns.Add("M", "M")
        tblOOP.Columns.Add("N", "N")
        tblOOP.Columns.Add("O", "O")
        tblOOP.Columns.Add("P", "P")

        tblOOP.Rows(4).Cells(5).Value = "Neg Rate"
        tblOOP.Rows(4).Cells(6).Value = "Copay"
        tblOOP.Rows(4).Cells(7).Value = "INN Deduct"
        tblOOP.Rows(4).Cells(8).Value = "OON Deduct"
        tblOOP.Rows(4).Cells(9).Value = "Paid"
        tblOOP.Rows(4).Cells(10).Value = "INN OOP"
        tblOOP.Rows(4).Cells(11).Value = "OON  OOP"
        tblOOP.Rows(4).Cells(12).Value = "Not Covered"
        tblOOP.Rows(4).Cells(13).Value = "Comments"
        tblOOP.Rows(4).Cells(14).Value = "ICN"
        tblOOP.Rows(4).Cells(15).Value = "Process Date"

    End Sub
    Sub Medicare_Headers()

        tblOOP.Columns.Add("F", "F")
        tblOOP.Columns.Add("G", "G")
        tblOOP.Columns.Add("H", "H")
        tblOOP.Columns.Add("I", "I")
        tblOOP.Columns.Add("J", "J")
        tblOOP.Columns.Add("K", "K")
        tblOOP.Columns.Add("L", "L")
        tblOOP.Columns.Add("M", "M")
        tblOOP.Columns.Add("N", "N")
        tblOOP.Columns.Add("O", "O")

        tblOOP.Rows(4).Cells(5).Value = "Medicare_Approved"
        tblOOP.Rows(4).Cells(6).Value = "UHC_Allowed"
        tblOOP.Rows(4).Cells(7).Value = "Deductible"
        tblOOP.Rows(4).Cells(8).Value = "Medicare_Paid"
        tblOOP.Rows(4).Cells(9).Value = "UHC Paid"
        tblOOP.Rows(4).Cells(10).Value = "OOP"
        tblOOP.Rows(4).Cells(11).Value = "Not_Covered"
        tblOOP.Rows(4).Cells(12).Value = "Comments"
        tblOOP.Rows(4).Cells(13).Value = "ICN"
        tblOOP.Rows(4).Cells(14).Value = "Process Date"


    End Sub
    Sub MedicarePPO_Headers()
        tblOOP.Columns.Add("F", "F")
        tblOOP.Columns.Add("G", "G")
        tblOOP.Columns.Add("H", "H")
        tblOOP.Columns.Add("I", "I")
        tblOOP.Columns.Add("J", "J")
        tblOOP.Columns.Add("K", "K")
        tblOOP.Columns.Add("L", "L")
        tblOOP.Columns.Add("M", "M")
        tblOOP.Columns.Add("N", "N")
        tblOOP.Columns.Add("O", "O")
        tblOOP.Columns.Add("P", "P")
        tblOOP.Columns.Add("Q", "Q")

        tblOOP.Rows(4).Cells(5).Value = "Medicare_Approved"
        tblOOP.Rows(4).Cells(6).Value = "UHC_Allowed"
        tblOOP.Rows(4).Cells(7).Value = "Copay"
        tblOOP.Rows(4).Cells(8).Value = "INN_Ded"
        tblOOP.Rows(4).Cells(9).Value = "OON_Ded"
        tblOOP.Rows(4).Cells(10).Value = "Medicare_Paid"
        tblOOP.Rows(4).Cells(11).Value = "INN_OOP"
        tblOOP.Rows(4).Cells(12).Value = "Not_Cov"
        tblOOP.Rows(4).Cells(13).Value = "Comments"
        tblOOP.Rows(4).Cells(14).Value = "ICN"
        tblOOP.Rows(4).Cells(15).Value = "Process Date"

    End Sub
    Sub OI_Headers()

        tblOOP.Columns.Add("F", "F")
        tblOOP.Columns.Add("G", "G")
        tblOOP.Columns.Add("H", "H")
        tblOOP.Columns.Add("I", "I")
        tblOOP.Columns.Add("J", "J")
        tblOOP.Columns.Add("K", "K")
        tblOOP.Columns.Add("L", "L")
        tblOOP.Columns.Add("M", "M")
        tblOOP.Columns.Add("N", "N")
        tblOOP.Columns.Add("O", "O")
        tblOOP.Columns.Add("P", "P")
        tblOOP.Columns.Add("Q", "Q")
        tblOOP.Columns.Add("R", "R")

        tblOOP.Rows(4).Cells(5).Value = "Primary Approved"
        tblOOP.Rows(4).Cells(6).Value = "UHC Allowed"
        tblOOP.Rows(4).Cells(7).Value = "Copay"
        tblOOP.Rows(4).Cells(8).Value = "INN Ded"
        tblOOP.Rows(4).Cells(9).Value = "OON Ded"
        tblOOP.Rows(4).Cells(10).Value = "Primary Paid"
        tblOOP.Rows(4).Cells(11).Value = "UHC Paid"
        tblOOP.Rows(4).Cells(12).Value = "INN OOP"
        tblOOP.Rows(4).Cells(13).Value = "OON OOP"
        tblOOP.Rows(4).Cells(14).Value = "Not Cov"
        tblOOP.Rows(4).Cells(15).Value = "Comments"
        tblOOP.Rows(4).Cells(16).Value = "ICN"
        tblOOP.Rows(4).Cells(17).Value = "Process Date"

    End Sub
    Sub OIMEDI_Headers()

        tblOOP.Columns.Add("F", "F")
        tblOOP.Columns.Add("G", "G")
        tblOOP.Columns.Add("H", "H")
        tblOOP.Columns.Add("I", "I")
        tblOOP.Columns.Add("J", "J")
        tblOOP.Columns.Add("K", "K")
        tblOOP.Columns.Add("L", "L")
        tblOOP.Columns.Add("M", "M")
        tblOOP.Columns.Add("N", "N")
        tblOOP.Columns.Add("O", "O")
        tblOOP.Columns.Add("P", "P")
        tblOOP.Columns.Add("Q", "Q")
        tblOOP.Columns.Add("R", "R")
        tblOOP.Columns.Add("S", "S")
        tblOOP.Columns.Add("T", "T")


        tblOOP.Rows(4).Cells(5).Value = "Medicare_Approved"
        tblOOP.Rows(4).Cells(6).Value = "Secondary_Approved"
        tblOOP.Rows(4).Cells(7).Value = "UHC_Approved"
        tblOOP.Rows(4).Cells(8).Value = "Copay"
        tblOOP.Rows(4).Cells(9).Value = "INN_Ded"
        tblOOP.Rows(4).Cells(10).Value = "OON_Ded"
        tblOOP.Rows(4).Cells(11).Value = "Medicare_Paid"
        tblOOP.Rows(4).Cells(12).Value = "Secondary_Paid"
        tblOOP.Rows(4).Cells(13).Value = "UHC_Paid"
        tblOOP.Rows(4).Cells(14).Value = "INN_OOP"
        tblOOP.Rows(4).Cells(15).Value = "OON_OOP"
        tblOOP.Rows(4).Cells(16).Value = "Not_Covered"
        tblOOP.Rows(4).Cells(17).Value = "Comments"
        tblOOP.Rows(4).Cells(18).Value = "ICN"
        tblOOP.Rows(4).Cells(19).Value = "Process_Date"

    End Sub

    Sub Claim_Combine()

        tblMHI.Refresh()

        Dim StrICN, StrDraft As String, dtFrstDos As Date, dtLastDos As Date, dtProcDate As Date
        Dim curCovered As Double, curNotCovd As Double, curCharge As Double
        Dim strProv As String, curCopay As Double, curPaid As Double
        Dim curInnDed As Double, curOonDed As Double, curInnOop As Double
        Dim curOonOop As Double, curMedCov As Double, curMedPaid As Double
        Dim curOICov As Double, curOIPaid As Double, strSvcCode As String
        Dim DedCode As String, blnInnNet As Boolean, blnQuestion As Boolean, OVQuestion As Boolean
        Dim strOvCode As String, strRemark As String, strPayee As String
        Dim blnTiered As Boolean

        Dim rCount As Integer = 0
        Dim rCount1 As Integer = 0
        Dim OOPRcnt As Long = 5
        'Range("A5").Select
        For rCount = 0 To tblMHI.Rows.Count - 1
            Threading.Thread.Sleep(10)
            blnInnNet = False
            blnQuestion = False
            OVQuestion = False
            Do
                Threading.Thread.Sleep(10)
                If IsDBNull(tblMHI.Rows(rCount1).Cells(38).Value) Then
                    rCount1 = rCount1 + 1
                End If
                If IsDBNull(tblMHI.Rows(rCount1).Cells(38).Value) Then
                    Exit Do
                End If

                strSvcCode = Trim(tblMHI.Rows(rCount1).Cells(2).Value)
                StrDraft = Trim(tblMHI.Rows(rCount1).Cells(25).Value)
                dtFrstDos = Format(CDate(tblMHI.Rows(rCount1).Cells(0).Value), "MM/dd/yyyy")
                dtLastDos = Format(CDate(tblMHI.Rows(rCount1).Cells(1).Value), "MM/dd/yyyy")
                StrICN = tblMHI.Rows(rCount1).Cells(30).Value
                Try
                    strProv = Trim(tblMHI.Rows(rCount1).Cells(51).Value)
                Catch ex As Exception
                End Try

                Try
                    If isNullOrEmpty(tblMHI.Rows(rCount1).Cells(26).Value) Then
                    Else
                        dtProcDate = Format(CDate(tblMHI.Rows(rCount1).Cells(26).Value), "MM/dd/yyyy")
                    End If
                Catch ex As Exception

                End Try



                'dtProcDate = Format(CDate(tblMHI.Rows(rCount1).Cells(26).Value), "MM/dd/yyyy")


                strOvCode = Trim(tblMHI.Rows(rCount1).Cells(5).Value)
                strRemark = Trim(tblMHI.Rows(rCount1).Cells(8).Value)

                'If isNullOrEmpty(tblMHI.Rows(rCount1).Cells(6).Value) Then
                'Else
                '    strPayee = tblMHI.Rows(rCount1).Cells(6).Value
                'End If

                strPayee = tblMHI.Rows(rCount1).Cells(6).Value

                If InStr(strSvcCode, "OIM") > 0 Then                            '''removing CCur from Int
                    curMedCov = tblMHI.Rows(rCount1).Cells(9).Value
                    curMedCov = Math.Round(curMedCov, 2)
                    curMedPaid = tblMHI.Rows(rCount1).Cells(16).Value
                    curMedPaid = Math.Round(curMedPaid, 2)
                    Try
                        curInnDed = curInnDed + tblMHI.Rows(rCount1).Cells(39).Value
                        curInnDed = Math.Round(curInnDed, 2)
                    Catch ex As Exception

                    End Try

                    curOonDed = curOonDed + tblMHI.Rows(rCount1).Cells(42).Value
                    curOonDed = Math.Round(curOonDed, 2)
                    curInnOop = curInnOop + tblMHI.Rows(rCount1).Cells(40).Value
                    curInnOop = Math.Round(curInnOop, 2)
                    curOonOop = curOonOop + tblMHI.Rows(rCount1).Cells(42).Value
                    curOonOop = Math.Round(curOonOop, 2)
                ElseIf InStr(strSvcCode, "OI") > 0 Then
                    curOICov = tblMHI.Rows(rCount1).Cells(12).Value
                    curOICov = Math.Round(curOICov, 2)
                    curOIPaid = tblMHI.Rows(rCount1).Cells(16).Value
                    curOIPaid = Math.Round(curOIPaid, 2)
                    Try
                        curInnDed = curInnDed + tblMHI.Rows(rCount1).Cells(39).Value
                    Catch ex As Exception
                    End Try
                    curInnDed = Math.Round(curInnDed, 2)
                    curOonDed = curOonDed + tblMHI.Rows(rCount1).Cells(41).Value
                    curOonDed = Math.Round(curOonDed, 2)
                    curInnOop = curInnOop + tblMHI.Rows(rCount1).Cells(40).Value
                    curInnOop = Math.Round(curInnOop, 2)
                    curOonOop = curOonOop + tblMHI.Rows(rCount1).Cells(42).Value
                    curInnOop = Math.Round(curInnOop, 2)
                ElseIf InStr(strSvcCode, "COPAY") > 0 Then
                    curCopay = tblMHI.Rows(rCount1).Cells(10).Value
                    curCopay = Math.Round(curCopay, 2)
                    Try
                        curInnOop = curInnOop + tblMHI.Rows(rCount1).Cells(40).Value
                        curInnOop = Math.Round(curInnOop, 2)
                    Catch ex As Exception

                    End Try
                    ''''''''''''need to work here if it is blank  updating by sanjeet
                    Try
                        curOonOop = curOonOop + tblMHI.Rows(rCount1).Cells(42).Value
                        curOonOop = Math.Round(curOonOop, 2)
                    Catch ex As Exception

                    End Try

                Else
                    If strOvCode <> "30" Or (strOvCode = "30" And InStr(conOvCode, strPayee) > 0) Then
                        curCharge = curCharge + tblMHI.Rows(rCount1).Cells(9).Value : Threading.Thread.Sleep(10)
                        curCharge = Math.Round(curCharge, 2) : Threading.Thread.Sleep(10)
                        curNotCovd = curNotCovd + tblMHI.Rows(rCount1).Cells(10).Value : Threading.Thread.Sleep(10)
                        curNotCovd = Math.Round(curNotCovd, 2) : Threading.Thread.Sleep(10)
                        curCovered = curCovered + tblMHI.Rows(rCount1).Cells(12).Value : Threading.Thread.Sleep(10)
                        curCovered = Math.Round(curCovered, 2) : Threading.Thread.Sleep(10)
                        curPaid = curPaid + tblMHI.Rows(rCount1).Cells(16).Value : Threading.Thread.Sleep(10)
                        curPaid = Math.Round(curPaid, 2) : Threading.Thread.Sleep(10)
                        Try
                            curInnDed = curInnDed + tblMHI.Rows(rCount1).Cells(39).Value : Threading.Thread.Sleep(10)
                        Catch ex As Exception
                        End Try

                        curInnDed = Math.Round(curInnDed, 2) : Threading.Thread.Sleep(10)

                        Try
                            Dim cellValue As Object = curOonDed + tblMHI.Rows(rCount1).Cells(41).Value : Threading.Thread.Sleep(10)
                            If Not DBNull.Value.Equals(cellValue) Then
                                'curOonDed += Convert.ToDouble(cellValue)
                                curOonDed = cellValue
                            End If
                        Catch ex As Exception

                        End Try
                        '   curOonDed = curOonDed + tblMHI.Rows(rCount).Cells(41).Value
                        curOonDed = Math.Round(curOonDed, 2)

                        Try
                            curInnOop = curInnOop + tblMHI.Rows(rCount1).Cells(40).Value
                            curInnOop = Math.Round(curInnOop, 2)
                        Catch ex As Exception

                        End Try
                        Try
                            curOonOop = curOonOop + tblMHI.Rows(rCount1).Cells(42).Value
                            curOonOop = Math.Round(curOonOop, 2)
                        Catch ex As Exception

                        End Try
                        DedCode = tblMHI.Rows(rCount1).Cells(14).Value

                    End If
                End If

                If InStr(DedCode, "M") > 0 Then blnInnNet = True
                If InStr(DedCode, "Z") > 0 Then
                    If blnTiered = True Then
                        blnInnNet = True
                    End If
                End If

                If DedCode = "?" Or DedCode = "" Then blnQuestion = True

                If strRemark = "69" Then                        ''''''''''''''''''''''''Added on 04/08/2023
                    If tblMHI.Rows(rCount1).Cells(25).Value = StrDraft And
                        tblMHI.Rows(rCount1).Cells(8).Value = "69" Then
                        'ActiveCell.Offset(1, 0).Select
                        rCount1 = rCount1 + 1
                    Else
                        rCount1 = rCount1 + 1
                        'ActiveCell.Offset(1, 0).Select
                        Exit Do
                    End If

                    Try
                        If IsDBNull(tblMHI.Rows(rCount1 + 1).Cells(25).Value) Then rCount1 = rCount1 + 1
                    Catch ex As Exception
                        rCount1 = rCount1 + 1
                    End Try

                    '''04/09/2023

                    If tblMHI.Rows(rCount1 + 1).Cells(25).Value = Nothing Then
                        Exit Do
                    End If


                ElseIf tblMHI.Rows(rCount1 + 1).Cells(25).Value = StrDraft And tblMHI.Rows(rCount1).Cells(8).Value <> "69" Then

                    rCount1 = rCount1 + 1
                    'ActiveCell.Offset(1, 0).Select
                Else
                    rCount1 = rCount1 + 1
                    'ActiveCell.Offset(1, 0).Select
                    Exit Do                                       '''Commented as per 04/08/2023
                End If

                If rCount1 >= tblMHI.Rows.Count - 1 Then            ''Ending loop ''Added 
                    Exit Do
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''Added on 04/08/2023

            Loop


            tblOOP.Rows.Insert(tblOOP.RowCount - 1)
            tblOOP.Update()

            tblOOP.Rows(OOPRcnt).Cells(0).Value = Format(dtFrstDos, "MM/dd/yyyy")
            tblOOP.Rows(OOPRcnt).Cells(1).Value = Format(dtLastDos, "MM/dd/yyyy")
            tblOOP.Rows(OOPRcnt).Cells(2).Value = strProv
            tblOOP.Rows(OOPRcnt).Cells(4).Value = "$" & curCharge

            If intChoice <> 2 Then
                If blnInnNet = True Then
                    tblOOP.Rows(OOPRcnt).Cells(3).Value = "Yes"
                Else
                    If blnQuestion = True Then
                        tblOOP.Rows(OOPRcnt).Cells(3).Value = "?"
                    Else
                        tblOOP.Rows(OOPRcnt).Cells(3).Value = "No"
                    End If
                End If
            End If

            '************************************************************************************
            Select Case intChoice
                Case 1
                    Threading.Thread.Sleep(100)
                    tblOOP.Rows(OOPRcnt).Cells(5).Value = "$" & curCovered
                    tblOOP.Rows(OOPRcnt).Cells(6).Value = "$" & curCopay
                    tblOOP.Rows(OOPRcnt).Cells(7).Value = "$" & curInnDed
                    tblOOP.Rows(OOPRcnt).Cells(8).Value = "$" & curOonDed
                    tblOOP.Rows(OOPRcnt).Cells(7).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(8).Style.BackColor = Color.Aqua

                    tblOOP.Rows(OOPRcnt).Cells(9).Value = "$" & curPaid

                    tblOOP.Rows(OOPRcnt).Cells(10).Value = "$" & curInnOop
                    tblOOP.Rows(OOPRcnt).Cells(11).Value = "$" & curOonOop
                    tblOOP.Rows(OOPRcnt).Cells(10).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(11).Style.BackColor = Color.Aqua

                    tblOOP.Rows(OOPRcnt).Cells(12).Value = "$" & curNotCovd
                    tblOOP.Rows(OOPRcnt).Cells(14).Value = StrICN

                    tblOOP.Rows(OOPRcnt).Cells(15).Value = Format(dtProcDate, "MM/dd/yyyy")

                Case 2
                    tblOOP.Rows(OOPRcnt).Cells(5).Value = "$" & curMedCov
                    tblOOP.Rows(OOPRcnt).Cells(6).Value = "$" & curCovered
                    tblOOP.Rows(OOPRcnt).Cells(7).Value = "$" & curInnDed
                    tblOOP.Rows(OOPRcnt).Cells(8).Value = "$" & curMedPaid
                    tblOOP.Rows(OOPRcnt).Cells(9).Value = "$" & curPaid
                    tblOOP.Rows(OOPRcnt).Cells(10).Value = "$" & curInnOop
                    tblOOP.Rows(OOPRcnt).Cells(11).Value = "$" & curNotCovd
                    tblOOP.Rows(OOPRcnt).Cells(13).Value = StrICN
                    Try
                        tblOOP.Rows(OOPRcnt).Cells(14).Value = Format(dtProcDate, "MM/dd/yyyy")
                    Catch ex As Exception

                    End Try


                Case 3
                    tblOOP.Rows(OOPRcnt).Cells(5).Value = "$" & curMedCov
                    tblOOP.Rows(OOPRcnt).Cells(6).Value = "$" & curCovered
                    tblOOP.Rows(OOPRcnt).Cells(7).Value = "$" & curCopay
                    tblOOP.Rows(OOPRcnt).Cells(8).Value = "$" & curInnDed
                    tblOOP.Rows(OOPRcnt).Cells(9).Value = "$" & curOonDed
                    tblOOP.Rows(OOPRcnt).Cells(8).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(9).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(10).Value = "$" & curMedPaid
                    tblOOP.Rows(OOPRcnt).Cells(11).Value = "$" & curPaid
                    tblOOP.Rows(OOPRcnt).Cells(12).Value = "$" & curInnOop
                    tblOOP.Rows(OOPRcnt).Cells(13).Value = "$" & curOonOop
                    tblOOP.Rows(OOPRcnt).Cells(12).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(13).Style.BackColor = Color.Aqua

                    tblOOP.Rows(OOPRcnt).Cells(14).Value = "$" & curNotCovd
                    tblOOP.Rows(OOPRcnt).Cells(16).Value = StrICN
                    Try
                        tblOOP.Rows(OOPRcnt).Cells(17).Value = Format(dtProcDate, "MM/dd/yyyy")
                    Catch ex As Exception

                    End Try


                Case 4
                    tblOOP.Rows(OOPRcnt).Cells(5).Value = "$" & curOICov
                    tblOOP.Rows(OOPRcnt).Cells(6).Value = "$" & curCovered
                    tblOOP.Rows(OOPRcnt).Cells(7).Value = "$" & curCopay
                    tblOOP.Rows(OOPRcnt).Cells(8).Value = "$" & curInnDed
                    tblOOP.Rows(OOPRcnt).Cells(9).Value = "$" & curOonDed
                    tblOOP.Rows(OOPRcnt).Cells(8).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(9).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(10).Value = "$" & curOIPaid
                    tblOOP.Rows(OOPRcnt).Cells(11).Value = "$" & curPaid
                    tblOOP.Rows(OOPRcnt).Cells(12).Value = "$" & curInnOop
                    tblOOP.Rows(OOPRcnt).Cells(13).Value = "$" & curOonOop
                    tblOOP.Rows(OOPRcnt).Cells(12).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(13).Style.BackColor = Color.Aqua
                    tblOOP.Rows(OOPRcnt).Cells(14).Value = "$" & curNotCovd
                    tblOOP.Rows(OOPRcnt).Cells(16).Value = StrICN
                    tblOOP.Rows(OOPRcnt).Cells(17).Value = Format(dtProcDate, "MM/dd/yyyy")

                Case 5

            End Select


            If Int(tblOOP.Rows(OOPRcnt).Cells(4).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(4).Style.ForeColor = Color.Red

            If Int(tblOOP.Rows(OOPRcnt).Cells(5).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(5).Style.ForeColor = Color.Red

            If Int(tblOOP.Rows(OOPRcnt).Cells(6).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(6).Style.ForeColor = Color.Red

            If Int(tblOOP.Rows(OOPRcnt).Cells(7).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(7).Style.ForeColor = Color.Red

            If Int(tblOOP.Rows(OOPRcnt).Cells(8).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(8).Style.ForeColor = Color.Red

            If Int(tblOOP.Rows(OOPRcnt).Cells(9).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(9).Style.ForeColor = Color.Red

            If Int(tblOOP.Rows(OOPRcnt).Cells(10).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(10).Style.ForeColor = Color.Red

            If Int(tblOOP.Rows(OOPRcnt).Cells(11).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(11).Style.ForeColor = Color.Red

            'If Int(tblOOP.Rows(OOPRcnt).Cells(12).Value) < 0 Then tblOOP.Rows(OOPRcnt).Cells(12).Style.ForeColor = Color.Red

            OOPRcnt = OOPRcnt + 1

            dtLastDos = Nothing
            strProv = Nothing
            curCharge = Nothing
            curCopay = Nothing
            curInnDed = Nothing
            curOonDed = Nothing
            curPaid = Nothing
            curInnOop = Nothing
            curOonOop = Nothing
            curNotCovd = Nothing
            StrICN = Nothing
            dtProcDate = Nothing
            curMedCov = Nothing
            curMedPaid = Nothing
            curOICov = Nothing
            curOIPaid = Nothing
            DedCode = Nothing
            curCovered = Nothing
            strRemark = Nothing

            If rCount1 >= tblMHI.Rows.Count - 1 Then            ''Ending loop 
                Exit For
            End If
        Next rCount

        Main(starttime)
    End Sub

    Sub Format_Sheet()

        RichTextBox1.AppendText("Gathering Provider Name And Type...   " & vbCrLf)
        RichTextBox1.SelectionBullet = True
        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4

        '''''--------------------------------------04/10/2023  deleted all blank row which was created for first calculation 
        Dim intLastrow As Integer = tblMHI.Rows.Count - 1
        For I = 0 To intLastrow
            If I > intLastrow Then Exit For
            Try
                If IsDBNull(tblMHI.Rows(I).Cells(0).Value) Then
                    tblMHI.Refresh()
                    tblMHI.Rows.RemoveAt(I)
                    tblMHI.Refresh()
                    intLastrow = intLastrow - 1
                    I = I - 1
                End If
            Catch ex As Exception

            End Try
        Next I
        ''--------------------------
        ProcDate_Only(True)  'Sorts claims by the processed date need to check Sanjeet
        Call Clean_Claims()        'Searches through History Detail (MHI) spreadsheet for claim scenarios _ 
        'that are not in scope for the final formatted sheet that is mailed out
        'to member.

        '        Exit Sub

        'Verifies that a deductible indicator (M or G) is listed on each claim line.

        'MHIActive
        Dim inxt As Integer

        For inxt = 0 To tblMHI.Rows.Count - 1

            If IsDBNull(tblMHI.Rows(inxt).Cells(38).Value) Then
                inxt = inxt + 1
            End If
            If IsDBNull(tblMHI.Rows(inxt).Cells(38).Value) Then
                Exit For
            End If


            'If tblMHI.Rows(inxt).Cells(14).Value Then
            '    MsgBox(True)
            'End If

            If isNullOrEmpty(tblMHI.Rows(inxt).Cells(14).Value) Then
                strDedCode = "0.00"
            Else
                strDedCode = Trim(tblMHI.Rows(inxt).Cells(14).Value)
            End If

            If strDedCode = "" Or strDedCode = "?" Then
                If strInputDeductible = "" Then
                    'DedIndSort()                       ''''commented for DedIndSort to multiple member             04/08/2023
                    'strDedCode = MsgBox("Please review the Deductible indicator On the" & vbCrLf &
                    '   "MHI tab And update As appropriate (M For INN, G For OON)." &
                    'vbCrLf & "Then run the macro again.", vbInformation, "Deductible Indicator")
                    'strInputDeductible = InputBox("Please enter value 'M' OR 'G' to update the Deductible Indicator." & vbCrLf & "Note: There are few blank values under Deductible Indicator column.") ''need to do etalic
                End If
                'tblMHI.Rows(inxt).Cells(14).Value = strInputDeductible

                tblMHI.Rows(inxt).Cells(14).Value = "0"

                'Exit Sub
            End If
        Next inxt


        pName = tblMHI.Rows(0).Cells(35).Value
        pRel = tblMHI.Rows(0).Cells(37).Value

        GetProvInfo()
        Format_Option()

        Format_Sort()  '''need to uncomment
        Claim_Combine()

        ''''Need to check Sanjeet

        'Formulas to add up Deductible/OOP columns for total, listed at bottom of formatted sheet.        
        'Select Case intChoice
        Call Sum_OOP_Case1()
        'Call Sum_Data()
        'End Select
        ''**************************************************************************

        RichTextBox1.SelectionIndent = 5
        RichTextBox1.BulletIndent = 4
        RichTextBox1.SelectionBullet = True
        RichTextBox1.AppendText("OOP Spreadsheet - Data Populated" & vbCrLf)

        TabControl1.SelectedIndex = 8

    End Sub
    Sub Sum_OOP_Case1()

        'Dim sum As Integer = 0
        'Dim icnt As Integer
        'For j As Integer = 7 To 12
        '    icnt = 5
        '    For icnt = 5 To tblOOP.Rows.Count() - 1
        '        If Trim(tblOOP.Rows(icnt).Cells(j).Value) = "" Then
        '            Exit For
        '        Else
        '            sum = sum + Int(tblOOP.Rows(icnt).Cells(j).Value)
        '        End If
        '    Next
        '    tblOOP.Rows(icnt).Cells(j).Value = sum
        '    tblOOP.Rows(icnt).Cells(j).Style.BackColor = Color.Gray
        '    tblOOP.Rows(icnt).Cells(6).Value = "Total"
        '    sum = 0.00
        'Next

        ' sumit singh
        Dim sum As Double = 0.00

        Dim icnt As Integer

        For j As Integer = 7 To 11

            icnt = 5

            For icnt = 5 To tblOOP.Rows.Count() - 1

                If Trim(tblOOP.Rows(icnt).Cells(j).Value) = "" Then

                    Exit For

                Else

                    sum = sum + CDbl(tblOOP.Rows(icnt).Cells(j).Value)


                End If

            Next

            tblOOP.Rows(icnt).Cells(j).Value = sum.ToString("F2")



            tblOOP.Rows(icnt).Cells(j).Style.BackColor = Color.Gray

            tblOOP.Rows(icnt).Cells(6).Value = "Total"

            sum = 0.00

        Next

        q = tblOOP.Rows.Count() - 1
        tblOOP.Rows(q).Cells(9).Value = ""


    End Sub
    Sub Format_Option()
        Dim strMyChoice As String

        Do

            strMyChoice = InputBox("Please choose from the following format options for " & pName &
                                    "/" & pRel & ":" & vbCrLf & vbCrLf & "1 - OOP Spreadsheet" & vbCrLf &
                                    "2 - Medicare" & vbCrLf & "3 - Medicare w/PPO" & vbCrLf &
                                    "4 - Other Insurance" & vbCrLf &
                                    "5 - Other Insurance and Medicare" & vbCrLf & vbCrLf & "", "OOP Spreadsheet Choices")


            '"5 - Other Insurance and Medicare"
            Select Case strMyChoice
                Case 1
                    RichTextBox1.AppendText("Gathering Data to OOP Spreadsheet...   " & vbCrLf)
                    RichTextBox1.SelectionBullet = True
                    RichTextBox1.SelectionIndent = 5
                    RichTextBox1.BulletIndent = 4

                    strMyChoice = "OOP Spreadsheet" & " - " & pName & " " & pRel
                    intChoice = 1
                Case 2
                    strMyChoice = "Medicare" & " - " & pName & " " & pRel
                    intChoice = 2
                Case 3
                    strMyChoice = "Medicare wPPO" & " - " & pName & " " & pRel
                    intChoice = 3
                Case 4
                    strMyChoice = "Other Insurance" & " - " & pName & " " & pRel
                    intChoice = 4
                Case 5
                    strMyChoice = "OIMEDI" & " - " & pName & " " & pRel
                    intChoice = 5
                Case Else
                    strMyChoice = MsgBox("Invalid entry, please try again.", vbOKOnly + vbExclamation)
                    strMyChoice = ""
            End Select
        Loop Until strMyChoice <> ""

        Dim strPtLast, strPayloc
        Dim intSheetType As Integer



        Setup_Headers(intSheetType, strPtLast, strPayloc, strMyChoice, intChoice)

        'Add_Sheet strMyChoice
    End Sub

    'Sub Setup_Headers(ByVal intSheetType As Integer, ByVal strPtLast As String,
    '                ByVal strPayloc As String)

    'End Sub



    '*******************************************************************************************************
    'Runs through claim history on MHI tab and removes claim scenarios that are not
    'in scope of project.
    '*******************************************************************************************************
    Sub Clean_Claims()
        Dim strRemark As String, strOvRide As String, strPayee As String
        Dim curCharge As Double, curNotCov As Double, strSvcCode As String
        Dim blnRowDeleted As Boolean, strPOS As String, curPaidAmt As Double
        Dim StrDraft As String, intClmLoop As Integer, strOIMLine As String
        Dim blnOIMLine As Boolean, intRowDelete As Integer, intStartRow As Integer
        Dim blnMultiDraft As Boolean, curDedAmt As Double, StrICN As String, intTestRow

        Call ICN_Draft_Sort()

        Verify_Draft()

        'ProcDate_Only(True)

        'Range("A5").Select
        For I = 0 To tblMHI.Rows.Count - 2              '''added 04/08/2023

            'Do
            If IsDBNull(tblMHI.Rows(I).Cells(38).Value) Then
                I = I + 1
            End If
            If IsDBNull(tblMHI.Rows(I).Cells(38).Value) Then
                Exit For
            End If

            blnRowDeleted = False
            blnOIMLine = False
            blnMultiDraft = False

            Try
                strSvcCode = tblMHI.Rows(I).Cells(2).Value.ToString()
                strPOS = tblMHI.Rows(I).Cells(3).Value.ToString()
                strOvRide = tblMHI.Rows(I).Cells(5).Value.ToString()
                strPayee = tblMHI.Rows(I).Cells(6).Value.ToString()
                strRemark = tblMHI.Rows(I).Cells(8).Value.ToString()
                curCharge = tblMHI.Rows(I).Cells(9).Value
                'If IsDBNull(Trim(tblMHI.Rows(I).Cells(10).Value)) Or Trim(tblMHI.Rows(I).Cells(10).Value) = Nothing Then
                '    tblMHI.Rows(I).Cells(10).Value = "0.00"
                'End If
                curNotCov = Trim(tblMHI.Rows(I).Cells(10).Value)              'CCur
                curPaidAmt = tblMHI.Rows(I).Cells(16).Value              'CCur
                StrDraft = tblMHI.Rows(I).Cells(15).Value
                curDedAmt = tblMHI.Rows(I).Cells(13).Value              'CCur            
                StrICN = tblMHI.Rows(I).Cells(30).Value
            Catch ex As Exception

            End Try


            intTestRow = I

            intClmLoop = 1
            strOIMLine = strSvcCode
            Do
                If tblMHI.Rows(intClmLoop).Cells(25).Value <> StrDraft Then
                    Exit Do
                Else
                    If InStr(strOIMLine, "OI") = 0 Then
                        curCharge = curCharge + tblMHI.Rows(intClmLoop).Cells(9).Value 'CCUR
                        curNotCov = curNotCov + tblMHI.Rows(intClmLoop).Cells(9).Value 'CCUR
                        strOIMLine = Trim(tblMHI.Rows(intClmLoop).Cells(2).Value)
                    Else
                        blnOIMLine = True
                    End If
                    intClmLoop = intClmLoop + 1
                End If
            Loop

            If strOvRide = "C" Or strOvRide = "P" Or strOvRide = "R" Then           ''need to check all delete
                If blnMultiDraft = True Then
                    For intRowDelete = 1 To intClmLoop
                        ' ActiveCell.Offset(-intRowDelete, 0).EntireRow.Delete
                        If intRowDelete = intClmLoop Then
                            '        ActiveCell.Offset(-intRowDelete, 0).Select
                            blnRowDeleted = True
                            Exit For
                        End If
                    Next
                Else
                    'Selection.EntireRow.Delete
                End If
                blnRowDeleted = True
            End If
            If strRemark = "70" And blnRowDeleted = False Then
                'Selection.EntireRow.Delete
                blnRowDeleted = True
            End If

            If blnRowDeleted = False And I <> 0 Then
                If Not IsDBNull(tblMHI.Rows(I - 1).Cells(22).Value) Then                ''Added 04/08/2023
                    If tblMHI.Rows(I).Cells(22).Value = "" Or (blnMultiDraft = True And tblMHI.Rows(I - 1).Cells(22).Value = "") Then
                        If (strPOS = "RX" Or strSvcCode = "RX") And (strPayee = "H" Or strPayee = "D" Or strPayee = "T") Then
                            If blnMultiDraft = True Then
                                For intRowDelete = 1 To intClmLoop
                                    'ActiveCell.Offset(-intRowDelete, 22).Value = "Pharmacy"
                                    tblMHI.Rows(-intRowDelete).Cells(22).Value = "Pharmacy"
                                    If intRowDelete = intClmLoop Then
                                        Exit For
                                    End If
                                Next
                                'ActiveCell.Offset(-1, 0).Select
                                I = I - 1
                            Else
                                tblMHI.Rows(intClmLoop).Cells(22).Value = "Pharmacy"
                            End If
                        End If
                    End If
                End If
                'ActiveCell.Offset(1, 0).Select
            End If

                strSvcCode = ""
            strPOS = ""
            strOvRide = ""
            strPayee = ""
            strRemark = ""
            curCharge = 0
            curNotCov = 0
            curPaidAmt = 0
            StrDraft = ""
            curDedAmt = 0
            StrICN = ""



            'Loop Until ActiveCell.Value = "" And ActiveCell.Offset(1, 0).Value = ""
        Next I
    End Sub

    Sub Verify_Draft()
        Dim StrDraft As String, StrICN As String
        Dim strReplDraft As String, strNewICN As String, strNewDraft As String

        strReplDraft = 1
        Dim q As Integer = 0
        Do
            StrDraft = tblMHI.Rows(q).Cells(25).Value.ToString()
            StrICN = tblMHI.Rows(q).Cells(30).Value.ToString()
            If StrDraft <> "0000000000" Then
                Exit Do
            ElseIf StrDraft = "0000000000" Then
                'strNewDraft = String(10 - Len(strReplDraft), "0") & strReplDraft  '' need to check 
                strNewDraft = String.Format(10 - Len(strReplDraft), "0") & strReplDraft  '' need to check 
                tblMHI.Rows(q).Cells(24).Value = strNewDraft
                q = q + 1
                Do
                    strNewICN = tblMHI.Rows(q).Cells(29).Value
                    If strNewICN = StrICN Then
                        tblMHI.Rows(q).Cells(24).Value = strNewDraft
                        q = q + 1
                    Else
                        strReplDraft = strReplDraft + 1
                        Exit Do
                    End If
                Loop
            End If
        Loop

    End Sub

    Private Sub tblMHI_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles tblMHI.CellEndEdit
        ' Save the changes to the cell
        Dim rowIndex As Integer = e.RowIndex
        Dim colIndex As Integer = e.ColumnIndex
        Dim cellValue As Object = tblMHI.Rows(rowIndex).Cells(colIndex).Value

        Call Build_Accural_Column_Formulas()

        ' Code to update the data source with the new value goes here
        ' For example, if you are using a DataTable as the data source:
        ' dataTable.Rows(rowIndex)(colIndex) = cellValue

    End Sub

    Private Sub DGridOverview_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGridOverview.CellEndEdit
        ' Save the changes to the cell
        Dim rowIndex As Integer = e.RowIndex
        Dim colIndex As Integer = e.ColumnIndex
        Dim cellValue As Object = DGridOverview.Rows(rowIndex).Cells(colIndex).Value

        ' Code to update the data source with the new value goes here
        ' For example, if you are using a DataTable as the data source:
        ' dataTable.Rows(rowIndex)(colIndex) = cellValue
    End Sub

    Private Sub DGrid_PG1_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGrid_PG1.CellEndEdit
        ' Save the changes to the cell
        Dim rowIndex As Integer = e.RowIndex
        Dim colIndex As Integer = e.ColumnIndex
        Dim cellValue As Object = DGrid_PG1.Rows(rowIndex).Cells(colIndex).Value

        ' Code to update the data source with the new value goes here
        ' For example, if you are using a DataTable as the data source:
        ' dataTable.Rows(rowIndex)(colIndex) = cellValue
    End Sub

    Private Sub DGrid_PG4_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGrid_PG4.CellEndEdit
        ' Save the changes to the cell
        Dim rowIndex As Integer = e.RowIndex
        Dim colIndex As Integer = e.ColumnIndex
        Dim cellValue As Object = DGrid_PG4.Rows(rowIndex).Cells(colIndex).Value

        ' Code to update the data source with the new value goes here
        ' For example, if you are using a DataTable as the data source:
        ' dataTable.Rows(rowIndex)(colIndex) = cellValue
    End Sub

    Private Sub DGrid_PG5_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGrid_PG5.CellEndEdit
        ' Save the changes to the cell
        Dim rowIndex As Integer = e.RowIndex
        Dim colIndex As Integer = e.ColumnIndex
        Dim cellValue As Object = DGrid_PG5.Rows(rowIndex).Cells(colIndex).Value

        ' Code to update the data source with the new value goes here
        ' For example, if you are using a DataTable as the data source:
        ' dataTable.Rows(rowIndex)(colIndex) = cellValue
    End Sub

    Private Sub DGrid_PG10_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGrid_PG10.CellEndEdit
        ' Save the changes to the cell
        Dim rowIndex As Integer = e.RowIndex
        Dim colIndex As Integer = e.ColumnIndex
        Dim cellValue As Object = DGrid_PG10.Rows(rowIndex).Cells(colIndex).Value

        ' Code to update the data source with the new value goes here
        ' For example, if you are using a DataTable as the data source:
        ' dataTable.Rows(rowIndex)(colIndex) = cellValue
    End Sub

    Private Sub DGridCEI_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGridCEI.CellEndEdit
        ' Save the changes to the cell
        Dim rowIndex As Integer = e.RowIndex
        Dim colIndex As Integer = e.ColumnIndex
        Dim cellValue As Object = DGridCEI.Rows(rowIndex).Cells(colIndex).Value

        ' Code to update the data source with the new value goes here
        ' For example, if you are using a DataTable as the data source:
        ' dataTable.Rows(rowIndex)(colIndex) = cellValue
    End Sub

    Private Sub DGridMInfo_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DGridMInfo.CellEndEdit
        ' Save the changes to the cell
        Dim rowIndex As Integer = e.RowIndex
        Dim colIndex As Integer = e.ColumnIndex
        Dim cellValue As Object = DGridMInfo.Rows(rowIndex).Cells(colIndex).Value

        ' Code to update the data source with the new value goes here
        ' For example, if you are using a DataTable as the data source:
        ' dataTable.Rows(rowIndex)(colIndex) = cellValue
    End Sub

    Dim myConn As SqlConnection
    Dim myCmd As SqlCommand
    Public Property Final_dt_main_SORT As Data.DataTable

    'Sub Main(args As String())
    Sub Main(args)

        Dim MacroName, ME_FileName, UserId, Citrix As String
        Dim endtime As Date
        MacroName = "UNET_Out_of_Pocket_Calculator"
        ME_FileName = "PCOMM_EI_NAT"
        'UserId = "Sample"
        UserId = Environment.UserName
        starttime = args
        endtime = Now
        FComments = "UAT-" & txt_Policy.Text & " " & txt_SSN.Text & " " & MHIstarttime & " " & MHIendtime & " " & MHIHistoryCnt & " " & MMIstarttime & " " & MMIendtime
        'FComments = "Testing"
        Citrix = ""

        myConn = New SqlConnection("Data Source=wp000075633; Initial Catalog=NMT;Integrated Security=SSPI;")

        myCmd = New SqlCommand("INSERT INTO MacroUtilization (MacroName, ME_FileName, NTID, StartTime,  EndTime, FreeFormComments, Citrix) VALUES ('" & MacroName & "','" & ME_FileName & "','" & UserId & "','" & starttime & "','" & endtime & "','" & FComments & "','" & Citrix & "')", myConn)

        Dim s As String = "INSERT INTO MacroUtilization (MacroName, ME_FileName, NTID, StartTime,  EndTime, FreeFormComments, Citrix) VALUES ('" & MacroName & "','" & ME_FileName & "','" & UserId & "','" & starttime & "','" & endtime & "','" & FComments & "','" & FComments & "')"
        Try
            myConn.Open()
            myCmd.ExecuteNonQuery()
            myConn.Close()
        Catch ex As Exception
            'Dim dFormat = Today
            Dim dfromat1 As String
            'dfromat1 = Replace(dFormat, "/", "-")

            'Dim Error_UsageLog As String = "\\nasgw013pn.uhc.com\National_Macro_Project\MasterCopies\User Data\MUtilization_Error_Log_NEW\" & dfromat1 & "-" & UserId & "-ErrorLog_OOP_Calculator.txt"
            'Dim ErrFilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Insight Software\Macro Express\ErrLog.txt"
            'Dim Log_Data As String = MacroName & Chr(9) & starttime & Chr(9) & endtime & Chr(9) & ME_FileName & Chr(9) & UserId & Chr(9) & starttime & Chr(9) & FComments & Err.Number & Chr(9) & Citrix
            'File.AppendAllText(Error_UsageLog, Log_Data)

            '      Dim ErrStr As String = "Message:" + "/n" + ex.StackTrace
            Dim dFormat = Today
            Dim dFormat1 As String
            dFormat1 = Replace(dFormat, "/", "-")
            Dim objNet = CreateObject("Wscript.Network")
            ' Dim UserId As String = objNet.UserName
            Dim Error_UsageLog As String = "\\nasgw013pn.uhc.com\National_Macro_Project\MasterCopies\User Data\MUtilization_Error_Log_NEW\" & dfromat1 & "-" & UserId & "-ErrorLog_OOP_Calculator.txt"
            'Dim ErrFilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Insight Software\Macro Express\-ErrorLog_OOP_Calculator.txt"
            If Not File.Exists(Error_UsageLog) Then
                Threading.Thread.Sleep(50)
                Dim context As String = "MACRONAME" & Chr(9) & "STARTTIME" & Chr(9) & "ENDTIME" & Chr(9) & "MACROFILE" & Chr(9) & "NTID" & Chr(9) & "SERVERTIME" & Chr(9) & "COMMENTS" & Chr(9) & "CITRIX"



                Using sw As New StreamWriter(File.Open(Error_UsageLog, FileMode.OpenOrCreate))
                    sw.WriteLine(context)
                    sw.Close()
                End Using
            End If
            Dim Log_Data As String = MacroName & Chr(9) & starttime & Chr(9) & endtime & Chr(9) & ME_FileName & Chr(9) & UserId & Chr(9) & starttime & Chr(9) & FComments & Err.Number & Chr(9) & Citrix & vbCrLf
            File.AppendAllText(Error_UsageLog, Log_Data)
        End Try

    End Sub



End Class