Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.Configuration
Imports AutConnMgrTypeLibrary
Imports AutSessTypeLibrary

Module Module1
    Public xlApp As Excel.Application = Nothing
    Public xlWorksheet As Excel.Worksheet = Nothing
    Public xlWorkbooks As Excel.Workbooks = Nothing
    Public xlWorkbook As Excel.Workbook = Nothing
    Public xlWorksheets As Excel.Sheets = Nothing

    Public CriteriaSet As Boolean
    Dim LineArray As String
    Dim strarr() As String
    Dim cols As String()
    Dim col As String
    Dim fileReader As String
    Dim reportsFolder As String
    Dim array
    Public URL = "https://doc360-rest-find.optum.com/doc360/auth/v1/token/generate"
    Public objForm As Form1 = New Form1
    Public COSMOS1 As AutSess = New AutSess
    Public BoolRWchk As Boolean
    Public Final_dt_main As New DataTable
    Public Final_Date As String()
    Public dict As New Dictionary(Of String, String)



    Function Get_CEIPlan(ByVal strMember As String, ByVal intRow As Integer)
        Dim x As Integer = 0
        Dim CeiPtName As String
        Dim CeiPTRel()
        Dim PTLastDOS
        Dim memCnt As Integer


        For memCnt = 0 To Form1.DGridCEI.Rows.Count - 2 Step 3

            '''checking for line breadk 


            If Not IsDBNull(Form1.DGridCEI.Rows(memCnt).Cells(2).Value) Then


                Dim stEdt = Format(CDate(Form1.endSelect.Text), "MM/dd/yyyy")
                Dim stSDate = Form1.DGridCEI.Rows(memCnt + 1).Cells(2).Value
                'stSDate = Mid(stSDate, 1, 2) & "/" & stSDate = Mid(stSDate, 3, 2) & "/" & stSDate = Mid(stSDate, 5, 2)

                CeiPtName = Form1.DGridCEI.Rows(memCnt).Cells(0).Value
                Dim stSdt = Right(Form1.DGridCEI.Rows(memCnt + 1).Cells(2).Value, 2)

                If strMember = CeiPtName Then

                    'MsgBox(Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value))

                    If Form1.endSelect.Text = "12/31/2023" And Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value) = "99-99-99" Then
                        Get_CEIPlan = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(0).Value)
                        Exit For
                    ElseIf Right(Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value), 2) = Right(stEdt, 2) Then
                        Get_CEIPlan = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(0).Value)
                        Exit For
                    ElseIf Right(Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value), 2) = Right(stEdt, 2) Then
                        Get_CEIPlan = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(0).Value)
                        Exit For
                    ElseIf IsDBNull(Right(Trim(Form1.DGridCEI.Rows(memCnt + 2).Cells(2).Value), 2)) Then
                        Continue For
                    ElseIf Right(Trim(Form1.DGridCEI.Rows(memCnt + 2).Cells(2).Value), 2) < Right(stEdt, 2) And Right(Trim(Form1.DGridCEI.Rows(memCnt + 2).Cells(3).Value), 2) > Right(stEdt, 2) Then
                        Get_CEIPlan = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(0).Value)
                        Exit For
                    End If

                    If stSdt = "" Then Exit For

                    If Int(stSdt) <= Int(Right(stEdt, 2)) And Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value) = "99-99-99" Then
                        Get_CEIPlan = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(0).Value)
                        Exit For
                    End If

                End If

                Threading.Thread.Sleep(500)
            End If
        Next
    End Function

    Function Get_CEIClass(ByVal strMember As String)
        Dim x As Integer = 0
        Dim CeiPtName As String
        Dim CeiPTRel()
        Dim PTLastDOS
        Dim memCnt As Integer


        For memCnt = 0 To Form1.DGridCEI.Rows.Count - 2 Step 3
            If Not IsDBNull(Form1.DGridCEI.Rows(memCnt).Cells(2).Value) Then


                Dim stEdt = Format(CDate(Form1.endSelect.Text), "MM/dd/yyyy")
                Dim stSDate = Form1.DGridCEI.Rows(memCnt + 1).Cells(2).Value
                'stSDate = Mid(stSDate, 1, 2) & "/" & stSDate = Mid(stSDate, 3, 2) & "/" & stSDate = Mid(stSDate, 5, 2)

                CeiPtName = Form1.DGridCEI.Rows(memCnt).Cells(0).Value
                Dim stSdt = Right(Form1.DGridCEI.Rows(memCnt + 1).Cells(2).Value, 2)

                If strMember = CeiPtName Then

                    'MsgBox(Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value))

                    If Form1.endSelect.Text = "12/31/2023" And Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value) = "99-99-99" Then
                        Get_CEIClass = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(1).Value)
                        Exit For
                    ElseIf Right(Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value), 2) = Right(stEdt, 2) Then
                        Get_CEIClass = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(1).Value)
                        Exit For
                    ElseIf Right(Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value), 2) = Right(stEdt, 2) Then
                        Get_CEIClass = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(1).Value)
                        Exit For
                    ElseIf IsDBNull(Right(Trim(Form1.DGridCEI.Rows(memCnt + 2).Cells(2).Value), 2)) Then
                        Continue For
                    ElseIf Right(Trim(Form1.DGridCEI.Rows(memCnt + 2).Cells(2).Value), 2) < Right(stEdt, 2) And Right(Trim(Form1.DGridCEI.Rows(memCnt + 2).Cells(3).Value), 2) > Right(stEdt, 2) Then
                        Get_CEIClass = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(1).Value)
                        Exit For
                    End If

                    If stSdt = "" Then Exit For

                    If Int(stSdt) <= Int(Right(stEdt, 2)) And Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(3).Value) = "99-99-99" Then
                        Get_CEIClass = Trim(Form1.DGridCEI.Rows(memCnt + 1).Cells(1).Value)
                        Exit For
                    End If
                End If

                Threading.Thread.Sleep(500)
            End If
        Next
    End Function

    Function COSMOS_Window_Selection(strpolicy, strssn) As Boolean

        Dim Logon_Check, Cosmos_Count
        Dim Emulator_Count, Window_Title, Logged_On
        Dim Session_Name As String
        Dim COSMOS_Connection = False
        Dim COSMOS_Session = ""

        Dim Connect_Manager As AutConnMgr = New AutConnMgr


        Connect_Manager.autECLConnList.refresh
        Emulator_Count = Connect_Manager.autECLConnList.count

        If Emulator_Count = 0 Then
            MsgBox("No UNET Sessions were found. Automation will abort.  Emulator count 0.", vbExclamation, "INFORMATION")
            COSMOS_Connection = False
        ElseIf Emulator_Count > 0 Then
            Cosmos_Count = 0
            For Count = 1 To Emulator_Count
                Session_Name = Connect_Manager.autECLConnList(Count).Name
                Dim COSMOS As AutSess = New AutSess
                COSMOS.SetConnectionByName(Session_Name)
                Logon_Check = Trim(GetText(COSMOS, 2, 1, 28))
                If (Logon_Check <> "UHC0010: UNITED HEALTH CARE") Then
                    Logged_On = "Yes"
                Else
                    Logged_On = "No"
                End If

                Window_Title = COSMOS.autECLWinMetrics.WindowTitle
                '                If (Right(Window_Title, 3) = "<" & Session_Name & ">" And Logged_On = "Yes") Then
                If Logged_On = "Yes" Then
                    Cosmos_Count = Cosmos_Count + 1
                    COSMOS_Session = Session_Name
                End If
            Next
            If (Cosmos_Count <> 0) And COSMOS_Session <> "" Then
                'COSMOS1 = CreateObject("pcomm.auteclsession")\

                Try
                    COSMOS1.SetConnectionByName(COSMOS_Session)
                Catch ex As Exception

                End Try

                COSMOS_Connection = True
            Else
                MsgBox("No UNET Sessions were found. Automation will abort.  Emulator count 0.", vbExclamation, "INFORMATION")
                COSMOS_Connection = False
            End If
        Else
            COSMOS_Connection = False
        End If
        If COSMOS_Connection = True Then


            Dim ceiContrilLine As String
            If strssn = "MCI" Then
                ceiContrilLine = "MCI," & strpolicy
                SendKeys(COSMOS1, ceiContrilLine, 2, 2)
                'Logon_Check = Trim(GetText(COSMOS1, 2, 1, 28))
            Else
                ceiContrilLine = "CEI," & strpolicy & "," & strssn
                SendKeys(COSMOS1, ceiContrilLine, 2, 2)
            End If

        End If
        Return COSMOS_Connection
    End Function
    Function GetText(ByRef COSMOS1 As AutSess, ByVal unetRow As Integer, ByVal unetCol As Integer, ByVal length As Integer) As String
        If IsFieldVisible(COSMOS1, unetRow, unetCol) = True Then
            Threading.Thread.Sleep(300)
            Return COSMOS1.autECLPS.gettext(unetRow, unetCol, length)
        Else
            Return ""
        End If
    End Function
    Function IsFieldVisible(ByRef COSMOS1 As AutSess, ByVal unetRow As Integer, ByVal unetCol As Integer) As Boolean
        'Returns the visible property for the text at the row/column location
        'A few Unet screens may contain text that is not visble, but the autECLPS.gettext method will gather/return it.  The bottom half of the MHI history screen is one such case.
        COSMOS1.autECLPS.autECLFieldList.Refresh()
        If COSMOS1.autECLPS.autECLFieldList.FindFieldByRowCol(unetRow, unetCol).display Then
            Return True
        Else
            Return False
        End If
    End Function
    Function SendKeys(ByRef COSMOS1 As AutSess, ByVal unetText As String, Optional ByVal unetRow As Integer = 0, Optional ByVal unetCol As Integer = 0) As Integer
        Dim x As Integer = 0
        Dim CeiPtName()
        Dim CeiPTRel()
        Dim PTLastDOS()

        If unetRow = 0 Or unetCol = 0 Then
            COSMOS1.autECLPS.sendkeys(unetText)
        Else
            COSMOS1.autECLPS.sendkeys("[clear]")
            COSMOS1.autECLPS.sendkeys(unetText, unetRow, unetCol)
            COSMOS1.autECLPS.sendkeys("[Enter]")
            Threading.Thread.Sleep(500)

            Dim strCEILogin As String
            Threading.Thread.Sleep(500)
            strCEILogin = GetText(COSMOS1, 24, 3, 24)

            If InStr(strCEILogin, "NOT SIGNED") > 0 Then
                MsgBox("Adjuster not signed on, Please SignOn and re-run the Tool ")
                Exit Function
            End If

            For memCnt = 9 To 21 Step 4

                ReDim Preserve CeiPtName(x)
                ReDim Preserve CeiPTRel(x)
                ReDim Preserve PTLastDOS(x)
                Threading.Thread.Sleep(500)
                CeiPtName(x) = GetText(COSMOS1, memCnt, 3, 13)
                Threading.Thread.Sleep(500)
                CeiPTRel(x) = GetText(COSMOS1, memCnt, 16, 2)
                Threading.Thread.Sleep(500)
                PTLastDOS(x) = GetText(COSMOS1, memCnt, 19, 6)
                Threading.Thread.Sleep(500)

                Form1.DGridCEI.Rows.Add(CeiPtName(x).ToString(),
                                  CeiPTRel(x).ToString(),
                                  PTLastDOS(x).ToString())
                If Trim(CeiPtName(x)) <> "" Then
                    Form1.memberList.Items.Add(Trim(CeiPtName(x)).ToString() & "/" & Trim(CeiPTRel(x)).ToString())
                End If

                x = x + 1
                Threading.Thread.Sleep(500)


            Next

            Form1.DGridMInfo.Rows.Add(2)
            Form1.DGridMInfo.Rows(0).Cells(0).Value = Trim(GetText(COSMOS1, 3, 14, 15))
            Form1.DGridMInfo.Rows(0).Cells(2).Value = Trim(GetText(COSMOS1, 3, 75, 4))
            Form1.DGridMInfo.Rows(0).Cells(1).Value = Trim(GetText(COSMOS1, 4, 1, 40))
            Form1.DGridMInfo.Rows(1).Cells(1).Value = Trim(GetText(COSMOS1, 5, 1, 40))
            Threading.Thread.Sleep(500)
            Form1.DGridOverview.Refresh()
        End If
        Return Nothing


    End Function

    Function GetTable() As DataTable

        Dim myvalue As String = "sanjeet"
        Dim lastvalue = Right(myvalue, 3)
        'Dim dt As New DataTable("Sample")
        'dt.Columns.Add("Id")
        'dt.Columns.Add("TimeStamp")

        'For i As Int32 = 0 To 200
        '    dt.Rows.Add(New Object() {i, DateTime.Now})
        'Next

        'Dim bs As New BindingSource
        'bs.DataSource = dt
        'Form1.DGridADJ.DataSource = bs
        'bs.Filter = "Id > 10 AND Id < 20"
    End Function
    Function creating_dt()



        Dim dt As New DataTable("Sample")



        dt.Columns.Add("From")
        dt.Columns.Add("Thru")
        dt.Columns.Add("Svc")
        dt.Columns.Add("PS")
        dt.Columns.Add("NBR")
        dt.Columns.Add("OV")
        dt.Columns.Add("P")
        dt.Columns.Add("N")
        dt.Columns.Add("RC")
        dt.Columns.Add("CHARGE")
        dt.Columns.Add("NotCov")
        dt.Columns.Add("BM")
        dt.Columns.Add("Covered")
        dt.Columns.Add("Deduct")
        dt.Columns.Add("D")
        dt.Columns.Add("Perc")
        dt.Columns.Add("Paid")
        dt.Columns.Add("S")
        dt.Columns.Add("DC")
        dt.Columns.Add("SANC")
        dt.Columns.Add("CauseCode")
        dt.Columns.Add("P1")
        dt.Columns.Add("TIN")
        dt.Columns.Add("Suffix")
        dt.Columns.Add("ClaimNumber")
        dt.Columns.Add("Draft")
        dt.Columns.Add("ProcDate")
        dt.Columns.Add("AdjNo")
        dt.Columns.Add("TotalBilled")
        dt.Columns.Add("TotalPaid")
        dt.Columns.Add("ICN")
        dt.Columns.Add("Suf")
        dt.Columns.Add("FLN")
        dt.Columns.Add("PRS")
        dt.Columns.Add("SI")
        dt.Columns.Add("PTName")
        dt.Columns.Add("Blank")
        dt.Columns.Add("PT_Rel")
        dt.Columns.Add("PT_Name")
        dt.Columns.Add("INN_DED")
        dt.Columns.Add("INN_OOP")
        dt.Columns.Add("OON_DED")
        dt.Columns.Add("ONN_OOP")
        dt.Columns.Add("INNDED")
        dt.Columns.Add("INNOOP")
        dt.Columns.Add("ONNDED")
        dt.Columns.Add("ONNOOP")
        dt.Columns.Add("OI_OIM")
        dt.Columns.Add("OOPCalcRun")
        dt.Columns.Add("ICNandSuffix")
        dt.Columns.Add("InpFacility")
        dt.Columns.Add("ProviderName")
        dt.Columns.Add("ProviderType")
        dt.Columns.Add("M1")
        dt.Columns.Add("M2")
        dt.Columns.Add("M3")
        dt.Columns.Add("M4")
        Return dt
    End Function

    Public Function getMHI(Optional ByRef dt_n As DataTable = Nothing) As DataTable
        'Dim dt As New DataTable("Sample")
        'dt.Columns.Add("From")
        'dt.Columns.Add("Thru")
        'dt.Columns.Add("Svc")
        'dt.Columns.Add("PS")
        'dt.Columns.Add("NBR")
        'dt.Columns.Add("OV")
        'dt.Columns.Add("P")
        'dt.Columns.Add("N")
        'dt.Columns.Add("RC")
        'dt.Columns.Add("CHARGE")
        'dt.Columns.Add("NotCov")
        'dt.Columns.Add("BM")
        'dt.Columns.Add("Covered")
        'dt.Columns.Add("Deduct")
        'dt.Columns.Add("D")
        'dt.Columns.Add("Perc")
        'dt.Columns.Add("Paid")
        'dt.Columns.Add("S")
        'dt.Columns.Add("DC")
        'dt.Columns.Add("SANC")
        'dt.Columns.Add("CauseCode")
        'dt.Columns.Add("P1")
        'dt.Columns.Add("TIN")
        'dt.Columns.Add("Suffix")
        'dt.Columns.Add("ClaimNumber")
        'dt.Columns.Add("Draft")
        'dt.Columns.Add("ProcDate")
        'dt.Columns.Add("AdjNo")
        'dt.Columns.Add("TotalBilled")
        'dt.Columns.Add("TotalPaid")
        'dt.Columns.Add("ICN")
        'dt.Columns.Add("Suf")
        'dt.Columns.Add("FLN")
        'dt.Columns.Add("PRS")
        'dt.Columns.Add("SI")
        'dt.Columns.Add("PTName")
        'dt.Columns.Add("Blank")
        'dt.Columns.Add("PT_Rel")
        'dt.Columns.Add("PT_Name")
        'dt.Columns.Add("INN_DED")
        'dt.Columns.Add("INN_OOP")
        'dt.Columns.Add("OON_DED")
        'dt.Columns.Add("ONN_OOP")
        'dt.Columns.Add("INNDED")
        'dt.Columns.Add("INNOOP")
        'dt.Columns.Add("ONNDED")
        'dt.Columns.Add("ONNOOP")
        'dt.Columns.Add("OI_OIM")
        'dt.Columns.Add("OOPCalcRun")
        'dt.Columns.Add("ICNandSuffix")
        'dt.Columns.Add("InpFacility")
        'dt.Columns.Add("ProviderName")
        'dt.Columns.Add("ProviderType")
        'dt.Columns.Add("M1")
        'dt.Columns.Add("M2")
        'dt.Columns.Add("M3")
        'dt.Columns.Add("M4")

        'Dim dt As New DataTable()
        'dt = creating_dt()

        'For i As Int32 = 0 To Form1.DGridMHI_II.Rows.Count - 1
        '    dt.Rows.Add(New Object() {Form1.DGridMHI_II.Rows(i).Cells(0).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(1).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(2).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(3).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(4).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(5).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(6).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(7).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(8).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(9).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(10).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(11).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(12).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(13).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(14).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(15).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(16).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(17).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(18).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(19).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(20).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(21).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(22).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(23).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(24).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(25).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(26).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(27).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(28).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(29).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(30).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(31).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(32).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(33).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(34).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(35).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(36).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(37).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(38).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(39).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(40).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(41).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(42).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(43).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(44).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(45).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(46).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(47).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(48).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(49).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(50).Value,
        '                Form1.DGridMHI_II.Rows(i).Cells(51).Value})
        'Next

        'Dim bs As New BindingSource
        'bs.DataSource = dt
        '''sorting data for multiple Columnt
        ''If Form1.memberList.Items.Count <= 0 Then
        'bs.Sort = "From DESC"
        'bs.Sort = "ProcDate Asc,ICNandSuffix Asc,ProcDate Asc"
        'Form1.tblMHI.DataSource = bs
        ''---------------------------------------------------------------------------------------

        Dim tblCount As Integer = 0

        Dim dttables_ As New List(Of DataTable)
        dttables_ = list(dt_n, dt_n.Rows.Count)
        Dim dt_count As Integer = dttables_.Count
        'Form1.tblMHI.DataSource = dttables_


        Final_dt_main = dt_n.Clone
        Dim increment = New DataColumn("increment", GetType(Integer))
        Final_dt_main.Columns.Add(increment)

        Dim cnt As Integer
        Dim drow As Integer = 0
        For Each datatable As DataTable In dttables_
            tblCount += 1
            Dim dtview As New DataView(datatable)
            dtview.Sort = "From DESC"
            dtview.Sort = "ProcDate Asc,ICNandSuffix Asc,ProcDate Asc"
            Dim dt_dt As DataTable = dtview.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt.Rows

                If dt_dt.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_main.ImportRow(rw)
                    rw = Final_dt_main.NewRow()
                    Final_dt_main.Rows.Add(rw)
                Else
                    Final_dt_main.ImportRow(rw)
                End If
                cnt += 1

            Next
            ReDim Preserve Final_Date(drow)
            Final_Date(drow) = dt_dt.Rows(dt_dt.Rows.Count - 1).Item("From").ToString & "," & dt_dt.Rows(dt_dt.Rows.Count - 1).Item("PT_Name").ToString
            drow += 1
            'dict.Add(dt_dt.Rows(dt_dt.Rows.Count - 1).Item("PT_Name").ToString, dt_dt.Rows(dt_dt.Rows.Count - 1).Item("From").ToString(0))

        Next


        For trow = 0 To Final_dt_main.Rows.Count - 1
            Final_dt_main.Rows(trow).Item("increment") = trow
        Next

        Form1.tblMHI.DataSource = Final_dt_main

        Purgemain.Hide()
    End Function
    Public Function list(dtclone As DataTable, numberOfRecords As Integer) As List(Of DataTable)
        Dim rowcheck As Boolean = False
        Dim dttables As New List(Of DataTable)

        Dim count As Integer = 0

        Dim table_cnt As Integer = 0

        Dim dt As DataTable

        For Each dr As DataRow In dtclone.Rows
            rowcheck = False
            'If (count Mod numberOfRecords = 0) Then

            If (dr.Item("From").ToString().Equals("")) Or count = 0 Then
                rowcheck = True
                If dtclone.Rows.Count.Equals(count) Then Exit Function
                dt = New DataTable()
                dt = dtclone.Clone()
                dt.TableName = "dtCustomer" + table_cnt.ToString()
                dttables.Add(dt)
                table_cnt += 1
            End If

            If rowcheck = False Or count = 0 Then
                dt.ImportRow(dr)
            End If
            count += 1
        Next

        Return dttables

    End Function




    Public Sub Initialize()

        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        Dim i As Int16, j As Int16

        'xlApp = New Excel.ApplicationClass

        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        For i = 0 To Form1.DGridCEI.RowCount - 2
            For j = 0 To Form1.DGridCEI.ColumnCount - 1
                xlWorkSheet.Cells(i + 1, j + 1) = Form1.DGridCEI(j, i).Value.ToString()
            Next
        Next
        xlApp.Visible = True
    End Sub
    Sub Export_MHI()
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim ncol As Integer = 1
        Dim i As Int16, j As Int16
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        xlWorkSheet.Cells(1, 1) = "From"
        xlWorkSheet.Cells(1, 2) = "Thru"
        xlWorkSheet.Cells(1, 3) = "Svc"
        xlWorkSheet.Cells(1, 4) = "PS"
        xlWorkSheet.Cells(1, 5) = "NBR"
        xlWorkSheet.Cells(1, 6) = "OV"
        xlWorkSheet.Cells(1, 7) = "P"
        xlWorkSheet.Cells(1, 8) = "N"
        xlWorkSheet.Cells(1, 9) = "RC"
        xlWorkSheet.Cells(1, 10) = "Charge"
        xlWorkSheet.Cells(1, 11) = "Not Cov"
        xlWorkSheet.Cells(1, 12) = "B/M"
        xlWorkSheet.Cells(1, 13) = "Covered"
        xlWorkSheet.Cells(1, 14) = "Deduct"
        xlWorkSheet.Cells(1, 15) = "D"
        xlWorkSheet.Cells(1, 16) = "%"
        xlWorkSheet.Cells(1, 17) = "Paid"
        xlWorkSheet.Cells(1, 18) = "S"
        xlWorkSheet.Cells(1, 19) = "D\C"
        xlWorkSheet.Cells(1, 20) = "Sanc"
        xlWorkSheet.Cells(1, 21) = "Cause Code"
        xlWorkSheet.Cells(1, 22) = "P1"
        xlWorkSheet.Cells(1, 23) = "Tin"
        xlWorkSheet.Cells(1, 24) = "Suffix"
        xlWorkSheet.Cells(1, 25) = "Claim Number"
        xlWorkSheet.Cells(1, 26) = "Draft"
        xlWorkSheet.Cells(1, 27) = "Proc Date"
        xlWorkSheet.Cells(1, 28) = "Adj#"
        xlWorkSheet.Cells(1, 29) = "Total Billed"
        xlWorkSheet.Cells(1, 30) = "Total Paid"
        xlWorkSheet.Cells(1, 31) = "ICN"
        xlWorkSheet.Cells(1, 32) = "Suf"
        xlWorkSheet.Cells(1, 33) = "FLN"
        xlWorkSheet.Cells(1, 34) = "PRS"
        xlWorkSheet.Cells(1, 35) = "SI"
        xlWorkSheet.Cells(1, 36) = "PT Name"
        xlWorkSheet.Cells(1, 37) = "."
        xlWorkSheet.Cells(1, 38) = "PT Rel"
        xlWorkSheet.Cells(1, 39) = "PT Name"
        xlWorkSheet.Cells(1, 40) = "INN Ded"
        xlWorkSheet.Cells(1, 41) = "INN OOP"
        xlWorkSheet.Cells(1, 42) = "OON Ded"
        xlWorkSheet.Cells(1, 43) = "OON OOP"
        xlWorkSheet.Cells(1, 44) = "INN Ded"
        xlWorkSheet.Cells(1, 45) = "INN OOP"
        xlWorkSheet.Cells(1, 46) = "OON Ded"
        xlWorkSheet.Cells(1, 47) = "OON OOP"
        xlWorkSheet.Cells(1, 48) = "OI/OIM"
        xlWorkSheet.Cells(1, 49) = "OOP Calc Run"
        xlWorkSheet.Cells(1, 50) = "ICN and Suffix"
        xlWorkSheet.Cells(1, 51) = "Inp Facility"
        xlWorkSheet.Cells(1, 52) = "Provider Name"
        xlWorkSheet.Cells(1, 53) = "Provider Type"
        xlWorkSheet.Cells(1, 54) = "M1"
        xlWorkSheet.Cells(1, 55) = "M2"
        xlWorkSheet.Cells(1, 56) = "M3"
        xlWorkSheet.Cells(1, 57) = "M4"

        For i = 0 To Form1.tblMHI.RowCount - 2
            For j = 0 To Form1.tblMHI.ColumnCount - 2
                xlWorkSheet.Cells(i + 2, j + 1) = Form1.tblMHI(j, i).Value.ToString()
            Next j
        Next i

        xlApp.Visible = True
    End Sub

    Sub Export_MMI_OVERVIEW()

        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim ncol As Integer = 1
        Dim i As Int16, j As Int16
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        xlWorkSheet.Cells(1, 1) = "A"
        xlWorkSheet.Cells(1, 2) = "B"
        xlWorkSheet.Cells(1, 3) = "C"
        xlWorkSheet.Cells(1, 4) = "D"
        xlWorkSheet.Cells(1, 5) = "E"
        xlWorkSheet.Cells(1, 6) = "F"
        xlWorkSheet.Cells(1, 7) = "G"
        xlWorkSheet.Cells(1, 8) = "H"
        xlWorkSheet.Cells(1, 9) = "I"
        xlWorkSheet.Cells(1, 10) = "J"
        xlWorkSheet.Cells(1, 11) = "K"
        xlWorkSheet.Cells(1, 12) = "L"
        xlWorkSheet.Cells(1, 13) = "M"

        Dim lastcolumn As Integer

        For j = 0 To Form1.DGridOverview.ColumnCount - 2

            If isNullOrEmpty(Form1.DGridOverview.Rows(2).Cells(j).Value) Then
                Exit For
            Else
                lastcolumn = j
                'Exit For
            End If
        Next

        For i = 0 To Form1.DGridOverview.RowCount - 2
            If isNullOrEmpty(Form1.DGridOverview.Rows(i).Cells(0).Value) Then
                i = i + 1
            End If
            For j = 0 To lastcolumn

                If isNullOrEmpty(Form1.DGridOverview(j, i).Value) Then
                Else
                    xlWorkSheet.Cells(i + 2, j + 1) = Form1.DGridOverview(j, i).Value.ToString()
                End If

            Next j
        Next i


        xlApp.Visible = True
    End Sub

    Sub OOPSpreadSheet_Export()
        Dim xlApp As New Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim ncol As Integer = 1
        Dim i As Int16, j As Int16
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")

        xlWorkSheet.Cells(1, 1) = "From"
        xlWorkSheet.Cells(1, 2) = "Thru"
        xlWorkSheet.Cells(1, 3) = "Svc"
        xlWorkSheet.Cells(1, 4) = "PS"
        xlWorkSheet.Cells(1, 5) = "NBR"
        xlWorkSheet.Cells(1, 6) = "OV"
        xlWorkSheet.Cells(1, 7) = "P"
        xlWorkSheet.Cells(1, 8) = "N"
        xlWorkSheet.Cells(1, 9) = "RC"
        xlWorkSheet.Cells(1, 10) = "Charge"
        xlWorkSheet.Cells(1, 11) = "Not Cov"
        xlWorkSheet.Cells(1, 12) = "B/M"
        xlWorkSheet.Cells(1, 13) = "Covered"
        xlWorkSheet.Cells(1, 14) = "Deduct"
        xlWorkSheet.Cells(1, 15) = "D"

        For i = 5 To Form1.tblOOP.RowCount - 2
            For j = 0 To Form1.tblOOP.ColumnCount - 2
                'xlWorkSheet.Cells(i + 2, j + 1) = Form1.tblOOP(j, i).Value.ToString()
                'Code upated by Sumit Singh
                If Form1.tblOOP(j, i).Value IsNot Nothing Then
                    xlWorkSheet.Cells(i + 2, j + 1) = Form1.tblOOP(j, i).Value.ToString()
                Else
                    xlWorkSheet.Cells(i + 2, j + 1) = ""
                End If

            Next j
        Next i
        xlApp.Visible = True
    End Sub

    'Form1.TabControl1.SelectedIndex = 1

    'With Form1.DGridOverview
    '.Rows(0).Cells(1).Value = "Sanjeet"
    '.Rows(0).Cells(1).Value = Form1.DGridMInfo.Rows(1).Cells(1).Value
    'End With


    'Form1.DGrid_PG5.Columns(0).Width = 220



    'Sub MMI_Page6()


    '    With Form1.DGrid_PG6
    '        Form1.DGrid_PG6.Columns(0).Width = (220)

    '        .Rows.Add(51)

    '        .Rows(0).Cells(0).Value = "Policy"
    '        .Rows(1).Cells(0).Value = "Plan Code/Reporting Code/Plan Var"
    '        .Rows(2).Cells(0).Value = "Year"
    '        .Rows(3).Cells(0).Value = "Patient Name"
    '        .Rows(4).Cells(0).Value = "(DED5)  Deductible Description"
    '        .Rows(5).Cells(0).Value = "(DED5) Network Type Ded"
    '        .Rows(6).Cells(0).Value = "(DED5)  Cost Containment Deductible"
    '        .Rows(7).Cells(0).Value = "(DED5)  Deductible Type Code"
    '        .Rows(8).Cells(0).Value = "(DED5)  Deductible Amount"
    '        .Rows(9).Cells(0).Value = "(DED5)  Frequency Code"
    '        .Rows(10).Cells(0).Value = "(DED5)  Deductible Benefit Period"
    '        .Rows(11).Cells(0).Value = "(DED5)  Deductible Carry-Over Code"
    '        .Rows(12).Cells(0).Value = "(DED5)  COB Deductible Exclusion"
    '        .Rows(13).Cells(0).Value = "(DED5)  X Semi Private Rate"
    '        .Rows(14).Cells(0).Value = "(DED5)  Deductible Accumulation Code"
    '        .Rows(15).Cells(0).Value = "(DED5)  Deductible Accumulation Period"
    '        .Rows(16).Cells(0).Value = "(DED5)  Deductible Maintenance Code"
    '        .Rows(17).Cells(0).Value = "(DED5)  Deductible Maintenance Period"
    '        .Rows(18).Cells(0).Value = "(DED5)  Deductible Maintenance Amount"
    '        .Rows(19).Cells(0).Value = "(DED6)  Deductible Description"
    '        .Rows(20).Cells(0).Value = "(DED6) Network Type Ded"
    '        .Rows(21).Cells(0).Value = "(DED6)  Cost Containment Deductible"
    '        .Rows(22).Cells(0).Value = "(DED6)  Deductible Type Code"
    '        .Rows(23).Cells(0).Value = "(DED6)  Deductible Amount"
    '        .Rows(24).Cells(0).Value = "(DED6)  Frequency Code"
    '        .Rows(25).Cells(0).Value = "(DED6)  Deductible Benefit Period"
    '        .Rows(26).Cells(0).Value = "(DED6)  Deductible Carry-Over Code"
    '        .Rows(27).Cells(0).Value = "(DED6)  COB Deductible Exclusion"
    '        .Rows(28).Cells(0).Value = "(DED6)  X Semi Private Rate"
    '        .Rows(29).Cells(0).Value = "(DED6)  Deductible Accumulation Code"
    '        .Rows(30).Cells(0).Value = "(DED6)  Deductible Accumulation Period"
    '        .Rows(31).Cells(0).Value = "(DED6)  Deductible Maintenance Code"
    '        .Rows(32).Cells(0).Value = "(DED6)  Deductible Maintenance Period"
    '        .Rows(33).Cells(0).Value = "(DED6)  Deductible Maintenance Amount"
    '        .Rows(34).Cells(0).Value = "(DED3) Family Deductible Description"
    '        .Rows(35).Cells(0).Value = "(DED3) Cost Containment Deductibles"
    '        .Rows(36).Cells(0).Value = "(DED3) Family Deductible Amount"
    '        .Rows(37).Cells(0).Value = "(DED3) Family Deductible Frequency"
    '        .Rows(38).Cells(0).Value = "(DED3) Family Deductible Carry Over"
    '        .Rows(39).Cells(0).Value = "(DED3) Family Deductible Multiplier"
    '        .Rows(40).Cells(0).Value = "(DED3) Family Deductible s"
    '        .Rows(41).Cells(0).Value = "HRA Access Point"
    '        .Rows(42).Cells(0).Value = "Family HRA Access Point"
    '        .Rows(43).Cells(0).Value = "Deductible Credit Type 1"
    '        .Rows(44).Cells(0).Value = "Deductible Credit Amount 1"
    '        .Rows(45).Cells(0).Value = "Deductible Credit Type 2"
    '        .Rows(46).Cells(0).Value = "Deductible Credit Amount 2"
    '        .Rows(47).Cells(0).Value = "Deductible Credit Type 3"
    '        .Rows(48).Cells(0).Value = "Deductible Credit Amount 3"
    '        .Rows(49).Cells(0).Value = "Deductible Credit Type 4"
    '        .Rows(50).Cells(0).Value = "Deductible Credit Amount 4"

    '    End With

    '    'Form1.RichTextBox1.AppendText("Completed to gather data for MMI Page 6")
    '    Form1.RichTextBox1.AppendText("Completed to gather data for MMI Page 6   " & vbCrLf)
    '    Form1.RichTextBox1.SelectionBullet = True
    '    Form1.RichTextBox1.SelectionIndent = 5
    '    Form1.RichTextBox1.BulletIndent = 4

    'End Sub

    'Sub MMI_Page8()
    '    With Form1.DGrid_PG8

    '        Form1.DGrid_PG8.Columns(0).Width = (220)


    '        .Rows.Add(718)

    '        .Rows(0).Cells(0).Value = "Policy"
    '        .Rows(1).Cells(0).Value = "Plan Code/Reporting Code/Plan Var"
    '        .Rows(2).Cells(0).Value = "Year"
    '        .Rows(3).Cells(0).Value = "Patient Name"
    '        .Rows(4).Cells(0).Value = "Group Table Number"
    '        .Rows(5).Cells(0).Value = "RECIPROCITY TABLE NUMBER"
    '        .Rows(6).Cells(0).Value = "Captiation Exclusion Indicator"
    '        .Rows(7).Cells(0).Value = "Managed Psych"
    '        .Rows(8).Cells(0).Value = "Excluded Provider Processing"
    '        .Rows(9).Cells(0).Value = "Non-Preferred Provider Processing"
    '        .Rows(10).Cells(0).Value = "Base Covered"
    '        .Rows(11).Cells(0).Value = "Obtained or Not Obtained (1)"
    '        .Rows(12).Cells(0).Value = "Base Percent (1)"
    '        .Rows(13).Cells(0).Value = "MM1 Covered"
    '        .Rows(14).Cells(0).Value = "Obtained or Not Obtained (2)"
    '        .Rows(15).Cells(0).Value = "Base Percent (2)"
    '        .Rows(16).Cells(0).Value = "New Coinsurance Apply Indicator (1)"
    '        .Rows(17).Cells(0).Value = "Deductible Description (1)"
    '        .Rows(18).Cells(0).Value = "Emergency Indicator"
    '        .Rows(19).Cells(0).Value = "Full Service Indicator"
    '        .Rows(20).Cells(0).Value = "FAMILY PLANNING INDICATOR"
    '        .Rows(21).Cells(0).Value = "Dollar Tolerance Indicator"
    '        .Rows(22).Cells(0).Value = "Managed Care Preferred Provider Processing Indicator"
    '        .Rows(23).Cells(0).Value = "Base Covered Percent or Coinsurance Percent"
    '        .Rows(24).Cells(0).Value = "Obtained or Not Obtained (3)"
    '        .Rows(25).Cells(0).Value = "Base Percent (3)"
    '        .Rows(26).Cells(0).Value = "MM2 COV-NC"
    '        .Rows(27).Cells(0).Value = "Obtained or Not Obtained (4)"
    '        .Rows(28).Cells(0).Value = "Base Percent (4)"
    '        .Rows(29).Cells(0).Value = "New Coinsurance Apply Indicator (2)"
    '        .Rows(30).Cells(0).Value = "Deductible Description (2)"
    '        .Rows(31).Cells(0).Value = "Out of Area"
    '        .Rows(32).Cells(0).Value = "Member Network Key"
    '        .Rows(33).Cells(0).Value = "ALTERNATE MARKET TYPE"
    '        .Rows(34).Cells(0).Value = "PCP COPAY INDICATOR"
    '        .Rows(35).Cells(0).Value = "NON-PCP COPAY AMOUNT 1"
    '        .Rows(36).Cells(0).Value = "Co-Pay Amount"
    '        .Rows(37).Cells(0).Value = "Co-Pay Maximum"
    '        .Rows(38).Cells(0).Value = "URGENT CARE AMOUNT 1"
    '        .Rows(39).Cells(0).Value = "OOP INDICATOR"
    '        .Rows(40).Cells(0).Value = "PPO Mental and Nervous Exclude"
    '        .Rows(41).Cells(0).Value = "PARS Mental and Nervous Exclude 1"
    '        .Rows(42).Cells(0).Value = "PARS Mental and Nervous Exclude 2"
    '        .Rows(43).Cells(0).Value = "Provision Order"
    '        .Rows(44).Cells(0).Value = "Cap Minimum Percent"
    '        .Rows(45).Cells(0).Value = "Contiguous market override"
    '        .Rows(46).Cells(0).Value = "PPO Cap Amount"
    '        .Rows(47).Cells(0).Value = "TIER 1 URGENT CARE AMOUNT"
    '        .Rows(48).Cells(0).Value = "TIER 1 URGENT CARE OOP INDICATOR"
    '        .Rows(49).Cells(0).Value = "MNNRP"
    '        .Rows(50).Cells(0).Value = "MNNRP Percent"
    '        .Rows(51).Cells(0).Value = "Facility MNNRP PCT"
    '        .Rows(52).Cells(0).Value = "Coalition of America Applies?"
    '        .Rows(53).Cells(0).Value = "OPT OUT UBH TIER"
    '        .Rows(54).Cells(0).Value = "PHY SSP"
    '        .Rows(55).Cells(0).Value = "Course of Treatment days"
    '        .Rows(56).Cells(0).Value = "IH COPAY TYPE"
    '        .Rows(57).Cells(0).Value = "Direct Access"
    '        .Rows(58).Cells(0).Value = "UBH Care Mgmt Indicator"
    '        .Rows(59).Cells(0).Value = "Inpatient Copay Days"
    '        .Rows(60).Cells(0).Value = "UBH office visit maximum remark code"
    '        .Rows(61).Cells(0).Value = "Medication Management Indicator"
    '        .Rows(62).Cells(0).Value = "OB/GYN PCP COPAY"
    '        .Rows(63).Cells(0).Value = "COIN/COPAY"
    '        .Rows(65).Cells(0).Value = "IPLAN"
    '        .Rows(66).Cells(0).Value = "PCP%"
    '        .Rows(67).Cells(0).Value = "Travel benefit pre-auth table number"
    '        .Rows(68).Cells(0).Value = "American Chiro Ntwork"
    '        .Rows(69).Cells(0).Value = "COPAY ID"
    '        .Rows(70).Cells(0).Value = "Travel benefit pre-auth market number"
    '        .Rows(71).Cells(0).Value = "COPAY WAIVE TABLE ID"
    '        .Rows(72).Cells(0).Value = "TIER 1 COPAY AMOUNT"
    '        .Rows(73).Cells(0).Value = "ENRP Non Network Emergency Reimbursement"
    '        .Rows(74).Cells(0).Value = "ENRP Non Network Non-Emergency/Gap Reimbursement"
    '        .Rows(75).Cells(0).Value = "Non ER ERNP & GA ENRP Reimbursement"
    '        .Rows(76).Cells(0).Value = "Extended Non Network Emergency Reimbursement"

    '    End With

    '    'Form1.RichTextBox1.AppendText("Completed to gather data for MMI Page 8")
    '    Form1.RichTextBox1.AppendText("Completed to gather data for MMI Page 8   " & vbCrLf)
    '    Form1.RichTextBox1.SelectionBullet = True
    '    Form1.RichTextBox1.SelectionIndent = 5
    '    Form1.RichTextBox1.BulletIndent = 4
    'End Sub




End Module








