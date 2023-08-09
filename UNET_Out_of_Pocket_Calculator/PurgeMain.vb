Imports System.Security.Claims
Imports DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing

Imports System.Collections
Imports System.Collections.Generic


Public Class Purgemain
    Public output As DataTable
    Public output2 As DataTable
    Public cnt_output As Integer = 1
    Public Property Final_dt_purge As DataTable

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        Form1.RichTextBox1.AppendText("Gathering data from Purge History..." & vbCrLf)
        Form1.RichTextBox1.SelectionBullet = True
        Form1.RichTextBox1.SelectionIndent = 5
        Form1.RichTextBox1.BulletIndent = 4



        Dim rwcount As Integer = Form1.DGridMHI.Rows.Count - 1
        Dim startTime As Date = DateTime.Now



        Call RefreshLblServer()



        txtEmployeeID.Text = txtEmployeeID.Text.Trim()



        If String.IsNullOrEmpty(txtEmployeeID.Text) Then
            MsgBox("You must fill in the Employee Id field.", vbInformation, "Validation")
        ElseIf txtEmployeeID.Text.Length <> 10 OrElse LCase(txtEmployeeID.Text.Substring(0, 1)) <> "s" Then
            MsgBox("The Employee Id must be 9 digits and be preceeded by an ""s"".", vbInformation, "Validation")
        Else
            startTime = DateTime.Now
            'Debug.Print("")
            'Debug.Print("")
            'Debug.Print(DateTime.Now.ToString() & ", Start parsing " & txtEmployeeID.Text)
            output = New DataTable()
            ' output.Columns.Add("S.no") :
            output.Columns.Add("From") : output.Columns.Add("Thru") : output.Columns.Add("Svc") : output.Columns.Add("PS") : output.Columns.Add("Nbr") : output.Columns.Add("OV")
            output.Columns.Add("P") : output.Columns.Add("N") : output.Columns.Add("RC") : output.Columns.Add("Charge") : output.Columns.Add("NotCov") : output.Columns.Add("BM")
            output.Columns.Add("Covered") : output.Columns.Add("Deduct") : output.Columns.Add("D") : output.Columns.Add("Perc") : output.Columns.Add("Paid") : output.Columns.Add("S")
            output.Columns.Add("DC") : output.Columns.Add("Sanc") : output.Columns.Add("CauseCode") : output.Columns.Add("P1") : output.Columns.Add("TIN") : output.Columns.Add("Suffix") : output.Columns.Add("ClaimNumber")
            output.Columns.Add("Draft") : output.Columns.Add("ProcDate") : output.Columns.Add("Adjno") : output.Columns.Add("TotalBilled") : output.Columns.Add("TotalPaid") : output.Columns.Add("ICN") : output.Columns.Add("Suf")
            output.Columns.Add("FLN") : output.Columns.Add("PRS") : output.Columns.Add("SI") : output.Columns.Add("ptname") : output.Columns.Add("DOT") : output.Columns.Add("PT_Rel")
            output.Columns.Add("PT_Name") : output.Columns.Add("Compound_Doc") : output.Columns.Add("Page_Cnt") : output.Columns.Add("Line_Cnt")


            Start(empID:=txtEmployeeID.Text, firstName:=txtFirstName.Text, inclBoneyard:=chkInclBoneyard.Checked, maintainRecordInfo:=chkMaintainRecordInfo.Checked _
               , inclRio:=Not chkExcludeRio.Checked, skipParsing:=chkSkipParsing.Checked, restFindBatchSize:=txtRestFindBatchSize.Text)
            'Dim OBJ As New outputgrid
            'OBJ.Visible = True
            'OBJ.Enabled = True
            'OBJ.DataGridView1.DataSource = output
            'Debug.Print(DateTime.Now.ToString() & ", Done, Duration " & Format(DateTime.Now - startTime, "hh:mm:ss"))
            '            Debug.Print("")
            '           Debug.Print("")



        End If



        'Dim strvalue As String = output.Rows(0).Item(0)
        Dim intRw As Integer 'PurgeStartDT.Text



        'Dim result() As DataRow = output.Select("From >= '" + Convert.ToDateTime(PurgeStartDT.Value.ToString("MM/dd/yyyy")) + "' And Thru <= '" + Convert.ToDateTime(PurgeEndDT.Value.ToString("MM/dd/yyyy")) + "'")



        Dim memlist As New ArrayList(Form1.memberList.CheckedItems)
        Dim mem_names As String = ""
        Dim mem_name As Boolean = False
        For Each item In memlist
            mem_names = mem_names & item & ";"
            'mm.Append(item)
        Next



        output2 = New DataTable
        output2 = output.Clone



        For Each row As DataRow In output.Rows



            For Each item In memlist
                If item.ToString().Contains("/") Then
                    If Trim(Replace(item, ";", "")) = Trim(row.Item("PT_Name")) Then
                        mem_name = True
                        Exit For
                    End If
                Else
                    If Trim(Replace(item, ";", "")) = Trim(row.Item("ptname")) Then
                        mem_name = True
                        Exit For
                    End If
                End If
            Next
            If mem_name Then

                If row.Item("From") >= Convert.ToDateTime(PurgeStartDT.Value.ToString("MM/dd/yyyy")) And row.Item("Thru") <= Convert.ToDateTime(PurgeEndDT.Value.ToString("MM/dd/yyyy")) Then
                    ' Dim newRow As DataRow = output2.NewRow()
                    'newRow = row
                    output2.ImportRow(row)
                End If
            End If
            mem_name = False
        Next




        Dim dt_n As New System.Data.DataTable
        dt_n = Form1.dgridmhi_()
        'dt_n = output2.Clone
        For Each row As DataRow In output2.Rows
            dt_n.ImportRow(row)
        Next
        For i As Integer = 0 To dt_n.Rows.Count - 1
            If dt_n.Rows(i).Item("D").ToString() = "" Then
                dt_n.Rows(i).Item("D") = "0"
            End If
        Next
        Dim duplicate As New System.Data.DataTable
        duplicate = dt_n.Clone
        Dim cnt_ As Integer = 0
        Try
            For i As Integer = 0 To dt_n.Rows.Count - 1
                If dt_n.Rows(i).Item("From").ToString().Trim().Equals("") Then
                    dt_n.Rows.RemoveAt(i)
                    cnt_ += 1
                    i -= cnt_
                End If
            Next
        Catch ex As Exception

        End Try


        Dim dtview As New DataView(dt_n)
        dtview.Sort = "ptname asc"
        'dtview.Sort = "ProcDate Asc,ICNandSuffix Asc,ProcDate Asc"
        Dim dt_dt As DataTable = dtview.ToTable()



        For i As Integer = 0 To dt_dt.Rows.Count - 1

            If (i + 1) = dt_dt.Rows.Count Then
                duplicate.ImportRow(dt_dt.Rows(i))
                Exit For
            End If

            If dt_dt.Rows(i).Item("ptname").ToString().Trim() < dt_dt.Rows(i + 1).Item("ptname").ToString().Trim() Then
                Dim new_ As DataRow = duplicate.NewRow()
                'new_ = Final_dt_.Rows(q)
                duplicate.Rows.Add(new_)
            Else
                duplicate.ImportRow(dt_dt.Rows(i))
            End If
        Next

        Dim duplicate2 As New System.Data.DataTable
        duplicate2 = duplicate.Clone
        For i As Integer = 0 To duplicate.Rows.Count - 1

            If (i + 1) = duplicate.Rows.Count Then
                duplicate2.ImportRow(duplicate.Rows(i))
                Exit For
            End If

            If IsDBNull(duplicate.Rows(i).Item(0)) And i <> duplicate.Rows.Count - 1 Then

                duplicate2.ImportRow(duplicate.Rows(i))
                'q = q + 1
                Continue For

            End If

            If i = duplicate.Rows.Count - 1 Then
                duplicate2.ImportRow(duplicate.Rows(i))
                Exit For
            End If


            If IsDBNull(duplicate.Rows(i + 1).Item(0)) Then
                duplicate2.ImportRow(duplicate.Rows(i))
                Continue For
            End If


            If Strings.Right(duplicate.Rows(i).Item("From"), 4).ToString().Trim() <> Strings.Right(duplicate.Rows(i + 1).Item("From"), 4).ToString().Trim() Then
                Dim new_ As DataRow = duplicate2.NewRow()
                'new_ = Final_dt_.Rows(q)
                duplicate2.Rows.Add(new_)
            Else
                duplicate2.ImportRow(duplicate.Rows(i))
            End If
        Next


        Dim tblCount As Integer = 0

        Dim dttables_ As New List(Of DataTable)
        dttables_ = list(duplicate2, duplicate2.Rows.Count)
        Dim dt_count As Integer = dttables_.Count
        'Form1.tblMHI.DataSource = dttables_

        Final_dt_purge = New DataTable
        Final_dt_purge = duplicate2.Clone
        Dim cnt As Integer
        For Each datatable As DataTable In dttables_
            tblCount += 1
            Dim dtview_ As New DataView(datatable)
            dtview_.Sort = "From Asc"
            'dtview.Sort = "ProcDate Asc,ICNandSuffix Asc,ProcDate Asc"
            Dim dt_dt_ As DataTable = dtview_.ToTable()
            cnt = 0
            For Each rw As DataRow In dt_dt_.Rows

                If dt_dt_.Rows.Count - 1 = cnt And tblCount <> dt_count Then
                    Final_dt_purge.ImportRow(rw)
                    rw = Final_dt_purge.NewRow()
                    Final_dt_purge.Rows.Add(rw)
                Else
                    Final_dt_purge.ImportRow(rw)
                End If
                cnt += 1
            Next
        Next

        ''Form1.tblMHI.DataSource = Final_dt_main

        'Form1.tblMHI.DataSource = Final_dt_purge

        Call getMHI(Final_dt_purge)
        'Me.Show()

        Form1.RichTextBox1.SelectionIndent = 5
        Form1.RichTextBox1.BulletIndent = 4
        Form1.RichTextBox1.SelectionBullet = True
        Form1.RichTextBox1.AppendText("Data fetched from purged History" & vbCrLf)


        'Try



        '    Form1.DGridMHI.Rows.Add(output2.Rows.Count - 1)
        '    For I = 0 To output2.Rows.Count - 1



        '        For J = 0 To 38
        '            Form1.DGridMHI.Rows(rwcount).Cells(J).Value = output2.Rows(I).Item(J)
        '        Next J
        '        rwcount = rwcount + 1
        '    Next



        '    Call getMHI()
        '    Me.Show()



        '    Form1.RichTextBox1.AppendText("Data fetched from Purge History..." & vbCrLf)
        '    Form1.RichTextBox1.SelectionBullet = True
        '    Form1.RichTextBox1.SelectionIndent = 5
        '    Form1.RichTextBox1.BulletIndent = 4



        'Catch ex As Exception
        '    MessageBox.Show("No information found for Purge History", "OOP Calculator", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'End Try
    End Sub
    Private Sub RefreshLblServer()
        Dim server As String = ""
        If IsDoc360Production(server) Then
            lblTestServer.Visible = False
        Else
            lblTestServer.Text = """" & server & """ TEST SERVER"
            lblTestServer.Visible = True
        End If
    End Sub



    Private Sub Purgemain_Load(sender As Object, e As EventArgs) Handles MyBase.Load



    End Sub
End Class