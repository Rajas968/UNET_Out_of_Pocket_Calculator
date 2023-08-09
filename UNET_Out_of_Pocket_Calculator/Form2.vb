
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports Newtonsoft.Json
Imports Microsoft.Office.Interop.Word


Public Class Form2
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = 780
        Threading.Thread.Sleep(500)
        'CorsdAddrs.Text = Form1.DGridMInfo.Rows(0).Cells(1).Value          ''member address
        Dim strMem As String
        Threading.Thread.Sleep(500)
        strMem = Trim(Form1.DGridCEI.Rows(0).Cells(0).Value)
        strMem = strMem & " " & Trim(Form1.DGridMInfo.Rows(0).Cells(0).Value)
        '
        Threading.Thread.Sleep(500)
        Dim strMember As String = strMem
        Threading.Thread.Sleep(500)
        'txtMember_Name.Text = StrConv(strMember.ToString, VbStrConv.ProperCase)
        txtMember_Name.Text = strMem
        Threading.Thread.Sleep(100)
        txtMemberAddrs1.Text = Trim(Form1.DGridMInfo.Rows(0).Cells(1).Value)          ''member address
        txtMemberCSZ.Text = Trim(Form1.DGridMInfo.Rows(1).Cells(1).Value)
        'txtMDEDamt.Text = Form1.DGridOverview.Rows(13).Cells(1).Value
        'txtDEDMax.Text = Form1.DGridOverview.Rows(14).Cells(1).Value
        'txtOOPMetDate.Text = Form1.DGridOverview.Rows(15).Cells(1).Value

        txtMemberFirstname.Text = Trim(Form1.DGridCEI.Rows(0).Cells(0).Value)
        txtPatientName.Text = Trim(Form1.DGridCEI.Rows(0).Cells(0).Value) & " " & Trim(Form1.DGridMInfo.Rows(0).Cells(0).Value)
        txtMember_Name.Text = Trim(Form1.DGridCEI.Rows(0).Cells(0).Value) & " " & Trim(Form1.DGridMInfo.Rows(0).Cells(0).Value)
        txtDEDMetDate.Text = "MM/DD/YYYY"

        ''Proper cases
        txtMember_Name.Text = Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtMember_Name.Text.ToLower)
        txtPatientName.Text = Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtPatientName.Text.ToLower)
        txtMemberFirstname.Text = Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtMemberFirstname.Text.ToLower)


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim word_app As Word._Application = New Word.Application
        Dim strORS As String

        word_app.Visible = True


        Dim Mydocpath As String

        Dim mydocpath2 As String
        mydocpath2 = ""

        If Trim(ComboBox2.Text) = "" Then
            MsgBox("Please select any Letter Type and re-submit")
            Exit Sub
        End If

        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MsgBox("Please select any of Family or Individual and re-submit")
            Exit Sub
        End If

        Try
            Dim inputletter As String
            Dim word_doc As Word._Document
            Dim para As Word.Paragraph
            Dim mydocpath1 As String
            inputletter = ComboBox2.Text

            If inputletter = "Caterpillar DED OOP Benefits Due" Or inputletter = "Caterpillar DED OOP No Benefits Due" Or inputletter = "Caterpillar OOP ONLY Benefits Due" Or inputletter = "Caterpillar OOP ONLY No Benefits Due" Or inputletter = "DED ONLY Benefits Due" Or inputletter = "DED ONLY No Benefits Due" _
             Or inputletter = "DED OOP Benefits Due" Or inputletter = "DED OOP No Benefits Due" Or inputletter = "OOP ONLY Benefits Due" Or inputletter = "OOP ONLY No Benefits Due" Then

                Mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\" + "Insight Software\Macro Express\Macro Files\NET\OOPCalculator" + "\" + "Caterpillar  DED OOP  benefits due ls.docx"
                'Dim word_doc As Word._Document = word_app.Documents.Add("C:\Users\kramesh2\Documents\Letter_Draft.docx")
                word_doc = word_app.Documents.Add(Mydocpath)
                'word_app.Documents.Add("C:\Users\kramesh2\Documents\Letter_Draft.docx")
                para = word_doc.Paragraphs.Add()
                'word_doc.Content.Find.Execute(FindText:="<Plan Correspondence Address 1> ", ReplaceWith:=PCorsdAddrs.Text, Replace:=Word.WdReplace.wdReplaceAll)
                word_doc.Content.Find.Execute(FindText:="<Member Full Name>", ReplaceWith:=txtMember_Name.Text, Replace:=Word.WdReplace.wdReplaceAll)
                word_doc.Content.Find.Execute(FindText:="<Member Address 1>", ReplaceWith:=txtMemberAddrs1.Text, Replace:=Word.WdReplace.wdReplaceAll)
                'While word_doc.Content.Find.Execute(FindText:="  ", Wrap:=Word.WdFindWrap.wdFindContinue)
                '    word_doc.Content.Find.Execute(FindText:="  ", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                'End While
                If txtMemberAddrs2.Text <> "" Then
                    word_doc.Content.Find.Execute(FindText:="<Member Address 2>", ReplaceWith:=txtMemberAddrs2.Text, Replace:=Word.WdReplace.wdReplaceAll)
                Else
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("<Member Address 2>")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                End If
                If txtMemberAddrs3.Text <> "" Then
                    word_doc.Content.Find.Execute(FindText:="<Member Address 3>", ReplaceWith:=txtMemberAddrs3.Text, Replace:=Word.WdReplace.wdReplaceAll)
                Else
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("<Member Address 3>")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                End If

                word_doc.Content.Find.Execute(FindText:="<Member City State Zip>", ReplaceWith:=txtMemberCSZ.Text, Replace:=Word.WdReplace.wdReplaceAll)
                word_doc.Content.Find.Execute(FindText:="<Date>", ReplaceWith:=txtDate.Text, Replace:=Word.WdReplace.wdReplaceAll)
                word_doc.Content.Find.Execute(FindText:="<Member First Name>", ReplaceWith:=txtMemberFirstname.Text, Replace:=Word.WdReplace.wdReplaceAll)
                'word_doc.Content.Find.Execute(FindText:="<Patient Name>,", ReplaceWith:="John Smith", Replace:=Word.WdReplace.wdReplaceAll)
                'word_doc.Content.Find.Execute(FindText:="<family / individual>,", ReplaceWith:="Family", Replace:=Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<Year>"
                    .Replacement.Text = txtYear.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                'If inputletter = "Caterpillar DED OOP Benefits Due" Or inputletter = "Caterpillar DED OOP No Benefits Due" Or inputletter = "DED ONLY Benefits Due" Or inputletter = "DED ONLY No Benefits Due" Or inputletter = "DED OOP Benefits Due" Or inputletter = "DED OOP No Benefits Due" Then
                '    With word_doc.Application.Selection.Find
                '        .Text = "<DEDMet Date>"
                '        .Replacement.Text = txtDEDMetDate.Text
                '        .Forward = True
                '        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                '        .Format = False
                '        .MatchCase = False
                '        .MatchWholeWord = False
                '        .MatchWildcards = False
                '        .MatchSoundsLike = False
                '        .MatchAllWordForms = False
                '    End With
                '    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                If inputletter = "Caterpillar DED OOP Benefits Due" Or inputletter = "Caterpillar DED OOP No Benefits Due" Or inputletter = "DED ONLY Benefits Due" Or inputletter = "DED ONLY No Benefits Due" Or inputletter = "DED OOP Benefits Due" Or inputletter = "DED OOP No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                        .Text = "<DEDMet Date>"
                        .Replacement.Text = txtDEDMetDate.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    With word_doc.Application.Selection.Find
                        .Text = "<MDEDAmount>"
                        .Replacement.Text = txtMDEDamt.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    With word_doc.Application.Selection.Find
                        .Text = "<DEDMax>"
                        .Replacement.Text = txtDEDMax.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                End If
                If inputletter = "Caterpillar DED OOP Benefits Due" Or inputletter = "Caterpillar DED OOP No Benefits Due" Or inputletter = "Caterpillar OOP ONLY No Benefits Due" Or inputletter = "Caterpillar OOP ONLY Benefits Due" Or inputletter = "DED OOP Benefits Due" Or inputletter = "DED OOP No Benefits Due" Or inputletter = "OOP ONLY Benefits Due" Or inputletter = "OOP ONLY No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                        .Text = "<OOPMet Date>"
                        .Replacement.Text = txtOOPMetDate.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                    With word_doc.Application.Selection.Find
                        .Text = "<MOOPAmount>"
                        .Replacement.Text = txtMOOPamt.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    With word_doc.Application.Selection.Find
                        .Text = "<OOPMax>"
                        .Replacement.Text = txtOOPamt.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                End If
                If inputletter = "Caterpillar DED OOP Benefits Due" Or inputletter = "Caterpillar DED OOP No Benefits Due" Or inputletter = "Caterpillar OOP ONLY No Benefits Due" Or inputletter = "Caterpillar OOP ONLY Benefits Due" Then
                    With word_doc.Application.Selection.Find
                        .Text = "<FOOPMet Date>"
                        .Replacement.Text = txtFOOPMdate.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    With word_doc.Application.Selection.Find
                        .Text = "<MFOOPAmount>"
                        .Replacement.Text = txtFOOPAmt.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    With word_doc.Application.Selection.Find
                        .Text = "<FOOPMax>"
                        .Replacement.Text = txtFOOPMax.Text
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                End If
                With word_doc.Application.Selection.Find
                    If Trim(txtPhonenumber.Text) <> "" Then
                        strORS = "ORS#: "
                    End If
                    .Text = "<phone number>"
                    .Replacement.Text = strORS & txtPhonenumber.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
            End If
            'With word_doc.Application.Selection.Find
            '        .Text = "<phone number>"
            '        .Replacement.Text = txtPhonenumber.Text
            '        .Forward = True
            '        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
            '        .Format = False
            '        .MatchCase = False
            '        .MatchWholeWord = False
            '        .MatchWildcards = False
            '        .MatchSoundsLike = False
            '        .MatchAllWordForms = False
            '    End With
            '    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
            'word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

            If inputletter = "Caterpillar DED OOP Benefits Due" Then
                While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                    word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                End While

                If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If


                mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "Caterpillar  DED OOP  benefits due ls"

                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "Caterpillar  DED OOP  benefits due ls.pdf"
                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)













                ElseIf inputletter = "Caterpillar DED OOP No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                    .Text = "claims and updated"
                    .Replacement.Text = "claims and confirmed"
                    .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                    If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "Caterpillar  DED OOP no beneifts due ls"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)

                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveOptions.wdSaveChanges)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    'word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    'word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "Caterpillar  DED OOP no beneifts due ls.pdf"

                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)


                ElseIf inputletter = "Caterpillar OOP ONLY Benefits Due" Then
                    With word_doc.Application.Selection.Find
                    .Text = "<family / individual> deductible and Out-of-pocket summary Is"
                    .Replacement.Text = "<family / individual> Out-of-pocket summary Is"
                    .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("Deductible")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next

                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <DEDMet Date> you met $<MDEDAmount> Of your $<DEDMax> deductible")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                    If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If

                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "Caterpillar OOP ONLY benefits due ls"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    'word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    'word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "Caterpillar OOP ONLY benefits due ls.pdf"

                    'MsgBox(System.IO.File.Exists(mydocpath2))
                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)

                ElseIf inputletter = "Caterpillar OOP ONLY No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                    .Text = "claims For <Patient Name> and updated"
                    .Replacement.Text = "claims For <Patient Name> and confirmed"
                    .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    With word_doc.Application.Selection.Find
                    .Text = "<family / individual> deductible and Out-of-pocket summary Is"
                    .Replacement.Text = "<family / individual> Out-of-pocket summary Is"
                    .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("Deductible")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next

                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <DEDMet Date> you met $<MDEDAmount> Of your $<DEDMax> deductible")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                    If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "Caterpillar  OOP ONLY no benefits due ls"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    'word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    'word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "Caterpillar  OOP ONLY no benefits due ls.pdf"
                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)

                ElseIf inputletter = "DED ONLY Benefits Due" Then
                    With word_doc.Application.Selection.Find
                    .Text = "<family / individual> deductible and Out-of-pocket summary is"
                    .Replacement.Text = "<family / individual> deductible summary Is"
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Out-of-pocket")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next

                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <OOPMet Date> you met $<MOOPAmount> Of your $<OOPMax> maximum")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Federal Out-of-pocket")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <FOOPMet Date> you met $<MFOOPAmount> Of your $<FOOPMax> maximum")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                    If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "DED only  benefits due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    ''word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "DED only  benefits due.pdf"
                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)

                ElseIf inputletter = "DED ONLY No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                    .Text = "claims For <Patient Name> and updated"
                    .Replacement.Text = "claims For <Patient Name> and confirmed"
                    .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    With word_doc.Application.Selection.Find
                    .Text = "<family / individual> deductible and Out-of-pocket summary Is"
                    .Replacement.Text = "<family / individual> deductible summary Is"
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Out-of-pocket")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <OOPMet Date> you met $<MOOPAmount> Of your $<OOPMax> maximum")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Federal Out-of-pocket")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <FOOPMet Date> you met $<MFOOPAmount> Of your $<FOOPMax> maximum")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                    If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If

                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "DED only no beneifts due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    ''word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "DED only no beneifts due.pdf"
                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)

                ElseIf inputletter = "DED OOP Benefits Due" Then
                    While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                    If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Federal Out-of-pocket")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next

                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <FOOPMet Date> you met $<MFOOPAmount> Of your $<FOOPMax> maximum")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next


                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "DED OOP  benefits due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()

                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "DED OOP  benefits due.pdf"

                    'MsgBox(System.IO.File.Exists(mydocpath2))
                    If System.IO.File.Exists(mydocpath2) = True Then

                        System.IO.File.Delete(mydocpath2)
                    End If
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)
                ElseIf inputletter = "DED OOP No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                    .Text = "claims For <Patient Name> and updated"
                    .Replacement.Text = "claims For <Patient Name> and confirmed"
                    .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Federal Out-of-pocket")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <FOOPMet Date> you met $<MFOOPAmount> Of your $<FOOPMax> maximum")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                    If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If

                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "DED OOP no beneifts due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "DED OOP no beneifts due.pdf"

                    'MsgBox(System.IO.File.Exists(mydocpath2))
                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)

                ElseIf inputletter = "OOP ONLY Benefits Due" Then
                    With word_doc.Application.Selection.Find
                    .Text = "<family / individual> deductible and Out-of-pocket summary Is"
                    .Replacement.Text = "<family / individual> Out-of-pocket summary Is"
                    .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("Deductible")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next

                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <DEDMet Date> you met $<MDEDAmount> Of your $<DEDMax> deductible")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Federal Out-of-pocket")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)

                        If UCase(para.Range.Text.Trim) = UCase(Trim("On <FOOPMet Date> you met $<MFOOPAmount> Of your $<FOOPMax> maximum")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                    While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                    If RadioButton1.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    If RadioButton2.Checked = True Then
                        While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                            word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                        End While
                    End If
                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "OOP ONLY benefits due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "OOP ONLY benefits due.pdf"

                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)

                ElseIf inputletter = "OOP ONLY No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                    .Text = "claims For <Patient Name> and updated"
                    .Replacement.Text = "claims For <Patient Name> and confirmed"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<family / individual> deductible and out-Of-pocket summary Is"
                    .Replacement.Text = "<family / individual> out-Of-pocket summary Is"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Deductible")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next

                For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)

                    If UCase(para.Range.Text.Trim) = UCase(Trim("On <DEDMet Date> you met $<MDEDAmount> Of your $<DEDMax> deductible")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next
                For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)
                    If UCase(para.Range.Text.Trim) = UCase(Trim("Federal Out-of-pocket")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next
                For Each para In word_doc.Paragraphs
                    'Debug.Print(p.Range.Text)

                    If UCase(para.Range.Text.Trim) = UCase(Trim("On <FOOPMet Date> you met $<MFOOPAmount> Of your $<FOOPMax> maximum")) Then
                        para.Range.Delete()
                        Exit For
                    End If
                Next
                While word_doc.Content.Find.Execute(FindText:="<Patient Name>", Wrap:=Word.WdFindWrap.wdFindContinue)
                    word_doc.Content.Find.Execute(FindText:="<Patient Name>", ReplaceWith:=txtPatientName.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                End While
                If RadioButton1.Checked = True Then
                    While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton1.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                End If
                If RadioButton2.Checked = True Then
                    While word_doc.Content.Find.Execute(FindText:="<family / individual>", Wrap:=Word.WdFindWrap.wdFindContinue)
                        word_doc.Content.Find.Execute(FindText:="<family / individual>", ReplaceWith:=RadioButton2.Text, Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                    End While
                End If
                mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "OOP ONLY no benefits due"
                word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                ''word_doc.Close()

                'word_app.Quit()
                mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "OOP ONLY no benefits due.pdf"

                If System.IO.File.Exists(mydocpath2) = True Then
                    'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                    System.IO.File.Delete(mydocpath2)
                End If
                'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                Threading.Thread.Sleep(2000)
                word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                Threading.Thread.Sleep(1000)
                word_app.Quit()


                'Me.SendToBack()
                Threading.Thread.Sleep(4000)
                'SendKeys.Send("%fa")
                'SendKeys.SendWait("%fa")
                Threading.Thread.Sleep(1000)

            End If
            'End If

            If inputletter = "INF ONLY Benefits Due" Or inputletter = "INF ONLY No Benefits Due" Then
                Mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\" + "Insight Software\Macro Express\Macro Files\NET\OOPCalculator" + "\" + "INF only  benefits due.docx"
                word_doc = word_app.Documents.Add(Mydocpath)
                para = word_doc.Paragraphs.Add()
                'word_doc.Content.Find.Execute(FindText:="<Plan Correspondence Address 1> ", ReplaceWith:=PCorsdAddrs.Text, Replace:=Word.WdReplace.wdReplaceAll)
                word_doc.Content.Find.Execute(FindText:="<Member Full Name>", ReplaceWith:=txtMember_Name.Text, Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member Address 1>", ReplaceWith:=txtMemberAddrs1.Text, Replace:=Word.WdReplace.wdReplaceAll)
                If txtMemberAddrs2.Text <> "" Then
                    word_doc.Content.Find.Execute(FindText:="<Member Address 2>", ReplaceWith:=txtMemberAddrs2.Text, Replace:=Word.WdReplace.wdReplaceAll)
                Else
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("<Member Address 2>")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                End If
                If txtMemberAddrs3.Text Then
                    word_doc.Content.Find.Execute(FindText:="<Member Address 3>", ReplaceWith:=txtMemberAddrs3.Text, Replace:=Word.WdReplace.wdReplaceAll)
                Else
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("<Member Address 3>")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                End If

                word_doc.Content.Find.Execute(FindText:="<Member City State Zip>", ReplaceWith:=txtMemberCSZ.Text, Replace:=Word.WdReplace.wdReplaceAll)
                word_doc.Content.Find.Execute(FindText:="<Date>", ReplaceWith:=txtDate.Text, Replace:=Word.WdReplace.wdReplaceAll)
                word_doc.Content.Find.Execute(FindText:="<Member First Name>", ReplaceWith:=txtMemberFirstname.Text, Replace:=Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<INFMet Date>"
                    .Replacement.Text = txtINFDate.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<INFAmount>"
                    .Replacement.Text = txtINFAmt.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<INF Max>"
                    .Replacement.Text = txtINFMax.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<phone number>"
                    .Replacement.Text = txtPhonenumber.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                'word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                If inputletter = "INF ONLY Benefits Due" Then
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "INF only  benefits due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "INF only  benefits due.pdf"

                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                    'SendKeys.Send("%fa")
                    'SendKeys.SendWait("%fa")
                    Threading.Thread.Sleep(1000)

                ElseIf inputletter = "INF ONLY No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                        .Text = "claims and updated"
                        .Replacement.Text = "claims and confirmed"
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "INF only no beneifts due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "INF only no beneifts due.pdf"

                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                    'SendKeys.Send("%fa")
                    'SendKeys.SendWait("%fa")
                    Threading.Thread.Sleep(1000)

                End If
            End If
            If inputletter = "LTM ONLY Benefits Due" Or inputletter = "LTM ONLY No Benefits Due" Then
                Mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/" + "LTM only  benefits due.docx"
                word_doc = word_app.Documents.Add(Mydocpath)
                para = word_doc.Paragraphs.Add()
                'word_doc.Content.Find.Execute(FindText:="<Plan Correspondence Address 1> ", ReplaceWith:=PCorsdAddrs.Text, Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member Full Name>", ReplaceWith:=txtMember_Name.Text, Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member Address 1>", ReplaceWith:=txtMemberAddrs1.Text, Replace:=Word.WdReplace.wdReplaceAll)
                'While word_doc.Content.Find.Execute(FindText:="  ", Wrap:=Word.WdFindWrap.wdFindContinue)
                '    word_doc.Content.Find.Execute(FindText:="  ", ReplaceWith:=" ", Replace:=Word.WdReplace.wdReplaceAll, Wrap:=Word.WdFindWrap.wdFindContinue)
                'End While

                If txtMemberAddrs2.Text <> "" Then
                    word_doc.Content.Find.Execute(FindText:="<Member Address 2>", ReplaceWith:=txtMemberAddrs2.Text, Replace:=Word.WdReplace.wdReplaceAll)
                Else
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("<Member Address 2>")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                End If

                If txtMemberAddrs3.Text <> "" Then
                    word_doc.Content.Find.Execute(FindText:="<Member Address 3>", ReplaceWith:=txtMemberAddrs3.Text, Replace:=Word.WdReplace.wdReplaceAll)
                Else
                    For Each para In word_doc.Paragraphs
                        'Debug.Print(p.Range.Text)
                        If UCase(para.Range.Text.Trim) = UCase(Trim("<Member Address 3>")) Then
                            para.Range.Delete()
                            Exit For
                        End If
                    Next
                End If


                word_doc.Content.Find.Execute(FindText:="<Member City State Zip>", ReplaceWith:=txtMemberCSZ.Text, Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Date>", ReplaceWith:=txtDate.Text, Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member First Name>", ReplaceWith:=txtMemberFirstname.Text, Replace:=Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<LTMMet Date>"
                    .Replacement.Text = txtLTMMetDate.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<LTMAmount>"
                    .Replacement.Text = txtLTMAmt.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<LTM Max>"
                    .Replacement.Text = txtLTMMax.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<phone number>"
                    .Replacement.Text = txtPhonenumber.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                'word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                If inputletter = "LTM ONLY Benefits Due" Then
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "LTM only  benefits due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    ''word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "LTM only  benefits due.pdf"

                    'MsgBox(System.IO.File.Exists(mydocpath2))
                    If System.IO.File.Exists(mydocpath2) = True Then

                        System.IO.File.Delete(mydocpath2)
                    End If
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                    'SendKeys.Send("%fa")
                    'SendKeys.SendWait("%fa")
                    Threading.Thread.Sleep(1000)
                ElseIf inputletter = "LTM ONLY No Benefits Due" Then
                    With word_doc.Application.Selection.Find
                        .Text = "claims and updated"
                        .Replacement.Text = "claims and confirmed"
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = False
                        .MatchCase = False
                        .MatchWholeWord = False
                        .MatchWildcards = False
                        .MatchSoundsLike = False
                        .MatchAllWordForms = False
                    End With
                    word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                    mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "LTM only no beneifts due"
                    word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    'word_doc.Close(WdSaveOptions.wdPromptToSaveChanges)
                    ''word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    ''word_doc.Close()

                    'word_app.Quit()
                    mydocpath2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "LTM only no beneifts due.pdf"

                    If System.IO.File.Exists(mydocpath2) = True Then
                        'System.IO.File.Open("C:\Users\kramesh2\Documents\NTIA\Final_Result.pdf", FileMode.Open)
                        System.IO.File.Delete(mydocpath2)
                    End If
                    'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                    word_doc.ExportAsFixedFormat(mydocpath1, WdExportFormat.wdExportFormatPDF, True)
                    Threading.Thread.Sleep(2000)
                    word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)

                    Threading.Thread.Sleep(1000)
                    word_app.Quit()


                    'Me.SendToBack()
                    Threading.Thread.Sleep(4000)
                    'SendKeys.Send("%fa")
                    'SendKeys.SendWait("%fa")
                    Threading.Thread.Sleep(1000)

                End If
            End If

            If inputletter = "Letter_Draft" Then


                Mydocpath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\" + "Insight Software\Macro Express\Macro Files\NET\OOPCalculator" + "\" + "Letter_Draft.docx"
                'Dim word_doc As Word._Document = word_app.Documents.Add("C:\Users\kramesh2\Documents\Letter_Draft.docx")
                word_doc = word_app.Documents.Add(Mydocpath)
                'word_app.Documents.Add("C:\Users\kramesh2\Documents\Letter_Draft.docx")
                para = word_doc.Paragraphs.Add()

                word_doc.Content.Find.Execute(FindText:="<Member Full Name>", ReplaceWith:="Vb.net", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member Address 1>", ReplaceWith:="123 xyz", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member Address 2>", ReplaceWith:="234 xyz", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member Address 3>", ReplaceWith:="234 abc", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member City State Zip>", ReplaceWith:="12345", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="Susan Smith", ReplaceWith:="Ramesh kodam", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="John Smith", ReplaceWith:="Ramesh kodam1", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="123456789", ReplaceWith:="891023657", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="Any Group, LLC", ReplaceWith:="abc LLC", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="11111", ReplaceWith:="56877", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="134567", ReplaceWith:="686777", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="LETTID-TBD", ReplaceWith:="letter1", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Date>", ReplaceWith:="1/12/2023", Replace:=Word.WdReplace.wdReplaceAll)

                word_doc.Content.Find.Execute(FindText:="<Member First Name>,", ReplaceWith:="abcdef", Replace:=Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<Year>"
                    .Replacement.Text = "2023"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<TITLE>"
                    .Replacement.Text = "abcd"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<DEDMet Date>"
                    .Replacement.Text = "1/12/2023"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<MDEDAmount>"
                    .Replacement.Text = 99
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<DEDMax>"
                    .Replacement.Text = 10
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<OOPMet Date>"
                    .Replacement.Text = "1/12/2023"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<MOOPAmount>"
                    .Replacement.Text = 98
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<OOPMax>"
                    .Replacement.Text = 9
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<LTMMet Date>"
                    .Replacement.Text = "1/12/2023"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<LTMAmount>"
                    .Replacement.Text = 97
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<LTM Max>"
                    .Replacement.Text = 7
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<INFMet Date>"
                    .Replacement.Text = "1/12/2023"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<INFAmount>"
                    .Replacement.Text = 90
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<INF Max>"
                    .Replacement.Text = 2
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<FOOPMet Date>"
                    .Replacement.Text = "1/12/2023"
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
                With word_doc.Application.Selection.Find
                    .Text = "<MFOOPAmount>"
                    .Replacement.Text = 95
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<FOOPMax>"
                    .Replacement.Text = 5
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<PHONE_NUMBER>"
                    .Replacement.Text = txtPhonenumber.Text
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<TTY_NUMBER>"
                    .Replacement.Text = 12345
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)

                With word_doc.Application.Selection.Find
                    .Text = "<OPERATING_HOURS>"
                    .Replacement.Text = 24
                    .Forward = True
                    .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                    .Format = False
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                word_doc.Application.Selection.Find.Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)


                mydocpath1 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "/NTIA/" + "ELGS Letter"
                word_doc.Protect(WdProtectionType.wdAllowOnlyReading, vbNull, "password")
                'word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatPDF)
                word_doc.SaveAs2(mydocpath1, Word.WdSaveFormat.wdFormatDocument)
                'Insert a paragraph at the beginning of the document.
                'word_doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                word_doc.Close()
                word_app.Quit()
            End If
        Catch ex As FileNotFoundException

            MsgBox(ex.Message.ToString)

            Dim ErrStr As String = "Message:" + "/n" + ex.StackTrace
            Dim ErrFilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Insight Software\Macro Express\ErrLog.txt"
            File.WriteAllText(ErrFilePath, ErrStr)
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
            Dim ErrStr As String = "Message:" + "/n" + ex.StackTrace
            Dim ErrFilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\Insight Software\Macro Express\ErrLog.txt"
            File.WriteAllText(ErrFilePath, ErrStr)
        End Try

        'word_app.Quit()

        'Dim para As Word.Paragraph = word_doc.Paragraphs.Add()
        'para.Range.InsertParagraphAfter()

        'para.Range.Text = "Member Full Name:   " & "XYZ"

        'para.Range.InsertParagraphAfter()


        'para.Range.InsertAfter("123 xyz")
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter("Patient:   " & Me.txtPNM.Text)


        'para.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
        'para.Range.Text = "Patient:   " & Me.txtPNM.Text
        'para.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
        'para.Range.InsertParagraphAfter()
        'para.Range.InsertAfter("456 xyz")
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter("Member ID:   " & Me.txtMID.Text)
        'para.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
        'para.Range.Text = "Member:   " & Me.txtMName.Text
        'para.Range.InsertParagraphAfter()
        'para.Range.Text = "Member ID:   " & Me.txtMID.Text
        'para.Range.InsertAfter("Member Address 3:  " & "956756 xyz")
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter(vbTab)
        'para.Range.InsertAfter("Member ID:   " & Me.txtMID.Text)
        'para.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
        'para.Range.InsertParagraphAfter()
        'para.Range.Text = "Group    :   " & Me.txtGRP.Text
        'para.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
        'para.Range.InsertParagraphAfter()

        'para.Range.InsertParagraphAfter()
        'para.Range.Text = "Refference Tracking#    :   " & Me.txtREF.Text
        'para.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
        'para.Range.InsertParagraphAfter()
        'para.Range.Text = "Template ID#    :   " & "LETTID-TDB-Thank you for contacting us about your <Year> |TITLE-one of: 1) Deductible 2) Out-of-Pocket amounts 3) Deductible and Out-of-Pocket amounts 4) policy’s lifetime maximum 5) policy’s infertility maximum|. We reviewed your <Year> claims and |when benefits are due: updated| |when no benefits are due: confirmed| your total amounts."
        'para.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify

        'para.Range.InsertParagraphAfter()
        'para.Range.Style = "Heading 1"





        'para.Range.InsertParagraphAfter()



        ' Add more text.





        ' para.Range.Text = "To make a chrysanthemum curve, use" &
        '"the following " &
        '"parametric equations as t goes from 0 to 21 * ? to" &
        '    "generate " &
        '"points and then connect them."
        'para.Range.InsertParagraphAfter()



        ' Save the current font and start using Courier New.
        'Dim old_font As String = para.Range.Font.Name
        'para.Range.Font.Name = "Courier New"





        ' Add the equations.
        'para.Range.Text =
        '    "  r = 5 * (1 + Sin(11 * t / 5)) -" & vbCrLf &
        '    "      4 * Sin(17 * t / 3) ^ 4 *" & vbCrLf &
        '    "      Sin(2 * Cos(3 * t) - 28 * t) ^ 8" & vbCrLf & _
        '                                                        _
        '    "  x = r * Cos(t)" & vbCrLf &
        '    "  y = r * Sin(t)"




        ' Start a new paragraph and then switch back to the
        ' original font.
        'para.Range.InsertParagraphAfter()
        'para.Range.Font.Name = old_font



        ' Save the document.
        'Dim filename As Object = Path.GetFullPath(
        'Path.Combine(Application.StartupPath, "..\\..")) &
        '"\\test.doc"
        'word_doc.SaveAs(FileName:=filename)



        '' Close.
        'Dim save_changes As Object = False
        'word_doc.Close(save_changes)
        'word_app.Quit(save_changes)



        'word_app.ProtectedViewWindows.ToString()
        'Try
        'System.IO.Path.ChangeExtension(word_app, "pdf")
        'word_doc.SaveAs2("C:\Users\kramesh2\Documents\rammy.pdf", 0,,,,, True,,,,,,,,,,)
        'word_doc.SaveAs2("C:\Users\kramesh2\Documents\rammy1", Word.WdSaveFormat.wdFormatPDF)

        'Catch ex As Exception

        'End Try
        'word_doc.SaveAs2("C:\Users\kramesh2\Documents\rammy", 0,,,,, True,,,,,,,,,,)
        'word_doc.Protect(WdProtectionType.wdAllowOnlyReading)
        'word_doc.Save()
        '      Document.Protect(ProtectionType.AllowOnlyReading, "password");
        ''Saves the Word document
        'Document.Save("Protection.docx", FormatType.Docx);

        'Me.Hide()

        Form1.RichTextBox1.SelectionIndent = 5
        Form1.RichTextBox1.BulletIndent = 4
        Form1.RichTextBox1.SelectionBullet = True
        Form1.RichTextBox1.AppendText("Letter Generated Successfully" & vbCrLf)

    End Sub


End Class
