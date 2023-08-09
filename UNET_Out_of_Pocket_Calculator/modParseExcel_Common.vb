Module modParseExcel_Common
    Private ReadOnly QUESTIONABLE_SINGLE_PAGE_LINE_COUNT = 980 ' Account for some trailing blank lines being truncated in web page
    Private ReadOnly QUESTIONABLE_SINGLE_PAGE_CHAR_COUNT = 5000 ' Don't need all the text, just enough to determined uniqueness

    Private collQuestionableSinglePage As Collection
    Private collMultiPage As Collection
    Public Function isTPRDDocumentComplete(ByRef claimText As String, ByVal pageCount As Integer, ByVal lineCount As Long, ByRef questionableSinglePage As Boolean) As Boolean
        Dim claimTextStripped As String
        Const END_OF_CLAIM_STRIPPED As String = "***ENDOFCLAIM***"

        ' The end of the document may or may not have trailing spaces or blank lines
        ' Remove them all just in case to insure END OF CLAIM is always at the end of the string
        claimTextStripped = Right(claimText, 1000)
        claimTextStripped = claimTextStripped.Replace(" ", "")
        claimTextStripped = claimTextStripped.Replace(vbCrLf, "")
        claimTextStripped = claimTextStripped.Replace(vbCr, "") 'LK#7783 - Downloaded files have this
        If claimTextStripped.EndsWith(END_OF_CLAIM_STRIPPED) Then
            isTPRDDocumentComplete = True
            If pageCount = 1 And lineCount > QUESTIONABLE_SINGLE_PAGE_LINE_COUNT Then
                questionableSinglePage = True
            Else
                questionableSinglePage = False
            End If
        Else
            isTPRDDocumentComplete = False
            questionableSinglePage = False
        End If

    End Function
    Private Function FixLine(line As String) As String
        ' Sometimes the last field in the line contains the new line character.
        ' Replace it with a string of spaces to insure the line has a valid character for every field
        ' (the blank string length is arbitrarily long)
        Return line.Replace(vbCr, New String(" "c, 73))
    End Function

    Public Sub ParseDocument(
    empID As String _
  , ByVal docText As String _
  , recNum As String _
  , compDoc As String _
  , pageCount As Integer _
  , lineCount As Long)
        Dim lines() As String  ' array of data lines
        Dim line As String, line1 As String, line2 As String, line3 As String, line4 As String
        Dim idx As Long ' index into lines array
        Dim origStatusBarText As String
        Dim questionableSinglePage As Boolean

        On Error GoTo ErrHandler

        'origStatusBarText = Application.StatusBar

        ' Collect questionable single page and all multi-page documents for post-processing
        Call AddToCollection(docText, compDoc, pageCount, lineCount, questionableSinglePage)

        ' There are cases where there is no new line before the start of a claim
        ' So can't simply split at new line characters
        Call SplitLines(empID, docText, lines)

        ''    ' TEST - save every text line
        ''    sht2Row = sht2Row + 1 ' write the next rep
        ''    sht2.Cells(sht2Row, 1) = "Rep Name = " & repName
        ''    For idx = LBound(lines) To UBound(lines)
        ''        sht2Row = sht2Row + 1
        ''        sht2.Cells(sht2Row, 1) = lines(idx)
        ''    Next idx

        idx = LBound(lines)
        While idx < UBound(lines)

            line = lines(idx)
            If Left(line, 1) = "1" Then
                line1 = FixLine(lines(idx))
                line2 = FixLine(lines(idx + 1))
                line3 = FixLine(lines(idx + 2))
                line4 = FixLine(lines(idx + 3))

                '' MSAs DON'T USE THIS MACRO ANYMORE
                Call modParseExcel_Alan.ParseData(DataTypeEnum.ClaimData, recNum, compDoc, pageCount, lineCount, line1, line2, line3, line4)

                idx = idx + 4

            ElseIf Len(Trim(line)) < 2 Then
                idx = idx + 1 ' skip blank lines

            ElseIf Left(line, 3) <> "   " Then
                idx = idx + 1   ' the data line must start with 3 spaces
                ' (there are some statuses beginning in these spaces to be skipped)

            ElseIf Mid(line, 4, 2) = "PS" Then
                idx = idx + 3 ' ignore the service header

            ElseIf InStr(line, "END OF CLAIM") > 0 Then
                idx = idx + 1 ' ignore the end of claim comment

            ElseIf Mid(line, 4, 8) = "PROVIDER" Then
                line1 = FixLine(lines(idx + 1)) ' get the next row which contains the data
                line2 = FixLine(lines(idx + 3)) ' get the PD data, skipping the PD header row

                '' MSAs DON'T USE THIS MACRO ANYMORE
                Call modParseExcel_Alan.ParseData(DataTypeEnum.ProviderData, recNum, compDoc, pageCount, lineCount, line1, line2)

                ''End If
                idx = idx + 4

            Else    ' anything else MIGHT be service data lines

                Dim FSTDT As String
                FSTDT = Mid(line, 17, 6) ' ** Same as in ParseData
                If Not IsNumeric(FSTDT) Or InStr(FSTDT, " ") Then ' Check if valid date
                    idx = idx + 1 ' skip this line (probably related to facility data)
                Else
                    line1 = FixLine(lines(idx))
                    line2 = FixLine(lines(idx + 1))
                    line3 = FixLine(lines(idx + 2))

                    '' MSAs DON'T USE THIS MACRO ANYMORE
                    ' Application.StatusBar = origStatusBarText & ", Processing service data line " & idx & " of " & UBound(lines, 1)
                    Call modParseExcel_Alan.ParseData(DataTypeEnum.ServiceData, recNum, compDoc, pageCount, lineCount, line1, line2, line3)

                    idx = idx + 3
                End If
            End If
        End While

        Exit Sub

ErrHandler:
        MsgBox("Error " & Err.Number & ", " & Err.Description)
    End Sub
    Private Sub SplitLines(ByVal empID As String, ByVal docText As String, ByRef lines() As String)
        Dim lfPos As Integer
        Dim idPos As Integer
        Dim cnt As Integer
        Dim ucaseEmpID As String
        Dim currPos As Integer

        On Error GoTo ErrHandler

        ReDim lines(0) ' Clear the caller's array

        ucaseEmpID = "1  " & UCase(empID) ' This is the expected string in the text
        cnt = 0
        currPos = 1
        Do
            cnt += 1
            ReDim Preserve lines(cnt)

            lfPos = InStr(currPos, docText, vbLf)

            ' If no more new lines, this is the last line
            If lfPos = 0 Then
                lines(cnt) = Mid(docText, currPos)
                currPos = Len(docText)
            Else
                ' Find if a new block starts in the middle of a line (before the next new line)
                ' skip checking the first column, where it's expected
                idPos = InStr(currPos + 1, docText, ucaseEmpID)
                If lfPos < idPos Or idPos = 0 Then
                    If lfPos > currPos + 1 Then
                        lines(cnt) = Mid(docText, currPos, lfPos - currPos - 1)
                    End If
                    currPos = lfPos + 1 ' Skip the vbLf character
                Else
                    lines(cnt) = Mid(docText, currPos, idPos - currPos)
                    currPos = idPos
                End If
            End If
        Loop Until currPos >= Len(docText)

        Exit Sub

ErrHandler:
        MsgBox("Error " & Err.Number & ", " & Err.Description, MsgBoxStyle.Critical, "SplitLines, currPos " & currPos)
        ReDim lines(0) ' Clear the caller's array
    End Sub
    Public Function AddToCollection(ByVal docText As String, ByVal compoundDoc As String, ByVal pageCount As Integer, ByVal lineCount As Long, ByRef questionableSinglePage As Boolean) As Boolean
        AddToCollection = False
        collMultiPage = New Collection
        collQuestionableSinglePage = New Collection
        If pageCount = 1 Then
            If isTPRDDocumentComplete(docText, pageCount, lineCount, questionableSinglePage) Then
                If questionableSinglePage Then
                    collQuestionableSinglePage.Add(compoundDoc, Left(docText, QUESTIONABLE_SINGLE_PAGE_CHAR_COUNT))
                    AddToCollection = True
                End If
            End If
        Else
            collMultiPage.Add(Left(docText, QUESTIONABLE_SINGLE_PAGE_CHAR_COUNT))
            AddToCollection = True
        End If
    End Function

    Public Enum DataTypeEnum
        ServiceData = 1
        ProviderData = 2
        ClaimData = 3
    End Enum

End Module
