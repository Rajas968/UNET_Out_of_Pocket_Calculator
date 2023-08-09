Imports Microsoft.Office.Interop
Module modStart
    Public Structure SearchResultsSummaryType
        Public startTime As Date
        Public ExpectedRecords As Long
        Public ActualTotalRecords As Long
    End Structure

    Public Sub Start(ByVal empID As String, ByVal firstName As String, ByVal inclBoneyard As Boolean, ByVal maintainRecordInfo As Boolean _
, ByVal inclRio As Boolean, ByVal skipParsing As Boolean, ByVal restFindBatchSize As Integer)
        Dim Success As Boolean
        Dim msg As String
        Dim URL As String
        Dim errorMessage As String
        Dim gotSearchResults As Boolean
        Dim dataWb As Excel.Workbook
        Dim arrResults() As Doc360ResultInfo
        Dim RioSearchResultsSummary As SearchResultsSummaryType
        Dim BoneyardSearchResultsSummary As SearchResultsSummaryType

        'On Error GoTo ErrHandler

        'Application.StatusBar = Now() & "Starting ..."

        ' Initialize function with indicated Employee ID
        Call GetEmpId(newEmpID:=empID)

        ReDim arrResults(0)

        If inclRio Then
            With RioSearchResultsSummary
                .startTime = Now()
                Call GetDoc360SearchResults(empIdTPRD:=empID, firstNameTPRD:=firstName, boneyard:=False, arrResults:=arrResults, retTotalRecords:= .ExpectedRecords _
                                    , restFindBatchSize:=restFindBatchSize)
                .ActualTotalRecords = UBound(arrResults)
            End With
        End If

        If inclBoneyard Then
            With BoneyardSearchResultsSummary
                .startTime = Now()
                Call GetDoc360SearchResults(empIdTPRD:=empID, firstNameTPRD:=firstName, boneyard:=True, arrResults:=arrResults, retTotalRecords:= .ExpectedRecords)
                .ActualTotalRecords = UBound(arrResults) - RioSearchResultsSummary.ActualTotalRecords
            End With
        End If

        If UBound(arrResults) > 0 Then

            'dataWb = Application.Workbooks.Add()
            Success = ParsePurgeHistory(arrResults, dataWb _
            , skipParsing, RioSearchResultsSummary:=RioSearchResultsSummary, BoneyardSearchResultsSummary:=BoneyardSearchResultsSummary)
            If Success Then

                '      Application.StatusBar = ""

                ' modParseExcel_Alan.Finalize(maintainRecordInfo)
                ' Don't save the file.
                ' Don't display "Done!". The status bar will be cleared when done.

                ' Display log sheet after completion if an error occurred
                'Call ShowLogSheetIfError(skipParsing)

            End If

            'ElseIf errorMessage <> "" Then ' Error during search
            '    Application.Visible = True ' Excel must be on top to see the message box
            '    MsgBox(errorMessage, vbCritical)
            'ElseIf UBound(arrResults) = 0 Then
            '    Application.Visible = True ' Excel must be on top to see the message box
            '    MsgBox("No results found", vbCritical)
            'Else
            '    Application.Visible = True ' Excel must be on top to see the message box
            '    MsgBox("Unknown error occurred", vbCritical)
        End If

        'ExitSub:
        '        'Application.StatusBar = ""
        '        Exit Sub

        'ErrHandler:
        '        MsgBox("Error " & Err.Number & ", " & Err.Description)
        '        Resume ExitSub

    End Sub

    'Public Function ParsePurgeHistory(ByRef arrResults() As Doc360ResultInfo, dataWb As Excel.Workbook, skipParsing As Boolean, RioSearchResultsSummary As SearchResultsSummaryType, BoneyardSearchResultsSummary As SearchResultsSummaryType) As Boolean
    '    Dim claimText As String
    '    Dim recNumber As String
    '    Dim compDoc As String
    '    Dim globalDocID As String
    '    Dim pageCount As Integer
    '    Dim restart As Boolean
    '    Dim statusBarText As String
    '    Dim claimTextStripped As String
    '    Dim processClaimText As Boolean
    '    Dim lineCount As Long
    '    Dim firstName As String
    '    Dim questionableSinglePage As Boolean
    '    Dim docEmployeeID As String
    '    Dim status As String
    '    'LK#7783 - nextCompDoc no longer necessary
    '    'Dim nextCompDoc As String
    '    Dim prevRecNum As String
    '    Dim startTime As Date

    '    'On Error GoTo ErrHandler

    '    'modParseExcel_Alan.InitializeWb(dataWb)

    '    startTime = CDate(Now)

    '    With RioSearchResultsSummary
    '        If .startTime <> CDate(Now) Then
    '            startTime = .startTime
    '            ' LogRecord(timeStamp:= .startTime, recNum:="", pageCnt:=0, lineCnt:=0, charCnt:=0, compDoc:="", globalDocID:="", firstName:="", boneyard:=False, status:="Rio record count: Expected " & .ExpectedRecords & ", Actual " & .ActualTotalRecords)
    '        End If
    '    End With

    '    With BoneyardSearchResultsSummary
    '        If .startTime <> CDate(Now) Then
    '            If startTime = CDate(Now) Then startTime = .startTime
    '            '      LogRecord(timeStamp:=BoneyardSearchResultsSummary.startTime, recNum:="", pageCnt:=0, lineCnt:=0, charCnt:=0, compDoc:="", globalDocID:="", firstName:="", boneyard:=False, status:="Boneyard record count: Expected " & .ExpectedRecords & ", Actual " & .ActualTotalRecords)
    '        End If
    '    End With

    '    If Not skipParsing Then
    '        restart = True ' This is updated by the called function
    '        'LK#7220 - Added nextCompDoc to support checking for Doc360 restarting the search results (Doc360 anomaly)
    '        'LK#7783 - nextCompDoc no longer necessary
    '        Do While GetNextRecord(arrResults, GetEmpId(), claimText, recNumber, compDoc, globalDocID, pageCount, firstName, restart)



    '            lineCount = claimText.Count(Function(c) c = vbLf)
    '            docEmployeeID = claimText.Substring(3, 10)

    '            statusBarText = "Rec#" & recNumber & ", PageCnt " & pageCount & ", LineCnt " & lineCount & ", " & compDoc & ", " & firstName

    '            ' Ignore document if incomplete (expected to be in multi-page document elsewhere)
    '            ' Note: Cannot just check right-most text since there could be spaces and new lines after it
    '            If docEmployeeID <> GetEmpId() Then
    '                If claimText.StartsWith("<html>") Then
    '                    Debug.Print($"{Now()}, *** ERROR *** {statusBarText}, ** PROBLEM WITH DOCUMENT: {Left(claimText, claimText.IndexOf(vbCrLf) - 1)}")
    '                    status = STATUS_ERROR_PREFIX & "PROBLEM WITH DOCUMENT: " & Left(claimText, claimText.IndexOf(vbCrLf) - 1)
    '                ElseIf lineCount > 1 Then
    '                    Debug.Print($"{Now()}, *** ERROR *** {statusBarText}, ** WRONG EMPLOYEE ID {docEmployeeID}")
    '                    status = STATUS_ERROR_PREFIX & "WRONG EMPLOYEE ID " & docEmployeeID
    '                Else
    '                    Debug.Print($"{Now()}, *** ERROR *** {statusBarText}, ** CAN'T GET DOCUMENT")
    '                    status = STATUS_ERROR_PREFIX & "CAN'T GET DOCUMENT"
    '                End If
    '            ElseIf pageCount = 1 Then
    '                If isTPRDDocumentComplete(claimText, pageCount, lineCount, questionableSinglePage) Then
    '                    If questionableSinglePage Then
    '                        statusBarText = $"{statusBarText}, QUESTIONABLE COMPLETE SINGLE PAGE"
    '                        status = STATUS_PROCESSING_PREFIX & ", QUESTIONABLE SINGLE PAGE"
    '                    Else
    '                        status = STATUS_PROCESSING_PREFIX
    '                    End If
    '                Else
    '                    Debug.Print($"{Now()}, * SKIPPING {statusBarText}, INCOMPLETE DOCUMENT")
    '                    status = STATUS_SKIP_PREFIX & "INCOMPLETE DOCUMENT"
    '                End If
    '            ElseIf lineCount > 1000 Then
    '                statusBarText = $"{statusBarText}, MULTI-PAGE"
    '                status = STATUS_PROCESSING_PREFIX & ", MULTI-PAGE"
    '            Else
    '                statusBarText = $"{statusBarText}, QUESTIONABLE MULTI-PAGE"
    '                status = STATUS_PROCESSING_PREFIX & ", QUESTIONABLE MULTI-PAGE"
    '            End If

    '            ' If this compDoc has already been processed, ignore this one
    '            ' Append to the unprocessed compDoc to prevent from finding this record as being processed
    '            If isCompDocProcessed(compDoc, prevRecNum) Then
    '                Debug.Print($"{Now()}, *** SKIPPING {statusBarText}, DUP, Rec#{prevRecNum}")
    '                status = "*** SKIPPED - DUP, Rec#" & prevRecNum & ", " & Replace(status, STATUS_PROCESSING_PREFIX, "Not Processed")
    '                compDoc = $"{compDoc}{COMP_DOC_DUP_POSTFIX}" ' Prevents finding this record's compDoc as being processed
    '            ElseIf status.StartsWith(STATUS_SKIP_PREFIX) Then
    '                compDoc = $"{compDoc}{COMP_DOC_SKIP_POSTFIX}"
    '            ElseIf status.StartsWith(STATUS_ERROR_PREFIX) Then
    '                compDoc = $"{compDoc}{COMP_DOC_ERROR_POSTFIX}"
    '            End If

    '            ' LogRecord(timeStamp:=Now(), recNum:=recNumber, pageCnt:=pageCount, lineCnt:=lineCount, charCnt:=claimText.Length, compDoc:=compDoc, globalDocID:=globalDocID, firstName:=firstName, boneyard:=arr
    '        Loop

    '    End If
    'End Function

    Public Function ParsePurgeHistory(
    ByRef arrResults() As Doc360ResultInfo _
  , dataWb As Excel.Workbook _
  , skipParsing As Boolean _
  , RioSearchResultsSummary As SearchResultsSummaryType _
  , BoneyardSearchResultsSummary As SearchResultsSummaryType
) As Boolean
        Dim claimText As String
        Dim recNumber As String
        Dim compDoc As String
        Dim globalDocID As String
        Dim pageCount As Integer
        Dim restart As Boolean
        Dim statusBarText As String
        Dim claimTextStripped As String
        Dim processClaimText As Boolean
        Dim lineCount As Long
        Dim firstName As String
        Dim questionableSinglePage As Boolean
        Dim docEmployeeID As String
        Dim status As String
        ''LK#7783 - nextCompDoc no longer necessary
        ''Dim nextCompDoc As String
        Dim prevRecNum As String
        Dim startTime As Date

        '  On Error GoTo ErrHandler

        ' Call modParseExcel_Alan.InitializeWb(dataWb)
        'startTime = CDate(0)

        'With RioSearchResultsSummary
        '    If .startTime <> CDate(0) Then
        '        startTime = .startTime
        '        Call LogRecord(timeStamp:= .startTime _
        '                 , recNum:="" _
        '                 , pageCnt:=0 _
        '                 , lineCnt:=0 _
        '                 , charCnt:=0 _
        '                 , compDoc:="" _
        '                 , globalDocID:="" _
        '                 , firstName:="" _
        '                 , boneyard:=False _
        '                 , status:="Rio record count: Expected " & .ExpectedRecords & ", Actual " & .ActualTotalRecords)
        '    End If
        'End With
        'With BoneyardSearchResultsSummary
        '    If .startTime <> CDate(0) Then
        '        If startTime = CDate(0) Then startTime = .startTime
        '        'Call LogRecord(timeStamp:=BoneyardSearchResultsSummary.startTime _
        '        '         , recNum:="" _
        '        '         , pageCnt:=0 _
        '        '         , lineCnt:=0 _
        '        '         , charCnt:=0 _
        '        '         , compDoc:="" _
        '        '         , globalDocID:="" _
        '        '         , firstName:="" _
        '        '         , boneyard:=False _
        '        '         , status:="Boneyard record count: Expected " & .ExpectedRecords & ", Actual " & .ActualTotalRecords)
        '    End If
        'End With

        If Not skipParsing Then
            restart = True ' This is updated by the called function
            ''LK#7220 - Added nextCompDoc to support checking for Doc360 restarting the search results (Doc360 anomaly)
            ''LK#7783 - nextCompDoc no longer necessary
            Do While GetNextRecord(arrResults, GetEmpId(), claimText, recNumber, compDoc, globalDocID, pageCount, firstName, restart)

                lineCount = Len(claimText) - Len(Replace(claimText, vbLf, ""))
                docEmployeeID = Mid(claimText, 4, 10)

                statusBarText = "Rec#" & recNumber & ", PageCnt " & pageCount & ", LineCnt " & lineCount & ", " & compDoc & ", " & firstName

                ' Ignore document if incomplete (expected to be in multi-page document elsewhere)
                ' Note: Cannot just check right-most text since there could be spaces and new lines after it
                If docEmployeeID <> GetEmpId() Then
                    If InStr(claimText, "<html>") = 1 Then
                        'Debug.Print Now() & ", *** ERROR *** " & statusBarText & ", ** PROBLEM WITH DOCUMENT: " & Left(claimText, InStr(claimText, vbCrLf) - 1)
                        status = STATUS_ERROR_PREFIX & "PROBLEM WITH DOCUMENT: " & Left(claimText, InStr(claimText, vbCrLf) - 1)
                    ElseIf lineCount > 1 Then
                        'Debug.Print Now() & ", *** ERROR *** " & statusBarText & ", ** WRONG EMPLOYEE ID " & docEmployeeID
                        status = STATUS_ERROR_PREFIX & "WRONG EMPLOYEE ID " & docEmployeeID
                    Else
                        '  Debug.Print Now() & ", *** ERROR *** " & statusBarText & ", ** CAN'T GET DOCUMENT"
                        status = STATUS_ERROR_PREFIX & "CAN'T GET DOCUMENT"
                    End If
                ElseIf pageCount = 1 Then
                    If isTPRDDocumentComplete(claimText, pageCount, lineCount, questionableSinglePage) Then
                        If questionableSinglePage Then
                            statusBarText = statusBarText & ", QUESTIONABLE COMPLETE SINGLE PAGE"
                            status = STATUS_PROCESSING_PREFIX & ", QUESTIONABLE SINGLE PAGE"
                        Else
                            status = STATUS_PROCESSING_PREFIX
                        End If
                    Else
                        '   Debug.Print Now() & ", * SKIPPING " & statusBarText & ", INCOMPLETE DOCUMENT"
                        status = STATUS_SKIP_PREFIX & "INCOMPLETE DOCUMENT"
                    End If
                ElseIf lineCount > 1000 Then
                    statusBarText = statusBarText & ", MULTI-PAGE"
                    status = STATUS_PROCESSING_PREFIX & ", MULTI-PAGE"
                Else
                    statusBarText = statusBarText & ", QUESTIONABLE MULTI-PAGE"
                    status = STATUS_PROCESSING_PREFIX & ", QUESTIONABLE MULTI-PAGE"
                End If

                ' If this compDoc has already been processed, ignore this one
                ' Append to the unprocessed compDoc to prevent from finding this record as being processed
                If IsCompDocProcessed(compDoc, prevRecNum) Then
                    'Debug.Print Now() & ", *** SKIPPING " & statusBarText & ", DUP, Rec#" & prevRecNum
                    status = "*** SKIPPED - DUP, Rec#" & prevRecNum & ", " & Replace(status, STATUS_PROCESSING_PREFIX, "Not Processed")
                    compDoc = compDoc & COMP_DOC_DUP_POSTFIX ' Prevents finding this record's compDoc as being processed
                    ''LK#7783 - nextCompDoc no longer necessary
                    ''ElseIf nextCompDoc <> "" And isCompDocProcessed(nextCompDoc, prevRecNum) Then
                    ''    Debug.Print Now() & ", *** SKIPPING " & statusBarText & ", prevRecNum#" & prevRecNum & ", nextCompDoc = " & nextCompDoc
                    ''    status = "*** IGNORED - NEXT COMP DOC EXISTS, Rec#" & prevRecNum & ", " & Replace(status, STATUS_PROCESSING_PREFIX, "Not Processed")
                    ''    compDoc = compDoc & COMP_DOC_ERROR_POSTFIX ' Prevents finding this record's compDoc as being processed
                ElseIf InStr(status, STATUS_SKIP_PREFIX) = 1 Then
                    compDoc = compDoc & COMP_DOC_SKIP_POSTFIX
                ElseIf InStr(status, STATUS_ERROR_PREFIX) = 1 Then
                    compDoc = compDoc & COMP_DOC_ERROR_POSTFIX
                End If

                'Call LogRecord(timeStamp:=Now() _
                '         , recNum:=recNumber _
                '         , pageCnt:=pageCount _
                '         , lineCnt:=lineCount _
                '         , charCnt:=Len(claimText) _
                '         , compDoc:=compDoc _
                '         , globalDocID:=globalDocID _
                '         , firstName:=firstName _
                '         , boneyard:=arrResults(recNumber).boneyard _
                '         , status:=status)

                If InStr(status, STATUS_PROCESSING_PREFIX) = 1 Then
                    'Application.StatusBar = "Processing " & statusBarText
                    'Debug.Print Now() & ", Processing " & statusBarText

                    Call modParseExcel_Common.ParseDocument(GetEmpId(), claimText, recNumber, compDoc, pageCount, lineCount)

                    'Application.StatusBar = "Completed " & statusBarText
                End If

            Loop

        End If

        'Application.StatusBar = "Almost done ..."

        'Call LogRecord(timeStamp:=Now() _
        '         , recNum:="" _
        '         , pageCnt:=0 _
        '         , lineCnt:=0 _
        '         , charCnt:=0 _
        '         , compDoc:="" _
        '         , globalDocID:="" _
        '         , firstName:="" _
        '         , boneyard:=False _
        '         , status:="Duration " & Format(Now() - startTime, "hh:mm:ss"))

        'P 'arsePurgeHistory = True

        Exit Function



    End Function
    Private Function GetNextRecord(ByRef arrResults() As Doc360ResultInfo, employeeID As String, ByRef docText As String, ByRef recNum As String, ByRef compDoc As String, ByRef globalDocID As String, ByRef pageCount As Integer, ByRef firstName As String, ByRef restart As Boolean) As Boolean
        Dim currCompDoc As String
        Dim Info As Doc360ResultInfo
        Dim bestInfo As Doc360ResultInfo
        Dim bestDocumentText As String
        Dim bestRecNum As Integer
        Dim keepLooking As Boolean

        Static prevEmployeeID As String
        Static indx As Integer

        On Error GoTo ErrHandler

        GetNextRecord = False

        If indx = -1 Then
            indx = 0
            Exit Function
        End If

        If restart Or prevEmployeeID <> employeeID Then
            prevEmployeeID = employeeID
            currCompDoc = ""
            indx = 1
            restart = False
        End If

        Info = arrResults(indx)

        currCompDoc = Info.compoundDocID
        bestDocumentText = ""

        keepLooking = True
        While indx <= UBound(arrResults) And keepLooking
            Info = arrResults(indx)

            'Dim txt As String
            'Debug.Print "Rec#" & indx & ", " & Info.globalDocId
            'txt = GetDoc360TPRDDocumentText(Info.globalDocId, Info.boneyard)
            'Debug.Print Left(txt, 1000)

            'Application.StatusBar = "Evaluating Rec#" & indx & ", " & Info.compoundDocID & ", " & Info.firstName
            If Info.compoundDocID = currCompDoc Then
                If bestDocumentText = "" Then
                    bestInfo = Info
                    '       Application.StatusBar = "Getting " & Info.pageCount & " pages, Rec#" & indx & ", " & Info.compoundDocID & ", " & Info.firstName
                    bestDocumentText = GetDoc360TPRDDocumentText(Info.globalDocID, Info.boneyard)
                    bestRecNum = indx
                End If
                indx = indx + 1
            Else
                keepLooking = False
            End If
        End While
        If indx > UBound(arrResults) Then
            indx = -1 ' Done
        End If

        ' If we found another document, return it (otherwise, False is returned and we're done)
        If bestDocumentText <> "" Then
            recNum = bestRecNum
            compDoc = bestInfo.compoundDocID
            globalDocID = bestInfo.globalDocID
            docText = bestDocumentText
            pageCount = bestInfo.pageCount
            firstName = bestInfo.firstName

            ''LK#7220 - Return the compDoc that indicated the end of the currCompDoc
            ''LK#7783 - nextCompDoc no longer necessary
            ''nextCompDoc = Info.compoundDocID

            GetNextRecord = True

        End If

        Exit Function

ErrHandler:
        MessageBox.Show(Err.Number & ", " & Err.Description)
        GetNextRecord = False
        Exit Function
        'Resume Next ' support testing
    End Function
    Public Function GetEmpId(Optional ByVal newEmpID As Object = Nothing) As String
        Static theEmpID As String

        If newEmpID IsNot Nothing Then
            theEmpID = CStr(newEmpID)
        End If

        GetEmpId = theEmpID
    End Function
    Public Function GetMyMSID() As String
        Return Environment.UserName
    End Function



End Module
