Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Text
Imports MSXML

Module modLibDoc360
    Private ReadOnly PLATFORM_ALAN = "Purged-Archived HX"
    Private Const DOC360_SERVER_PROD As String = "PROD"
    Private Const DOC360_SERVER_UAT As String = "UAT"
    Private Const DOC360_SERVER_STG As String = "STG"
    Private Const DOC360_SERVER_DEV As String = "DEV"
    Private Const DOC360_SERVER_SYS As String = "SYS"
    Private Const DOC360_SERVER_INT As String = "INT"
    Private Const DOC360_SERVER = DOC360_SERVER_PROD

    Private ReadOnly PRODUCTION_APPLICATION_ID As String = "b80df20c-7c91-79d0-df4a-a9b1d7dfedbe"
    Private ReadOnly TEST_APPLICATION_ID As String = "f8235fb1-2f3e-4fa8-8f42-da63f165e464"
    Public Const DOC360_DATAGROUP_EDI_CLAIMS = "tedi"
    Public Const DOC360_DATAGROUP_TOPS_PURGE = "tprd"
    Public Const DOC360_DATAGROUP_CLAIMS_CORSP = "CLAIMS"
    Public Const DOC360_DATAGROUP_TREASURY_CHECKS = "TRSY"

    Private ReadOnly DOC360_CLASS_EDI_CLAIMS = "u_edi_claim"
    Private ReadOnly DOC360_CLASS_TOPS_PURGE = "u_tops_purge"
    Private ReadOnly DOC360_CLASS_CLAIMS_CORSP = "u_clm_corsp_lwso_doc"
    Private ReadOnly DOC360_CLASS_TREASURY_CHECKS = "u_treasury_doc"

    Private ReadOnly DOC360_CLASSES_CSV As String = DOC360_CLASS_EDI_CLAIMS & "," &
                                                    DOC360_CLASS_TOPS_PURGE & "," &
                                                    DOC360_CLASS_CLAIMS_CORSP & "," &
                                                    DOC360_CLASS_TREASURY_CHECKS
    Private ReadOnly DoCmd As Object

    Public Const DOC360_SCREEN_HOME = 0
    Public Const DOC360_SCREEN_DATAGROUP = 1
    Public Const DOC360_SCREEN_RESULTS = 2
    Public Const DOC360_SCREEN_VIEW = 3
    Public Const DOC360_REST_FIND_DEFAULT_BATCH_SIZE = 50


    Public Structure Doc360ResultInfo
        Public employeeID As String 'TPRD
        Public firstName As String ' TPRD
        Public fln As String ' TEDI
        Public receivedDate As Date 'TEDI
        Public globalDocID As String
        Public compoundDocID As String
        Public pageCount As Integer
        Public boneyard As Boolean
    End Structure

    Private Function GetDoc360URLServerPostfix(Optional ByVal doc360Server As String = DOC360_SERVER) As String
        Select Case UCase(doc360Server)
            Case DOC360_SERVER_PROD
                GetDoc360URLServerPostfix = ""
            Case DOC360_SERVER_UAT
                GetDoc360URLServerPostfix = "-uat"
            Case DOC360_SERVER_STG
                GetDoc360URLServerPostfix = "-stg"
            Case DOC360_SERVER_DEV
                GetDoc360URLServerPostfix = "-dev"
            Case DOC360_SERVER_SYS
                GetDoc360URLServerPostfix = "-sys"
            Case DOC360_SERVER_INT
                GetDoc360URLServerPostfix = "-int"
            Case Else
                GetDoc360URLServerPostfix = ""
        End Select
    End Function

    Private Function GetDoc360BaseURL(Optional ByVal doc360Server As String = DOC360_SERVER) As String
        Const URL As String = "https://doc360<server>.optum.com/doc360-ui/"
        GetDoc360BaseURL = URL.Replace("<server>", GetDoc360URLServerPostfix(doc360Server))
    End Function

    Public Function GetDoc360TPRDDocumentText(ByVal globalDocID As String, ByVal boneyard As Boolean) As String
        GetDoc360TPRDDocumentText = GetDoc360DocumentText(DOC360_CLASS_TOPS_PURGE, globalDocID, boneyard)
    End Function


    Public Function GetDoc360SearchResults(boneyard As Boolean, ByRef arrResults() As Doc360ResultInfo, Optional ByRef retTotalRecords As Long = 0, Optional empIdTPRD As String = "", Optional firstNameTPRD As String = "", Optional flnTEDI As String = "", Optional restFindBatchSize As Integer = DOC360_REST_FIND_DEFAULT_BATCH_SIZE) As Boolean
        Dim doc360Class As String = ""
        Dim initialUbound As Long
        Dim totalRecords As Long
        Dim scrollID As String = ""

        If empIdTPRD <> "" Then
            doc360Class = DOC360_CLASS_TOPS_PURGE
        ElseIf flnTEDI <> "" Then
            doc360Class = DOC360_CLASS_EDI_CLAIMS
        End If

        If doc360Class <> "" Then
            initialUbound = UBound(arrResults)
            retTotalRecords = 0

            Do
                ''Note: Since Application.StatusBar is not available in VB.NET, you need to replace it with an appropriate status message display mechanism.
                'Application.StatusBar = Now() & ", GetNextSearchResults (" & doc360Class & If(boneyard, ", Boneyard", "") & ")" & ", " & (UBound(arrResults) - initialUbound) & If(totalRecords = 0, "", " out of " & totalRecords)

                GetNextSearchResults(doc360Class, boneyard:=boneyard, arrResults:=arrResults, totalRecords:=totalRecords, scrollID:=scrollID, restFindBatchSize:=restFindBatchSize, empIdTPRD:=empIdTPRD, firstNameTPRD:=firstNameTPRD, flnTEDI:=flnTEDI)

                If retTotalRecords = 0 Then
                    retTotalRecords = totalRecords
                End If
            Loop Until (scrollID = "")
        End If
    End Function

    Private Function GetNextSearchResults(ByVal doc360Class As String, ByVal boneyard As Boolean, ByRef arrResults() As Doc360ResultInfo, ByRef scrollID As String, ByVal restFindBatchSize As Integer, Optional ByRef totalRecords As Long = 0, Optional ByVal empIdTPRD As String = "", Optional ByVal firstNameTPRD As String = "", Optional ByVal flnTEDI As String = "") As Boolean
        Dim URL As String
        Dim body As String
        Dim respText As String
        Dim respData() As String
        Dim i As Integer
        Dim startScrollId As Long
        Dim endScrollId As Long
#If MY_APP_TYPE = APP_TYPE_ACCESS Then
        '  DoCmd.Hourglass(True)
#Else ' APP_TYPE_EXCEL
        Cursor.Current = Cursors.WaitCursor
#End If
        URL = "https://doc360-rest-find<server>.optum.com/doc360/api/v1/types/" & doc360Class & "/documents/find" & GetBoneyardURLParam(boneyard)
        URL = URL.Replace("<server>", GetDoc360URLServerPostfix(DOC360_SERVER))
        Debug.Print(Now() & ", " & URL)

        body = ""
        body = body & ",""totalRecords"":" & restFindBatchSize
        body = body & If(scrollID <> "", ",""scrollId"":""" & scrollID & """", "")

        body = body & ",""indexName"":""" & doc360Class & """"
        body = body & ",""criteria"":"
        body = body & "{""filterClauses"":["
        Select Case doc360Class
            Case DOC360_CLASS_TOPS_PURGE
                body = body & "{""type"":""equal"",""name"":""u_emp_id"",""value"":""" & UCase(empIdTPRD) & """}"
                If firstNameTPRD <> "" Then
                    body = body & ",{""type"":""equal"",""name"":""u_first_name"",""value"":""" & UCase(firstNameTPRD) & """}"
                End If
            Case DOC360_CLASS_EDI_CLAIMS
                body = body & "{""type"":""equal"",""name"":""u_fln_dcc"",""value"":""" & UCase(flnTEDI) & """}"
        End Select
        body = body & "]}" ' end of criteria, filterClauses
        body = "{" & Mid(body, 2) & "}" ' strip off leading command and enclose in braces
        Debug.Print(Now() & ", " & body)

        Dim request As HttpWebRequest = WebRequest.Create(URL)
        request.Method = "POST"
        request.Headers.Add("JWT", GetJWT())
        request.ContentType = "application/json"

        Using writer As New StreamWriter(request.GetRequestStream())
            writer.Write(body)
        End Using

        Dim response As HttpWebResponse = request.GetResponse()

        respText = ""

        Using reader As New StreamReader(response.GetResponseStream())
            respText = reader.ReadToEnd()
        End Using

        If respText = "" Then
            MsgBox("*** .responseText is blank")
        Else
            Debug.Print(Now() & ", " & Left(respText, 1000))

            totalRecords = Val(Mid(respText, InStr(respText, """totalRecords"":") + Len("""totalRecords"":")))

            scrollID = ""
            startScrollId = InStr(respText, "{\""sourceSystem\"":")
            If startScrollId > 0 Then
                endScrollId = InStr(Mid(respText, startScrollId), "}")
                If startScrollId > 0 And startScrollId < endScrollId Then
                    scrollID = Mid(respText, startScrollId, endScrollId)
                End If
            End If

            respData = Split(respText, """objectId"":")
            For i = LBound(respData, 1) + 1 To UBound(respData, 1)
                Call ParseDownloadMetaData(respData(i), boneyard, arrResults)
            Next i
        End If

        '#If MY_APP_TYPE = APP_TYPE_ACCESS Then
        '        DoCmd.Hourglass False
        '#Else ' APP_TYPE_EXCEL
        '    Application.Cursor = xlDefault


    End Function


    'Public Sub ParseDownloadMetaData(meta As String, boneyard As Boolean, ByRef arrResults() As Doc360ResultInfo)
    '    Dim labelNames = meta.Split("""labelName"":")
    '    Dim i As Integer
    '    Dim value As String
    '    Dim recNum As Integer
    '    'labelNames = meta.Split("""labelName"":")

    '    recNum = UBound(arrResults, 1) + 1
    '    ReDim Preserve arrResults(recNum) ' Add new record to array

    '    With arrResults(recNum)
    '        .boneyard = boneyard
    '        For i = LBound(labelNames, 1) To UBound(labelNames, 1)
    '            MsgBox(LBound(labelNames, 1))
    '            MsgBox(UBound(labelNames, 1))
    '            value = Mid(labelNames(i), InStr(labelNames(i), """value"":") + Len("""value"":"""))
    '            value = Left(value, InStr(value, """") - 1)
    '            Select Case Left(labelNames(i), InStr(labelNames(i), """") - 1)
    '                Case "u_emp_id"
    '                    .employeeID = value
    '                Case "u_gbl_doc_id"
    '                    .globalDocID = value
    '                Case "u_compound_doc"
    '                    .compoundDocID = value
    '                Case "u_first_name"
    '                    .firstName = value
    '                Case "r_page_cnt"
    '                    .pageCount = value
    '                Case "u_received_dt"
    '                    .receivedDate = CDate(Left(value, 10)) ' Extract yyyy-mm-dd from UTC datetime format (e.g., "2019-07-24T00:00:00.000Z")
    '            End Select
    '        Next i
    '    End With

    'End Sub
    Public Sub ParseDownloadMetaData(meta As String, boneyard As Boolean, ByRef arrResults() As Doc360ResultInfo)
        Dim labelNames() As String
        Dim i As Integer
        Dim value As String
        Dim recNum As Integer

        labelNames = Split(meta, """labelName"",""")
        labelNames = Split(meta, """labelName"":""")


        recNum = UBound(arrResults, 1) + 1

        ReDim Preserve arrResults(recNum) ' Add new record to array

        With arrResults(recNum)
            .boneyard = boneyard
            For i = LBound(labelNames, 1) To UBound(labelNames, 1)
                value = Mid(labelNames(i), InStr(labelNames(i), """value"":") + Len("""value"":"""))

                value = Left(value, InStr(value, """") - 1)

                Select Case Left(labelNames(i), InStr(labelNames(i), """") - 1)
                    Case "u_emp_id"
                        .employeeID = value
                    Case "u_gbl_doc_id"
                        .globalDocID = value
                    Case "u_compound_doc"
                        .compoundDocID = value
                    Case "u_first_name"
                        .firstName = value
                    Case "r_page_cnt"
                        .pageCount = value
                    Case "u_received_dt"
                        .receivedDate = CDate(Left(value, 10)) ' Extract yyyy-mm-dd from UTC datetime format (e.g., "2019-07-24T00:00:00.000Z")
                End Select
            Next i
        End With

    End Sub


    Public Function IsDoc360Production(<Out> Optional ByRef server As String = "") As Boolean
        server = DOC360_SERVER
        Return DOC360_SERVER = DOC360_SERVER_PROD
    End Function

    Public Function GetJWT() As String
        Dim URL As String
        Dim body As String

        URL = "https://doc360-rest-find<server>.optum.com/doc360/auth/v1/token/generate"
        URL = URL.Replace("<server>", GetDoc360URLServerPostfix(DOC360_SERVER))
        body = "{"
        body = body & """appId"":""" & If(isDoc360Production(), PRODUCTION_APPLICATION_ID, TEST_APPLICATION_ID) & """"
        body = body & ",""domain"":""MS"""
        body = body & ",""userId"":""" & GetMyMSID() & """"
        body = body & "}"

        Dim request As HttpWebRequest = CType(WebRequest.Create(URL), HttpWebRequest)
        request.Method = "POST"
        request.ContentType = "application/json"

        Dim bytes As Byte() = Encoding.UTF8.GetBytes(body)
        Using requestStream As Stream = request.GetRequestStream()
            requestStream.Write(bytes, 0, bytes.Length)
        End Using

        Dim responseText As String = ""
        Using response As HttpWebResponse = CType(request.GetResponse(), HttpWebResponse)
            Using responseStream As Stream = response.GetResponseStream()
                Using reader As New StreamReader(responseStream)
                    responseText = reader.ReadToEnd()
                End Using
            End Using
        End Using

        Const TOKEN_STRING As String = "{""token"":"""
        If responseText = "" Then
            MsgBox("*** .responseText is blank")
        ElseIf responseText.IndexOf(TOKEN_STRING) = -1 Then
            MsgBox(responseText)
        Else
            Dim jwt As String
            jwt = responseText.Substring(responseText.IndexOf(TOKEN_STRING) + TOKEN_STRING.Length)
            GetJWT = jwt.Replace("""}", "")
        End If
    End Function


    Public Function GetDoc360DocumentText(ByVal doc360Class As String, ByVal globalDocID As String, ByVal boneyard As Boolean) As String
        Dim docText As String
        On Error GoTo ErrHandler

        Dim URL As String
        URL = "https://doc360-rest-getcontent<server>.optum.com/doc360/api/v1/types/" & doc360Class & "/document/" & globalDocID & "/content" & GetBoneyardURLParam(boneyard)
        URL = Replace(URL, "<server>", GetDoc360URLServerPostfix(DOC360_SERVER))
        'URL = "https://doc360-rest-getcontent.optum.com/doc360/api/v1/types/u_tops_purge/document/eeba703d-ff34-4bc5-b49d-6ab939c23c0a|u_tops_purge_2022-03_v1/content?sourceSystem=R"
        Debug.Print(Now() & ", URL: " & URL)

        'Dim xmlhttp As New MSXML.XMLHTTP
        'xmlhttp.Open("GET", URL, True)
        'xmlhttp.setRequestHeader("JWT", GetJWT())
        'xmlhttp.send()
        'Do While xmlhttp.readyState <> 4 ' 0:Unsent, 1:Opened, 2:HeadersReceived, 3:Loading, 4:Done
        '    DoEvents()
        'Loop

        'docText = xmlhttp.responseText

        With CreateObject("MSXML2.XMLHTTP")
            '.Open("GET", strURL, True)
            '.setRequestHeader("JWT", accessToken)
            '.Send
            '' While .readyState <> 4 ' 0:Unsent, 1:Opened, 2:HeadersReceived, 3:Loading, 4:Done
            ''    '    'DoEvents
            'System.Threading.Thread.Sleep(120 * 1000)
            ''End While
            'docText = .responseText
            .Open("GET", URL, False)
            .setRequestHeader("JWT", GetJWT())
            .Send
            If .Status = 200 Then
                docText = .responseText
            Else
                'error
                Debug.Print("Error")
            End If

            If InStr(.responseText, vbCrLf) Then
                Debug.Print(Now() & ", Left(ResponseText,n): " & Left(.responseText, InStr(.responseText, vbCrLf) - 1))
            Else
                Debug.Print(Now() & ", ResponseText: " & .responseText)
            End If
        End With

        If InStr(docText, vbCrLf) Then
            Debug.Print(Now() & ", Left(ResponseText,n): " & Left(docText, InStr(docText, vbCrLf) - 1))
        Else
            Debug.Print(Now() & ", ResponseText: " & docText)
        End If

        Debug.Print(Now() & ", Len(docText) = " & Len(docText))

        ' Large documents have page breaks that should be removed
        docText = Replace(docText, vbFormFeed, "")
        GetDoc360DocumentText = docText
ExitFunction:
        Exit Function

ErrHandler:
        MsgBox("Error " & Err.Number & ", " & Err.Description, vbCritical, "GetDoc360DocumentText")
        Err.Clear()
        Resume Next ' support testing

    End Function
    Private Function GetBoneyardURLParam(ByVal boneyard As Boolean) As String
        If boneyard Then
            Return "?sourceSystem=B" ' boneyard (>8 yrs)
        Else
            Return "?sourceSystem=R" ' rio (<=8 yrs)
        End If
    End Function

End Module
