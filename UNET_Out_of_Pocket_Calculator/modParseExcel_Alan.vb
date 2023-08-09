Imports System.Security.Claims
Imports Microsoft.Office.Interop.Excel

Module modParseExcel_Alan
    Public Const STATUS_PROCESSING_PREFIX = "Processing"
    Public Const STATUS_SKIP_PREFIX = "SKIPPED - "
    Public Const STATUS_ERROR_PREFIX = "*** ERROR - "

    Public Const COMP_DOC_SKIP_POSTFIX = " **SKIPPED"
    Public Const COMP_DOC_ERROR_POSTFIX = " **ERROR"
    Public Const COMP_DOC_DUP_POSTFIX = " **DUP"
    Dim rwcount As Integer = 1
    Dim ColCount As Integer = 1

    Private Structure MHIType
        Public BCV As String
        Public BP As String
        Public CD_AMOUNT As String ' ("C/D", value)
        Public CD_CODE As String ' ("C/D", 1st letter)
        Public CHARGE As String
        Public D1 As String ' ("D" 1st one)
        Public D2 As String ' ("D" 2nd one)
        Public DD1 As String
        Public DD2 As String ' ("DD" 2nd one)
        Public FSTDT As String
        Public LSTDT As String
        Public MCV As String ' allowed amount
        Public N As String
        Public NBR As String
        Public NOTCOV As String
        Public OV As String
        Public P As String
        Public PCT1 As String ' ("%" 1st one)
        Public PCT2 As String ' ("%" 2nd one)
        Public PS As String
        Public rc As String
        Public S As String
        Public SVC As String
        Public TOTAL_PAID As String
        Public OPTIONAL_REC_NUM As String
        Public OPTIONAL_REP_NAME As String
        Public OPTIONAL_PAGE_COUNT As Integer
        Public OPTIONAL_LINE_COUNT As Long 'LK#7017 - Added
    End Structure

    Private Structure ProvDataType
        Public ACT_DRAFT As String
        Public ADJ As String
        Public DRAFT As String
        Public fln As String ' needed to help determine duplicitly
        Public PROCDATE As String
        Public TIN_FULL As String
        Public TIN_PREFIX As String
        Public TIN As String
        Public TIN_SUFFIX As String
        Public TOTAL_BILLED As String
        Public TOTAL_PAID As String
    End Structure

    Private Structure ClaimDataType
        Public ClaimNumber As String
        Public DiagCode1 As String
        Public DiagCode4 As String
        Public DiagDesc As String
        Public OfficeNumber As String 'need this with FLN
        Public PatientFirstName As String
        Public PatientLastName As String
        Public PatientName As String
        Public Relationship As String
    End Structure


    Private MHI() As MHIType
    Private prov As ProvDataType
    Private Claim As ClaimDataType

    Private svcCnt As Long
    Private latestRow As Long

    Private wb As Workbook
    Private sht As Worksheet
    Private shtLog As Worksheet


    Public Function IsCompDocProcessed(ByVal compDoc As String, ByRef retRecNum As String) As Boolean
        Dim rng As Range

        'rng = shtLog.Columns(colLogCompDoc).Find(What:=compDoc, LookIn:=xlValues, LookAt:=xlWhole)
        If rng Is Nothing Then
            retRecNum = ""
            IsCompDocProcessed = False
        Else
            ' retRecNum = shtLog.Cells(rng.Row, colLogRecNum)
            IsCompDocProcessed = True
        End If

    End Function
    Public Enum DataTypeEnum
        ServiceData = 1
        ProviderData = 2
        ClaimData = 3
    End Enum
    Public Sub ParseData(dataType As DataTypeEnum, recNumber As String, compDoc As String, pageCount As Integer, lineCount As Long, line1 As String, line2 As String, Optional line3 As String = "", Optional line4 As String = "")

        'Dim MHI() As MHIType
        ' There will be several sets of service data before its provider data is sent
        ' so they must be kept until then.
        If dataType = DataTypeEnum.ServiceData Then
            svcCnt = svcCnt + 1
            ReDim Preserve MHI(svcCnt)

            ' Note: Dollar amounts could be negative so need to account for
            ' the sign on the left.

            ' The following are in field position order

            MHI(svcCnt).PS = Mid(line1, 4, 2) ' ("PS")
            MHI(svcCnt).SVC = Mid(line1, 8, 6) ' ("SVC")
            MHI(svcCnt).FSTDT = Mid(line1, 17, 6) ' ("FST DT") ** same as in ParseDocument
            MHI(svcCnt).LSTDT = Mid(line1, 25, 6) ' ("LST DT")
            MHI(svcCnt).NBR = Mid(line1, 34, 3) ' ("NBR")
            MHI(svcCnt).OV = Mid(line1, 40, 2) ' ("OV")
            MHI(svcCnt).P = Mid(line1, 44, 1) ' ("P")
            MHI(svcCnt).N = Mid(line1, 46, 1) ' ("N")
            MHI(svcCnt).rc = Mid(line1, 49, 2) ' ("RC")
            MHI(svcCnt).CHARGE = Mid(line1, 52, 9) ' ("CHARGE")
            MHI(svcCnt).NOTCOV = Mid(line1, 64, 9) ' (NOT COV")

            '** DD1 and all subsequent fields in Line 2 are expanded by 1 character and then trimmed
            '** to accomodate the anomaly where DD1 and DD2 are sometimes 8 characters, not 9
            MHI(svcCnt).BCV = Mid(line2, 3, 9) ' ("B CV") ''LK#751
            MHI(svcCnt).DD1 = Trim(Mid(line2, 14, 9)) ' ("DD", 1st one)
            MHI(svcCnt).D1 = Trim(Mid(line2, 24, 2)) ' ("D", 1st one)
            MHI(svcCnt).PCT1 = Trim(Mid(line2, 31, 4)) ' ("%", 1st one)
            MHI(svcCnt).BP = Trim(Mid(line2, 36, 10)) ' ("BP")
            MHI(svcCnt).S = Trim(Mid(line2, 47, 10)) ' ("S")
            MHI(svcCnt).MCV = Trim(Mid(line2, 58, 10)) ' ("MCV")
            MHI(svcCnt).DD2 = Trim(Mid(line2, 69, 10)) ' ("DD", 2nd one)

            MHI(svcCnt).D2 = Mid(line3, 4, 2) ' ("D", 2nd one)
            MHI(svcCnt).PCT2 = Mid(line3, 11, 3) ' ("%", 2nd one)
            MHI(svcCnt).TOTAL_PAID = Mid(line3, 16, 9) ' ("P")
            MHI(svcCnt).CD_CODE = Mid(line3, 27, 1) ' ("C/D", 1st letter)
            MHI(svcCnt).CD_AMOUNT = Mid(line3, 29, 9) ' ("C/D", value)

            ' Always collect record info
            MHI(svcCnt).OPTIONAL_REC_NUM = recNumber
            MHI(svcCnt).OPTIONAL_REP_NAME = compDoc
            MHI(svcCnt).OPTIONAL_PAGE_COUNT = pageCount
            MHI(svcCnt).OPTIONAL_LINE_COUNT = lineCount

        ElseIf dataType = DataTypeEnum.ClaimData Then
            ' Note: Although there are several claim headers for each claim,
            ' the diagnostic codes should be the same, so OK to just overwrite
            ' with the same info

            ' The following are in field position order

            Claim.PatientName = Trim(Mid(line1, 20, 11)) & " " & Trim(Mid(line1, 31, 21))
            Claim.PatientFirstName = Trim(Mid(line1, 20, 11))
            Claim.PatientLastName = Trim(Mid(line1, 31, 21))

            Claim.Relationship = Trim(Mid(line2, 76, 5)) ' ("REL")

            Claim.OfficeNumber = Trim(Mid(line3, 34, 5)) ' ("OFFICE NUMBER")

            Claim.ClaimNumber = Trim(Mid(line4, 9, 6)) ' ("CLM #")
            Claim.DiagDesc = Mid(line4, 18, 9) ' ("DX")
            Claim.DiagCode1 = Mid(line4, 53, 1) ' ("CAUSE") 1st character)
            Claim.DiagCode4 = Mid(line4, 54, 4) ' ("CAUSE") last 4 characters)

        ElseIf dataType = DataTypeEnum.ProviderData Then

            ' The following are in field position order

            ' Note: Dollars are 10 digits here, although 9 digits in claim data
            prov.TIN_FULL = Mid(line1, 4, 17) ' ("PROVIDER NBR")
            prov.TIN_PREFIX = Left(prov.TIN_FULL, 1)
            prov.TIN = Mid(prov.TIN_FULL, 3, 9)
            prov.TIN_SUFFIX = Right(prov.TIN_FULL, 5)

            prov.DRAFT = Mid(line1, 40, 10) ' ("DR#")
            prov.PROCDATE = Mid(line1, 51, 6) ' ("DATE")
            prov.ADJ = Mid(line1, 58, 9) ' ("ADJ")
            prov.TOTAL_BILLED = Mid(line1, 69, 10) ' ("CHG")

            prov.TOTAL_PAID = Trim(Mid(line2, 3, 10)) ' ("PD")
            prov.fln = Mid(line2, 15, 10) ' ("FLN")
            prov.ACT_DRAFT = Mid(line2, 40, 10) ' ("ACT DR#")

            writeoutputtable()
            '  Call WriteData   ' Write collected data to Excel sheet

            ' Reset for new service (not necessarily a new claim)
            svcCnt = 0
            Erase MHI ' clear array

        Else
            MsgBox("Unexpected data", vbCritical)
        End If



    End Sub

    Private Sub writeoutputtable()

        Dim cnt As Long
        Dim arr_row_string() As String
        'Dim mhiobj As MHIType
        If latestRow = 0 Then latestRow = 2
        cnt = 1
        Dim row_string, COL_FROM, COL_THRU, COL_SVC, COL_PS, COL_NBR, COL_OV, COL_P, COL_N, COL_RC, COL_CHARGE, COL_NOT_COV, COL_B_M, COL_COVERED, COL_DEDUCT, COL_D, COL_PCT, COL_PAID, COL_S, COL_D_C As String
        Dim COL_SANC, COL_CAUSE_CODE, COL_P1, COL_TIN, COL_SUFFIX, COL_DRAFT, COL_PROC_DATE, COL_ADJ, COL_TOTAL_BILLED, COL_TOTAL_PAID, COL_ICN, COL_SUF, COL_FLN, COL_PRS, COL_SI, COL_PT_NAME As String
        Dim COL_DOT, COL_PT_REL, COL_PT_NAME2, COL_OPTIONAL_REC_NUM, COL_OPTIONAL_REP_NAME, COL_OPTIONAL_PAGE_COUNT, COL_OPTIONAL_LINE_COUNT As String


        While cnt <= svcCnt

            Dim newRow As DataRow = Purgemain.output.NewRow()

            COL_FROM = FormatDate(MHI(cnt).FSTDT)
            COL_THRU = FormatDate(MHI(cnt).LSTDT)
            COL_SVC = MHI(cnt).SVC
            COL_PS = MHI(cnt).PS
            COL_NBR = CInt(MHI(cnt).NBR) ' strips off leading zeroes
            COL_OV = MHI(cnt).OV
            COL_P = MHI(cnt).P
            COL_N = MHI(cnt).N
            COL_RC = MHI(cnt).rc
            COL_CHARGE = MHI(cnt).CHARGE


            COL_NOT_COV = MHI(cnt).NOTCOV
            If Trim(COL_NOT_COV) = "" Then
                COL_NOT_COV = "0.00"
            End If

            COL_B_M = "NA"
            COL_COVERED = MHI(cnt).MCV
            COL_DEDUCT = MHI(cnt).DD2
            COL_D = MHI(svcCnt).D2
            COL_D = Trim(COL_D)
            If COL_D = "" Then
                COL_D = "0.00"
            End If
            COL_PCT = CInt(MHI(cnt).PCT2) & "%" ' format as text
            COL_PAID = MHI(cnt).TOTAL_PAID
            COL_S = MHI(cnt).S
            COL_D_C = MHI(svcCnt).BP
            COL_SANC = "N"
            COL_CAUSE_CODE = Claim.DiagCode1
            COL_P1 = prov.TIN_PREFIX
            COL_TIN = prov.TIN
            COL_SUFFIX = prov.TIN_SUFFIX
            COL_DRAFT = prov.DRAFT
            COL_PROC_DATE = FormatDate(prov.PROCDATE)
            COL_ADJ = prov.ADJ
            COL_TOTAL_BILLED = prov.TOTAL_BILLED
            COL_TOTAL_PAID = prov.TOTAL_PAID
            COL_ICN = "NA" ' Alan Satin says don't fill in any ICN value ' "9999999999"
            COL_SUF = "01"
            COL_FLN = "800 " & prov.fln ' 800 is in lieu of the office number
            COL_PRS = "N"
            COL_SI = "N"
            COL_PT_NAME = Claim.PatientFirstName
            COL_DOT = "NA"
            COL_PT_REL = Claim.Relationship
            COL_PT_NAME2 = Claim.PatientFirstName & "/" & Claim.Relationship
            COL_OPTIONAL_REC_NUM = MHI(cnt).OPTIONAL_REC_NUM
            COL_OPTIONAL_REP_NAME = MHI(cnt).OPTIONAL_REP_NAME
            COL_OPTIONAL_PAGE_COUNT = MHI(cnt).OPTIONAL_PAGE_COUNT
            COL_OPTIONAL_LINE_COUNT = MHI(cnt).OPTIONAL_LINE_COUNT
            'Form1.cnt_output & ";" &
            row_string = COL_FROM & ";" & COL_THRU & ";" & COL_SVC & ";" & COL_PS & ";" & COL_NBR & ";" & COL_OV & ";" & COL_P & ";" & COL_N & ";" & COL_RC & ";" & COL_CHARGE & ";" & COL_NOT_COV & ";" & COL_B_M & ";" & COL_COVERED & ";" & COL_DEDUCT & ";" & COL_D & ";" & COL_PCT & ";" & COL_PAID & ";" & COL_S & ";" & COL_D_C & ";"

            row_string = row_string & COL_SANC & ";" & COL_CAUSE_CODE & ";" & COL_P1 & ";" & COL_TIN & ";" & COL_SUFFIX & ";" & "" & ";" & COL_DRAFT & ";" & COL_PROC_DATE & ";" & COL_ADJ & ";" & COL_TOTAL_BILLED & ";" & COL_TOTAL_PAID & ";" & COL_ICN & ";" & COL_SUF & ";" & COL_FLN & ";" & COL_PRS & ";" & COL_SI & ";" & COL_PT_NAME & ";"

            row_string = row_string & COL_DOT & ";" & COL_PT_REL & ";" & COL_PT_NAME2 & ";" & COL_OPTIONAL_REC_NUM & ";" & COL_OPTIONAL_REP_NAME & ";" & COL_OPTIONAL_PAGE_COUNT & ";" & COL_OPTIONAL_LINE_COUNT
            arr_row_string = Split(row_string, ";")
            For innerIndex = 0 To UBound(arr_row_string) - 1
                'If Trim(arr_row_string(innerIndex)) = "" Then
                '    newRow(innerIndex) = 0
                'Else
                newRow(innerIndex) = Trim(arr_row_string(innerIndex))
                'End If

                'Form1.DGridMHI.Rows(rwcount).Cells(ColCount).Value = arr_row_string(innerIndex)
                'ColCount = ColCount + 1
            Next
            '            rwcount = rwcount + 1            
            Purgemain.output.Rows.Add(newRow)

            Purgemain.cnt_output += 1
            cnt = cnt + 1
            latestRow = latestRow + 1

        End While


    End Sub
    Public Function FormatDate(dt As String) As String
        If Len(Trim(dt)) = 6 Then
            FormatDate = Left(dt, 2) & "/" & Mid(dt, 3, 2) & "/20" & Right(dt, 2)
        Else
            FormatDate = ""
        End If
    End Function
End Module
