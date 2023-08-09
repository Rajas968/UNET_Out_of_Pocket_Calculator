Module ModMXIParse
    Public strMMIPG1 As String
    Public mxiEffDt As String
    Public mxiCanDt As String
    Sub Parse_MXI(strMMIPG1)
        mxiEffDt = ParseJsonResponseElement(strMMIPG1, 1, "effDt", 0)
        mxiCanDt = ParseJsonResponseElement(strMMIPG1, 1, "cancDt", 0)
    End Sub
    Function ParseJsonResponseElement(strResponse, intStartPos, strTagText, strErrorMsg) As String
        Dim intBeginPos, intEndPos, strValue

        'Determine the beginning location of value
        intBeginPos = InStr(intStartPos, strResponse, strTagText)
        If intBeginPos = 0 Then
            strErrorMsg = "Unable to find beginning of " & strTagText
            Exit Function
        End If
        intBeginPos = intBeginPos + Len(strTagText)
        If Mid(strResponse, intBeginPos, 1) = "\" Then
            intBeginPos = intBeginPos + 1
        End If
        If Mid(strResponse, intBeginPos, 1) = Chr(34) Then 'quotation mark
            intBeginPos = intBeginPos + 1
        End If
        If Mid(strResponse, intBeginPos, 1) = ": " Then
            intBeginPos = intBeginPos + 1
        End If

        'Determine the ending location of value
        intEndPos = MinPos(InStr(intBeginPos, strResponse, ","), InStr(intBeginPos, strResponse, "}"))
        If intEndPos = 0 Then
            strErrorMsg = "Unable to find ending of " & strTagText
            Exit Function
        End If
        intEndPos = intEndPos - 1

        'Gather value
        strValue = Mid(strResponse, intBeginPos, intEndPos - intBeginPos + 1)
        strValue = Replace(strValue, "\" & Chr(34), "") 'remove \"
        strValue = Replace(strValue, ":", "") 'remove \"
        strValue = Trim(strValue)

        ParseJsonResponseElement = Replace(strValue, Chr(34), "")
    End Function

    Function MinPos(pos1, pos2)
        'Return the minimum position value between two instr values that is not zero
        Dim posReturn

        If pos1 = 0 And pos2 = 0 Then
            posReturn = 0
        ElseIf pos2 = 0 Then
            posReturn = pos1
        ElseIf pos1 = 0 Then
            posReturn = pos2
        ElseIf pos1 < pos2 Then
            posReturn = pos1
        Else
            posReturn = pos2
        End If

        MinPos = posReturn
    End Function
End Module
