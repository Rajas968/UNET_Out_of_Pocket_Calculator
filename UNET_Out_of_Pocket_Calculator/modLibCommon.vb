Module modLibCommon
    Public Const LENGTH_MPIN = 9
    Public Const LENGTH_TIN = 9
    Public Function isNullOrEmpty(ByVal var As Object) As Boolean
        If IsDBNull(var) OrElse var Is Nothing OrElse var.ToString().Trim() = "" Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function Nz(ByVal var As Object, ByVal ifNull As Object) As Object
        If IsDBNull(var) OrElse var Is Nothing Then
            Return ifNull
        Else
            Return var
        End If

    End Function


End Module
