Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Module Tracking
    Public ceiDetails
    Public filePath
    Public ceiFile As Boolean
    Public pmiFile As Boolean
    Public ceiTextFile As String
    Public pmiTextFile As String
    Dim inputPFile As String
    Dim strParameter As String

    Public Local_Function_Folder = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & "Insight Software\Macro Express\Macro Files\TXT\"


    Function Get_AHI() As Boolean

        Dim SourcePath As String = Local_Function_Folder & "AHI_detail.txt" 'This is just an example string and could be anything, it maps to fileToCopy in your code.        
        Dim Filename As String = System.IO.Path.GetFileName(SourcePath) 'get the filename of the original file without the directory on it
        Dim filePath As String = System.IO.Path.Combine(Local_Function_Folder, Filename) 'combines the saveDirectory and the filename to get a fully qualified path.

        If System.IO.File.Exists(filePath) Then
            pmiTextFile = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & "Insight Software\Macro Express\Macro Files\TXT\" & "AHI_detail.txt"
            pmiFile = True
        Else
            Get_AHI = False
            MsgBox("AHI Details not available") 'the file doesn't exist
        End If
    End Function
    Function Get_MXI() As Boolean
        Dim SourcePath As String = Local_Function_Folder & "AHI_detail.txt" 'This is just an example string and could be anything, it maps to fileToCopy in your code.        
        Dim Filename As String = System.IO.Path.GetFileName(SourcePath) 'get the filename of the original file without the directory on it
        Dim filePath As String = System.IO.Path.Combine(Local_Function_Folder, Filename) 'combines the saveDirectory and the filename to get a fully qualified path.

        If System.IO.File.Exists(filePath) Then
            pmiTextFile = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & "Insight Software\Macro Express\Macro Files\TXT\" & "AHI_detail.txt"
            pmiFile = True
        Else
            MsgBox("MXI Details not available") 'the file doesn't exist
        End If
    End Function
    Function Get_MRI() As Boolean
        Dim SourcePath As String = Local_Function_Folder & "MRI_detail.txt" 'This is just an example string and could be anything, it maps to fileToCopy in your code.        
        Dim Filename As String = System.IO.Path.GetFileName(SourcePath) 'get the filename of the original file without the directory on it
        Dim filePath As String = System.IO.Path.Combine(Local_Function_Folder, Filename) 'combines the saveDirectory and the filename to get a fully qualified path.

        If System.IO.File.Exists(filePath) Then
            pmiTextFile = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & "Insight Software\Macro Express\Macro Files\TXT\" & "MRI_detail.txt"
            pmiFile = True
        Else
            MsgBox("MRI Details not available") 'the file doesn't exist
        End If
    End Function
    Public Function Input_Parameter(ByVal strParameter As String)

        Dim SourcePath As String = Local_Function_Folder & "API_Input.txt" 'This is just an example string and could be anything, it maps to fileToCopy in your code.
        Dim Filename As String = System.IO.Path.GetFileName(SourcePath) 'get the filename of the original file without the directory on it
        Dim filePath As String = System.IO.Path.Combine(Local_Function_Folder, Filename) 'combines the saveDirectory and the filename to get a fully qualified path.

        If System.IO.File.Exists(filePath) Then
            Dim objWriter As New StreamWriter(filePath)
            objWriter.Write(strParameter)
            objWriter.Close()
            Call LaunchAPI_Exe()
        ElseIf Not System.IO.File.Exists(filePath) Then
            System.IO.File.Create(filePath).Dispose()
            Dim objWriter As New StreamWriter(filePath)
            objWriter.Write(strParameter)
            objWriter.Close()
            Call LaunchAPI_Exe()
        End If

    End Function

End Module