Imports System.Windows.Forms.VisualStyles
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Newtonsoft.Json.Linq

Module ModExeFile
    Public Local_Path = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    Sub LaunchAPI_Exe()
        Dim exeFilePath As String = Local_Path & "\Insight Software\Macro Express\Macro Files\NET\API\UNET_API.exe"
        System.Diagnostics.Process.Start(exeFilePath)
    End Sub

    Sub Launch_vbs()
        Dim scriptFilePath As String = Local_Path & "\_n\ntr\Production Files\VBS\E_Signon.vbs"
        System.Diagnostics.Process.Start("WScript.exe", Chr(34) + scriptFilePath + Chr(34))
    End Sub

End Module
