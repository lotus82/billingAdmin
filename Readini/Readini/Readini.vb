
Public Class Readini
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal Section As String, ByVal Key As String, ByVal putStr As String, ByVal INIfile As String) As Integer

    Public Shared Function ReadIni(ByVal Section As String, ByVal Param As String) As String

        Dim app As String
        app = My.Application.Info.AssemblyName
        Dim exit_code As String
        If FileIO.FileSystem.FileExists(My.Application.Info.DirectoryPath & "\" & app & ".ini") = False Then
            MsgBox("Не удалось загрузить файл конфигурации", MsgBoxStyle.Critical)
            Err.Description = "Не удалось загрузить файл конфигурации"
            Err.Number = 101
            exit_code = "101"

            Return exit_code
            Exit Function
        End If
        Try
            Dim rc As String = Strings.StrDup(255, vbNullChar)
            Dim x As Integer

            x = GetPrivateProfileString(Section, Param, "", rc, 255, My.Application.Info.DirectoryPath & "\" & app & ".ini")
            If x <> 0 Then rc = Strings.Left(rc, x)
            Return rc
            If x = 0 Then WritePrivateProfileString("APP", "APP_GUID", Guid.NewGuid.ToString, My.Application.Info.DirectoryPath & "\" & app & ".ini")
        Catch ex As Exception

            MsgBox(ex.Message)
            Return Nothing
        End Try
        Exit Function

    End Function
    Public Shared Function WriteIni(ByVal Section As String, ByVal Param As String, ByVal value As String) As String
        Dim app As String
        app = My.Application.Info.AssemblyName
        WritePrivateProfileString(Section, Param, value, My.Application.Info.DirectoryPath & "\" & app & ".ini")
        Return Nothing
    End Function
End Class
