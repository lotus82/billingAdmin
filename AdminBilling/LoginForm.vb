Imports System.Globalization
Imports ClassAuth
Imports Readini
Imports Cr2015




Public Class LoginForm

    Function getMD5Hash(ByVal strToHash As String) As String
        Dim md5Obj As New Security.Cryptography.MD5CryptoServiceProvider
        Dim ByteHash() As Byte = System.Text.Encoding.Default.GetBytes(strToHash)
        ByteHash = md5Obj.ComputeHash(ByteHash)
        Dim strResult As String = ""
        For Each b As Byte In ByteHash
            strResult += b.ToString("x2")
        Next
        Return strResult
    End Function

    ' TODO: вставить код для настраиваемой аутентификации с использованием переданного имени пользователя и пароля 
    ' (См. http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' Пользовательский принципал можно затем подключить к принципалу потока следующим образом: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' где CustomPrincipal - реализация интерфейса IPrincipal, используемая для аутентификации. 
    ' Впоследствии My.User будет возвращать идентификационную информацию, заключенную в объекте CustomPrincipal,
    ' например, имя пользователя, отображаемое имя и т.д.
    Dim app_key As String = "m-irc.ru"

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        On Error GoTo err
        If String.IsNullOrEmpty(UsernameTextBox.Text) Then
            Err.Description = "Введите имя польлователя!!!"
            GoTo err
        ElseIf String.IsNullOrEmpty(PasswordTextBox.Text) Then
            Err.Description = "Введите пароль!!!"
            GoTo err
        End If
        If Err.Number = 0 Then
            If (Readini.Readini.ReadIni("Connect", "ip") = "101") Then
                Application.Exit()
                Me.Finalize()
                ' Me.Close()
            End If

            Dim ip As String = Readini.Readini.ReadIni("Connect", "ip")
            Dim ip_history As String = Readini.Readini.ReadIni("Connect", "ip_history")

            Dim catalog As String = Readini.Readini.ReadIni("Connect", "catalog")
            Dim catalog_history As String = Readini.Readini.ReadIni("Connect", "catalog_history")
            'ClassAuth.ClassAuth.Auth(UsernameTextBox.Text, getMD5Hash(PasswordTextBox.Text), ip, catalog, app_key)

            Dim Ses_list As ArrayList = New ArrayList
            Dim Ses_history As String = ""

            Ses_list = ClassAuth.ClassAuth.Auth(UsernameTextBox.Text, getMD5Hash(PasswordTextBox.Text), ip, catalog, app_key)
            Ses_history = ClassAuth.ClassAuth.Auth_history(UsernameTextBox.Text, getMD5Hash(PasswordTextBox.Text), ip_history, catalog_history, app_key)

            Dim Ses_id As String = Cr2015.CDCrypt.Encrypt(Ses_list.Item(0).ToString(), app_key)
            Dim Ses_id_history As String = Cr2015.CDCrypt.Encrypt(Ses_history, app_key)

            Dim Oper_FIO As String = Ses_list.Item(1).ToString()
            Dim Oper_GUID As String = Ses_list.Item(2).ToString()

            My.Settings.Sess_id = Ses_id
            My.Settings.Ses_id_history = Ses_id_history

            My.Settings.Oper_FIO = Oper_FIO
            My.Settings.Oper_GUID = Oper_GUID
            'MsgBox("Криптованная сессия history в параметре: " & My.Settings.Ses_id_history.ToString()) '& ", Oper_GUID: " & Ses_list.Item(2).ToString())
            'MsgBox(ClassAuth.ClassAuth.con_history)
            FormAdmin.Show()
            'Dim f As New FormAdmin
            Me.Close()
        Else
            GoTo err
        End If
        Exit Sub
err:
        MsgBox(Err.Description, MsgBoxStyle.Critical)
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UsernameTextBox.Text = "1"
        PasswordTextBox.Text = "test"
    End Sub

    Private Sub PasswordTextBox_Enter(sender As Object, e As EventArgs) Handles PasswordTextBox.Enter
        InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(New CultureInfo("en-US"))
    End Sub

    Private Sub UsernameTextBox_Enter(sender As Object, e As EventArgs) Handles UsernameTextBox.Enter
        InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(New CultureInfo("en-US"))
    End Sub
End Class
