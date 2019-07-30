Imports System.Data.SqlClient
Imports Cr2015




Public Class ClassAuth

    '-------------------------------------------------Криптованная часть строки подключения----------------------------
    Public Shared conCr As String = "WKLDV1QU1ZLV2MwoQGitC3YsEskbCrG+gg07gQPHACMZ4erXUTQQ9VTAU1T+icH3AsgMfVbXt11yf1yxMBsBnA=="
    Public Shared con_historyCr As String = "WKLDV1QU1ZLV2MwoQGitC3YsEskbCrG+gg07gQPHACMZ4erXUTQQ9VTAU1T+icH3AsgMfVbXt11yf1yxMBsBnA=="

    '-------------------------------------------------Расшифровка строки подключения-----------------------------------
    Public Shared Function Decrypt(key As String)
        Dim con2 As String
        con2 = Cr2015.CDCrypt.Decrypt(conCr, key)
        Return con2
    End Function

    Public Shared Function Decrypt_history(key As String)
        Dim con2_history As String
        con2_history = Cr2015.CDCrypt.Decrypt(con_historyCr, key)
        Return con2_history
    End Function

    '------------------------------------------------Получение GUID приложения-----------------------------------------
    Public Shared Function App_GUID()
        Dim APP As String = "EB5E8E10-798E-4379-99B5-003F8C4DDFA5"
        Return APP
    End Function

    Public Shared Function Host()
        Dim HOST_NAME As String = System.Net.Dns.GetHostName()
        Return HOST_NAME
    End Function

    '------------------------------------------------Переменная для строки подключения---------------------------------
    Public Shared con As String = ""
    Public Shared con_history As String = ""

    '-------------------------------------------------------Авторизация HISTORY-----------------------------------------------
    Public Shared Function Auth_history(login As String, password As String, ip As String, catlog As String, key As String)
        con_history = "Data Source=" + ip + ";Initial Catalog=" + catlog + ";Persist Security Info=True;" + Decrypt_history(key)
        Return con_history
    End Function


    '-------------------------------------------------------Авторизация------------------------------------------------
    Public Shared Function Auth(login As String, password As String, ip As String, catlog As String, key As String)
        con = "Data Source=" + ip + ";Initial Catalog=" + catlog + ";Persist Security Info=True;" + Decrypt(key)
        'con_history = "Data Source=" + ip + ";Initial Catalog=" + catlog + ";Persist Security Info=True;" + Decrypt(key)
        Dim DS As DataSet
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(con)
        MyDataAdapter = New SqlDataAdapter("select TOP 1 * from Cities", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.Text
        DS = New DataSet()
        MyDataAdapter.Fill(DS)
        MyConnection.Close()
        MyConnection.Dispose()
        Dim message As String = "Строка=" & con & " логин: " & login & " пароль: " & password & DS.Tables(0).Rows.Count
        Dim Session As Guid
        Dim HostName As String
        HostName = System.Net.Dns.GetHostName()
        Dim ip_user As String = System.Net.Dns.GetHostByName(HostName).AddressList(0).ToString
        Dim Oper As ArrayList = New ArrayList
        Dim dr As SqlDataReader
        Dim conStr As String = con
        Using con1 As New SqlConnection(conStr)
            Dim User_login As New SqlCommand("Oper_login", con1)
            User_login.CommandType = CommandType.StoredProcedure
            Dim prmUser_name = New SqlParameter("@Oper_name", SqlDbType.VarChar, 50)
            prmUser_name.Value = login
            User_login.Parameters.Add(prmUser_name)
            Dim prmUser_pass = New SqlParameter("@Oper_pass", SqlDbType.VarChar, 150)
            prmUser_pass.Value = password
            User_login.Parameters.Add(prmUser_pass)
            Dim prmUser_ip = New SqlParameter("@Oper_ip", SqlDbType.VarChar, 50)
            prmUser_ip.Value = ip_user
            User_login.Parameters.Add(prmUser_ip)
            Dim prmApp_guid = New SqlParameter("@app_guid", SqlDbType.VarChar, 64)
            prmApp_guid.Value = App_GUID()
            User_login.Parameters.Add(prmApp_guid)
            con1.Open()
            dr = User_login.ExecuteReader
            If dr.HasRows = True Then
                While dr.Read
                    Oper.Add(dr.Item(0))
                    Oper.Add(dr.Item(1))
                    Oper.Add(dr.Item(2))
                End While
            End If
            dr.Close()
            con1.Dispose()
        End Using
        Session = Oper.Item(0)
        'If (Oper.Count > 0) Then
        '    MsgBox("Пользователь: " & Oper.Item(1).ToString() & " успешно авторизован." & " Сессия: " & Session.ToString())
        'Else MsgBox("Пользователь не найден. Проверьте правильность пары Логин/Пароль!")
        'End If
        If Err.Number = 0 Then
            Return Oper
        Else
            GoTo err
        End If
        Exit Function
err:
        MsgBox(Err.Description, MsgBoxStyle.Critical)
        Return 0
    End Function

    '--------------------------------------------Выборка улиц---------------------------------------------------
    Public Shared Function Streets() As DataTable
        Dim Streets_table As DataTable = New DataTable
        For i = 0 To 10
            Streets_table.Columns.Add()
        Next
        Dim dr As SqlDataReader
        Dim conStr As String = con
        Using con1 As New SqlConnection(conStr)
            Dim Streets_select As New SqlCommand("select * from Streets", con1)
            Streets_select.CommandType = CommandType.Text
            con1.Open()
            dr = Streets_select.ExecuteReader
            If dr.HasRows = True Then
                While dr.Read
                    Streets_table.Rows.Add(dr.Item(0), dr.Item(1))
                End While
            End If
            dr.Close()
            con1.Dispose()
            Return Streets_table
        End Using
    End Function


End Class
