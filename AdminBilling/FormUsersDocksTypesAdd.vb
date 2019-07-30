Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormUsersDocksTypesAdd
    Public Shared FUDAClosed = False

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormUsersDocksTypesAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    '-------------------------------------ДОБАВЛЕНИЕ ТИПА ДОКУМЕНТА-----------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim errVvod As Integer = 0
        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_DocksTypes_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 3
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_doc_type", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_doc_type").Value = vbNull
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TypeDoc", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@TypeDoc").Value = TextBox1.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@OPER_GUID", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@OPER_GUID").Value = My.Settings.Oper_GUID
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@APP_GUID", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@APP_GUID").Value = ClassAuth.ClassAuth.App_GUID()
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@HOST", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@HOST").Value = ClassAuth.ClassAuth.Host().ToString()
                '---------------------------------------------------------------------------------------
                MyConnection.Open()
                MyDataAdapter.ExecuteNonQuery()
                MyConnection.Close()
                TextBox1.Text = ""
                Me.Close()
                Exit Sub
            Catch ex As Exception
                MsgBox(Err.Description & " " & Err.Number, MsgBoxStyle.Critical)
            End Try
        Else Exit Sub
        End If
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormUsersDocksTypes_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        FUDAClosed = True
    End Sub


End Class