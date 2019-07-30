Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormStreetAdd

    Public Shared FSAClosed = False

    '-------------------------------------ДОБАВЛЕНИЕ УЛИЦЫ--------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlCommand
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlCommand("Adm_Street_Edit", MyConnection)
        MyDataAdapter.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
        MyDataAdapter.Parameters("@type").Value = 3
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@id_street", SqlDbType.Int))
        MyDataAdapter.Parameters("@id_street").Value = vbNull
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@STREET_NAME", SqlDbType.VarChar, 50))
        MyDataAdapter.Parameters("@STREET_NAME").Value = TextBox1.Text
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@KLADR", SqlDbType.VarChar, 50))
        MyDataAdapter.Parameters("@KLADR").Value = TextBox2.Text
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@OKATO", SqlDbType.VarChar, 50))
        MyDataAdapter.Parameters("@OKATO").Value = TextBox3.Text
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@OKTMO", SqlDbType.VarChar, 50))
        MyDataAdapter.Parameters("@OKTMO").Value = TextBox4.Text
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
        Try
            MyConnection.Open()
            MyDataAdapter.ExecuteNonQuery()
            MyConnection.Close()
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormStreetAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        FormStreets.Grid()
        FSAClosed = True
    End Sub

End Class