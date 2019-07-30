Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormTechAdd

    Public Shared FTAClosed = False


    '-------------------------------------ДОБАВЛЕНИЕ ТЕХУЧАСТКА---------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlCommand
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlCommand("Adm_Tech_Edit", MyConnection)
        MyDataAdapter.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
        MyDataAdapter.Parameters("@type").Value = 3
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@id_tech", SqlDbType.Int))
        MyDataAdapter.Parameters("@id_tech").Value = vbNull
        '---------------------------------------------------------------------------------------
        MyDataAdapter.Parameters.Add(New SqlParameter("@NAME", SqlDbType.VarChar, 50))
        If TextBox1.Text = "" Then
            MsgBox("Поле Название не может быть пустым")
            Exit Sub
        Else
            MyDataAdapter.Parameters("@NAME").Value = TextBox1.Text
        End If

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
    Private Sub FormTechAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        FormTech.Grid()
        FTAClosed = True
    End Sub


End Class