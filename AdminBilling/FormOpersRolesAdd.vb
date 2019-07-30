Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Class FormOpersRolesAdd
    Public Shared FSAClosed = False

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormOpersRolesAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox(FormOpers.GUID_OPER.ToString())
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlDataAdapter
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlDataAdapter("Adm_Opers_Roles", MyConnection)
            MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GUID_oper", SqlDbType.UniqueIdentifier))
            MyDataAdapter.SelectCommand.Parameters("@GUID_oper").Value = FormOpers.GUID_OPER
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
            '---------------------------------------------------------------------------------------
            Dim DS As DataSet
            DS = New DataSet()
            MyDataAdapter.Fill(DS)
            ComboBox1.DataSource = DS.Tables(1)
            ComboBox1.DisplayMember = "ROLE_NAME"
            ComboBox1.ValueMember = "ROLE_ID"
            'ComboBox1.SelectedIndex = 0
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
        End Try
    End Sub

    '-------------------------------------НАЖАТИЕ НА КНОПКУ "ДОБАВИТЬ"----------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Label_Error.Text = ""
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlCommand
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlCommand("Adm_Opers_Roles_Edit", MyConnection)
            MyDataAdapter.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.Parameters("@type").Value = 3
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@GUID_oper", SqlDbType.UniqueIdentifier))
            MyDataAdapter.Parameters("@GUID_oper").Value = FormOpers.GUID_OPER
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ROLE_ID", SqlDbType.Int))
            MyDataAdapter.Parameters("@ROLE_ID").Value = ComboBox1.SelectedValue
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
            Me.Close()
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
        End Try
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormOpersRolesAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormOpersRoles.Grid()
        FSAClosed = True
    End Sub


End Class