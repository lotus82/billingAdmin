﻿Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Class FormRolesAdd
    Public Shared FSAClosed = False

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormRolesAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label_Error.Text = ""
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormRolesAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormRoles.Grid()
        FSAClosed = True
    End Sub

    '-------------------------------------НАЖАТИЕ НА КНОПКУ "ДОБАВИТЬ"--------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Label_Error.Text = ""
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlCommand
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlCommand("Adm_Roles_Edit", MyConnection)
            MyDataAdapter.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.Parameters("@type").Value = 3
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ROLE_NAME", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@ROLE_NAME").Value = TextBox1.Text
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


End Class