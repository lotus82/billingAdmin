Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception


Public Class FormSuppliersAdd
    Public Shared FSAClosed = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlCommand
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlCommand("Adm_Suppliers_Edit", MyConnection)
            MyDataAdapter.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.Parameters("@type").Value = 3
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPLIER_ID", SqlDbType.Int))
            MyDataAdapter.Parameters("@SUPPLIER_ID").Value = vbNull
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPLIER_NAME", SqlDbType.VarChar, 100))
            MyDataAdapter.Parameters("@SUPPLIER_NAME").Value = TextBox1.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@INN", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@INN").Value = TextBox2.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@KPP", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@KPP").Value = TextBox3.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@BIK", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@BIK").Value = TextBox4.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@RS", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@RS").Value = TextBox5.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@KS", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@KS").Value = TextBox6.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@BANK_NAME", SqlDbType.VarChar, 100))
            MyDataAdapter.Parameters("@BANK_NAME").Value = TextBox7.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ADDRESS", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@ADDRESS").Value = TextBox8.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@PHONE", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@PHONE").Value = TextBox9.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@FAX", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@FAX").Value = TextBox10.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@MAIL", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@MAIL").Value = TextBox11.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@WEB", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@WEB").Value = TextBox12.Text
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@WORK_DAYS", SqlDbType.VarChar, 100))
            MyDataAdapter.Parameters("@WORK_DAYS").Value = TextBox13.Text
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
            Exit Sub
        Catch ex As Exception
            MsgBox(Err.Description & " " & Err.Number, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub FormSuppliersAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        FSAClosed = True
    End Sub
End Class