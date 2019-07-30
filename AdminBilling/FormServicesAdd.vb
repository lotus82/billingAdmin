Imports ClassAuth
Imports System.Data.SqlClient


Public Class FormServicesAdd

    Public Shared FServicesAClosed = False

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormServicesAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Services", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS As DataSet
        DS = New DataSet()
        MyDataAdapter.Fill(DS)

        ComboBox1.DataSource = DS.Tables(1)
        ComboBox1.DisplayMember = "RL"
        ComboBox1.ValueMember = "ID_RL"
        ComboBox2.DataSource = DS.Tables(2)
        ComboBox2.DisplayMember = "SUPPLIER_NAME"
        ComboBox2.ValueMember = "SUPPLIER_ID"
        ComboBox3.DataSource = DS.Tables(3)
        ComboBox3.DisplayMember = "SUPPLIER_PAYEE_NAME"
        ComboBox3.ValueMember = "SUPPLIER_PAYEE_ID"
    End Sub

    '-------------------------------------ДОБАВЛЕНИЕ УСЛУГИ-------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim errVvod As Integer = 0
        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Services_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 3
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_Service", SqlDbType.Char, 4))
                If TextBox1.Text.Length < 4 Then
                    TextBox1.Text = Replace(TextBox1.Text.PadLeft(4), " ", "0")
                    MyDataAdapter.Parameters("@ID_Service").Value = TextBox1.Text
                Else
                    MyDataAdapter.Parameters("@ID_Service").Value = TextBox1.Text
                End If
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Service_Name", SqlDbType.VarChar, 64))
                MyDataAdapter.Parameters("@Service_Name").Value = TextBox2.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Short_Name", SqlDbType.VarChar, 32))
                MyDataAdapter.Parameters("@Short_Name").Value = TextBox3.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@IS_ODN", SqlDbType.Int))
                MyDataAdapter.Parameters("@IS_ODN").Value = ComboBox1.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@IS_DEPENDS", SqlDbType.Int))
                MyDataAdapter.Parameters("@IS_DEPENDS").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NORM_COEFF", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NORM_COEFF").Value = NumericUpDown1.Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NDS", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NDS").Value = NumericUpDown2.Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@USPN_CODE", SqlDbType.Int))
                MyDataAdapter.Parameters("@USPN_CODE").Value = NumericUpDown3.Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPL_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@SUPPL_ID").Value = ComboBox2.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPL_PAYEE_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@SUPPL_PAYEE_ID").Value = ComboBox3.SelectedValue
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
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
        Else Exit Sub
        End If
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormServicesAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        FServicesAClosed = True
    End Sub

    '-------------------------------------ВВОД ЗНАЧЕНИЯ В ПОЛЕ "ID"-----------------------------------
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim Start As Integer
        Start = TextBox1.SelectionStart
        If TextBox1.Text.Length > 0 Then
            If (Char.IsLetterOrDigit(TextBox1.Text(TextBox1.Text.Length - 1))) Then

            Else
                TextBox1.Text = Replace(TextBox1.Text, TextBox1.Text(TextBox1.Text.Length - 1), "")
                TextBox1.SelectionStart = Start
            End If
        End If
    End Sub

End Class