Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormCountersAdd

    Public Shared FCountersAddClosed = False

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormCountersAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Counters", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS As DataSet
        DS = New DataSet()
        MyDataAdapter.Fill(DS)
        ComboBox_RES_ID.DataSource = DS.Tables(1)
        ComboBox_RES_ID.DisplayMember = "RES_NAME"
        ComboBox_RES_ID.ValueMember = "CNTR_RES_ID"
        ComboBox_SERVICE_ID.DataSource = DS.Tables(2)
        ComboBox_SERVICE_ID.DisplayMember = "ID_Service"
        ComboBox_SERVICE_ID.ValueMember = "ID_Service"
    End Sub

    '-------------------------------------ДОБАВЛЕНИЕ ПРИБОРА------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim errVvod As Integer = 0
        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Counters_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 3
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@CNTR_RES_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@CNTR_RES_ID").Value = ComboBox_RES_ID.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SERVICE_ID", SqlDbType.Char, 4))
                MyDataAdapter.Parameters("@SERVICE_ID").Value = ComboBox_SERVICE_ID.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@MODEL_NAME", SqlDbType.VarChar, 100))
                MyDataAdapter.Parameters("@MODEL_NAME").Value = TextBox_Model_Name.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@DIGIT_CAPACITY", SqlDbType.Int))
                MyDataAdapter.Parameters("@DIGIT_CAPACITY").Value = NumericUpDown_Digit_Capacity.Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@PRECISION", SqlDbType.Int))
                MyDataAdapter.Parameters("@PRECISION").Value = NumericUpDown_Precision.Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@COEFF", SqlDbType.Decimal, 18, 2))
                MyDataAdapter.Parameters("@COEFF").Value = NumericUpDown_Coeff.Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@PERIOD_VERIFICATION", SqlDbType.Int))
                MyDataAdapter.Parameters("@PERIOD_VERIFICATION").Value = NumericUpDown_Period_Ver.Value
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
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
        Else Exit Sub
        End If
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormCountersAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox_Model_Name.Text = ""
        FCountersAddClosed = True
    End Sub

End Class