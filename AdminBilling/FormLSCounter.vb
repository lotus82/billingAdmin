Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormLSCounter

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormLSCounter_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        '--------------------------------заполнение типа прибора--------------------------------------
        MyDataAdapter = New SqlDataAdapter("Adm_Counters", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS1 As DataSet
        DS1 = New DataSet()
        MyDataAdapter.Fill(DS1)
        ComboBox1.DataSource = DS1.Tables(0)
        ComboBox1.DisplayMember = "Name"
        ComboBox1.ValueMember = "ID_counter"

        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        '--------------------------------заполнение данных о приборе и показаний----------------
        MyDataAdapter = New SqlDataAdapter("Adm_LS", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LS", SqlDbType.BigInt))
        MyDataAdapter.SelectCommand.Parameters("@LS").Value = FormLS.LS
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LS_max", SqlDbType.BigInt))
        MyDataAdapter.SelectCommand.Parameters("@LS_max").Value = FormLS.LS
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Street_LS", SqlDbType.Int))
        MyDataAdapter.SelectCommand.Parameters("@Street_LS").Value = vbNull
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Build_LS", SqlDbType.Int))
        MyDataAdapter.SelectCommand.Parameters("@Build_LS").Value = vbNull
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Flat_LS", SqlDbType.Int))
        MyDataAdapter.SelectCommand.Parameters("@Flat_LS").Value = vbNull
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS2 As DataSet
        DS2 = New DataSet()
        MyDataAdapter.Fill(DS2)
        ComboBox1.SelectedValue = DS2.Tables(6).Rows.Item(0).Item(5)
        TextBox1.Text = DS2.Tables(6).Rows.Item(0).Item(6)
        DateTimePicker1.Value = DS2.Tables(6).Rows.Item(0).Item(7)
        NumericUpDown1.Value = DS2.Tables(6).Rows.Item(0).Item(11)
        NumericUpDown2.Value = DS2.Tables(6).Rows.Item(0).Item(12)
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub


End Class