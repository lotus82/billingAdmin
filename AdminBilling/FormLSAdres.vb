Imports ClassAuth
Imports System.Data.SqlClient


Public Class FormLSAdres

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormLSAdres_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        '--------------------------------заполнение улиц----------------------------------------
        MyDataAdapter = New SqlDataAdapter("Adm_streets", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS1 As DataSet
        DS1 = New DataSet()
        MyDataAdapter.Fill(DS1)
        ComboBox1.DataSource = DS1.Tables(0)
        ComboBox1.DisplayMember = "Street"
        ComboBox1.ValueMember = "ID_streets"
        ComboBox1.SelectedIndex = -1
        Adres() ' получение адреса л/с
    End Sub

    '-------------------------------------ПОЛУЧЕНИЕ АДРЕСА Л/С----------------------------------------
    Private Sub Adres()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_LS_Adres", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LS", SqlDbType.BigInt))
        MyDataAdapter.SelectCommand.Parameters("@LS").Value = FormLS.LS
        '---------------------------------------------------------------------------------------
        Dim DS4 As DataSet
        DS4 = New DataSet()
        MyDataAdapter.Fill(DS4)
        TextBox1.ReadOnly = True
        TextBox1.Text = "ул. " + DS4.Tables(0).Rows.Item(0).Item(0).ToString() + ", д. " + DS4.Tables(1).Rows.Item(0).Item(0).ToString() + ", кв. " + DS4.Tables(2).Rows.Item(0).Item(0).ToString()
    End Sub

    '-------------------------------------ВЫБОР УЛИЦЫ-------------------------------------------------
    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        If ComboBox1.SelectedIndex > -1 Then
            ComboBox2.Enabled = True
            ComboBox3.Enabled = False
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlDataAdapter
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            '--------------------------------заполнение домов---------------------------------------
            MyDataAdapter = New SqlDataAdapter("Adm_buildings", MyConnection)
            MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@street_id", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@street_id").Value = ComboBox1.SelectedValue
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
            '---------------------------------------------------------------------------------------
            Dim DS2 As DataSet
            DS2 = New DataSet()
            MyDataAdapter.Fill(DS2)
            ComboBox2.DataSource = DS2.Tables(0)
            ComboBox2.DisplayMember = "Build"
            ComboBox2.ValueMember = "ID_build"
            ComboBox2.SelectedIndex = -1
        End If
    End Sub

    '-------------------------------------ВЫБОР ДОМА--------------------------------------------------
    Private Sub ComboBox2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox2.SelectionChangeCommitted
        If ComboBox2.SelectedIndex > -1 Then
            ComboBox3.Enabled = True
            ComboBox3.Text = ""
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlDataAdapter
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            '--------------------------------заполнение квартир---------------------------------------
            MyDataAdapter = New SqlDataAdapter("Adm_Flats", MyConnection)
            MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@id_builds", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@id_builds").Value = ComboBox2.SelectedValue
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
            '---------------------------------------------------------------------------------------
            Dim DS3 As DataSet
            DS3 = New DataSet()
            MyDataAdapter.Fill(DS3)
            ComboBox3.DataSource = DS3.Tables(0)
            ComboBox3.DisplayMember = "Flat"
            ComboBox3.ValueMember = "ID_flats"
            ComboBox3.SelectedIndex = -1
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

End Class