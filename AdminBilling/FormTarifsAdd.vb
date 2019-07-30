Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormTarifsAdd

    Public Shared FTarifsAClosed = False

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormTarifsAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Tarifs", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
        MyDataAdapter.SelectCommand.Parameters("@type").Value = 3
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SUPPL_ID", SqlDbType.Int))
        MyDataAdapter.SelectCommand.Parameters("@SUPPL_ID").Value = vbNull
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ID_Service", SqlDbType.VarChar, 4))
        MyDataAdapter.SelectCommand.Parameters("@ID_Service").Value = vbNull
        '---------------------------------------------------------------------------------------
        Dim DS As DataSet
        DS = New DataSet()
        MyDataAdapter.Fill(DS)
        'ComboBox1.DataSource = DS.Tables(1)
        'ComboBox1.DisplayMember = "Short_Name"
        'ComboBox1.ValueMember = "ID_Service"

        ComboBox2.DataSource = DS.Tables(1)
        ComboBox2.DisplayMember = "UNITS_TYPE_ED"
        ComboBox2.ValueMember = "ID"

        ComboBox3.DataSource = DS.Tables(2)
        ComboBox3.DisplayMember = "UNITS_TYPE_ED"
        ComboBox3.ValueMember = "ID"

    End Sub

    '-------------------------------------ДОБАВЛЕНИЕ ТАРИФА------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim errVvod As Integer = 0
        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Tarifs_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 3
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TAR_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@TAR_ID").Value = vbNull
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SERVICE_ID", SqlDbType.Char, 4))
                MyDataAdapter.Parameters("@SERVICE_ID").Value = FormTarifs.id_service
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TAR_NAME", SqlDbType.VarChar, 100))
                If TextBox1.Text = "" Then
                    MsgBox("Поле Название не может быть пустым")
                    Exit Sub
                Else
                    MyDataAdapter.Parameters("@TAR_NAME").Value = TextBox1.Text
                End If
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@UNITS_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@UNITS_ID").Value = ComboBox2.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@UNITS_CNTR_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@UNITS_CNTR_ID").Value = ComboBox3.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TARIFF", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@TARIFF").Value = NumericUpDown1.Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TARIFF_CNTR", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@TARIFF_CNTR").Value = NumericUpDown2.Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NORM", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NORM").Value = NumericUpDown3.Value
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
    Private Sub FormTarifsAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        FTarifsAClosed = True
    End Sub
End Class