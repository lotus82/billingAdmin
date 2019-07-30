Imports ClassAuth
Imports System.Data.SqlClient


Public Class FormUsersAdd
    Public Shared FUAClosed = False

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormUsersAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Users", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS As DataSet
        DS = New DataSet()
        MyDataAdapter.Fill(DS)
        ComboBox1.DataSource = DS.Tables(1)
        ComboBox1.DisplayMember = "S"
        ComboBox1.ValueMember = "ID_S"
        ComboBox2.DataSource = DS.Tables(2)
        ComboBox2.DisplayMember = "Type"
        ComboBox2.ValueMember = "ID_reg_type"
    End Sub

    '-------------------------------------ДОБАВЛЕНИЕ ЧЕЛОВЕКА---------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim errVvod As Integer = 0

        ''--------------------------проверка на соответствие поля FIAS_GUD формату GUID--------------------------------------------
        'If e.ColumnIndex = DataGridView1.Columns("FIAS_GUD").Index Then
        '    If ClassVvod.Format_Validation(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("FIAS_GUD").Index).Value.ToString, 1) Then
        '        errVvod = 0
        '    Else
        '        errVvod = 1
        '        'Grid() 'обновление грида
        '        MsgBox("Не верный формат GUID")
        '        'DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
        '    End If
        'End If

        ''--------------------------проверка на соответствие поля OKATO целочисленному формату-------------------------------------
        'If e.ColumnIndex = DataGridView1.Columns("OKATO").Index Then
        '    If ClassVvod.Format_Validation(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKATO").Index).Value.ToString, 2) Then
        '        errVvod = 0
        '    Else
        '        errVvod = 1
        '        DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
        '        MsgBox("Не верный формат ОКАТО")
        '        DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
        '    End If
        'End If

        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Users_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 3
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@GUID_user", SqlDbType.UniqueIdentifier))
                MyDataAdapter.Parameters("@GUID_user").Value = Guid.NewGuid()
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Surname", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Surname").Value = TextBox1.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Name", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Name").Value = TextBox2.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Patronymic", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Patronymic").Value = TextBox3.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Email", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Email").Value = TextBox4.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Date_birth", SqlDbType.Date))
                MyDataAdapter.Parameters("@Date_birth").Value = Date.Now()
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Sex", SqlDbType.Int))
                MyDataAdapter.Parameters("@Sex").Value = ComboBox1.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Phone", SqlDbType.VarChar, 12))
                MyDataAdapter.Parameters("@Phone").Value = TextBox6.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_doc", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_doc").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_registration_type", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_registration_type").Value = ComboBox2.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Date_registration", SqlDbType.Date))
                MyDataAdapter.Parameters("@Date_registration").Value = Date.Now()
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
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                Me.Close()
                Exit Sub
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
        Else Exit Sub
        End If
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormUsersAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        FUAClosed = True
    End Sub

End Class