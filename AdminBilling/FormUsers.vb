Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception


Public Class FormUsers
    Dim X As Integer
    Dim Y As Integer
    Dim V As String

    '-------------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------
    Private Sub FormUsers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid() 'вызов функции заполнения грида
    End Sub

    '-------------------------------------------ЗАПОЛНЕНИЕ ГРИДА ЛЮДЕЙ--------------------------
    Private Sub Grid()
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
        DataGridView1.DataSource = DS.Tables(0).DefaultView
        DataGridView1.Columns.Clear()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.ReadOnly = False
        '---------------------------------------------------------------------------------------
        Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col1.Name = "GUID_user"
        col1.ReadOnly = True
        col1.HeaderText = "GUID"
        col1.DataPropertyName = "GUID_user"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.Visible = False
        col1.Width = 60
        DataGridView1.Columns.Add(col1)
        '---------------------------------------------------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Surname"
        col2.ReadOnly = False
        col2.HeaderText = "Фамилия"
        col2.DataPropertyName = "Surname"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 73
        DataGridView1.Columns.Add(col2)
        '---------------------------------------------------------------------------------------
        Dim col3 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col3.Name = "Name"
        col3.ReadOnly = False
        col3.HeaderText = "Имя"
        col3.DataPropertyName = "Name"
        col3.SortMode = DataGridViewColumnSortMode.NotSortable
        col3.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col3.Visible = True
        col3.MinimumWidth = 73
        DataGridView1.Columns.Add(col3)
        '---------------------------------------------------------------------------------------
        Dim colEntrance As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colEntrance.Name = "Patronymic"
        colEntrance.ReadOnly = False
        colEntrance.HeaderText = "Отчество"
        colEntrance.DataPropertyName = "Patronymic"
        colEntrance.SortMode = DataGridViewColumnSortMode.NotSortable
        colEntrance.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colEntrance.Visible = True
        colEntrance.MinimumWidth = 73
        DataGridView1.Columns.Add(colEntrance)
        '---------------------------------------------------------------------------------------
        Dim BtnColDoc As DataGridViewImageColumn = New DataGridViewImageColumn()
        BtnColDoc.Name = "DocBtn"
        BtnColDoc.SortMode = DataGridViewColumnSortMode.NotSortable
        BtnColDoc.ToolTipText = "Документ"
        Dim inImg As Image = PictureBox1.Image
        BtnColDoc.Image = inImg
        BtnColDoc.HeaderText = "Документ"
        BtnColDoc.DefaultCellStyle.ForeColor = Color.Black
        BtnColDoc.DefaultCellStyle.BackColor = Color.LightYellow
        BtnColDoc.DefaultCellStyle.SelectionBackColor = Color.Green
        BtnColDoc.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnColDoc.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnColDoc.Width = 65
        DataGridView1.Columns.Add(BtnColDoc)
        '---------------------------------------------------------------------------------------
        Dim colEmail As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colEmail.Name = "Email"
        colEmail.ReadOnly = False
        colEmail.HeaderText = "Email"
        colEmail.DataPropertyName = "Email"
        colEmail.SortMode = DataGridViewColumnSortMode.NotSortable
        colEmail.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colEmail.Visible = True
        colEmail.MinimumWidth = 73
        DataGridView1.Columns.Add(colEmail)
        '---------------------------------------------------------------------------------------
        Dim colDate_birth As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDate_birth.Name = "Date_birth"
        colDate_birth.ReadOnly = False
        colDate_birth.HeaderText = "Дата рождения"
        colDate_birth.DataPropertyName = "Date_birth"
        colDate_birth.SortMode = DataGridViewColumnSortMode.NotSortable
        colDate_birth.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colDate_birth.Visible = True
        colDate_birth.MinimumWidth = 73
        DataGridView1.Columns.Add(colDate_birth)
        '---------------------------------------------------------------------------------------
        Dim combColIndex As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColIndex.DataSource = DS.Tables(1).DefaultView
        combColIndex.HeaderText = "Пол"
        combColIndex.Name = "Sex"
        combColIndex.DisplayMember = "S"
        combColIndex.ValueMember = "ID_S"
        combColIndex.DataPropertyName = "Sex"
        combColIndex.FlatStyle = FlatStyle.Flat
        combColIndex.MinimumWidth = 80
        combColIndex.Width = 80
        DataGridView1.Columns.Add(combColIndex)
        '---------------------------------------------------------------------------------------
        Dim col4 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col4.Name = "Phone"
        col4.ReadOnly = False
        col4.HeaderText = "Телефон"
        col4.DataPropertyName = "Phone"
        col4.SortMode = DataGridViewColumnSortMode.NotSortable
        col4.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col4.Visible = True
        col4.MinimumWidth = 79
        DataGridView1.Columns.Add(col4)
        '---------------------------------------------------------------------------------------
        Dim col6 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col6.Name = "ID_doc"
        col6.ReadOnly = False
        col6.HeaderText = "Документ"
        col6.DataPropertyName = "ID_doc"
        col6.SortMode = DataGridViewColumnSortMode.NotSortable
        col6.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col6.Visible = False
        col6.MinimumWidth = 30
        DataGridView1.Columns.Add(col6)
        '---------------------------------------------------------------------------------------
        Dim combColReg As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColReg.DataSource = DS.Tables(2).DefaultView
        combColReg.HeaderText = "Регистрация"
        combColReg.Name = "ID_registration_type"
        combColReg.DisplayMember = "Type"
        combColReg.ValueMember = "ID_reg_type"
        combColReg.DataPropertyName = "ID_registration_type"
        combColReg.FlatStyle = FlatStyle.Flat
        combColReg.MinimumWidth = 100
        combColReg.Width = 100
        DataGridView1.Columns.Add(combColReg)
        '---------------------------------------------------------------------------------------
        Dim colDate_registration As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDate_registration.Name = "Date_registration"
        colDate_registration.ReadOnly = False
        colDate_registration.HeaderText = "Дата регистрации"
        colDate_registration.DataPropertyName = "Date_registration"
        colDate_registration.SortMode = DataGridViewColumnSortMode.NotSortable
        colDate_registration.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colDate_registration.Visible = True
        colDate_registration.MinimumWidth = 79
        DataGridView1.Columns.Add(colDate_registration)
        '---------------------------------------------------------------------------------------
        Dim colDate_registration_end As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDate_registration_end.Name = "Date_registration_end"
        colDate_registration_end.ReadOnly = False
        colDate_registration_end.HeaderText = "Окончание регистрации"
        colDate_registration_end.DataPropertyName = "Date_registration_end"
        colDate_registration_end.SortMode = DataGridViewColumnSortMode.NotSortable
        colDate_registration_end.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colDate_registration_end.Visible = True
        colDate_registration_end.MinimumWidth = 79
        DataGridView1.Columns.Add(colDate_registration_end)
        '---------------------------------------------------------------------------------------
        Dim BtnCol As DataGridViewButtonColumn = New DataGridViewButtonColumn
        BtnCol.Name = "DelRow"
        BtnCol.ReadOnly = True
        BtnCol.SortMode = DataGridViewColumnSortMode.NotSortable
        BtnCol.ToolTipText = "Удалить запись"
        BtnCol.FlatStyle = FlatStyle.Popup
        BtnCol.UseColumnTextForButtonValue = True
        BtnCol.Text = "Х"
        BtnCol.HeaderText = ""
        BtnCol.DefaultCellStyle.ForeColor = Color.Black
        BtnCol.DefaultCellStyle.BackColor = Color.Red
        BtnCol.DefaultCellStyle.SelectionBackColor = Color.Green
        BtnCol.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnCol.Width = 30
        DataGridView1.Columns.Add(BtnCol)
        '-------------------------сортировка по фамилии по возрастанию--------------------
        DataGridView1.Sort(DataGridView1.Columns("Surname"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    '-------------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЕК-----------------------------------------------------------------------------------------------------------------------------------
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
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
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@GUID_user", SqlDbType.UniqueIdentifier))
                MyDataAdapter.Parameters("@GUID_user").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("GUID_user").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Surname", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Surname").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Surname").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Name", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Name").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Name").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Patronymic", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Patronymic").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Patronymic").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Email", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Email").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Email").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Date_birth", SqlDbType.Date))
                MyDataAdapter.Parameters("@Date_birth").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Date_birth").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Sex", SqlDbType.Int))
                MyDataAdapter.Parameters("@Sex").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Sex").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Phone", SqlDbType.VarChar, 12))
                MyDataAdapter.Parameters("@Phone").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Phone").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_doc", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_doc").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_doc").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_registration_type", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_registration_type").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_registration_type").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Date_registration", SqlDbType.Date))
                MyDataAdapter.Parameters("@Date_registration").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Date_registration").Index).Value
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
                Exit Sub
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)

                DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
                    SendKeys.Send("{ESC}")

            End Try
        Else Exit Sub
        End If
    End Sub

    '-------------------------------------------КЛИК ПО ЯЧЕЙКЕ "УДАЛИТЬ"------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Y = DataGridView1.CurrentCellAddress.Y
            Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить человека " & DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("Surname").Index).Value & "?", vbYesNo)
            If DRes = DialogResult.Yes Then
                Try
                    Dim MyConnection As SqlConnection
                    Dim MyDataAdapter As SqlCommand
                    MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                    MyDataAdapter = New SqlCommand("Adm_Users_Edit", MyConnection)
                    MyDataAdapter.CommandType = CommandType.StoredProcedure
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                    MyDataAdapter.Parameters("@type").Value = 2
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@GUID_user", SqlDbType.UniqueIdentifier))
                    MyDataAdapter.Parameters("@GUID_user").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("GUID_user").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@Surname", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@Surname").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Surname").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@Name", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@Name").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Name").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@Patronymic", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@Patronymic").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Patronymic").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@Email", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@Email").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Email").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@Date_birth", SqlDbType.Date))
                    MyDataAdapter.Parameters("@Date_birth").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Date_birth").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@Sex", SqlDbType.Int))
                    MyDataAdapter.Parameters("@Sex").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Sex").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@Phone", SqlDbType.VarChar, 12))
                    MyDataAdapter.Parameters("@Phone").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Phone").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@ID_doc", SqlDbType.Int))
                    MyDataAdapter.Parameters("@ID_doc").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_doc").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@ID_registration_type", SqlDbType.Int))
                    MyDataAdapter.Parameters("@ID_registration_type").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_registration_type").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@Date_registration", SqlDbType.Date))
                    MyDataAdapter.Parameters("@Date_registration").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Date_registration").Index).Value
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
                    Grid()
                Catch ex As Exception
                    MsgBox(Err.Description & " " & Err.Number & " " & DataGridView1.CurrentCell.EditedFormattedValue, MsgBoxStyle.Critical)
                    Grid()
                End Try
            Else
                MsgBox("Выход " & Y)
            End If
        End If
        Exit Sub

    End Sub

    '-------------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ ЧЕЛОВЕКА"--------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormUsersAdd.ShowDialog()
    End Sub

    '-------------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------
    Private Sub FormUsers_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormUsers_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormUsersAdd.FUAClosed IsNot Nothing Then
            If FormUsersAdd.FUAClosed = True Then
                Grid()
                FormUsersAdd.FUAClosed = False
            End If
        End If
    End Sub


End Class