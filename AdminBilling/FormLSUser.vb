Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormLSUser

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormLSUser_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid1()

        Grid2()
        'MsgBox(FormLS.Who_ID.ToString())
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА 1------------------------------------------
    Public Sub Grid1()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Users", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Who_ID", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@Who_ID").Value = FormLS.Who_ID.ToString
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
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Surname"
        col2.ReadOnly = False
        col2.HeaderText = "Фамилия"
        col2.DataPropertyName = "Surname"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        col2.Visible = True
        col2.Width = 80
        col2.MinimumWidth = 80
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
        BtnColDoc.Width = 63
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
        colDate_birth.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colDate_birth.Visible = True
        colDate_birth.Width = 73
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
        combColIndex.AutoSizeMode = DataGridViewAutoSizeColumnsMode.None
        combColIndex.MinimumWidth = 85
        combColIndex.Width = 85
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
        combColReg.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        combColReg.FlatStyle = FlatStyle.Flat
        combColReg.MinimumWidth = 90
        combColReg.Width = 90
        DataGridView1.Columns.Add(combColReg)
        '---------------------------------------------------------------------------------------
        Dim colDate_registration As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDate_registration.Name = "Date_registration"
        colDate_registration.ReadOnly = False
        colDate_registration.HeaderText = "Дата регистрации"
        colDate_registration.DataPropertyName = "Date_registration"
        colDate_registration.SortMode = DataGridViewColumnSortMode.NotSortable
        colDate_registration.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colDate_registration.Visible = True
        colDate_registration.Width = 75
        colDate_registration.MinimumWidth = 75
        DataGridView1.Columns.Add(colDate_registration)
        '---------------------------------------------------------------------------------------
        Dim colDate_registration_end As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDate_registration_end.Name = "Date_registration_end"
        colDate_registration_end.ReadOnly = False
        colDate_registration_end.HeaderText = "Окончание регистрации"
        colDate_registration_end.DataPropertyName = "Date_registration_end"
        colDate_registration_end.SortMode = DataGridViewColumnSortMode.NotSortable
        colDate_registration_end.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colDate_registration_end.Visible = True
        colDate_registration_end.Width = 75
        colDate_registration_end.MinimumWidth = 75
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

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА 2------------------------------------------
    Public Sub Grid2()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Propiska", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LS_ID", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@LS_ID").Value = FormLS.LS
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS As DataSet
        DS = New DataSet()
        MyDataAdapter.Fill(DS)
        DataGridView2.DataSource = DS.Tables(0).DefaultView
        DataGridView2.Columns.Clear()
        DataGridView2.AutoGenerateColumns = False
        DataGridView2.ReadOnly = False
        '---------------------------------------------------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Surname"
        col2.ReadOnly = False
        col2.HeaderText = "Фамилия"
        col2.DataPropertyName = "Surname"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        col2.Visible = True
        col2.Width = 80
        col2.MinimumWidth = 80
        DataGridView2.Columns.Add(col2)
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
        DataGridView2.Columns.Add(col3)
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
        DataGridView2.Columns.Add(colEntrance)
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
        BtnColDoc.DefaultCellStyle.Font = New Font(DataGridView2.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnColDoc.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnColDoc.Width = 63
        DataGridView2.Columns.Add(BtnColDoc)
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
        DataGridView2.Columns.Add(colEmail)
        '---------------------------------------------------------------------------------------
        Dim colDate_birth As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDate_birth.Name = "Date_birth"
        colDate_birth.ReadOnly = False
        colDate_birth.HeaderText = "Дата рождения"
        colDate_birth.DataPropertyName = "Date_birth"
        colDate_birth.SortMode = DataGridViewColumnSortMode.NotSortable
        colDate_birth.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colDate_birth.Visible = True
        colDate_birth.Width = 73
        colDate_birth.MinimumWidth = 73
        DataGridView2.Columns.Add(colDate_birth)
        '---------------------------------------------------------------------------------------
        Dim combColIndex As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColIndex.DataSource = DS.Tables(1).DefaultView
        combColIndex.HeaderText = "Пол"
        combColIndex.Name = "Sex"
        combColIndex.DisplayMember = "S"
        combColIndex.ValueMember = "ID_S"
        combColIndex.DataPropertyName = "Sex"
        combColIndex.FlatStyle = FlatStyle.Flat
        combColIndex.AutoSizeMode = DataGridViewAutoSizeColumnsMode.None
        combColIndex.MinimumWidth = 85
        combColIndex.Width = 85
        DataGridView2.Columns.Add(combColIndex)
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
        DataGridView2.Columns.Add(col4)
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
        DataGridView2.Columns.Add(col6)
        '---------------------------------------------------------------------------------------
        Dim combColReg As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColReg.DataSource = DS.Tables(2).DefaultView
        combColReg.HeaderText = "Регистрация"
        combColReg.Name = "ID_registration_type"
        combColReg.DisplayMember = "Type"
        combColReg.ValueMember = "ID_reg_type"
        combColReg.DataPropertyName = "ID_registration_type"
        combColReg.FlatStyle = FlatStyle.Flat
        combColReg.AutoSizeMode = DataGridViewAutoSizeColumnsMode.None
        combColReg.MinimumWidth = 90
        combColReg.Width = 90
        DataGridView2.Columns.Add(combColReg)
        '---------------------------------------------------------------------------------------
        Dim colDate_registration As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDate_registration.Name = "Date_registration"
        colDate_registration.ReadOnly = False
        colDate_registration.HeaderText = "Дата регистрации"
        colDate_registration.DataPropertyName = "Date_registration"
        colDate_registration.SortMode = DataGridViewColumnSortMode.NotSortable
        colDate_registration.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colDate_registration.Visible = True
        colDate_registration.Width = 75
        colDate_registration.MinimumWidth = 75
        DataGridView2.Columns.Add(colDate_registration)
        '---------------------------------------------------------------------------------------
        Dim colDate_registration_end As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDate_registration_end.Name = "Date_registration_end"
        colDate_registration_end.ReadOnly = False
        colDate_registration_end.HeaderText = "Окончание регистрации"
        colDate_registration_end.DataPropertyName = "Date_registration_end"
        colDate_registration_end.SortMode = DataGridViewColumnSortMode.NotSortable
        colDate_registration_end.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colDate_registration_end.Visible = True
        colDate_registration_end.Width = 75
        colDate_registration_end.MinimumWidth = 75
        DataGridView2.Columns.Add(colDate_registration_end)
        '---------------------------------------------------------------------------------------
        Dim colRelated As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        colRelated.DataSource = DS.Tables(3).DefaultView
        colRelated.Name = "ID_relation_type"
        colRelated.ReadOnly = False
        colRelated.HeaderText = "Родственные отношения"
        colRelated.DisplayMember = "Type"
        colRelated.ValueMember = "ID_relation_type"
        colRelated.DataPropertyName = "ID_relation_type"
        colRelated.SortMode = DataGridViewColumnSortMode.NotSortable
        colRelated.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colRelated.Visible = True
        colRelated.Width = 80
        colRelated.MinimumWidth = 80
        DataGridView2.Columns.Add(colRelated)
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
        BtnCol.DefaultCellStyle.Font = New Font(DataGridView2.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnCol.Width = 30
        DataGridView2.Columns.Add(BtnCol)
        '-------------------------сортировка по фамилии по возрастанию--------------------
        DataGridView2.Sort(DataGridView2.Columns("Surname"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub


    Private Sub RadioButton4_Click(sender As Object, e As EventArgs) Handles RadioButton4.Click
        If (RadioButton4.Checked = True) And (RadioButton3.Checked = False) Then
            RadioButton4.Checked = False
            RadioButton3.Checked = True
            DataGridView2.Enabled = False
            DataGridView1.Enabled = True
        Else
            RadioButton4.Checked = True
            RadioButton3.Checked = False
            DataGridView2.Enabled = True
            DataGridView1.Enabled = False
        End If
    End Sub

    Private Sub RadioButton3_Click(sender As Object, e As EventArgs) Handles RadioButton3.Click
        If (RadioButton3.Checked = True) And (RadioButton4.Checked = False) Then
            RadioButton3.Checked = False
            RadioButton4.Checked = True
            DataGridView1.Enabled = False
            DataGridView2.Enabled = True
        Else
            RadioButton3.Checked = True
            RadioButton4.Checked = False
            DataGridView1.Enabled = True
            DataGridView2.Enabled = False
        End If
    End Sub


End Class