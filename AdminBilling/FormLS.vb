Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception


Public Class FormLS
    Dim X As Integer
    Dim Y As Integer
    Dim V As String
    Public Shared LS As Int64
    Public Shared Who_ID As Guid

    '-------------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------
    Private Sub FormLS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Select()
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
        ComboBox1.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        Grid() 'вызов функции заполнения грида
    End Sub

    '-------------------------------------------ЗАПОЛНЕНИЕ ГРИДА ЛИЦЕВЫХ СЧЕТОВ-----------------
    Private Sub Grid()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_LS", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LS", SqlDbType.BigInt))
        If TextBox1.Text.Length < 1 Then
            MyDataAdapter.SelectCommand.Parameters("@LS").Value = vbNull
        Else
            MyDataAdapter.SelectCommand.Parameters("@LS").Value = Convert.ToInt64(Replace(TextBox1.Text.PadRight(14), " ", "0"))
        End If
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@LS_max", SqlDbType.BigInt))
        If TextBox1.Text.Length < 1 Then
            MyDataAdapter.SelectCommand.Parameters("@LS_max").Value = vbNull
        Else
            MyDataAdapter.SelectCommand.Parameters("@LS_max").Value = Convert.ToInt64(Replace(TextBox1.Text.PadRight(14), " ", "9"))
        End If
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Street_LS", SqlDbType.Int))
        If ComboBox1.SelectedIndex > -1 Then
            MyDataAdapter.SelectCommand.Parameters("@Street_LS").Value = Convert.ToInt32(ComboBox1.SelectedValue)
        Else
            MyDataAdapter.SelectCommand.Parameters("@Street_LS").Value = vbNull
        End If
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Build_LS", SqlDbType.Int))
        If ComboBox2.SelectedIndex > -1 Then
            MyDataAdapter.SelectCommand.Parameters("@Build_LS").Value = Convert.ToInt32(ComboBox2.SelectedValue)
        Else
            MyDataAdapter.SelectCommand.Parameters("@Build_LS").Value = vbNull
        End If
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@Flat_LS", SqlDbType.Int))
        If ComboBox3.SelectedIndex > -1 Then
            MyDataAdapter.SelectCommand.Parameters("@Flat_LS").Value = Convert.ToInt32(ComboBox3.SelectedValue)
        Else
            MyDataAdapter.SelectCommand.Parameters("@Flat_LS").Value = vbNull
        End If
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS As DataSet
        DS = New DataSet()
        MyDataAdapter.Fill(DS)
        If RadioButton1.Checked And TextBox1.Text.Length < 1 Then
            DataGridView1.DataSource = DS.Tables(0).DefaultView
        End If
        If RadioButton1.Checked And TextBox1.Text.Length > 0 Then
            DataGridView1.DataSource = DS.Tables(6).DefaultView
        End If
        If RadioButton2.Checked And ComboBox1.SelectedIndex > -1 Then
            DataGridView1.DataSource = DS.Tables(3).DefaultView
        End If
        If RadioButton2.Checked And ComboBox1.SelectedIndex > -1 And ComboBox2.SelectedIndex > -1 Then
            DataGridView1.DataSource = DS.Tables(4).DefaultView
        End If
        If RadioButton2.Checked And ComboBox1.SelectedIndex > -1 And ComboBox2.SelectedIndex > -1 And ComboBox3.SelectedIndex > -1 Then
            DataGridView1.DataSource = DS.Tables(5).DefaultView
        End If
        DataGridView1.Columns.Clear()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.ReadOnly = False
        '---------------------------------------------------------------------------------------
        Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col1.Name = "ID_LS"
        col1.ReadOnly = True
        col1.HeaderText = "Лицевой счет"
        col1.DataPropertyName = "ID_LS"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        col1.Visible = True
        col1.Width = 130
        DataGridView1.Columns.Add(col1)
        '---------------------------------------------------------------------------------------
        Dim BtnColFlat As DataGridViewImageColumn = New DataGridViewImageColumn()
        BtnColFlat.Name = "UserBtn"
        BtnColFlat.SortMode = DataGridViewColumnSortMode.NotSortable
        BtnColFlat.ToolTipText = ""
        Dim inImg As Image = PictureBox1.Image
        BtnColFlat.Image = inImg
        BtnColFlat.HeaderText = "User"
        BtnColFlat.DefaultCellStyle.ForeColor = Color.Black
        BtnColFlat.DefaultCellStyle.BackColor = Color.LightYellow
        BtnColFlat.DefaultCellStyle.SelectionBackColor = Color.Green
        BtnColFlat.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnColFlat.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnColFlat.Width = 50
        DataGridView1.Columns.Add(BtnColFlat)
        '---------------------------------------------------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Who_ID"
        col2.ReadOnly = True
        col2.HeaderText = "Потребитель"
        col2.DataPropertyName = "Who_ID"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        col2.Visible = True
        col2.Width = 310
        col2.MinimumWidth = 310
        DataGridView1.Columns.Add(col2)
        '---------------------------------------------------------------------------------------
        Dim col3 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col3.Name = "Login"
        col3.ReadOnly = False
        col3.HeaderText = "Логин"
        col3.DataPropertyName = "Login"
        col3.SortMode = DataGridViewColumnSortMode.NotSortable
        col3.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        col3.Visible = True
        col3.Width = 100
        col3.MinimumWidth = 100
        DataGridView1.Columns.Add(col3)
        '---------------------------------------------------------------------------------------
        Dim colEntrance As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colEntrance.Name = "Password"
        colEntrance.ReadOnly = False
        colEntrance.HeaderText = "Пароль"
        colEntrance.DataPropertyName = "Password"
        colEntrance.SortMode = DataGridViewColumnSortMode.NotSortable
        colEntrance.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        colEntrance.Visible = True
        colEntrance.Width = 100
        colEntrance.MinimumWidth = 100
        DataGridView1.Columns.Add(colEntrance)
        '---------------------------------------------------------------------------------------
        Dim combColLSStatus As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColLSStatus.DataSource = DS.Tables(1).DefaultView
        combColLSStatus.HeaderText = "Статус"
        combColLSStatus.Name = "ID_LS_status"
        combColLSStatus.DisplayMember = "Status"
        combColLSStatus.ValueMember = "ID_LS_status"
        combColLSStatus.DataPropertyName = "ID_LS_status"
        combColLSStatus.FlatStyle = FlatStyle.Flat
        combColLSStatus.MinimumWidth = 150
        combColLSStatus.Width = 150
        DataGridView1.Columns.Add(combColLSStatus)
        '---------------------------------------------------------------------------------------
        Dim BtnColAdres As DataGridViewImageColumn = New DataGridViewImageColumn()
        BtnColAdres.Name = "AdresBtn"
        BtnColAdres.SortMode = DataGridViewColumnSortMode.NotSortable
        BtnColAdres.ToolTipText = ""
        Dim inImg2 As Image = PictureBox2.Image
        BtnColAdres.Image = inImg2
        BtnColAdres.HeaderText = "Адрес"
        BtnColAdres.DefaultCellStyle.ForeColor = Color.Black
        BtnColAdres.DefaultCellStyle.BackColor = Color.LightYellow
        BtnColAdres.DefaultCellStyle.SelectionBackColor = Color.Green
        BtnColAdres.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnColAdres.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnColAdres.Width = 50
        DataGridView1.Columns.Add(BtnColAdres)
        '---------------------------------------------------------------------------------------
        Dim BtnColCounter As DataGridViewImageColumn = New DataGridViewImageColumn()
        BtnColCounter.Name = "CounterBtn"
        BtnColCounter.SortMode = DataGridViewColumnSortMode.NotSortable
        BtnColCounter.ToolTipText = ""
        Dim inImg3 As Image = PictureBox3.Image
        BtnColCounter.Image = inImg3
        BtnColCounter.HeaderText = "Прибор"
        BtnColCounter.DefaultCellStyle.ForeColor = Color.Black
        BtnColCounter.DefaultCellStyle.BackColor = Color.LightYellow
        BtnColCounter.DefaultCellStyle.SelectionBackColor = Color.Green
        BtnColCounter.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnColCounter.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnColCounter.Width = 50
        DataGridView1.Columns.Add(BtnColCounter)
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
        '-------------------------сортировка по номеру дома по возрастанию--------------------
        DataGridView1.Sort(DataGridView1.Columns("ID_LS"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    '-------------------------------------ВЫБОР "ПОИСК ПО НОМЕРУ Л/С"---------------------------------
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            TextBox1.Enabled = True
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            ComboBox1.SelectedIndex = -1
            ComboBox2.SelectedIndex = -1
            ComboBox3.SelectedIndex = -1
        Else
            TextBox1.Enabled = False
        End If
        Grid()
    End Sub

    '-------------------------------------ВЫБОР "ПОИСК Л/С ПО АДРЕСУ"--------------------------------
    Private Sub RadioButton2_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            ComboBox1.Enabled = True
            ComboBox1.SelectedIndex = -1
            ComboBox2.Enabled = False
            ComboBox3.Enabled = False
            TextBox1.Text = ""
            TextBox1.Enabled = False
        Else
            ComboBox1.SelectedIndex = -1
            ComboBox1.Enabled = False
        End If
    End Sub

    '-------------------------------------ВВОД НОМЕРА Л/С--------------------------------------------
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim Start As Integer
        Start = TextBox1.SelectionStart
        If TextBox1.Text.Length > 0 Then
            If (Char.IsDigit(TextBox1.Text(TextBox1.Text.Length - 1))) Then

            Else
                TextBox1.Text = Replace(TextBox1.Text, TextBox1.Text(TextBox1.Text.Length - 1), "")
                TextBox1.SelectionStart = Start
            End If
        End If
        Grid()
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
        Grid()
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
        Grid()
    End Sub

    '-------------------------------------ВЫБОР КВАРТИРЫ-------------------------------------------------
    Private Sub ComboBox3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox3.SelectionChangeCommitted
        Grid()
    End Sub

    '-------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ Л/С"-------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormLSAdd.ShowDialog()
    End Sub

    '-------------------------------------КЛИК ПО КНОПКАМ "АДРЕС", "ПРИБОРЫ", "ЛЮДИ"------------------
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("AdresBtn").Index Then
            LS = DataGridView1.CurrentRow.Cells.Item(0).Value
            FormLSAdres.ShowDialog()
        End If
        If e.ColumnIndex = DataGridView1.Columns("CounterBtn").Index Then
            LS = DataGridView1.CurrentRow.Cells.Item(0).Value
            FormLSCounter.ShowDialog()
        End If
        If e.ColumnIndex = DataGridView1.Columns("UserBtn").Index Then
            LS = DataGridView1.CurrentRow.Cells.Item(0).Value
            Who_ID = DataGridView1.CurrentRow.Cells.Item(2).Value
            FormLSUser.ShowDialog()
        End If
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormLS_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub


End Class