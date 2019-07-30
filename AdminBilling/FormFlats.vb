Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception


Public Class FormFlats
    Dim X As Integer
    Dim Y As Integer
    Dim V As String
    Public Shared StrID As Integer
    Public Shared StrName As String
    Public Shared id_builds As Integer
    Public Shared BuildName As String
    Public Shared id_flat As Integer



    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormFlats_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        id_builds = FormBuilds.id_builds
        StrName = FormBuilds.StrName
        BuildName = FormBuilds.BuildName
        Label2.Text = "Улица """ & StrName & """, дом № " & BuildName
        Grid() 'вызов функции заполнения грида
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА КВАРТИР------------------------------------
    Private Sub Grid()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_flats", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@id_builds", SqlDbType.Int))
        MyDataAdapter.SelectCommand.Parameters("@id_builds").Value = FormBuilds.id_builds
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
        col1.Name = "ID_flats"
        col1.ReadOnly = True
        col1.HeaderText = "ID"
        col1.DataPropertyName = "ID_flats"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col1.Visible = False
        DataGridView1.Columns.Add(col1)
        ''---------------------------------------------------------------------------------------
        Dim BtnColFlat As DataGridViewImageColumn = New DataGridViewImageColumn()
        BtnColFlat.Name = "FlatBtn"
        BtnColFlat.SortMode = DataGridViewColumnSortMode.NotSortable
        BtnColFlat.ToolTipText = "Комнаты"
        Dim inImg As Image = PictureBox1.Image
        BtnColFlat.Image = inImg
        BtnColFlat.HeaderText = "Комнаты"
        BtnColFlat.DefaultCellStyle.ForeColor = Color.Black
        BtnColFlat.DefaultCellStyle.BackColor = Color.LightYellow
        BtnColFlat.DefaultCellStyle.SelectionBackColor = Color.Green
        BtnColFlat.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnColFlat.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnColFlat.Width = 60
        DataGridView1.Columns.Add(BtnColFlat)
        ''---------------------------------------------------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Flat"
        col2.ReadOnly = False
        col2.HeaderText = "№ квартиры"
        col2.DataPropertyName = "Flat"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 100
        DataGridView1.Columns.Add(col2)
        '---------------------------------------------------------------------------------------
        Dim col3 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col3.Name = "Entrance"
        col3.ReadOnly = False
        col3.HeaderText = "Подъезд"
        col3.DataPropertyName = "Entrance"
        col3.SortMode = DataGridViewColumnSortMode.Automatic
        col3.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col3.Visible = True
        col3.MinimumWidth = 100
        DataGridView1.Columns.Add(col3)
        '---------------------------------------------------------------------------------------
        Dim combColFloor As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColFloor.DataSource = DS.Tables(1).DefaultView()
        combColFloor.HeaderText = "Этаж"
        combColFloor.Name = "COMBO_FloorFlat"
        combColFloor.DisplayMember = "number"
        combColFloor.ValueMember = "number"
        combColFloor.DataPropertyName = "FloorFlat"
        combColFloor.Visible = True
        combColFloor.MinimumWidth = 110
        DataGridView1.Columns.Add(combColFloor)
        '---------------------------------------------------------------------------------------
        Dim BtnCol As DataGridViewButtonColumn = New DataGridViewButtonColumn
        BtnCol.Name = "DelRow"
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
        '-------------------------сортировка по номеру квартиры по возрастанию--------------------
        DataGridView1.Sort(DataGridView1.Columns("Flat"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    '-----------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЕК------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlCommand
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlCommand("Adm_Flats_Edit", MyConnection)
            MyDataAdapter.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.Parameters("@type").Value = 1
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@id_street", SqlDbType.Int))
            MyDataAdapter.Parameters("@id_street").Value = Convert.ToInt32(FormBuilds.StrID)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@id_flats", SqlDbType.Int))
            MyDataAdapter.Parameters("@id_flats").Value = Convert.ToInt32(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_flats").Index).Value)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@FLAT", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@FLAT").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Flat").Index).Value
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ENTRANCE", SqlDbType.Int))
            MyDataAdapter.Parameters("@ENTRANCE").Value = Convert.ToInt32(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Entrance").Index).Value)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@FLOORFLAT", SqlDbType.Int))
            MyDataAdapter.Parameters("@FLOORFLAT").Value = Convert.ToInt32(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_FloorFlat").Index).Value)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ID_BUILD", SqlDbType.Int))
            MyDataAdapter.Parameters("@ID_BUILD").Value = id_builds
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
            'DataGridView1.CancelEdit()
            SendKeys.Send("{ESC}")
            'DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
            'DataGridView1.CancelEdit()
        End Try

    End Sub

    '--------------------------------------------ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ЯЧЕЙКИ ДО РЕДАКТИРОВАНИЯ----------
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        X = DataGridView1.CurrentCellAddress.X
        Y = DataGridView1.CurrentCellAddress.Y
        V = DataGridView1(X, Y).Value.ToString
    End Sub

    '-------------------------------------------КЛИК ПО ЯЧЕЙКЕ "УДАЛИТЬ"------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить данную квартиру?", vbYesNo)
            If DRes = DialogResult.Yes Then
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_flats_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 2
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_street", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_street").Value = FormBuilds.StrID
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_flats", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_flats").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_flats").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FLAT", SqlDbType.Int))
                MyDataAdapter.Parameters("@FLAT").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Flat").Index).Value
                '                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ENTRANCE", SqlDbType.Int))
                MyDataAdapter.Parameters("@ENTRANCE").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Entrance").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FLOORFLAT", SqlDbType.Int))
                MyDataAdapter.Parameters("@FLOORFLAT").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_FloorFlat").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_BUILD", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_BUILD").Value = id_builds
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
                Try
                    MyConnection.Open()
                    MyDataAdapter.ExecuteNonQuery()
                    MyConnection.Close()
                    Grid()
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical)
                    Grid()
                End Try
            End If
        End If
    End Sub

    '-------------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ КВАРТИРУ"---------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormFlatsAdd.ShowDialog()
    End Sub

    '-------------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------
    Private Sub FormFlats_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        'DataGridView1.Dispose()
        'FormBuilds.Grid()
        'FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormFlats_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormFlatsAdd.FFAClosed IsNot Nothing Then
            If FormFlatsAdd.FFAClosed = True Then
                Grid()
                FormFlatsAdd.FFAClosed = False
            End If
        End If
    End Sub

    '-------------------------------------ОШИБКА ФОРМАТА ДАННЫХ ГРИДА---------------------------------
    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        If e.Exception IsNot Nothing Then
            e.Cancel = True
            MsgBox("********" & e.Exception.Message, MsgBoxStyle.Critical)
            SendKeys.Send("{ESC}")
        End If
    End Sub

    '------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЙКИ С COMBOBOX ПРИ ВЫБОРЕ ИЗ ВЫПАДАЮЩЕГО СПИСКА-------------
    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty And (DataGridView1.Columns.Item(DataGridView1.CurrentCell.ColumnIndex).Name Like "COMBO_*") Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

End Class