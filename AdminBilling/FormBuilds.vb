Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception

Public Class FormBuilds
    Dim Street_table As DataTable 'таблица с id и названиями улиц
    Dim Builds_table As DataTable
    Dim X As Integer
    Dim Y As Integer
    Dim V As String
    Public Shared StrID As Integer
    Public Shared StrName As String
    Public Shared id_builds As Integer
    Public Shared BuildName As String

    '-------------------------------------------ЗАГРУЗКА ФОРМЫ--------------------------------
    Private Sub FormBuilds_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Visible = False
        Label3.Visible = False
        Street_table = Streets()
        Label1.Text = "Выберите улицу"
        For i As Integer = 0 To Street_table.Rows.Count - 1 'Заполнение ComboBox
            ComboBox1.Items.Add(Street_table.Rows.Item(i).Item(1))
        Next
        If ComboBox1.SelectedIndex > 0 Then 'если выбрана улица
            Grid() 'вызов функции заполнения грида
        End If
    End Sub

    '-------------------------------------------ЗАПОЛНЕНИЕ ГРИДА ДОМОВ--------------------------
    Private Sub Grid()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_buildings", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@street_id", SqlDbType.Int))
        MyDataAdapter.SelectCommand.Parameters("@street_id").Value = StrID
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
        col1.Name = "ID_build"
        col1.ReadOnly = True
        col1.HeaderText = "ID дома"
        col1.DataPropertyName = "ID_build"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col1.Visible = False
        DataGridView1.Columns.Add(col1)
        '---------------------------------------------------------------------------------------
        Dim BtnColFlat As DataGridViewImageColumn = New DataGridViewImageColumn()
        BtnColFlat.Name = "FlatBtn"
        BtnColFlat.SortMode = DataGridViewColumnSortMode.NotSortable
        BtnColFlat.ToolTipText = "Квартиры"
        Dim inImg As Image = PictureBox1.Image
        BtnColFlat.Image = inImg
        BtnColFlat.HeaderText = "Квартиры"
        BtnColFlat.DefaultCellStyle.ForeColor = Color.Black
        BtnColFlat.DefaultCellStyle.BackColor = Color.LightYellow
        BtnColFlat.DefaultCellStyle.SelectionBackColor = Color.Green
        BtnColFlat.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
        BtnColFlat.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        BtnColFlat.Width = 73
        DataGridView1.Columns.Add(BtnColFlat)
        '---------------------------------------------------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Build"
        col2.ReadOnly = False
        col2.HeaderText = "№ дома"
        col2.DataPropertyName = "Build"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 73
        DataGridView1.Columns.Add(col2)
        '---------------------------------------------------------------------------------------
        Dim col3 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col3.Name = "Korpus"
        col3.ReadOnly = False
        col3.HeaderText = "Корпус"
        col3.DataPropertyName = "Korpus"
        col3.SortMode = DataGridViewColumnSortMode.NotSortable
        col3.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col3.Visible = True
        col3.MinimumWidth = 73
        DataGridView1.Columns.Add(col3)
        '---------------------------------------------------------------------------------------
        Dim colEntrance As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colEntrance.Name = "Entrances"
        colEntrance.ReadOnly = False
        colEntrance.HeaderText = "Подъезды"
        colEntrance.DataPropertyName = "Entrances"
        colEntrance.SortMode = DataGridViewColumnSortMode.NotSortable
        colEntrance.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colEntrance.Visible = True
        colEntrance.MinimumWidth = 73
        DataGridView1.Columns.Add(colEntrance)
        '---------------------------------------------------------------------------------------
        Dim colFloors As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colFloors.Name = "Floors"
        colFloors.ReadOnly = False
        colFloors.HeaderText = "Этажы"
        colFloors.DataPropertyName = "Floors"
        colFloors.SortMode = DataGridViewColumnSortMode.NotSortable
        colFloors.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colFloors.Visible = True
        colFloors.MinimumWidth = 73
        DataGridView1.Columns.Add(colFloors)
        '---------------------------------------------------------------------------------------
        Dim combColIndex As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColIndex.DataSource = DS.Tables(1).DefaultView
        combColIndex.HeaderText = "Индекс"
        combColIndex.Name = "COMBO_Index_id"
        combColIndex.DisplayMember = "Index_number"
        combColIndex.ValueMember = "ID_Index"
        combColIndex.DataPropertyName = "Index_id"
        combColIndex.FlatStyle = FlatStyle.Flat
        combColIndex.MinimumWidth = 73
        combColIndex.Width = 73
        DataGridView1.Columns.Add(combColIndex)
        '---------------------------------------------------------------------------------------
        Dim col4 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col4.Name = "OKATO"
        col4.ReadOnly = False
        col4.HeaderText = "ОКАТО"
        col4.DataPropertyName = "OKATO"
        col4.SortMode = DataGridViewColumnSortMode.NotSortable
        col4.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col4.Visible = True
        col4.MinimumWidth = 79
        DataGridView1.Columns.Add(col4)
        '---------------------------------------------------------------------------------------
        Dim combCol As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combCol.DataSource = DS.Tables(2).DefaultView
        combCol.HeaderText = "Техучасток"
        combCol.Name = "COMBO_ID_Tech"
        combCol.DisplayMember = "Name"
        combCol.ValueMember = "ID_Tech"
        combCol.DataPropertyName = "ID_Tech"
        combCol.FlatStyle = FlatStyle.Flat
        combCol.MinimumWidth = 138
        DataGridView1.Columns.Add(combCol)
        '---------------------------------------------------------------------------------------
        'Dim col6 As JThomas.Controls.DataGridViewMaskedTextColumn = New JThomas.Controls.DataGridViewMaskedTextColumn
        Dim col6 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col6.Name = "FIAS_GUD"
        col6.ReadOnly = False
        col6.HeaderText = "ФИАС"
        'col6.Mask = "AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA"
        col6.DataPropertyName = "FIAS_GUD"
        col6.SortMode = DataGridViewColumnSortMode.NotSortable
        col6.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col6.Visible = True
        'col6.ValueType = GetType(String)
        col6.MinimumWidth = 218
        DataGridView1.Columns.Add(col6)
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
        DataGridView1.Sort(DataGridView1.Columns("Build"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    '-------------------------------------------ЗАГРУЗКА СПИСКА УЛИЦ ИЗ БД-----------------------------------------------------------------------------------------------------------------------------
    Private Function Streets()
        Dim Street_table As DataTable = New DataTable
        Street_table = ClassAuth.ClassAuth.Streets()
        Return Street_table
    End Function

    '-------------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЕК-----------------------------------------------------------------------------------------------------------------------------------
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim errVvod As Integer = 0

        '--------------------------проверка на соответствие поля FIAS_GUD формату GUID--------------------------------------------
        If e.ColumnIndex = DataGridView1.Columns("FIAS_GUD").Index Then
            If ClassVvod.Format_Validation(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("FIAS_GUD").Index).Value.ToString, 1) Then
                errVvod = 0
            Else
                errVvod = 1
                'Grid() 'обновление грида
                MsgBox("Не верный формат GUID")
                'DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
            End If
        End If

        '--------------------------проверка на соответствие поля OKATO целочисленному формату-------------------------------------
        If e.ColumnIndex = DataGridView1.Columns("OKATO").Index Then
            If ClassVvod.Format_Validation(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKATO").Index).Value.ToString, 2) Then
                errVvod = 0
            Else
                errVvod = 1
                DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
                MsgBox("Не верный формат ОКАТО")
                DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
            End If
        End If

        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Bldn_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_bldn", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_bldn").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_build").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@BLDN_NO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@BLDN_NO").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Build").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@KORPUS", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@KORPUS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Korpus").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ENTRANCES", SqlDbType.Int))
                MyDataAdapter.Parameters("@ENTRANCES").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Entrances").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FLOORS", SqlDbType.Int))
                MyDataAdapter.Parameters("@FLOORS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Floors").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@INDEX_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@INDEX_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_Index_id").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@OKATO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@OKATO").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKATO").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@street_id", SqlDbType.Int))
                MyDataAdapter.Parameters("@street_id").Value = Street_table.Rows.Item(ComboBox1.SelectedIndex).Item(0)
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_TECH", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_TECH").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_ID_tech").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FIAS", SqlDbType.UniqueIdentifier))
                MyDataAdapter.Parameters("@FIAS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("FIAS_GUD").Index).Value
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
                If e.ColumnIndex = DataGridView1.Columns("FIAS_GUD").Index Then
                    Grid()
                Else
                    DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
                    SendKeys.Send("{ESC}")
                End If
            End Try
        Else Exit Sub
        End If
    End Sub

    '-------------------------------------------ВЫБОР УЛИЦЫ--------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Button1.Visible = True
        StrID = Street_table.Rows.Item(ComboBox1.SelectedIndex).Item(0)
        StrName = Street_table.Rows.Item(ComboBox1.SelectedIndex).Item(1)
        Label3.Text = "Выбрана улица """ & StrName & """"
        Label3.Visible = True
        Grid()
    End Sub

    '-------------------------------------------ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ЯЧЕЙКИ ДО РЕДАКТИРОВАНИЯ----------
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        X = DataGridView1.CurrentCellAddress.X
        Y = DataGridView1.CurrentCellAddress.Y
        V = DataGridView1(X, Y).Value.ToString
    End Sub

    '-------------------------------------------КЛИК ПО ЯЧЕЙКЕ "УДАЛИТЬ"------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Y = DataGridView1.CurrentCellAddress.Y
            Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить дом № " & DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("Build").Index).Value & " (ID=" & DataGridView1.Rows.Item(Y).Cells.Item(0).Value & ")?", vbYesNo)
            If DRes = DialogResult.Yes Then
                On Error GoTo Err
                'MsgBox("Id_BLDN= " & DataGridView1.Rows.Item(Y).Cells.Item(0).Value)
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Bldn_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 2
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_bldn", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_bldn").Value = DataGridView1.Rows.Item(Y).Cells.Item(0).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@BLDN_NO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@BLDN_NO").Value = DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("Build").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@KORPUS", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@KORPUS").Value = DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("Korpus").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ENTRANCES", SqlDbType.Int))
                MyDataAdapter.Parameters("@ENTRANCES").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Entrances").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FLOORS", SqlDbType.Int))
                MyDataAdapter.Parameters("@FLOORS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Floors").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@INDEX_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@INDEX_ID").Value = DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("COMBO_Index_id").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@OKATO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@OKATO").Value = DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("OKATO").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@street_id", SqlDbType.Int))
                MyDataAdapter.Parameters("@street_id").Value = Street_table.Rows.Item(ComboBox1.SelectedIndex).Item(0)
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_TECH", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_TECH").Value = DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("COMBO_ID_tech").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FIAS", SqlDbType.UniqueIdentifier))
                MyDataAdapter.Parameters("@FIAS").Value = DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("FIAS_GUD").Index).Value
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
            Else
                MsgBox("Выход " & Y)
            End If
        End If
        Exit Sub
Err:
        MsgBox(Err.Description & " " & Err.Number & " " & DataGridView1.CurrentCell.EditedFormattedValue, MsgBoxStyle.Critical)
        Grid()
    End Sub

    '-------------------------------------------КЛИК ПО КНОПКЕ "КВАРТИРЫ"-------------------------------
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        '---------------------------------------ПЕРЕХОД К ФОРМЕ КВАРТИР-------------------------------------------------
        If e.ColumnIndex = DataGridView1.Columns("FlatBtn").Index Then
            id_builds = DataGridView1.CurrentRow.Cells.Item(0).Value
            BuildName = DataGridView1.CurrentRow.Cells.Item(2).Value
            FormFlats.ShowDialog()
        End If
    End Sub

    '-------------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ ДОМ"---------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormBuildsAdd.ShowDialog()
    End Sub

    '-------------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------
    Private Sub FormBuilds_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormBuilds_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormBuildsAdd.FBAClosed IsNot Nothing Then
            If FormBuildsAdd.FBAClosed = True Then
                Grid()
                FormBuildsAdd.FBAClosed = False
            End If
        End If
    End Sub

    '-------------------------------------------ОШИБКА ФОРМАТА ДАННЫХ ГРИДА---------------------------------
    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        If e.Exception IsNot Nothing Then
            e.Cancel = True
            MsgBox(e.Exception.Message, MsgBoxStyle.Critical)
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