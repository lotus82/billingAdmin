Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormCounters
    Dim X As Integer
    Dim Y As Integer
    Dim V As String

    '-------------------------------------ЗАГРУЗКА ФОРМЫ------------------------------------------------------
    Private Sub FormCounters_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid() 'Вызов функции заполнения грида
        ' MsgBox(Math.)
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА МОДЕЛЕЙ ПРИБОРОВ УЧЕТА-----------------------------
    Public Sub Grid()
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
        DataGridView1.DataSource = DS.Tables(0).DefaultView
        DataGridView1.Columns.Clear()
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.ReadOnly = False
        '------------------------id прибора--------------------------------------------------
        Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col1.Name = "MODEL_ID"
        col1.ReadOnly = True
        col1.HeaderText = "ID"
        col1.DataPropertyName = "MODEL_ID"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col1.Visible = False
        DataGridView1.Columns.Add(col1)
        '------------------------Название прибора---------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "MODEL_NAME"
        col2.ReadOnly = False
        col2.HeaderText = "Название прибора"
        col2.DataPropertyName = "MODEL_NAME"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 300
        DataGridView1.Columns.Add(col2)
        '-------------------------Тип прибора-------------------------------------------------------
        Dim combColType As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColType.DataSource = DS.Tables(1).DefaultView
        combColType.HeaderText = "Тип прибора"
        combColType.Name = "CNTR_RES_ID"
        combColType.DisplayMember = "RES_NAME"
        combColType.ValueMember = "CNTR_RES_ID"
        combColType.DataPropertyName = "CNTR_RES_ID"
        combColType.FlatStyle = FlatStyle.Flat
        combColType.MinimumWidth = 100
        combColType.Width = 100
        DataGridView1.Columns.Add(combColType)
        '-------------------------Услуга-------------------------------------------------------
        Dim combColService As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColService.DataSource = DS.Tables(2).DefaultView
        combColService.HeaderText = "Услуга"
        combColService.Name = "Service_ID"
        combColService.DisplayMember = "ID_Service"
        combColService.ValueMember = "ID_Service"
        combColService.DataPropertyName = "Service_ID"
        combColService.FlatStyle = FlatStyle.Flat
        combColService.MinimumWidth = 100
        combColService.Width = 100
        DataGridView1.Columns.Add(combColService)
        '------------------------период поверки------------------------------------------------------
        Dim col3 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col3.Name = "PERIOD_VERIFICATION"
        col3.ReadOnly = False
        col3.HeaderText = "Период поверки (мес)"
        col3.DataPropertyName = "PERIOD_VERIFICATION"
        col3.SortMode = DataGridViewColumnSortMode.NotSortable
        col3.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col3.Visible = True
        col3.MinimumWidth = 80
        DataGridView1.Columns.Add(col3)
        '------------------------разрядность------------------------------------------------------
        Dim col4 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col4.Name = "DIGIT_CAPACITY"
        col4.ReadOnly = False
        col4.HeaderText = "Разрядность"
        col4.DataPropertyName = "DIGIT_CAPACITY"
        col4.SortMode = DataGridViewColumnSortMode.NotSortable
        col4.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col4.Visible = True
        col4.MinimumWidth = 80
        DataGridView1.Columns.Add(col4)
        '------------------------точность------------------------------------------------------
        Dim col_precision As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col_precision.Name = "PRECISION"
        col_precision.ReadOnly = False
        col_precision.HeaderText = "Точность"
        col_precision.DataPropertyName = "PRECISION"
        col_precision.SortMode = DataGridViewColumnSortMode.NotSortable
        col_precision.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col_precision.Visible = True
        col_precision.MinimumWidth = 80
        DataGridView1.Columns.Add(col_precision)
        '------------------------Коэффициент-----------------------------------------
        Dim col5 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col5.Name = "COEFF"
        col5.ReadOnly = False
        col5.HeaderText = "Коэффициент"
        col5.DataPropertyName = "COEFF"
        col5.SortMode = DataGridViewColumnSortMode.NotSortable
        col5.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col5.Visible = True
        col5.MinimumWidth = 80
        DataGridView1.Columns.Add(col5)
        '------------------------Активность-----------------------------------------
        Dim col_Active As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col_Active.Name = "ACTIVE"
        col_Active.ReadOnly = False
        col_Active.HeaderText = "Активность"
        col_Active.DataPropertyName = "ACTIVE"
        col_Active.SortMode = DataGridViewColumnSortMode.NotSortable
        col_Active.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col_Active.Visible = True
        col_Active.MinimumWidth = 80
        DataGridView1.Columns.Add(col_Active)
        '------------------------------кнопка удаления--------------------------------------
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
        '-------------------------сортировка по названию улицы по возрастанию--------------------
        DataGridView1.Sort(DataGridView1.Columns("ACTIVE"), System.ComponentModel.ListSortDirection.Descending)
        For i As Integer = 0 To DataGridView1.RowCount - 1
            If DataGridView1.Rows.Item(i).Cells.Item(8).Value = 0 Then
                Dim Cell As DataGridViewCell
                For Each Cell In DataGridView1.Rows(i).Cells
                    Cell.Style.BackColor = Color.LightGray
                Next
            End If
        Next

    End Sub

    '-------------------------------------ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ЯЧЕЙКИ ДО РЕДАКТИРОВАНИЯ-------------------------
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        X = DataGridView1.CurrentCellAddress.X
        Y = DataGridView1.CurrentCellAddress.Y
        V = DataGridView1(X, Y).Value.ToString
    End Sub

    '-------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЕК------------------------------------------------
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
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
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@MODEL_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@MODEL_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("MODEL_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@CNTR_RES_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@CNTR_RES_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("CNTR_RES_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SERVICE_ID", SqlDbType.Char, 4))
                MyDataAdapter.Parameters("@SERVICE_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("SERVICE_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@MODEL_NAME", SqlDbType.VarChar, 100))
                MyDataAdapter.Parameters("@MODEL_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("MODEL_NAME").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@DIGIT_CAPACITY", SqlDbType.Int))
                MyDataAdapter.Parameters("@DIGIT_CAPACITY").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("DIGIT_CAPACITY").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@PRECISION", SqlDbType.Int))
                MyDataAdapter.Parameters("@PRECISION").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("PRECISION").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@COEFF", SqlDbType.Decimal, 18, 2))
                MyDataAdapter.Parameters("@COEFF").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COEFF").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@PERIOD_VERIFICATION", SqlDbType.Int))
                MyDataAdapter.Parameters("@PERIOD_VERIFICATION").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("PERIOD_VERIFICATION").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ACTIVE", SqlDbType.Int))
                MyDataAdapter.Parameters("@ACTIVE").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ACTIVE").Index).Value
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

    '-------------------------------------УДАЛЕНИЕ МОДЕЛИ ПРИБОРА---------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim errVvod As Integer = 0
        If errVvod = 0 Then
            If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
                Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить данную модель прибора?", vbYesNo)
                If DRes = DialogResult.Yes Then
                    Try
                        Dim MyConnection As SqlConnection
                        Dim MyDataAdapter As SqlCommand
                        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                        MyDataAdapter = New SqlCommand("Adm_Counters_Edit", MyConnection)
                        MyDataAdapter.CommandType = CommandType.StoredProcedure
                        '---------------------------------------------------------------------------------------
                        MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                        MyDataAdapter.Parameters("@type").Value = 2
                        '---------------------------------------------------------------------------------------
                        MyDataAdapter.Parameters.Add(New SqlParameter("@MODEL_ID", SqlDbType.Int))
                        MyDataAdapter.Parameters("@MODEL_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("MODEL_ID").Index).Value
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
                        Exit Sub
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.Critical)
                        Grid()
                    End Try
                End If
            End If
        Else Exit Sub
        End If
    End Sub

    '-------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ МОДЕЛЬ ПРИБОРА"----------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormCountersAdd.ShowDialog()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ------------------------------------------------------
    Private Sub FormCounters_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ-----------------------------------------------------
    Private Sub FormCounters_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormCountersAdd.FCountersAddClosed IsNot Nothing Then
            If FormCountersAdd.FCountersAddClosed = True Then
                Grid()
                FormCountersAdd.FCountersAddClosed = False
            End If
        End If
    End Sub

    '-------------------------------------ОШИБКА ФОРМАТА ДАННЫХ ГРИДА-----------------------------------------
    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        If e.Exception IsNot Nothing Then
            e.Cancel = True
            MsgBox(e.Exception.Message, MsgBoxStyle.Critical)
            SendKeys.Send("{ESC}")
        End If
    End Sub

End Class