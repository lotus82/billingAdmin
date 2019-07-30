Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormTarifs

    Dim X As Integer
    Dim Y As Integer
    Dim V As String

    Public Shared id_service As String

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormTarifs_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox2.Visible = False
        Label3.Visible = False
        Button1.Enabled = False
        Grid(1) 'Вызов функции заполнения ComboBox1
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ COMBOBOX1, COMBOBOX2 И ГРИДА ТАРИФОВ---------------------------------------
    Public Sub Grid(type As Integer)
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Tarifs", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------

        If type = 1 Then
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@type").Value = 1
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SUPPL_ID", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@SUPPL_ID").Value = vbNull
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ID_Service", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@ID_Service").Value = vbNull
            '---------------------------------------------------------------------------------------
            Dim DS As DataSet
            DS = New DataSet()
            MyDataAdapter.Fill(DS)
            ComboBox1.DataSource = DS.Tables(0).DefaultView
            ComboBox1.DisplayMember = "SUPPLIER_NAME"
            ComboBox1.ValueMember = "SUPPLIER_ID"
            ComboBox1.SelectedIndex = 0
        End If

        If type = 2 Then
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@type").Value = 2
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SUPPL_ID", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@SUPPL_ID").Value = ComboBox1.SelectedValue
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ID_Service", SqlDbType.VarChar, 4))
            MyDataAdapter.SelectCommand.Parameters("@ID_Service").Value = vbNull
            '---------------------------------------------------------------------------------------
            Dim DS As DataSet
            DS = New DataSet()
            MyDataAdapter.Fill(DS)
            If DS.Tables(0).Rows.Count > 0 Then
                ComboBox2.DataSource = DS.Tables(0)
                ComboBox2.DisplayMember = "Service_Name"
                ComboBox2.ValueMember = "ID_Service"
                Button1.Enabled = True
            Else
                ComboBox2.DataSource = Nothing
                ComboBox2.Items.Clear()
                Button1.Enabled = False
            End If
        End If

        If type = 3 Then
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@type").Value = 3
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@SUPPL_ID", SqlDbType.Int))
            MyDataAdapter.SelectCommand.Parameters("@SUPPL_ID").Value = ComboBox1.SelectedValue
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@ID_Service", SqlDbType.VarChar, 4))
            MyDataAdapter.SelectCommand.Parameters("@ID_Service").Value = ComboBox2.SelectedValue
            '---------------------------------------------------------------------------------------
            Dim DS As DataSet
            DS = New DataSet()
            MyDataAdapter.Fill(DS)
            DataGridView1.DataSource = DS.Tables(0).DefaultView
            DataGridView1.Columns.Clear()
            DataGridView1.AutoGenerateColumns = False
            DataGridView1.ReadOnly = False
            '------------------------id тарифа---------------------------------------------------
            Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
            col1.Name = "TAR_ID"
            col1.ReadOnly = True
            col1.HeaderText = "ID"
            col1.DataPropertyName = "TAR_ID"
            col1.SortMode = DataGridViewColumnSortMode.NotSortable
            col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            col1.Visible = False
            DataGridView1.Columns.Add(col1)
            '------------------------Название тарифа--------------------------------------------------------
            Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
            col2.Name = "TAR_NAME"
            col2.ReadOnly = False
            col2.HeaderText = "Название тарифа"
            col2.DataPropertyName = "TAR_NAME"
            col2.SortMode = DataGridViewColumnSortMode.Automatic
            col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            col2.Visible = True
            col2.Width = 160
            col2.MinimumWidth = 160
            DataGridView1.Columns.Add(col2)
            '------------------------единица измерения без прибора------------------------------------------
            Dim combColUnit As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
            combColUnit.DataSource = DS.Tables(1).DefaultView()
            combColUnit.HeaderText = "Ед. без прибора"
            combColUnit.Name = "UNITS_ID"
            combColUnit.DisplayMember = "UNITS_TYPE_ED"
            combColUnit.ValueMember = "ID"
            combColUnit.DataPropertyName = "UNITS_ID"
            combColUnit.Width = 110
            combColUnit.MinimumWidth = 110
            DataGridView1.Columns.Add(combColUnit)
            '------------------------единица измерения с прибором-------------------------------------------
            Dim combColUniCntr As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
            combColUniCntr.DataSource = DS.Tables(2).DefaultView()
            combColUniCntr.HeaderText = "Ед. с прибором"
            combColUniCntr.Name = "UNITS_CNTR_ID"
            combColUniCntr.DisplayMember = "UNITS_TYPE_ED"
            combColUniCntr.ValueMember = "ID"
            combColUniCntr.DataPropertyName = "UNITS_CNTR_ID"
            combColUniCntr.Width = 110
            combColUniCntr.MinimumWidth = 110
            DataGridView1.Columns.Add(combColUniCntr)
            '------------------------Значение тарифа без прибора-------------------------------------------
            Dim col5 As DataGridViewColumn = New DataGridViewTextBoxColumn
            col5.Name = "TARIFF"
            col5.ReadOnly = False
            col5.HeaderText = "Тариф без прибора"
            col5.DataPropertyName = "TARIFF"
            col5.SortMode = DataGridViewColumnSortMode.NotSortable
            col5.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            col5.Visible = True
            col5.Width = 110
            col5.MinimumWidth = 110
            DataGridView1.Columns.Add(col5)
            '------------------------Значение тарифа с прибором------------------------------
            Dim colTarifCntr As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colTarifCntr.Name = "TARIFF_CNTR"
            colTarifCntr.ReadOnly = False
            colTarifCntr.HeaderText = "Тариф с прибором"
            colTarifCntr.DataPropertyName = "TARIFF_CNTR"
            colTarifCntr.SortMode = DataGridViewColumnSortMode.NotSortable
            colTarifCntr.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colTarifCntr.Visible = True
            colTarifCntr.Width = 110
            colTarifCntr.MinimumWidth = 110
            DataGridView1.Columns.Add(colTarifCntr)
            '------------------------Значение Норма------------------------------
            Dim colNorm As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colNorm.Name = "NORM"
            colNorm.ReadOnly = False
            colNorm.HeaderText = "Норма"
            colNorm.DataPropertyName = "NORM"
            colNorm.SortMode = DataGridViewColumnSortMode.NotSortable
            colNorm.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colNorm.Visible = True
            colNorm.Width = 110
            colNorm.MinimumWidth = 110
            DataGridView1.Columns.Add(colNorm)
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
            '-------------------------сортировка по названию тарифа по возрастанию--------------------
            DataGridView1.Sort(DataGridView1.Columns("TAR_NAME"), System.ComponentModel.ListSortDirection.Ascending)
        End If
    End Sub

    '-------------------------------------ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ЯЧЕЙКИ ДО РЕДАКТИРОВАНИЯ-----------------
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        X = DataGridView1.CurrentCellAddress.X
        Y = DataGridView1.CurrentCellAddress.Y
        V = DataGridView1(X, Y).Value.ToString
    End Sub

    '-------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЕК----------------------------------------
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
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
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TAR_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@TAR_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TAR_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SERVICE_ID", SqlDbType.Char, 4))
                MyDataAdapter.Parameters("@SERVICE_ID").Value = id_service
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TAR_NAME", SqlDbType.VarChar, 100))
                MyDataAdapter.Parameters("@TAR_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TAR_NAME").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@UNITS_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@UNITS_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("UNITS_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@UNITS_CNTR_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@UNITS_CNTR_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("UNITS_CNTR_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TARIFF", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@TARIFF").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TARIFF").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TARIFF_CNTR", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@TARIFF_CNTR").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TARIFF_CNTR").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NORM", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NORM").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("NORM").Index).Value
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
                SendKeys.Send("{TAB}")
                Exit Sub

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
                DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
                SendKeys.Send("{ESC}")
            End Try
        Else : Exit Sub
        End If
    End Sub

    '-------------------------------------УДАЛЕНИЕ ТАРИФА----------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить данный тариф?", vbYesNo)
            If DRes = DialogResult.Yes Then
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Tarifs_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 2
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TAR_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@TAR_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TAR_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SERVICE_ID", SqlDbType.Char, 4))
                MyDataAdapter.Parameters("@SERVICE_ID").Value = ComboBox2.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TAR_NAME", SqlDbType.VarChar, 100))
                MyDataAdapter.Parameters("@TAR_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TAR_NAME").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@UNITS_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@UNITS_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("UNITS_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@UNITS_CNTR_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@UNITS_CNTR_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("UNITS_CNTR_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TARIFF", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@TARIFF").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TARIFF").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TARIFF_CNTR", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@TARIFF_CNTR").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TARIFF_CNTR").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NORM", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NORM").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("NORM").Index).Value
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
                    'Grid(3)
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Critical)
                    ' Grid(3)
                End Try
            End If
        End If
    End Sub

    '-------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ ТАРИФ"--------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormTarifsAdd.ShowDialog()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormTarifs_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormTarifs_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormTarifsAdd.FTarifsAClosed IsNot Nothing Then
            If FormTarifsAdd.FTarifsAClosed = True Then
                Grid(3)
                FormTarifsAdd.FTarifsAClosed = False
            End If
        End If
    End Sub

    '-------------------------------------ОШИБКА ФОРМАТА ДАННЫХ ГРИДА---------------------------------
    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError
        If e.Exception IsNot Nothing Then
            e.Cancel = True
            MsgBox(e.Exception.Message, MsgBoxStyle.Critical)
            SendKeys.Send("{ESC}")
        End If
    End Sub

    '-------------------------------------ВЫБОР ПОСТАВЩИКА ИЗ СПИСКА-----------------------------------
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            DataGridView1.DataSource = Nothing
            DataGridView1.Columns.Clear()
            Grid(2)
            ComboBox2.Visible = True
            Label3.Visible = True
            Label2.Text = ComboBox2.Items.Count
            Grid(3)
            If ComboBox2.Items.Count = 0 Then
                ComboBox2.Text = ""
                Button1.Enabled = False
            Else
                id_service = ComboBox2.SelectedValue
                ComboBox2.Visible = True
                DataGridView1.Visible = True
                Button1.Enabled = True
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            id_service = ComboBox2.SelectedValue
            Grid(3)
            Button1.Enabled = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox1_ValueMemberChanged(sender As Object, e As EventArgs) Handles ComboBox1.ValueMemberChanged
        Try
            
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox2_ValueMemberChanged(sender As Object, e As EventArgs) Handles ComboBox2.ValueMemberChanged
        Try
            
        Catch ex As Exception

        End Try
    End Sub

    '------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЙКИ С COMBOBOX ПРИ ВЫБОРЕ ИЗ ВЫПАДАЮЩЕГО СПИСКА-------------
    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        Dim idx1 As Integer = DataGridView1.Columns.Item("UNITS_ID").Index
        Dim idx2 As Integer = DataGridView1.Columns.Item("UNITS_CNTR_ID").Index
        If DataGridView1.IsCurrentCellDirty And (DataGridView1.CurrentCell.ColumnIndex = idx1 Xor DataGridView1.CurrentCell.ColumnIndex = idx2) Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

   
   
   

End Class