Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormServices

    Dim X As Integer
    Dim Y As Integer
    Dim V As String
    Dim errorCount As Integer = 0

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormServices_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid() 'Вызов функции заполнения грида
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА УСЛУГ---------------------------------------
    Public Sub Grid()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Services", MyConnection)
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
        '------------------------id услуги---------------------------------------------------
        Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col1.Name = "ID_Service"
        col1.ReadOnly = True
        col1.HeaderText = "ID"
        col1.DataPropertyName = "ID_Service"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col1.Visible = False
        DataGridView1.Columns.Add(col1)
        '------------------------Название услуги---------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Service_Name"
        col2.ReadOnly = False
        col2.HeaderText = "Название ууслуги"
        col2.DataPropertyName = "Service_Name"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 273
        DataGridView1.Columns.Add(col2)
        '------------------------сокращенное название------------------------------------------------------
        Dim col3 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col3.Name = "Short_Name"
        col3.ReadOnly = False
        col3.HeaderText = "сокращенное название"
        col3.DataPropertyName = "Short_Name"
        col3.SortMode = DataGridViewColumnSortMode.NotSortable
        col3.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col3.Visible = True
        col3.MinimumWidth = 150
        DataGridView1.Columns.Add(col3)
        '-------------------------ОДН-----------------------------------------------------
        Dim combColType As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColType.DataSource = DS.Tables(1).DefaultView()
        combColType.HeaderText = "ОДН"
        combColType.Name = "COMBO_IS_ODN"
        combColType.DisplayMember = "RL"
        combColType.ValueMember = "ID_RL"
        combColType.DataPropertyName = "IS_ODN"
        combColType.Width = 50
        combColType.MinimumWidth = 50
        DataGridView1.Columns.Add(combColType)
        '-------------------------зависимая-----------------------------------------------------
        Dim colDepend As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colDepend.Name = "IS_DEPENDS"
        colDepend.ReadOnly = False
        colDepend.HeaderText = "Зависимая"
        colDepend.DataPropertyName = "IS_DEPENDS"
        colDepend.SortMode = DataGridViewColumnSortMode.NotSortable
        colDepend.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colDepend.Visible = False
        colDepend.MinimumWidth = 50
        DataGridView1.Columns.Add(colDepend)
        '------------------------норм. коэфф.-----------------------------------------
        Dim col5 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col5.Name = "NORM_COEFF"
        col5.ReadOnly = False
        col5.HeaderText = "Норм. коэфф."
        col5.DataPropertyName = "NORM_COEFF"
        col5.SortMode = DataGridViewColumnSortMode.NotSortable
        col5.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col5.Visible = True
        col5.Width = 50
        col5.MinimumWidth = 50
        DataGridView1.Columns.Add(col5)
        '------------------------НДС-----------------------------------------
        Dim colNDS As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colNDS.Name = "NDS"
        colNDS.ReadOnly = False
        colNDS.HeaderText = "НДС"
        colNDS.DataPropertyName = "NDS"
        colNDS.SortMode = DataGridViewColumnSortMode.NotSortable
        colNDS.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colNDS.Visible = True
        colNDS.MinimumWidth = 40
        DataGridView1.Columns.Add(colNDS)
        '------------------------Код УСПН-----------------------------------------
        Dim colUSPN As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colUSPN.Name = "USPN_CODE"
        colUSPN.ReadOnly = False
        colUSPN.HeaderText = "УСПН"
        colUSPN.DataPropertyName = "USPN_CODE"
        colUSPN.SortMode = DataGridViewColumnSortMode.NotSortable
        colUSPN.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colUSPN.Visible = True
        colUSPN.MinimumWidth = 40
        DataGridView1.Columns.Add(colUSPN)
        '-------------------------Исполнитель услуги-----------------------------------------------------
        Dim combColSuppliers As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColSuppliers.DataSource = DS.Tables(2).DefaultView()
        combColSuppliers.HeaderText = "Исполнитель услуги"
        combColSuppliers.Name = "COMBO_SUPPL_ID"
        combColSuppliers.DisplayMember = "SUPPLIER_NAME"
        combColSuppliers.ValueMember = "SUPPLIER_ID"
        combColSuppliers.DataPropertyName = "SUPPL_ID"
        combColSuppliers.MinimumWidth = 140
        DataGridView1.Columns.Add(combColSuppliers)
        '-------------------------Получатель платежа-----------------------------------------------------
        Dim combColSuppliersPayee As DataGridViewComboBoxColumn = New DataGridViewComboBoxColumn
        combColSuppliersPayee.DataSource = DS.Tables(3).DefaultView()
        combColSuppliersPayee.HeaderText = "Получатель платежа"
        combColSuppliersPayee.Name = "COMBO_SUPPL_PAYEE_ID"
        combColSuppliersPayee.DisplayMember = "SUPPLIER_PAYEE_NAME"
        combColSuppliersPayee.ValueMember = "SUPPLIER_PAYEE_ID"
        combColSuppliersPayee.DataPropertyName = "SUPPL_PAYEE_ID"
        combColSuppliersPayee.MinimumWidth = 140
        DataGridView1.Columns.Add(combColSuppliersPayee)
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
        '-------------------------сортировка по названию услуги по возрастанию--------------------
        DataGridView1.Sort(DataGridView1.Columns("Service_Name"), System.ComponentModel.ListSortDirection.Ascending)
        Dim Cell As DataGridViewCell
        Dim Row As DataGridViewRow
        For Each Row In DataGridView1.Rows
            If Row.Cells.Item("IS_DEPENDS").Value > 0 Then
                For Each Cell In Row.Cells
                    Cell.Style.BackColor = Color.Yellow
                Next
                Row.Cells.Item("DelRow").Style.BackColor = Color.Red
            End If
        Next
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
                MyDataAdapter = New SqlCommand("Adm_Services_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_Service", SqlDbType.Char, 4))
                MyDataAdapter.Parameters("@ID_Service").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_Service").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Service_Name", SqlDbType.VarChar, 64))
                MyDataAdapter.Parameters("@Service_Name").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Service_Name").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Short_Name", SqlDbType.VarChar, 32))
                MyDataAdapter.Parameters("@Short_Name").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Short_Name").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@IS_ODN", SqlDbType.Int))
                MyDataAdapter.Parameters("@IS_ODN").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_IS_ODN").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@IS_DEPENDS", SqlDbType.Int))
                MyDataAdapter.Parameters("@IS_DEPENDS").Value = vbNull
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NORM_COEFF", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NORM_COEFF").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("NORM_COEFF").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NDS", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NDS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("NDS").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@USPN_CODE", SqlDbType.Int))
                MyDataAdapter.Parameters("@USPN_CODE").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("USPN_CODE").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPL_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@SUPPL_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_SUPPL_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPL_PAYEE_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@SUPPL_PAYEE_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_SUPPL_PAYEE_ID").Index).Value
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

    '-------------------------------------УДАЛЕНИЕ УСЛУГИ----------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить данную услугу?", vbYesNo)
            If DRes = DialogResult.Yes Then
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Services_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 2
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_Service", SqlDbType.Char, 4))
                MyDataAdapter.Parameters("@ID_Service").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_Service").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Service_Name", SqlDbType.VarChar, 64))
                MyDataAdapter.Parameters("@Service_Name").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Service_Name").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Short_Name", SqlDbType.VarChar, 32))
                MyDataAdapter.Parameters("@Short_Name").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Short_Name").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@IS_ODN", SqlDbType.Int))
                MyDataAdapter.Parameters("@IS_ODN").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_IS_ODN").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@IS_DEPENDS", SqlDbType.Int))
                MyDataAdapter.Parameters("@IS_DEPENDS").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NORM_COEFF", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NORM_COEFF").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("NORM_COEFF").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NDS", SqlDbType.Decimal))
                MyDataAdapter.Parameters("@NDS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("NDS").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@USPN_CODE", SqlDbType.Int))
                MyDataAdapter.Parameters("@USPN_CODE").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("USPN_CODE").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPL_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@SUPPL_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_SUPPL_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPL_PAYEE_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@SUPPL_PAYEE_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("COMBO_SUPPL_PAYEE_ID").Index).Value
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

    '-------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ УСЛУГУ"--------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormServicesAdd.ShowDialog()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormStreets_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormStreets_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormServicesAdd.FServicesAClosed IsNot Nothing Then
            If FormServicesAdd.FServicesAClosed = True Then
                Grid()
                FormServicesAdd.FServicesAClosed = False
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

    '------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЙКИ С COMBOBOX ПРИ ВЫБОРЕ ИЗ ВЫПАДАЮЩЕГО СПИСКА-------------
    Private Sub DataGridView1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellDirtyStateChanged
        If DataGridView1.IsCurrentCellDirty And (DataGridView1.Columns.Item(DataGridView1.CurrentCell.ColumnIndex).Name Like "COMBO_*") Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        Dim Cell As DataGridViewCell
        Dim Row As DataGridViewRow
        For Each Row In DataGridView1.Rows
            If Row.Cells.Item("IS_DEPENDS").Value > 0 Then
                For Each Cell In Row.Cells
                    Cell.Style.BackColor = Color.Yellow
                Next
                Row.Cells.Item("DelRow").Style.BackColor = Color.Red
            End If
        Next
    End Sub
End Class