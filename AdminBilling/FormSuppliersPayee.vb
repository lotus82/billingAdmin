Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception

Public Class FormSuppliersPayee
    Dim X As Integer
    Dim Y As Integer
    Dim V As String

    '-------------------------------------------ЗАГРУЗКА ФОРМЫ--------------------------------------
    Private Sub FormSuppliers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid()
    End Sub

    '-------------------------------------------ЗАПОЛНЕНИЕ ГРИДА ИСПОЛНИТЕЛЕЙ------------------------
    Private Sub Grid()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_Suppliers_Payee", MyConnection)
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
        col1.Name = "SUPPLIER_PAYEE_ID"
        col1.ReadOnly = True
        col1.HeaderText = "ID"
        col1.DataPropertyName = "SUPPLIER_PAYEE_ID"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col1.Visible = False
        DataGridView1.Columns.Add(col1)
        '---------------------------------------------------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "SUPPLIER_PAYEE_NAME"
        col2.ReadOnly = False
        col2.HeaderText = "Название"
        col2.DataPropertyName = "SUPPLIER_PAYEE_NAME"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 70
        DataGridView1.Columns.Add(col2)
        '---------------------------------------------------------------------------------------
        Dim colINN As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colINN.Name = "INN"
        colINN.ReadOnly = False
        colINN.HeaderText = "ИНН"
        colINN.DataPropertyName = "INN"
        colINN.SortMode = DataGridViewColumnSortMode.NotSortable
        colINN.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colINN.Visible = True
        colINN.MinimumWidth = 70
        DataGridView1.Columns.Add(colINN)
        '---------------------------------------------------------------------------------------
        Dim colKPP As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colKPP.Name = "KPP"
        colKPP.ReadOnly = False
        colKPP.HeaderText = "КПП"
        colKPP.DataPropertyName = "KPP"
        colKPP.SortMode = DataGridViewColumnSortMode.NotSortable
        colKPP.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colKPP.Visible = True
        colKPP.MinimumWidth = 70
        DataGridView1.Columns.Add(colKPP)
        '---------------------------------------------------------------------------------------
        Dim colBIK As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colBIK.Name = "BIK"
        colBIK.ReadOnly = False
        colBIK.HeaderText = "БИК"
        colBIK.DataPropertyName = "BIK"
        colBIK.SortMode = DataGridViewColumnSortMode.NotSortable
        colBIK.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colBIK.Visible = True
        colBIK.MinimumWidth = 50
        DataGridView1.Columns.Add(colBIK)
        '---------------------------------------------------------------------------------------
        Dim colRS As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colRS.Name = "RS"
        colRS.ReadOnly = False
        colRS.HeaderText = "Рсчетный счет"
        colRS.DataPropertyName = "RS"
        colRS.SortMode = DataGridViewColumnSortMode.NotSortable
        colRS.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colRS.Visible = True
        colRS.MinimumWidth = 70
        DataGridView1.Columns.Add(colRS)
        '---------------------------------------------------------------------------------------
        Dim colKS As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colKS.Name = "KS"
        colKS.ReadOnly = False
        colKS.HeaderText = "Кор. счет"
        colKS.DataPropertyName = "KS"
        colKS.SortMode = DataGridViewColumnSortMode.NotSortable
        colKS.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colKS.Visible = True
        colKS.MinimumWidth = 70
        DataGridView1.Columns.Add(colKS)
        '---------------------------------------------------------------------------------------
        Dim colBank As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colBank.Name = "BANK_NAME"
        colBank.ReadOnly = False
        colBank.HeaderText = "Название банка"
        colBank.DataPropertyName = "BANK_NAME"
        colBank.SortMode = DataGridViewColumnSortMode.NotSortable
        colBank.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colBank.Visible = True
        colBank.Width = 70
        colBank.MinimumWidth = 70
        DataGridView1.Columns.Add(colBank)
        '---------------------------------------------------------------------------------------
        Dim colAdres As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colAdres.Name = "ADDRESS"
        colAdres.ReadOnly = False
        colAdres.HeaderText = "Адрес исполнителя"
        colAdres.DataPropertyName = "ADRES"
        colAdres.SortMode = DataGridViewColumnSortMode.NotSortable
        colAdres.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colAdres.Visible = True
        colAdres.Width = 70
        colAdres.MinimumWidth = 70
        DataGridView1.Columns.Add(colAdres)
        '---------------------------------------------------------------------------------------
        Dim colPhone As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colPhone.Name = "PHONE"
        colPhone.ReadOnly = False
        colPhone.HeaderText = "Телефон"
        colPhone.DataPropertyName = "PHONE"
        colPhone.SortMode = DataGridViewColumnSortMode.NotSortable
        colPhone.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colPhone.Visible = True
        colPhone.MinimumWidth = 50
        DataGridView1.Columns.Add(colPhone)
        '---------------------------------------------------------------------------------------
        Dim colFax As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colFax.Name = "FAX"
        colFax.ReadOnly = False
        colFax.HeaderText = "Факс"
        colFax.DataPropertyName = "FAX"
        colFax.SortMode = DataGridViewColumnSortMode.NotSortable
        colFax.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colFax.Visible = True
        colFax.MinimumWidth = 50
        DataGridView1.Columns.Add(colFax)
        '---------------------------------------------------------------------------------------
        Dim colEmail As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colEmail.Name = "MAIL"
        colEmail.ReadOnly = False
        colEmail.HeaderText = "Email"
        colEmail.DataPropertyName = "MAIL"
        colEmail.SortMode = DataGridViewColumnSortMode.NotSortable
        colEmail.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colEmail.Visible = True
        colEmail.MinimumWidth = 50
        DataGridView1.Columns.Add(colEmail)
        '---------------------------------------------------------------------------------------
        Dim colWeb As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colWeb.Name = "WEB"
        colWeb.ReadOnly = False
        colWeb.HeaderText = "Сайт"
        colWeb.DataPropertyName = "WEB"
        colWeb.SortMode = DataGridViewColumnSortMode.NotSortable
        colWeb.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colWeb.Visible = True
        colWeb.MinimumWidth = 50
        DataGridView1.Columns.Add(colWeb)
        '---------------------------------------------------------------------------------------
        Dim colWorkDays As DataGridViewColumn = New DataGridViewTextBoxColumn()
        colWorkDays.Name = "WORK_DAYS"
        colWorkDays.ReadOnly = False
        colWorkDays.HeaderText = "Режим работы"
        colWorkDays.DataPropertyName = "WORK_DAYS"
        colWorkDays.SortMode = DataGridViewColumnSortMode.NotSortable
        colWorkDays.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        colWorkDays.Visible = True
        colWorkDays.Width = 70
        colWorkDays.MinimumWidth = 70
        DataGridView1.Columns.Add(colWorkDays)
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
        '-------------------------сортировка по названию по возрастанию--------------------
        DataGridView1.Sort(DataGridView1.Columns("SUPPLIER_PAYEE_NAME"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    '-------------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЕК-----------------------------------------------------------------------------------------------------------------------------------
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim errVvod As Integer = 0

        '--------------------------проверка на соответствие поля FIAS_GUD формату GUID--------------------------------------------
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

        '--------------------------проверка на соответствие поля OKATO целочисленному формату-------------------------------------
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
                MyDataAdapter = New SqlCommand("Adm_Suppliers_Payee_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPLIER_PAYEE_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@SUPPLIER_PAYEE_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("SUPPLIER_PAYEE_ID").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPLIER_PAYEE_NAME", SqlDbType.VarChar, 100))
                MyDataAdapter.Parameters("@SUPPLIER_PAYEE_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("SUPPLIER_PAYEE_NAME").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@INN", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@INN").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("INN").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@KPP", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@KPP").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("KPP").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@BIK", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@BIK").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("BIK").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@RS", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@RS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("RS").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@KS", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@KS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("KS").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@BANK_NAME", SqlDbType.VarChar, 100))
                MyDataAdapter.Parameters("@BANK_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("BANK_NAME").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ADDRESS", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@ADDRESS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ADDRESS").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@PHONE", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@PHONE").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("PHONE").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FAX", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@FAX").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("FAX").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@MAIL", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@MAIL").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("MAIL").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@WEB", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@WEB").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("WEB").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@WORK_DAYS", SqlDbType.VarChar, 100))
                MyDataAdapter.Parameters("@WORK_DAYS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("WORK_DAYS").Index).Value
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

    '-------------------------------------------ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ЯЧЕЙКИ ДО РЕДАКТИРОВАНИЯ---------
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        X = DataGridView1.CurrentCellAddress.X
        Y = DataGridView1.CurrentCellAddress.Y
        V = DataGridView1(X, Y).Value.ToString
    End Sub

    '-------------------------------------------КЛИК ПО ЯЧЕЙКЕ "УДАЛИТЬ"------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Try
                Y = DataGridView1.CurrentCellAddress.Y
                Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить исполнителя """ & DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns("SUPPLIER_PAYEE_NAME").Index).Value & """?", vbYesNo)
                If DRes = DialogResult.Yes Then
                    Dim MyConnection As SqlConnection
                    Dim MyDataAdapter As SqlCommand
                    MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                    MyDataAdapter = New SqlCommand("Adm_Suppliers_Edit", MyConnection)
                    MyDataAdapter.CommandType = CommandType.StoredProcedure
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                    MyDataAdapter.Parameters("@type").Value = 2
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPLIER_PAYEE_ID", SqlDbType.Int))
                    MyDataAdapter.Parameters("@SUPPLIER_PAYEE_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("SUPPLIER_PAYEE_ID").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@SUPPLIER_PAYEE_NAME", SqlDbType.VarChar, 100))
                    MyDataAdapter.Parameters("@SUPPLIER_PAYEE_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("SUPPLIER_PAYEE_NAME").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@INN", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@INN").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("INN").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@KPP", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@KPP").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("KPP").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@BIK", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@BIK").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("BIK").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@RS", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@RS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("RS").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@KS", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@KS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("KS").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@BANK_NAME", SqlDbType.VarChar, 100))
                    MyDataAdapter.Parameters("@BANK_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("BANK_NAME").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@ADDRESS", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@ADDRESS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ADDRESS").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@PHONE", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@PHONE").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("PHONE").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@FAX", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@FAX").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("FAX").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@MAIL", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@MAIL").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("MAIL").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@WEB", SqlDbType.VarChar, 50))
                    MyDataAdapter.Parameters("@WEB").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("WEB").Index).Value
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@WORK_DAYS", SqlDbType.VarChar, 100))
                    MyDataAdapter.Parameters("@WORK_DAYS").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("WORK_DAYS").Index).Value
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

                Exit Sub
            Catch ex As Exception
                MsgBox(Err.Description & " " & Err.Number & " " & DataGridView1.CurrentCell.EditedFormattedValue, MsgBoxStyle.Critical)
                Grid()
            End Try
        End If
    End Sub

    '-------------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ ИСПОЛНИТЕЛЯ"-----------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormSuppliersPayeeAdd.ShowDialog()
    End Sub

    '-------------------------------------------ЗАКРЫТИЕ ФОРМЫ--------------------------------------
    Private Sub FormSuppliers_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormSuppliersPayee_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormSuppliersPayeeAdd.FSPAClosed IsNot Nothing Then
            If FormSuppliersPayeeAdd.FSPAClosed = True Then
                Grid()
                FormSuppliersPayeeAdd.FSPAClosed = False
            End If
        End If
    End Sub




End Class