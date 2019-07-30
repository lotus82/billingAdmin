Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormStreets

    Dim X As Integer
    Dim Y As Integer
    Dim V As String
    Dim errorCount As Integer = 0

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormStreets_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid() 'Вызов функции заполнения грида
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА УЛИЦ---------------------------------------
    Public Sub Grid()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_streets", MyConnection)
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
        '------------------------id улицы---------------------------------------------------
        Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col1.Name = "ID_streets"
        col1.ReadOnly = True
        col1.HeaderText = "ID улицы"
        col1.DataPropertyName = "ID_streets"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col1.Visible = False
        DataGridView1.Columns.Add(col1)
        '------------------------Название улицы---------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Street"
        col2.ReadOnly = False
        col2.HeaderText = "Название улицы"
        col2.DataPropertyName = "Street"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 424
        DataGridView1.Columns.Add(col2)
        '------------------------кладр------------------------------------------------------
        Dim col3 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col3.Name = "Code"
        col3.ReadOnly = False
        col3.HeaderText = "КЛАДР"
        col3.DataPropertyName = "Code"
        col3.SortMode = DataGridViewColumnSortMode.NotSortable
        col3.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col3.Visible = True
        col3.MinimumWidth = 150
        DataGridView1.Columns.Add(col3)
        '------------------------окато------------------------------------------------------
        Dim col4 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col4.Name = "OKATO"
        col4.ReadOnly = False
        col4.HeaderText = "ОКАТО"
        col4.DataPropertyName = "OKATO"
        col4.SortMode = DataGridViewColumnSortMode.NotSortable
        col4.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col4.Visible = True
        col4.MinimumWidth = 150
        DataGridView1.Columns.Add(col4)
        '------------------------октмо-----------------------------------------
        Dim col5 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col5.Name = "OKTMO"
        col5.ReadOnly = False
        col5.HeaderText = "ОКТМО"
        col5.DataPropertyName = "OKTMO"
        col5.SortMode = DataGridViewColumnSortMode.NotSortable
        col5.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col5.Visible = True
        col5.MinimumWidth = 150
        DataGridView1.Columns.Add(col5)
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
        DataGridView1.Sort(DataGridView1.Columns("Street"), System.ComponentModel.ListSortDirection.Ascending)
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

        '--------------------------проверка на соответствие поля OKATO целочисленному формату-------------------------------------
        If e.ColumnIndex = DataGridView1.Columns("OKATO").Index Then
            If ClassVvod.Format_Validation(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKATO").Index).Value.ToString, 2) Then
                errVvod = 0
            Else
                errVvod = 1
                DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
                MsgBox("Не верный формат ОКАТО")
            End If
        End If

        '--------------------------проверка на соответствие поля OKTMO целочисленному формату-------------------------------------
        If e.ColumnIndex = DataGridView1.Columns("OKTMO").Index Then
            If ClassVvod.Format_Validation(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKTMO").Index).Value.ToString, 2) Then
                errVvod = 0
            Else
                errVvod = 1
                DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
                MsgBox("Не верный формат OKTMO")
            End If
        End If

        '--------------------------проверка на соответствие поля КЛАДР целочисленному формату-------------------------------------
        If e.ColumnIndex = DataGridView1.Columns("Code").Index Then
            If ClassVvod.Format_Validation(DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Code").Index).Value.ToString, 2) Then
                errVvod = 0
            Else
                errVvod = 1
                DataGridView1.CurrentCell.Value = V 'возвращение предыдущего значения ячейки
                MsgBox("Не верный формат КЛАДР")
            End If
        End If

        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Street_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_street", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_street").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_streets").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@STREET_NAME", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@STREET_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Street").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@KLADR", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@KLADR").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Code").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@OKATO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@OKATO").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKATO").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@OKTMO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@OKTMO").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKTMO").Index).Value
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

    '-------------------------------------УДАЛЕНИЕ УЛИЦЫ----------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить данную улицу?", vbYesNo)
            If DRes = DialogResult.Yes Then
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Street_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 2
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_street", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_street").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_streets").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@STREET_NAME", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@STREET_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Street").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@KLADR", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@KLADR").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Code").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@OKATO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@OKATO").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKATO").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@OKTMO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@OKTMO").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("OKTMO").Index).Value
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

    '-------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ УЛИЦУ"--------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormStreetAdd.ShowDialog()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormStreets_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormStreets_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormStreetAdd.FSAClosed IsNot Nothing Then
            If FormStreetAdd.FSAClosed = True Then
                Grid()
                FormStreetAdd.FSAClosed = False
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
End Class