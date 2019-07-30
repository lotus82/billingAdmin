Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception


Public Class FormUsersDocksTypes
    Dim X As Integer
    Dim Y As Integer
    Dim V As String

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormUsersDocksTypes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid() 'вызов функции заполнения грида
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА ЛЮДЕЙ--------------------------------------
    Private Sub Grid()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_DocksTypes", MyConnection)
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
        col1.Name = "ID_doc_type"
        col1.ReadOnly = True
        col1.HeaderText = "ID"
        col1.DataPropertyName = "ID_doc_type"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col1.Visible = False
        DataGridView1.Columns.Add(col1)
        '---------------------------------------------------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "TypeDoc"
        col2.ReadOnly = False
        col2.HeaderText = "Тип документа"
        col2.DataPropertyName = "Type"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 919
        col2.Width = 919
        DataGridView1.Columns.Add(col2)
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
        '-------------------------сортировка по типу документа по возрастанию--------------------
        DataGridView1.Sort(DataGridView1.Columns("TypeDoc"), System.ComponentModel.ListSortDirection.Ascending)
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
        ' MsgBox("id: " & " тип: " & DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TypeDoc").Index).Value)
        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_DocksTypes_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_doc_type", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_doc_type").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_doc_type").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TypeDoc", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@TypeDoc").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TypeDoc").Index).Value
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

    '-------------------------------------УДАЛЕНИЕ ТИПА ДОКУМЕНТА-------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить данный тип документа?", vbYesNo)
            If DRes = DialogResult.Yes Then
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_DocksTypes_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 2
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_doc_type", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_doc_type").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_doc_type").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@TypeDoc", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@TypeDoc").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("TypeDoc").Index).Value
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

    '-------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ ТИП ДОКУМЕНТА"---------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormUsersDocksTypesAdd.ShowDialog()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormUsersDocksTypes_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormStreets_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormUsersDocksTypesAdd.FUDAClosed IsNot Nothing Then
            If FormUsersDocksTypesAdd.FUDAClosed = True Then
                Grid()
                FormUsersDocksTypesAdd.FUDAClosed = False
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