Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormTech
    Dim X As Integer
    Dim Y As Integer
    Dim V As String
    Dim errorCount As Integer = 0

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormTech_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid() 'Вызов функции заполнения грида
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА--------------------------------------------
    Public Sub Grid()
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_tech", MyConnection)
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
        '------------------------id техучастка---------------------------------------------------
        Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col1.Name = "ID_tech"
        col1.ReadOnly = True
        col1.HeaderText = "ID техучастка"
        col1.DataPropertyName = "ID_tech"
        col1.SortMode = DataGridViewColumnSortMode.NotSortable
        col1.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col1.Visible = False
        DataGridView1.Columns.Add(col1)
        '------------------------Название техучастка---------------------------------------------
        Dim col2 As DataGridViewColumn = New DataGridViewTextBoxColumn()
        col2.Name = "Name"
        col2.ReadOnly = False
        col2.HeaderText = "Название техучастка"
        col2.DataPropertyName = "Name"
        col2.SortMode = DataGridViewColumnSortMode.Automatic
        col2.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        col2.Visible = True
        col2.MinimumWidth = 869
        DataGridView1.Columns.Add(col2)
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
        '-------------------------сортировка по названию техучастка по возрастанию--------------------
        DataGridView1.Sort(DataGridView1.Columns("Name"), System.ComponentModel.ListSortDirection.Ascending)
    End Sub

    '-------------------------------------ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ЯЧЕЙКИ ДО РЕДАКТИРОВАНИЯ-----------------
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        X = DataGridView1.CurrentCellAddress.X
        Y = DataGridView1.CurrentCellAddress.Y
        V = DataGridView1(X, Y).Value.ToString
    End Sub

    '-------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЕК----------------------------------------
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlCommand
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlCommand("Adm_Tech_Edit", MyConnection)
            MyDataAdapter.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.Parameters("@type").Value = 1
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@id_tech", SqlDbType.Int))
            MyDataAdapter.Parameters("@id_tech").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_tech").Index).Value
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@NAME", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Name").Index).Value
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
    End Sub

    '-------------------------------------УДАЛЕНИЕ ТЕХУЧАСТКА-----------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
            Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить данный техучасток?", vbYesNo)
            If DRes = DialogResult.Yes Then
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Tech_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 2
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_tech", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_tech").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ID_tech").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NAME", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Name").Index).Value
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

    '-------------------------------------ДОБАВЛЕНИЕ ТЕХУЧАСТКА---------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormTechAdd.ShowDialog()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormStreets_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormStreets_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormTechAdd.FTAClosed IsNot Nothing Then
            If FormTechAdd.FTAClosed = True Then
                Grid()
                FormTechAdd.FTAClosed = False
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