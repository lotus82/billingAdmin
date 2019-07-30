Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Class FormRoles
    Dim X As Integer
    Dim Y As Integer
    Dim V As String

    '-------------------------------------ЗАГРУЗКА ФОРМЫ--------------------------------------------------------------------------------------
    Private Sub FormRoles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ--------------------------------------------------------------------------------------
    Private Sub FormRoles_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ-------------------------------------------------------------------------------------
    Private Sub FormRoles_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormRolesAdd.FSAClosed IsNot Nothing Then
            If FormRolesAdd.FSAClosed = True Then
                Grid()
                FormRolesAdd.FSAClosed = False
            End If
        End If
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА РОЛЕЙ------------------------------------------------------------------------------
    Public Sub Grid()
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlDataAdapter
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlDataAdapter("Adm_Roles", MyConnection)
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
            '------------------------ID Роли---------------------------------------------------
            Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
            col1.Name = "ROLE_ID"
            col1.DataPropertyName = "ROLE_ID"
            DataGridView1.Columns.Add(col1)
            col1.Visible = False
            '------------------------Фамилия опера---------------------------------------------
            Dim colL_NAME As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colL_NAME.Name = "ROLE_NAME"
            colL_NAME.ReadOnly = False
            colL_NAME.HeaderText = "Роли"
            colL_NAME.DataPropertyName = "ROLE_NAME"
            colL_NAME.SortMode = DataGridViewColumnSortMode.Automatic
            colL_NAME.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colL_NAME.Visible = True
            colL_NAME.MinimumWidth = 900
            colL_NAME.Width = 800
            DataGridView1.Columns.Add(colL_NAME)
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
            '-------------------------сортировка по фамилии по возрастанию--------------------
            DataGridView1.Sort(DataGridView1.Columns("ROLE_NAME"), System.ComponentModel.ListSortDirection.Ascending)
            Dim Row As DataGridViewRow
            For Each Row In DataGridView1.Rows
                Row.Cells.Item("DelRow").Style.BackColor = Color.Red
            Next
            Label_Error.ForeColor = Color.Blue
            Label_Error.Text = "Количество ролей: " + DS.Tables(0).Rows.Count.ToString() + " шт."
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
        End Try
    End Sub

    '-------------------------------------ПОЛУЧЕНИЕ ЗНАЧЕНИЯ ЯЧЕЙКИ ДО РЕДАКТИРОВАНИЯ---------------------------------------------------------
    Private Sub DataGridView1_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        X = DataGridView1.CurrentCellAddress.X
        Y = DataGridView1.CurrentCellAddress.Y
        V = DataGridView1(X, Y).Value.ToString
    End Sub

    '-------------------------------------РЕДАКТИРОВАНИЕ ЯЧЕЕК--------------------------------------------------------------------------------
    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlCommand
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlCommand("Adm_Roles_Edit", MyConnection)
            MyDataAdapter.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.Parameters("@type").Value = 1
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ROLE_ID", SqlDbType.Int))
            MyDataAdapter.Parameters("@ROLE_ID").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ROLE_ID").Index).Value
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ROLE_NAME", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@ROLE_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ROLE_NAME").Index).Value
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

    '-------------------------------------НАЖАТИЕ НА КНОПКУ "ДОБАВИТЬ РОЛЬ"-------------------------------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormRolesAdd.ShowDialog()
    End Sub

    '-------------------------------------КЛИК ПО ЯЧЕЙКЕ ГРИДА--------------------------------------------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        '-------------------------------------Клик по кнопке удалить------------------------------------------
        Try
            If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
                Dim Y_del As Integer = DataGridView1.CurrentCellAddress.Y
                Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить роль: """ & DataGridView1.Rows.Item(Y_del).Cells.Item(DataGridView1.Columns(1).Index).Value & """?", vbYesNo)
                If DRes = DialogResult.Yes Then
                    Label_Error.Text = ""
                    Dim MyConnection As SqlConnection
                    Dim MyDataAdapter As SqlCommand
                    MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                    MyDataAdapter = New SqlCommand("Adm_Roles_Edit", MyConnection)
                    MyDataAdapter.CommandType = CommandType.StoredProcedure
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                    MyDataAdapter.Parameters("@type").Value = 2
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@ROLE_ID", SqlDbType.Int))
                    MyDataAdapter.Parameters("@ROLE_ID").Value = DataGridView1.Rows.Item(Y_del).Cells.Item(DataGridView1.Columns(0).Index).Value
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
                    MsgBox("Удаление роли """ & DataGridView1.Rows.Item(Y_del).Cells.Item(DataGridView1.Columns(1).Index).Value & """ отменено")
                End If
            End If
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
            MsgBox(ex.Message)
        End Try
    End Sub


End Class