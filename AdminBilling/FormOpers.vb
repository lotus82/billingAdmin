Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Class FormOpers

    Public Shared GUID_OPER As Guid
    Dim X As Integer
    Dim Y As Integer
    Dim V As String
    Dim errorCount As Integer = 0

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormOpersLoad(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormOpersClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------АКТИВАЦИЯ ФОРМЫ---------------------------------------------
    Private Sub FormOpers_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        If FormOpersAdd.FSAClosed IsNot Nothing Then
            If FormOpersAdd.FSAClosed = True Then
                Grid()
                FormOpersAdd.FSAClosed = False
            End If
        End If
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА ОПЕРОВ-----------------
    Public Sub Grid()
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlDataAdapter
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlDataAdapter("Adm_Opers", MyConnection)
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
            '------------------------GUID опера---------------------------------------------------
            Dim col1 As DataGridViewColumn = New DataGridViewTextBoxColumn()
            col1.Name = "GUID_oper"
            col1.DataPropertyName = "GUID_oper"
            DataGridView1.Columns.Add(col1)
            col1.Visible = False
            '-------------------------Кнопка "Роли опера"-----------------------------------------------------
            Dim BtnColRole As DataGridViewImageColumn = New DataGridViewImageColumn()
            BtnColRole.Name = "OperRole"
            BtnColRole.SortMode = DataGridViewColumnSortMode.NotSortable
            BtnColRole.ToolTipText = "Роли"
            Dim inImg As Image = PictureBox1.Image
            BtnColRole.Image = inImg
            BtnColRole.HeaderText = "Роли"
            BtnColRole.DefaultCellStyle.ForeColor = Color.Black
            BtnColRole.DefaultCellStyle.BackColor = Color.LightYellow
            BtnColRole.DefaultCellStyle.SelectionBackColor = Color.Green
            BtnColRole.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font.Name, 10, FontStyle.Bold)
            BtnColRole.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
            BtnColRole.Width = 65
            DataGridView1.Columns.Add(BtnColRole)
            '------------------------Фамилия опера---------------------------------------------
            Dim colL_NAME As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colL_NAME.Name = "L_NAME"
            colL_NAME.ReadOnly = False
            colL_NAME.HeaderText = "Фамилия"
            colL_NAME.DataPropertyName = "L_NAME"
            colL_NAME.SortMode = DataGridViewColumnSortMode.Automatic
            colL_NAME.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colL_NAME.Visible = True
            colL_NAME.MinimumWidth = 100
            DataGridView1.Columns.Add(colL_NAME)
            '------------------------Имя опера------------------------------------------------------
            Dim colF_NAME As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colF_NAME.Name = "F_NAME"
            colF_NAME.ReadOnly = False
            colF_NAME.HeaderText = "Имя"
            colF_NAME.DataPropertyName = "F_NAME"
            colF_NAME.SortMode = DataGridViewColumnSortMode.NotSortable
            colF_NAME.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colF_NAME.Visible = True
            colF_NAME.MinimumWidth = 100
            DataGridView1.Columns.Add(colF_NAME)
            '-------------------------Отчество опера-----------------------------------------------------
            Dim colS_NAME As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colS_NAME.Name = "S_NAME"
            colS_NAME.ReadOnly = False
            colS_NAME.HeaderText = "Отчество"
            colS_NAME.DataPropertyName = "S_NAME"
            colS_NAME.SortMode = DataGridViewColumnSortMode.NotSortable
            colS_NAME.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colS_NAME.Visible = True
            colS_NAME.MinimumWidth = 100
            DataGridView1.Columns.Add(colS_NAME)
            '-------------------------Логин опера-----------------------------------------------------
            Dim colLogin As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colLogin.Name = "Login"
            colLogin.ReadOnly = False
            colLogin.HeaderText = "Логин"
            colLogin.DataPropertyName = "Login"
            colLogin.SortMode = DataGridViewColumnSortMode.NotSortable
            colLogin.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colLogin.Visible = True
            colLogin.MinimumWidth = 150
            DataGridView1.Columns.Add(colLogin)
            '------------------------Пароль опера-----------------------------------------
            Dim colPassword As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colPassword.Name = "Password"
            colPassword.ReadOnly = False
            colPassword.HeaderText = "Пароль"
            colPassword.DataPropertyName = "Password"
            colPassword.SortMode = DataGridViewColumnSortMode.NotSortable
            colPassword.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colPassword.Visible = True
            colPassword.Width = 50
            colPassword.MinimumWidth = 50
            DataGridView1.Columns.Add(colPassword)
            '------------------------Активный-----------------------------------------
            Dim colActive As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colActive.Name = "Active"
            colActive.ReadOnly = False
            colActive.HeaderText = "Активный"
            colActive.DataPropertyName = "Active"
            colActive.SortMode = DataGridViewColumnSortMode.NotSortable
            colActive.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colActive.Visible = True
            colActive.MinimumWidth = 40
            DataGridView1.Columns.Add(colActive)
            '------------------------Только чтение-----------------------------------------
            Dim colReadOnly As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colReadOnly.Name = "ReadOnly"
            colReadOnly.ReadOnly = False
            colReadOnly.HeaderText = "Только чтение"
            colReadOnly.DataPropertyName = "ReadOnly"
            colReadOnly.SortMode = DataGridViewColumnSortMode.NotSortable
            colReadOnly.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colReadOnly.Visible = True
            colReadOnly.MinimumWidth = 60
            DataGridView1.Columns.Add(colReadOnly)
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
            DataGridView1.Sort(DataGridView1.Columns("L_NAME"), System.ComponentModel.ListSortDirection.Ascending)
            Dim Row As DataGridViewRow
            For Each Row In DataGridView1.Rows
                Row.Cells.Item("DelRow").Style.BackColor = Color.Red
            Next
            Label_Error.ForeColor = Color.Blue
            Label_Error.Text = "Количество операторов БД: " + DS.Tables(0).Rows.Count.ToString() + " чел."
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
        End Try
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
                MyDataAdapter = New SqlCommand("Adm_Opers_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 1
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@GUID_oper", SqlDbType.UniqueIdentifier))
                MyDataAdapter.Parameters("@GUID_oper").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("GUID_oper").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@L_NAME", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@L_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("L_NAME").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@F_NAME", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@F_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("F_NAME").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@S_NAME", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@S_NAME").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("S_NAME").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Login", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Login").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Login").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Password", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Password").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Password").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@Active", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@Active").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("Active").Index).Value
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ReadOnly", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@ReadOnly").Value = DataGridView1.CurrentRow.Cells(DataGridView1.Columns("ReadOnly").Index).Value
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

    '-------------------------------------КЛИК ПО КРТИНКЕ "РОЛИ"--------------------------------------------------------------
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        '---------------------------------------Переход к форме "Роли"-------------------------------------------------
        If e.ColumnIndex = DataGridView1.Columns("OperRole").Index Then
            GUID_OPER = DataGridView1.CurrentRow.Cells.Item(0).Value
            FormOpersRoles.ShowDialog()
        End If
    End Sub

    '-------------------------------------НАЖАТИЕ НА КНОПКУ "ДОБАВИТЬ ОПЕРАТОРА"----------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormOpersAdd.ShowDialog()
    End Sub

    '-------------------------------------КЛИК ПО ЯЧЕЙКЕ ГРИДА----------------------------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        '-------------------------------------Клик по кнопке удалить------------------------------------------
        Try
            If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
                Dim Y_del As Integer = DataGridView1.CurrentCellAddress.Y
                Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить оператора: """ & DataGridView1.Rows.Item(Y_del).Cells.Item(DataGridView1.Columns(2).Index).Value & """?", vbYesNo)
                If DRes = DialogResult.Yes Then
                    Label_Error.Text = ""
                    Dim MyConnection As SqlConnection
                    Dim MyDataAdapter As SqlCommand
                    MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                    MyDataAdapter = New SqlCommand("Adm_Opers_Edit", MyConnection)
                    MyDataAdapter.CommandType = CommandType.StoredProcedure
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                    MyDataAdapter.Parameters("@type").Value = 2
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@GUID_oper", SqlDbType.UniqueIdentifier))
                    MyDataAdapter.Parameters("@GUID_oper").Value = DataGridView1.Rows.Item(Y_del).Cells.Item(DataGridView1.Columns(0).Index).Value
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
                    MsgBox("Удаление опратора """ & DataGridView1.Rows.Item(Y_del).Cells.Item(DataGridView1.Columns(2).Index).Value & """ отменено")
                End If
            End If
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
            MsgBox(ex.Message)
        End Try
    End Sub
End Class