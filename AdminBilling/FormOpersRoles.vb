Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Class FormOpersRoles

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormOpersRoles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid()
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА РОЛЕЙ ОПЕРА-----------------
    Public Sub Grid()
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlDataAdapter
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlDataAdapter("Adm_Opers_Roles", MyConnection)
            MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@GUID_oper", SqlDbType.UniqueIdentifier))
            MyDataAdapter.SelectCommand.Parameters("@GUID_oper").Value = FormOpers.GUID_OPER
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
            '------------------------ID роли опера---------------------------------------------
            Dim colRole_ID As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colRole_ID.Name = "ROLE_ID"
            'colRole_ID.ReadOnly = True
            colRole_ID.HeaderText = "ID"
            colRole_ID.DataPropertyName = "ROLE_ID"
            'colRole_ID.SortMode = DataGridViewColumnSortMode.Automatic
            'colRole_ID.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colRole_ID.Visible = False
            'colRole_ID.MinimumWidth = 20
            colRole_ID.Width = 20
            DataGridView1.Columns.Add(colRole_ID)
            '------------------------Роль опера---------------------------------------------
            Dim colL_NAME As DataGridViewColumn = New DataGridViewTextBoxColumn()
            colL_NAME.Name = "ROLE_NAME"
            colL_NAME.ReadOnly = True
            colL_NAME.HeaderText = "Роль"
            colL_NAME.DataPropertyName = "ROLE_NAME"
            colL_NAME.SortMode = DataGridViewColumnSortMode.Automatic
            colL_NAME.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            colL_NAME.Visible = True
            colL_NAME.MinimumWidth = 800
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
            '-------------------------сортировка по роли по возрастанию--------------------
            DataGridView1.Sort(DataGridView1.Columns("ROLE_NAME"), System.ComponentModel.ListSortDirection.Ascending)
            'Dim Cell As DataGridViewCell
            Dim Row As DataGridViewRow
            For Each Row In DataGridView1.Rows
                Row.Cells.Item("DelRow").Style.BackColor = Color.Red
            Next
            Label_Error.ForeColor = Color.Blue
            Label_Error.Text = "Количество ролей данного оператора БД: " + DS.Tables(0).Rows.Count.ToString() + " шт."
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
        End Try
    End Sub




    Private Sub FormOpersRoles_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

    End Sub

    '-------------------------------------НАЖАТИЕ НА КНОПКУ "ДОБАВИТЬ РОЛЬ"----------------------------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormOpersRolesAdd.ShowDialog()
    End Sub

    '-------------------------------------КЛИК ПО ЯЧЕЙКЕ ГРИДА----------------------------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        '-------------------------------------Клик по кнопке удалить------------------------------------------
        Try
            If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
                Dim Y_del As Integer = DataGridView1.CurrentCellAddress.Y
                Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить роль оператора: """ & DataGridView1.Rows.Item(Y_del).Cells.Item(DataGridView1.Columns(1).Index).Value & """?", vbYesNo)
                If DRes = DialogResult.Yes Then
                    Label_Error.Text = ""
                    Dim MyConnection As SqlConnection
                    Dim MyDataAdapter As SqlCommand
                    MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                    MyDataAdapter = New SqlCommand("Adm_Opers_Roles_Edit", MyConnection)
                    MyDataAdapter.CommandType = CommandType.StoredProcedure
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                    MyDataAdapter.Parameters("@type").Value = 2
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@GUID_oper", SqlDbType.UniqueIdentifier))
                    MyDataAdapter.Parameters("@GUID_oper").Value = FormOpers.GUID_OPER
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
                    MsgBox("Удаление роли опратора """ & DataGridView1.Rows.Item(Y_del).Cells.Item(DataGridView1.Columns(1).Index).Value & """ отменено")
                End If
            End If
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
            MsgBox(ex.Message)
        End Try
    End Sub
End Class