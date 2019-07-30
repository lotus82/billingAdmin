Imports ClassAuth
Imports System.Data.SqlClient
Imports System.Exception
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Collections.Generic

Public Class FormDbRestoreBackap

    Dim File_backap As String = ""

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------------------------------------
    Private Sub FormDbRestoreBackap_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Grid()
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------------------------------------
    Private Sub FormDbRestoreBackap_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub

    '-------------------------------------КЛИК ПО ЯЧЕЙКЕ ГРИДА------------------------------------------------------------
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Try
            '-------------------------------------Клик по кнопке удалить----------------------------------
            If e.ColumnIndex = DataGridView1.Columns("DelRow").Index Then
                Dim Y As Integer = DataGridView1.CurrentCellAddress.Y
                Dim DRes As DialogResult = MsgBox("Вы уверены, что хотите удалить файл: """ & DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns(0).Index).Value & """?", vbYesNo)
                If DRes = DialogResult.Yes Then
                    'Kill(DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns(0).Index).Value)
                    Dim MyConnection As SqlConnection
                    Dim MyDataAdapter As SqlCommand
                    MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                    MyDataAdapter = New SqlCommand("Adm_File_Delete", MyConnection)
                    MyDataAdapter.CommandType = CommandType.StoredProcedure
                    '---------------------------------------------------------------------------------------
                    MyDataAdapter.Parameters.Add(New SqlParameter("@File_Path", SqlDbType.VarChar, 500))
                    MyDataAdapter.Parameters("@File_Path").Value = DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns(0).Index).Value
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
                    MsgBox("Удаление " & DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns(0).Index).Value & " отменено")
                End If
            End If
            '-------------------------------------Клик по имени файла--------------------------------------
            If e.ColumnIndex = DataGridView1.Columns("Имя файла").Index Then
                Dim Y As Integer = DataGridView1.CurrentCellAddress.Y
                Dim DRes2 As DialogResult = MsgBox("Вы выбрали файл: """ & DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns(0).Index).Value & """. Восстановить БД из этого файла", vbYesNo)
                If DRes2 = DialogResult.Yes Then
                    File_backap = DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns(0).Index).Value
                Else
                    MsgBox("Восстановление из файла """ & DataGridView1.Rows.Item(Y).Cells.Item(DataGridView1.Columns(0).Index).Value & """ отменено")
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '-------------------------------------ЗАПОЛНЕНИЕ ГРИДА АРХИВОВ------------------------------------------------------------------
    Private Sub Grid()
        DataGridView1.Columns.Clear()
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlDataAdapter
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlDataAdapter("Adm_Backup_List", MyConnection)
            MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@db_name", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@db_name").Value = Readini.Readini.ReadIni("Connect", "catalog")
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
            '---------------------------------------------------------------------------------------
            Dim DS As DataSet
            DS = New DataSet()
            MyDataAdapter.Fill(DS)
            Dim st As New DataGridViewCellStyle With {.BackColor = Color.Red}
            For i As Integer = 0 To DS.Tables(0).Rows.Count - 1
                'If IO.File.Exists(DS.Tables(0).Rows.Item(i).Item(3).ToString()) Then
                DataGridView1.ColumnCount = 5
                DataGridView1.Columns(0).Name = "Имя файла"
                DataGridView1.Columns(0).Width = 400
                DataGridView1.Columns(0).ReadOnly = True
                DataGridView1.Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(1).Name = "Комментарий"
                DataGridView1.Columns(1).Width = 320
                DataGridView1.Columns(1).ReadOnly = True
                DataGridView1.Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(2).Name = "Размер (Kb)"
                DataGridView1.Columns(2).Width = 91
                DataGridView1.Columns(2).ReadOnly = True
                DataGridView1.Columns(2).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(3).Name = "Дата создания"
                DataGridView1.Columns(3).Width = 120
                DataGridView1.Columns(3).ReadOnly = True
                DataGridView1.Columns(3).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridView1.Columns(4).Name = "DelRow"
                DataGridView1.Columns(4).CellTemplate.Style = st
                DataGridView1.Columns(4).Width = 20
                DataGridView1.Columns(4).ReadOnly = True
                DataGridView1.Columns(4).HeaderText = ""
                Dim row As String() = New String() {DS.Tables(0).Rows.Item(i).Item(3).ToString(), DS.Tables(0).Rows.Item(i).Item(0).ToString(), DS.Tables(0).Rows.Item(i).Item(2).ToString(), DS.Tables(0).Rows.Item(i).Item(1).ToString(), "X"}
                DataGridView1.Rows.Add(row)
                'Else
                'MsgBox("Файл не существует")
                'End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    '*****************************************************************************************************************************************************************************************
    '*****************************************************************************************************************************************************************************************
    '-------------------------------------ВЫЗОВ ПРОЦЕССА ВОССТАНОВЛЕНИЯ АРХИВА------------------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ProgressBar1.Step = 1 ' шаг
        ProgressBar1.Maximum = 100
        ProgressBar1.Minimum = 0
        Timer1.Interval = 100
        Button1.Enabled = False ' отключаем доступность кнопки
        BackgroundWorker1.RunWorkerAsync() 'Запускаем асинхронную операцию
        Timer1.Enabled = True
    End Sub

    '-------------------------------------ПРОЦЕСС ВОССТАНОВЛЕНИЯ АРХИВА-------------------------------------------------------------
    Private Sub BackgroundWorker1_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlDataAdapter
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con_history)
            MyDataAdapter = New SqlDataAdapter("Adm_Restore_Backap_History", MyConnection)
            MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@fname", SqlDbType.VarChar, 300))
            MyDataAdapter.SelectCommand.Parameters("@fname").Value = File_backap
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@db_name", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@db_name").Value = Readini.Readini.ReadIni("Connect", "catalog")
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@OPER_GUID", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@OPER_GUID").Value = My.Settings.Oper_GUID
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@HOST", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@HOST").Value = ClassAuth.ClassAuth.Host().ToString()
            '---------------------------------------------------------------------------------------
            MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
            MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
            '---------------------------------------------------------------------------------------
            Dim DS As DataSet
            DS = New DataSet()
            MyDataAdapter.Fill(DS)
            MsgBox("Архив успешно восстановлен!")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    '-------------------------------------ОТОБРАЖЕНИЕ ПРОГРЕССБАРА ВОССТАНОВЛЕНИЯ ИЗ АРХИВА-----------------------------------------
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            Dim MyConnection2 As SqlConnection
            Dim MyDataAdapter2 As SqlDataAdapter
            MyConnection2 = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter2 = New SqlDataAdapter("Adm_Backap_Progress", MyConnection2)
            MyDataAdapter2.SelectCommand.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter2.SelectCommand.Parameters.Add(New SqlParameter("@comanda", SqlDbType.VarChar, 50))
            MyDataAdapter2.SelectCommand.Parameters("@comanda").Value = "RESTORE DATABASE"
            '---------------------------------------------------------------------------------------
            Dim DS2 As DataSet
            DS2 = New DataSet()
            MyDataAdapter2.Fill(DS2)
            If DS2.Tables(0).Rows.Count > 0 Then
                If DS2.Tables(0).Rows.Item(0).Item(0) = 100 Then
                    Timer1.Enabled = False
                    Timer1.Stop()
                    Label_Error.ForeColor = Color.Blue
                    Label_Error.Text = "Архив успешно восстановлен"
                Else
                    ProgressBar1.Value = DS2.Tables(0).Rows.Item(0).Item(0)
                    Label_Error.ForeColor = Color.Blue
                    Label_Error.Text = "Обработано:" & DS2.Tables(0).Rows(0).Item(0) & "%" & "  осталось времени: " & DS2.Tables(0).Rows(0).Item(1) & " мин."
                    Timer1.Interval = DS2.Tables(0).Rows(0).Item(2)
                End If
            Else
                ProgressBar1.Value = 100
                Label_Error.Text = "Операция завершена"
            End If
        Catch ex As Exception
            Label_Error.ForeColor = Color.Red
            Label_Error.Text = ex.Message
            MsgBox(ex.Message)
        End Try
    End Sub

    '-------------------------------------ЗАВЕРШЕНИЕ ПРОЦЕССА ВОССТАНОВЛЕНИЯ ИЗ АРХИВА----------------------------------------------
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        ProgressBar1.Visible = False
        'MsgBox("Восстановление базы завершено")
        Button1.Enabled = True
        Grid()
    End Sub
    '*****************************************************************************************************************************************************************************************
    '*****************************************************************************************************************************************************************************************
End Class