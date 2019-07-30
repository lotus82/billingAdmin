Imports System.IO
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports System.Linq

Public Class FormExcel

    Dim SelectedSheet As String

    'Private Sub ShowSelectForm(ByVal Sheets() As String, ByVal FileName As String)
    '    If Sheets Is Nothing Then SelectedSheet = Nothing : Return
    '    Dim F As New Form, T As New TreeView, I As New ImageList, P As New Panel, B1, B2 As New Button
    '    With I
    '        .ColorDepth = ColorDepth.Depth32Bit
    '        .ImageSize = New Size(16, 16)
    '        .Images.Add("Книга", My.Resources.excel)  'рисунок 16x16 px для названия файла excel
    '        .Images.Add("Документ", My.Resources.excel)  'рисунок 16x16 px для названия листа excel
    '    End With
    '    With T
    '        .Name = "T"
    '        .ImageList = I : .Font = New Font("Arial", 10, FontStyle.Bold)
    '        .ShowPlusMinus = False : .ShowLines = False : .ShowRootLines = False
    '        Dim N As TreeNode = .Nodes.Add("Книга", FileIO.FileSystem.GetName(FileName), 0, 0)
    '        Dim nn As TreeNode
    '        For ni As Int16 = 0 To Sheets.Length - 1
    '            nn = N.Nodes.Add("Лист" & ni.ToString, Sheets(ni), 1, 1)
    '            nn.NodeFont = New Font(.Font, FontStyle.Regular)
    '        Next
    '        .ExpandAll()
    '    End With
    '    With F
    '        .Text = " Документ Excel"
    '        .ShowInTaskbar = False
    '        .StartPosition = FormStartPosition.CenterParent
    '        .FormBorderStyle = FormBorderStyle.FixedToolWindow
    '        .Height = .Height * 1.3
    '        .AcceptButton = B1 : .CancelButton = B2

    '        With P
    '            .Name = "P" : .Parent = F
    '            .Dock = DockStyle.Bottom
    '        End With

    '        With T
    '            .Parent = F
    '            .Dock = DockStyle.Fill
    '            .BringToFront()

    '            AddHandler .NodeMouseDoubleClick, AddressOf T_NodeDoubleClick
    '        End With

    '        With B2
    '            .Name = "B2"
    '            .Text = "Отмена"
    '            .Parent = P : .Top = 5
    '            .Left = P.Width - .Width - 5

    '            AddHandler .Click, AddressOf B2_Click
    '        End With

    '        With B1
    '            .Name = "B1"
    '            .Text = "Открыть"
    '            .Parent = P : .Top = 5
    '            .Left = B2.Left - .Width - 5
    '            .Font = New Font(.Font, FontStyle.Bold)

    '            AddHandler .Click, AddressOf B1_Click
    '        End With

    '        P.Height = B1.Height + 10

    '        AddHandler .Load, AddressOf F_Load
    '        .ShowDialog()
    '    End With
    'End Sub

    'Private Function GetExcelSheetNames(ByVal connection As OleDb.OleDbConnection) As String()
    '    Dim Table As DataTable

    '    Try
    '        Table = connection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
    '        If Table Is Nothing Then Return Nothing

    '        With Table
    '            Dim i As Integer = 0, SheetsArray() As String = Nothing, s As String

    '            For n As Integer = 0 To .Rows.Count - 1
    '                s = .Rows(n).Item("TABLE_NAME").ToString.Trim(New Char() {"'"})

    '                If Strings.Right(s, 1) = "$" Then
    '                    ReDim Preserve SheetsArray(i)
    '                    SheetsArray(i) = s.Trim(New Char() {"$"})
    '                    i += 1
    '                End If
    '            Next

    '            Table.Dispose() : Return SheetsArray
    '        End With

    '    Catch ex As Exception
    '        MsgBox(ex.Message, MsgBoxStyle.Exclamation)
    '        Return Nothing
    '    End Try
    'End Function





    'Public Sub OpenSheetCntr3(ByVal Grid As DataGridView)

    '    Dim connect As String = CDCrypt.Decrypt(My.Settings.conn, "Payments+2018!")
    '    Dim consString As String = connect
    '    Dim File_Type, Delimiter As String
    '    'Определяем тип файла
    '    File_Type = FormExcel.version.Rows.Item(FormExcel.ComboBox1.SelectedIndex).Item(2)
    '    'Определяем разделитель в файле, если файл текстовый
    '    If File_Type = "txt" Then
    '        Delimiter = FormExcel.version.Rows.Item(FormExcel.ComboBox1.SelectedIndex).Item(3)
    '    End If
    '    'Чтение таблицы INI и запись в массив
    '    Dim ini As New ArrayList
    '    Using con As New SqlConnection(consString)
    '        Dim cmIniVibor As New SqlCommand("Ini_vibor", con)
    '        cmIniVibor.CommandType = CommandType.StoredProcedure
    '        Dim drIni As SqlDataReader
    '        Dim prmBankId = New SqlParameter("@bank_id", SqlDbType.Int)
    '        prmBankId.Value = FormExcel.ComboBox2.SelectedIndex + 1
    '        cmIniVibor.Parameters.Add(prmBankId)
    '        Dim prmVersionId = New SqlParameter("@version_id", SqlDbType.Int)
    '        prmVersionId.Value = FormExcel.version.Rows.Item(FormExcel.ComboBox1.SelectedIndex).Item(1)
    '        cmIniVibor.Parameters.Add(prmVersionId)
    '        con.Open()
    '        'MsgBox(prmBankId.Value.ToString())
    '        'MsgBox(prmVersionId.Value.ToString())
    '        drIni = cmIniVibor.ExecuteReader()
    '        If drIni.HasRows = True Then
    '            While drIni.Read
    '                For i As Integer = 0 To drIni.FieldCount - 1
    '                    ini.Add(drIni.GetValue(i))
    '                Next
    '            End While
    '            drIni.Close()
    '            con.Close()
    '        End If
    '    End Using

    '    Dim OpenDialog As New OpenFileDialog
    '    Dim FileName As String = ""
    '    Dim FileNameNew As String = ""
    '    Dim NewPath As String = ""
    '    Dim Table As New DataTable
    '    'Открываем диалог выбора файла в соответствии с расширением и получаем имя файла
    '    If (File_Type = "xls" Or File_Type = "xlsx") Then
    '        'Открытие файла типа XLS
    '        With OpenDialog
    '            .Title = "Открыть документ Excel"
    '            .Filter = "Документы Excel|*.xls;*.xlsx"
    '            If .ShowDialog = Windows.Forms.DialogResult.OK Then
    '                FileName = .FileName : Application.DoEvents()
    '            Else
    '                Return
    '            End If
    '        End With
    '        Dim f As String = IO.Path.GetFileName(FileName)
    '        Dim fpatch As String = IO.Path.GetDirectoryName(FileName)
    '        Try
    '            My.Computer.FileSystem.RenameFile(FileName, "tmp_" & f)
    '            My.Computer.FileSystem.RenameFile(fpatch + "/" + "tmp_" & f, f)
    '        Catch ex As Exception
    '            If ex.TargetSite.Name = "WinIOError" Then MsgBox("Файл занят другим процессом", MsgBoxStyle.Critical)
    '            Exit Sub
    '        End Try
    '        ' Подключение к Excel. 
    '        Dim connection As OleDb.OleDbConnection, connectionString As String
    '        Try
    '            'Для Excel 12.0 
    '            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + FileName + "; Extended Properties=""Excel 12.0 Xml;HDR=" & ini(2) & """;"
    '            connection = New OleDb.OleDbConnection(connectionString)
    '            connection.Open()
    '        Catch ex12 As Exception
    '            Try
    '                'Для более ранних версий 
    '                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FileName + "; Extended Properties=""Excel 8.0;HDR=" & ini(2) & ";IMEX=1"";"
    '                connection = New OleDb.OleDbConnection(connectionString)
    '                connection.Open()
    '            Catch ex11 As Exception
    '                MsgBox("Неизвестная версия Excel или файл имеет неизвестный формат!", MsgBoxStyle.Exclamation)
    '                Return
    '            End Try
    '        End Try
    '        'Отобразить форму выбора листов
    '        ShowSelectForm(GetExcelSheetNames(connection), FileName)
    '        If SelectedSheet Is Nothing Then Return
    '        'Выборка данных 
    '        Dim command As OleDb.OleDbCommand = connection.CreateCommand()
    '        command.CommandText = "Select * From [" & SelectedSheet & "$A:Z] "
    '        Dim Adapter As New OleDb.OleDbDataAdapter(command)
    '        Adapter.Fill(Table) : connection.Close()
    '        FormExcel.Import.Text = "Выбран файл: " & FileName
    '    End If

    '    'Заполнение таблицы SQL из DataTable
    '    If Table.Rows.Count > 0 Then
    '        Using con As New SqlConnection(consString)
    '            Dim delString As String = "set rowcount 0 DELETE FROM dbo.Tmp_Payments"
    '            Dim command2 As New SqlCommand(delString, con)
    '            con.Open()
    '            command2.ExecuteNonQuery()
    '            con.Close()
    '            Using sqlBulkCopy As New SqlBulkCopy(con)
    '                sqlBulkCopy.DestinationTableName = "dbo.Tmp_Payments"
    '                con.Open()
    '                sqlBulkCopy.WriteToServer(Table)
    '                con.Close()
    '            End Using
    '        End Using
    '    End If








    'End Sub

End Class