Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormBuildsAdd

    Public Shared FBAClosed = False
    Dim StrName As String

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormBuildsAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim MyConnection As SqlConnection
        Dim MyDataAdapter As SqlDataAdapter
        MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
        MyDataAdapter = New SqlDataAdapter("Adm_buildings", MyConnection)
        MyDataAdapter.SelectCommand.CommandType = CommandType.StoredProcedure
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@street_id", SqlDbType.Int))
        MyDataAdapter.SelectCommand.Parameters("@street_id").Value = FormBuilds.StrID
        '---------------------------------------------------------------------------------------
        MyDataAdapter.SelectCommand.Parameters.Add(New SqlParameter("@app_guid", SqlDbType.VarChar, 50))
        MyDataAdapter.SelectCommand.Parameters("@app_guid").Value = ClassAuth.ClassAuth.App_GUID()
        '---------------------------------------------------------------------------------------
        Dim DS As DataSet
        DS = New DataSet()
        MyDataAdapter.Fill(DS)
        ComboBox1.DataSource = DS.Tables(1)
        ComboBox1.DisplayMember = "Index_number"
        ComboBox1.ValueMember = "ID_Index"
        ComboBox2.DataSource = DS.Tables(2)
        ComboBox2.DisplayMember = "Name"
        ComboBox2.ValueMember = "ID_Tech"
        StrName = FormBuilds.StrName
        Label9.Text = "Улица """ & StrName & """"

    End Sub

    '-------------------------------------ДОБАВЛЕНИЕ ДОМА---------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim errVvod As Integer = 0

        If TextBox1.Text = "" Then
            errVvod = 1
            MsgBox("Поле номер дома не может быть пустым")
            Exit Sub
        End If

        '--------------------------проверка на соответствие поля FIAS_GUD формату GUID--------------------------------------------
        If ClassVvod.Format_Validation(TextBox6.Text, 1) Then
            errVvod = 0
        Else
            errVvod = 1
            TextBox6.Text = "" 'обновление TextBox
            MsgBox("Не верный формат GUID")
            Exit Sub
        End If

        '--------------------------проверка на соответствие поля OKATO целочисленному формату-------------------------------------
        If ClassVvod.Format_Validation(TextBox4.Text, 2) Then
            errVvod = 0
        Else
            errVvod = 1
            TextBox4.Text = "" 'обновление TextBox
            MsgBox("Не верный формат ОКАТО")
            Exit Sub
        End If

        '--------------------------проверка на соответствие поля ПОДЪЕЗДОВ целочисленному формату-------------------------------------
        If ClassVvod.Format_Validation(TextBox3.Text, 2) Then
            errVvod = 0
        Else
            errVvod = 1
            TextBox3.Text = "" 'обновление TextBox
            MsgBox("Не верный формат поля ПОДЪЕЗДОВ")
            Exit Sub
        End If


        '--------------------------проверка на соответствие поля ЭТАЖЕЙ целочисленному формату-------------------------------------
        If ClassVvod.Format_Validation(TextBox7.Text, 2) Then
            errVvod = 0
        Else
            errVvod = 1
            TextBox7.Text = "" 'обновление TextBox
            MsgBox("Не верный формат поля ЭТАЖЕЙ")
            Exit Sub
        End If

        ''--------------------------проверка на соответствие поля ID_tech целочисленному формату-------------------------------------
        'If ClassVvod.Format_Validation(TextBox5.Text, 2) Then
        '    errVvod = 0
        'Else
        '    errVvod = 1
        '    TextBox5.Text = "" 'обновление TextBox
        '    MsgBox("Не верный формат ID_tech")
        '    Exit Sub
        'End If

        If errVvod = 0 Then
            Try
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Bldn_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 3
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_bldn", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_bldn").Value = vbNull
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@street_id", SqlDbType.Int))
                MyDataAdapter.Parameters("@street_id").Value = FormBuilds.StrID
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@BLDN_NO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@BLDN_NO").Value = TextBox1.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@KORPUS", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@KORPUS").Value = TextBox2.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ENTRANCES", SqlDbType.Int))
                If TextBox3.Text = "" Then
                    MyDataAdapter.Parameters("@ENTRANCES").Value = 0
                Else MyDataAdapter.Parameters("@ENTRANCES").Value = TextBox3.Text
                End If
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FLOORS", SqlDbType.Int))
                If TextBox7.Text = "" Then
                    MyDataAdapter.Parameters("@FLOORS").Value = 0
                Else MyDataAdapter.Parameters("@FLOORS").Value = TextBox7.Text
                End If
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@INDEX_ID", SqlDbType.Int))
                MyDataAdapter.Parameters("@INDEX_ID").Value = ComboBox1.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@OKATO", SqlDbType.VarChar, 50))
                MyDataAdapter.Parameters("@OKATO").Value = TextBox4.Text
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_TECH", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_TECH").Value = ComboBox2.SelectedValue
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FIAS", SqlDbType.UniqueIdentifier))
                If TextBox6.Text = "" Then
                    MyDataAdapter.Parameters("@FIAS").Value = vbEmpty ' Guid.Empty 'vbNull  
                Else
                    MyDataAdapter.Parameters("@FIAS").Value = Guid.Parse(TextBox6.Text)
                End If
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
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox6.Text = ""
                TextBox7.Text = ""
                Me.Close()
                Exit Sub
            Catch ex As Exception
                MsgBox(Err.Description & " " & Err.Number, MsgBoxStyle.Critical)
            End Try
        Else Exit Sub
        End If

    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormBuildsAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        'FormBuilds.Grid()
        FBAClosed = True
    End Sub


End Class