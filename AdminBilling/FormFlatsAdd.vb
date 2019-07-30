Imports ClassAuth
Imports System.Data.SqlClient

Public Class FormFlatsAdd

    Public Shared FFAClosed = False
    Dim id_builds As Integer
    Dim BuildsName As String
    Dim StrName As String

    '-------------------------------------ЗАГРУЗКА ФОРМЫ----------------------------------------------
    Private Sub FormFlatsAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        id_builds = FormFlats.id_builds
        BuildsName = FormFlats.BuildName
        StrName = FormFlats.StrName
        Label12.Text = "Адрес: улица """ & StrName & """, дом № " & BuildsName
    End Sub

    '--------------------------ДОБАВЛЕНИЕ ДИАПАЗОНА------------------------------------
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            For i As Integer = Convert.ToInt32(TextBox2.Text) To Convert.ToInt32(TextBox3.Text)
                Dim MyConnection As SqlConnection
                Dim MyDataAdapter As SqlCommand
                MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
                MyDataAdapter = New SqlCommand("Adm_Flats_Edit", MyConnection)
                MyDataAdapter.CommandType = CommandType.StoredProcedure
                MyConnection.Open()
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
                MyDataAdapter.Parameters("@type").Value = 3
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_street", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_street").Value = Convert.ToInt32(FormBuilds.StrID)
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@id_flats", SqlDbType.Int))
                MyDataAdapter.Parameters("@id_flats").Value = vbNull
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FLAT", SqlDbType.Int))
                MyDataAdapter.Parameters("@FLAT").Value = i
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_FLAT_TYPE", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_FLAT_TYPE").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ENTRANCE", SqlDbType.Int))
                MyDataAdapter.Parameters("@ENTRANCE").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@FLOORFLAT", SqlDbType.Int))
                MyDataAdapter.Parameters("@FLOORFLAT").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@AREA", SqlDbType.Float))
                MyDataAdapter.Parameters("@AREA").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NUMBER_REG", SqlDbType.Int))
                MyDataAdapter.Parameters("@NUMBER_REG").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@NUMBER_ROOMS", SqlDbType.Int))
                MyDataAdapter.Parameters("@NUMBER_ROOMS").Value = 0
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_BUILD", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_BUILD").Value = id_builds
                '---------------------------------------------------------------------------------------
                MyDataAdapter.Parameters.Add(New SqlParameter("@ID_FLAT_STATUS", SqlDbType.Int))
                MyDataAdapter.Parameters("@ID_FLAT_STATUS").Value = 0
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
                MyDataAdapter.ExecuteNonQuery()
                MyConnection.Close()
            Next

            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox6.Text = ""
            Me.Close()
            Exit Sub
        Catch ex As Exception
            MsgBox(Err.Description & " " & Err.Number, MsgBoxStyle.Critical)
        End Try
    End Sub

    '--------------------------ДОБАВЛЕНИЕ ОДНОЙ КВАРТИРЫ-------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim MyConnection As SqlConnection
            Dim MyDataAdapter As SqlCommand
            MyConnection = New SqlConnection(ClassAuth.ClassAuth.con)
            MyDataAdapter = New SqlCommand("Adm_Flats_Edit", MyConnection)
            MyDataAdapter.CommandType = CommandType.StoredProcedure
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@type", SqlDbType.Int))
            MyDataAdapter.Parameters("@type").Value = 3
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@id_street", SqlDbType.Int))
            MyDataAdapter.Parameters("@id_street").Value = FormBuilds.StrID
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@id_flats", SqlDbType.Int))
            MyDataAdapter.Parameters("@id_flats").Value = vbNull
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@FLAT", SqlDbType.Int))
            MyDataAdapter.Parameters("@FLAT").Value = Convert.ToInt32(TextBox1.Text)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ID_FLAT_TYPE", SqlDbType.Int))
            MyDataAdapter.Parameters("@ID_FLAT_TYPE").Value = Convert.ToInt32(ComboBox1.SelectedValue)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ENTRANCE", SqlDbType.Int))
            MyDataAdapter.Parameters("@ENTRANCE").Value = Convert.ToInt32(ComboBox2.SelectedValue)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@FLOORFLAT", SqlDbType.Int))
            MyDataAdapter.Parameters("@FLOORFLAT").Value = Convert.ToInt32(ComboBox3.SelectedValue)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@AREA", SqlDbType.VarChar, 50))
            MyDataAdapter.Parameters("@AREA").Value = Convert.ToInt32(TextBox4.Text)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@NUMBER_REG", SqlDbType.Int))
            MyDataAdapter.Parameters("@NUMBER_REG").Value = Convert.ToInt32(TextBox5.Text)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@NUMBER_ROOMS", SqlDbType.Int))
            MyDataAdapter.Parameters("@NUMBER_ROOMS").Value = Convert.ToInt32(TextBox6.Text)
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ID_BUILD", SqlDbType.Int))
            MyDataAdapter.Parameters("@ID_BUILD").Value = id_builds
            '---------------------------------------------------------------------------------------
            MyDataAdapter.Parameters.Add(New SqlParameter("@ID_FLAT_STATUS", SqlDbType.Int))
            MyDataAdapter.Parameters("@ID_FLAT_STATUS").Value = Convert.ToInt32(ComboBox4.SelectedValue)
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
            Me.Close()
            Exit Sub
        Catch ex As Exception
            MsgBox(Err.Description & " " & Err.Number, MsgBoxStyle.Critical)
        End Try
    End Sub

    '-------------------------------------ЗАКРЫТИЕ ФОРМЫ----------------------------------------------
    Private Sub FormFlatsAdd_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        'FormFlats.Grid()
        FFAClosed = True
    End Sub


End Class