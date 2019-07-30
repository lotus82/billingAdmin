Imports Readini

Public Class FormAdmin

    '-------------------------------------------ЗАГРУЗКА ФОРМЫ--------------------------------
    Private Sub FormAdmin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Авторизованный пользователь: "
        Label2.Text = My.Settings.Oper_FIO
        Label3.Text = "Права:"
        Label4.Text = "Администратор"
        Label5.Text = "Город:"
        Label6.Text = Readini.Readini.ReadIni("Connect", "city")
    End Sub

    Private Sub ГородаToolStripMenuItem_Click(sender As Object, e As EventArgs) 
        Dim FormCities As New FormCities
        FormCities.ShowDialog()
    End Sub

    Private Sub УлицыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УлицыToolStripMenuItem.Click
        Dim FormStreets As New FormStreets
        FormStreets.Show()
        Me.Close()
    End Sub

    Private Sub ДомаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДомаToolStripMenuItem.Click
        Dim FormBuilds As New FormBuilds
        FormBuilds.Show()
        Me.Close()
    End Sub

    Private Sub ТехучасткиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТехучасткиToolStripMenuItem.Click
        Dim FormTech As New FormTech
        FormTech.Show()
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) 
        Dim FormUsers As New FormUsers
        FormUsers.Show()
        Me.Close()
    End Sub

    Private Sub ДокументыToolStripMenuItem_Click(sender As Object, e As EventArgs) 
        Dim FormUsersDocks As New FormUsersDocks
        FormUsersDocks.Show()
        Me.Close()
    End Sub

    Private Sub ОтношенияToolStripMenuItem_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub ТипыДокументовToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТипыДокументовToolStripMenuItem.Click
        Dim FormUsersDocksTypes As New FormUsersDocksTypes
        FormUsersDocksTypes.Show()
        Me.Close()
    End Sub

    Private Sub ТипыОтношенийToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТипыОтношенийToolStripMenuItem.Click
        Dim FormRelationsTypes As New FormRelationsTypes
        FormRelationsTypes.Show()
        Me.Close()
    End Sub

    Private Sub ТипРегистрацииToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТипРегистрацииToolStripMenuItem.Click
        Dim FormRegistrationTypes As New FormRegistrationTypes
        FormRegistrationTypes.Show()
        Me.Close()
    End Sub

    Private Sub ПоставщикиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПоставщикиToolStripMenuItem.Click
        Dim FormSuppliers As New FormSuppliers
        FormSuppliers.Show()
        Me.Close()
    End Sub

    Private Sub ИсполнителиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ИсполнителиToolStripMenuItem.Click
        Dim FormSuppliersPayee As New FormSuppliersPayee
        FormSuppliersPayee.Show()
        Me.Close()
    End Sub

    Private Sub УслугиToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles УслугиToolStripMenuItem1.Click
        Dim FormServices As New FormServices
        FormServices.Show()
        Me.Close()
    End Sub

    Private Sub ТарифыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТарифыToolStripMenuItem.Click
        Dim FormTarifs As New FormTarifs
        FormTarifs.Show()
        Me.Close()
    End Sub

    Private Sub ЕдиницыИзмеренияToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim FormServicesUnitsTypes As New FormServicesUnitsTypes
        FormServicesUnitsTypes.Show()
        Me.Close()
    End Sub

    Private Sub ПриборыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПриборыToolStripMenuItem.Click
        Dim FormCounters As New FormCounters
        FormCounters.Show()
        Me.Close()
    End Sub

    Private Sub ЛицевыеСчетаToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Dim FormLS As New FormLS
        FormLS.Show()
        Me.Close()
    End Sub

    Private Sub АрхивацияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles АрхивацияToolStripMenuItem.Click
        Dim FormDbCreateBackap As New FormDbCreateBackap
        FormDbCreateBackap.Show()
        Me.Close()
    End Sub

    Private Sub ВосстановлениеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВосстановлениеToolStripMenuItem.Click
        Dim FormDbRestoreBackap As New FormDbRestoreBackap
        FormDbRestoreBackap.Show()
        Me.Close()
    End Sub

    Private Sub ДанныеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДанныеToolStripMenuItem.Click
        Dim FormOpers As New FormOpers
        FormOpers.Show()
        Me.Close()
    End Sub

    Private Sub РолиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles РолиToolStripMenuItem.Click
        Dim FormRoles As New FormRoles
        FormRoles.Show()
        Me.Close()
    End Sub
End Class
