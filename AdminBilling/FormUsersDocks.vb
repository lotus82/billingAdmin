

Public Class FormUsersDocks

    '-------------------------------------------ЗАГРУЗКА ФОРМЫ-----------------------------------------
    Private Sub FormUsersDocks_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    '-------------------------------------------КЛИК ПО КНОПКЕ "ДОБАВИТЬ ДОМ"---------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FormUsersDocksAdd.ShowDialog()
    End Sub

    '-------------------------------------------ЗАКРЫТИЕ ФОРМЫ------------------------------------------
    Private Sub FormUsersDocks_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        FormAdmin.Show()
    End Sub


End Class