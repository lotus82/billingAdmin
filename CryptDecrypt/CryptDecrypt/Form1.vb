Imports Cr2015


Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox3.Text = Cr2015.CDCrypt.Encrypt(TextBox1.Text, TextBox2.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox6.Text = Cr2015.CDCrypt.Decrypt(TextBox4.Text, TextBox5.Text)
    End Sub
End Class
