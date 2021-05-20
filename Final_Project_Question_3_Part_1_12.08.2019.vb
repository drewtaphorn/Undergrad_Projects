Public Class Form1
    Private Sub Register_Click(sender As Object, e As EventArgs) Handles Register.Click

        Form2.Show()
        Me.Hide()

    End Sub

    Private Sub Login_Click(sender As Object, e As EventArgs) Handles Login.Click


        Form3.Show()
        Me.Hide()

    End Sub
End Class
