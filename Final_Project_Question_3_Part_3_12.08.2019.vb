Public Class Form3
    Private Sub Login_Click(sender As Object, e As EventArgs) Handles Login.Click
        Dim Username, Password As String
        Username = Username_Txt.Text
        Password = Password_Txt.Text
        For i As Integer = 0 To Form2.Count - 1
            If Username = Form2.Customers(i, 0) And Password = Form2.Customers(i, 1) Then
                MsgBox("Welcome " & Form2.Customers(i, 2) & "!")
                Exit Sub
            End If
        Next
        Username_Txt.Clear()
        Password_Txt.Clear()
        MsgBox("The username/password is incorrect")
    End Sub


    Private Sub Home_Click(sender As Object, e As EventArgs) Handles Home.Click
        Form1.Show()
        Me.Hide()

        Username_Txt.Clear()
        Password_Txt.Clear()
    End Sub

    Private Sub Forgot_Password_Click(sender As Object, e As EventArgs) Handles Forgot_Password.Click

        Form4.Show()
        Me.Hide()

    End Sub
End Class