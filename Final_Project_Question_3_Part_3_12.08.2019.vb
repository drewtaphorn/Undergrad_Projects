Public Class Final_Project_Question_3_Part_3_12.08.2019
    Private Sub Login_Click(sender As Object, e As EventArgs) Handles Login.Click
        Dim Username, Password As String
        Username = Username_Txt.Text
        Password = Password_Txt.Text
        For i As Integer = 0 To Final_Project_Question_3_Part_2_12.08.2019.Count - 1
            If Username = Final_Project_Question_3_Part_2_12.08.2019.Customers(i, 0) And Password = Final_Project_Question_3_Part_2_12.08.2019.Customers(i, 1) Then
                MsgBox("Welcome " & Final_Project_Question_3_Part_2_12.08.2019.Customers(i, 2) & "!")
                Exit Sub
            End If
        Next
        Username_Txt.Clear()
        Password_Txt.Clear()
        MsgBox("The username/password is incorrect")
    End Sub


    Private Sub Home_Click(sender As Object, e As EventArgs) Handles Home.Click
        Final_Project_Question_3_Part_1_12.08.2019.Show()
        Me.Hide()

        Username_Txt.Clear()
        Password_Txt.Clear()
    End Sub

    Private Sub Forgot_Password_Click(sender As Object, e As EventArgs) Handles Forgot_Password.Click

        Final_Project_Question_3_Part_4_12.08.2019.Show()
        Me.Hide()

    End Sub
End Class
