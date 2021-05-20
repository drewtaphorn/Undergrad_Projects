Public Class Form4
    Private Sub Retrieve_Password_Click(sender As Object, e As EventArgs) Handles Retrieve_Password.Click
        Dim Username, SecurityAnswer As String
        Username = Username_Txt.Text
        SecurityAnswer = Security_Answer_Txt.Text
        For i As Integer = 0 To Form2.Count - 1
            If Username = Form2.Customers(i, 0) And SecurityAnswer = Form2.Customers(i, 4) Then
                MsgBox("Your password is " & Form2.Customers(i, 1))
            Else
                MsgBox("Your answer to the Security Question is incorrect")
            End If
            Exit Sub
        Next
        Username_Txt.Clear()
        Security_Answer_Txt.Clear()
        Security_Question_Txt.Clear()


    End Sub

    Private Sub Home_Click(sender As Object, e As EventArgs) Handles Home.Click
        Form1.Show()
        Me.Hide()

        Username_Txt.Clear()
        Security_Question_Txt.Clear()
        Security_Answer_Txt.Clear()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim Username As String
        Username = Username_Txt.Text
        'SecurityAnswer = Security_Answer_Txt.Text
        For i As Integer = 0 To Form2.Count - 1
            If Username = Form2.Customers(i, 0) Then
                Security_Question_Txt.Text = Form2.Customers(i, 3)
                Exit Sub
            End If
        Next

        MsgBox("Username incorrect")


    End Sub
End Class