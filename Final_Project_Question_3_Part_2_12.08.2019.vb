Public Class Final_Project_Question_3_Part_2_12.08.2019
    Public Shared Customers(99, 4) As String
    Public Shared Count As Integer
    Public Shared CustomerName, Password, SecurityQuestion, SecurityAnswer As String
    Private Sub Register_Click(sender As Object, e As EventArgs) Handles Register.Click
        Dim Username As String

        CustomerName = Customer_Name_Txt.Text
        Username = User_Name_Txt.Text
        Password = Password_Txt.Text
        SecurityQuestion = Security_Question_Txt.Text
        SecurityAnswer = Security_Answer_Txt.Text

        For i As Integer = 0 To Count - 1
            If Username = Customers(i, 0) Then
                MsgBox("Username already exist")
                Exit Sub
            End If
        Next
        Customers(Count, 0) = Username
        Customers(Count, 1) = Password
        Customers(Count, 2) = CustomerName
        Customers(Count, 3) = SecurityQuestion
        Customers(Count, 4) = SecurityAnswer
        Count += 1
        Customer_Name_Txt.Clear()
        User_Name_Txt.Clear()
        Password_Txt.Clear()
        Security_Question_Txt.Clear()
        Security_Answer_Txt.Clear()

    End Sub

    Private Sub Home_Click(sender As Object, e As EventArgs) Handles Home.Click
        Final_Project_Question_3_Part_1_12.08.2019.Show()
        Me.Hide()
    End Sub
End Class
