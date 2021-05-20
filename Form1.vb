Public Class Form1
    Public Shared HomeworkGrades(9) As Decimal
    Public Shared QuizGrades(4) As Decimal
    Public Shared TestGrades(1) As Decimal
    Public Shared TotalGrade As Decimal
    Dim Homework, Quiz, Test As Decimal

    Private Sub Homework_Grades_Click(sender As Object, e As EventArgs) Handles Homework_Grades.Click
        For i As Integer = 0 To 9
            HomeworkGrades(i) = InputBox("Please Enter Homework Grades")
        Next
    End Sub

    Private Sub Quiz_Grades_Click(sender As Object, e As EventArgs) Handles Quiz_Grades.Click
        For i As Integer = 0 To 4
            QuizGrades(i) = InputBox("Please Enter Quiz Grades")
        Next
    End Sub

    Private Sub Test_Grades_Click(sender As Object, e As EventArgs) Handles Test_Grades.Click
        For i As Integer = 0 To 1
            TestGrades(i) = InputBox("Please Enter Test Grades")
        Next
    End Sub

    Private Sub Calculate_Final_Grade_Click(sender As Object, e As EventArgs) Handles Calculate_Final_Grade.Click
        Dim Calculate As String
        Calculate = GetLetterGrade(HomeworkGrades, QuizGrades, TestGrades)
        OutPut(Calculate)
    End Sub
    Sub OutPut(Calculate As String)
        Final_Grade_Txt.Text = Calculate
    End Sub
    Function GetLetterGrade(HomeworkGrades, QuizGrades, TestGrades)

        For i As Integer = 0 To 9
            Homework += HomeworkGrades(i)
        Next
        For i As Integer = 0 To 4
            Quiz += QuizGrades(i)
        Next
        For i As Integer = 0 To 1
            Test += TestGrades(i)
        Next

        TotalGrade = (Homework + Quiz + Test)

        If TotalGrade >= 900 Then Return "A"
        If TotalGrade >= 800 Then Return "B"
        If TotalGrade >= 700 Then Return "C"
        If TotalGrade >= 600 Then Return "D"
        If TotalGrade < 600 Then Return "F"
    End Function
End Class
