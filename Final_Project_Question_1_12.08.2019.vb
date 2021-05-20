Public Class Final_Project_Question_1_12.08.2019
    Private Sub Find_Keyword_Click(sender As Object, e As EventArgs) Handles Find_Keyword.Click

        Dim StringKeywords() As String = IO.File.ReadAllLines("Words.txt")

        For i As Integer = 0 To StringKeywords.Count - 1
            StringKeywords(i) = (StringKeywords(i))
        Next

        Dim Str As String
        Dim Word As String

        Word = InputData()
        Str = Find(Word, StringKeywords)
        OutPut(Str)
    End Sub

    Sub OutPut(str As String)
        MsgBox(str)
    End Sub

    Function InputData() As String
        Dim i As String
        i = InputBox("Please enter a word")
        Return i
    End Function

    Function Find(Word As String, StringKeywords() As String) As String

        For i As Integer = 0 To StringKeywords.Count - 1
            If StringKeywords(i) = Word Then
                Return (Word & " found at place " & i + 1)
            End If
        Next

        Return ("Word not found!")

    End Function

End Class



