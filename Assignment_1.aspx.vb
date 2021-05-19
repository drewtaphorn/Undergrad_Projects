Public Class Assingment_1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Background_Image.ImageUrl = "Tablet.JPG"
    End Sub

    Dim FirstNumber As Decimal
    Dim SecondNumber As Decimal
    Dim CalculatedNumber As Decimal


    Protected Sub Add_Button_Click(sender As Object, e As EventArgs) Handles Add_Button.Click
        If IsNumeric(First_Num_Txt.Text) Then
            FirstNumber = Convert.ToDecimal(First_Num_Txt.Text)
        Else
            Response.Write("Please type a number in the First Number Text Box")
            Exit Sub
        End If
        If IsNumeric(Second_Num_Txt.Text) Then
            SecondNumber = Convert.ToDecimal(Second_Num_Txt.Text)
        Else
            Response.Write("Please type a number in the Second Number Text Box")
            Exit Sub
        End If
        CalculatedNumber = FirstNumber + SecondNumber
        Calculated_Txt.Text = "Result of the calculation: " & vbNewLine & vbNewLine & CalculatedNumber.ToString("N2")
    End Sub

    Protected Sub Subtract_Button_Click(sender As Object, e As EventArgs) Handles Subtract_Button.Click
        If IsNumeric(First_Num_Txt.Text) Then
            FirstNumber = Convert.ToDecimal(First_Num_Txt.Text)
        Else
            Response.Write("Please type a number in the First Number Text Box")
            Exit Sub
        End If
        If IsNumeric(Second_Num_Txt.Text) Then
            SecondNumber = Convert.ToDecimal(Second_Num_Txt.Text)
        Else
            Response.Write("Please type a number in the Second Number Text Box")
            Exit Sub
        End If
        CalculatedNumber = FirstNumber - SecondNumber
        Calculated_Txt.Text = "Result of the calculation: " & vbNewLine & vbNewLine & CalculatedNumber.ToString("N2")
    End Sub

    Protected Sub Multiply_Button_Click(sender As Object, e As EventArgs) Handles Multiply_Button.Click
        If IsNumeric(First_Num_Txt.Text) Then
            FirstNumber = Convert.ToDecimal(First_Num_Txt.Text)
        Else
            Response.Write("Please type a number in the First Number Text Box")
            Exit Sub
        End If
        If IsNumeric(Second_Num_Txt.Text) Then
            SecondNumber = Convert.ToDecimal(Second_Num_Txt.Text)
        Else
            Response.Write("Please type a number in the Second Number Text Box")
            Exit Sub
        End If
        CalculatedNumber = FirstNumber * SecondNumber
        Calculated_Txt.Text = "Result of the calculation: " & vbNewLine & vbNewLine & CalculatedNumber.ToString("N2")
    End Sub

    Protected Sub Divide_Button_Click(sender As Object, e As EventArgs) Handles Divide_Button.Click
        If IsNumeric(First_Num_Txt.Text) Then
            FirstNumber = Convert.ToDecimal(First_Num_Txt.Text)
        Else
            Response.Write("Please type a number in the First Number Text Box")
            Exit Sub
        End If
        If IsNumeric(Second_Num_Txt.Text) Then
            SecondNumber = Convert.ToDecimal(Second_Num_Txt.Text)
        Else
            Response.Write("Please type a number in the Second Number Text Box")
            Exit Sub
        End If
        CalculatedNumber = FirstNumber / SecondNumber
        Calculated_Txt.Text = "Result of the calculation: " & vbNewLine & vbNewLine & CalculatedNumber.ToString("N2")
    End Sub

    Protected Sub Change_Image_Button_Click(sender As Object, e As EventArgs) Handles Change_Image_Button.Click
        Background_Image.ImageUrl = "Redbird.png"
        Call Clear()
    End Sub

    Protected Sub Clear()
        First_Num_Txt = Nothing
        Second_Num_Txt = Nothing
        Calculated_Txt = Nothing
    End Sub

End Class