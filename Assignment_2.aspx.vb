Public Class Assignment_2
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Dim MagazineQuantity As Decimal
    Dim MagazinePrice = 6.0
    Dim MagazineTotal As Decimal
    Dim BooksQuantity As Decimal
    Dim BooksPrice = 20.0
    Dim BooksTotal As Decimal
    Dim ToysQuantity As Decimal
    Dim ToysPrice = 15.0
    Dim ToysTotal As Decimal
    Dim SubTotal As Decimal
    Dim Tax As Decimal
    Dim TotalPrice As Decimal

    Protected Sub Calculate_Button_Click(sender As Object, e As EventArgs) Handles Calculate_Button.Click
        If IsNumeric(Magazine_Txt_Box.Text) Then
            MagazineQuantity = Convert.ToDecimal(Magazine_Txt_Box.Text)
        Else
            Response.Write("Please type a number value in the Magazine Field")
            Exit Sub
        End If
        If IsNumeric(Books_Txt_Box.Text) Then
            BooksQuantity = Convert.ToDecimal(Books_Txt_Box.Text)
        Else
            Response.Write("Please type a number value in the Books Field")
            Exit Sub
        End If
        If IsNumeric(Toys_Txt_Box.Text) Then
            ToysQuantity = Convert.ToDecimal(Toys_Txt_Box.Text)
        Else
            Response.Write("Please type a number value in the Toys Field")
            Exit Sub
        End If

        MagazineTotal = MagazinePrice * MagazineQuantity
        BooksTotal = BooksPrice * BooksQuantity
        ToysTotal = ToysPrice * ToysQuantity
        SubTotal = MagazineTotal + BooksTotal + ToysTotal
        Tax = SubTotal * 0.08
        TotalPrice = Tax + SubTotal

        Sub_Total_Txt_Box.Text = SubTotal.ToString("C2")
        Tax_Txt_Box.Text = Tax.ToString("C2")

        Total_Txt_Box.Text = MagazineQuantity & " Magazine(s) @ $6.00 = " & MagazineTotal.ToString("C2") & vbNewLine & BooksQuantity & " Book(s) @ $20.00 = " & BooksTotal.ToString("C2") & vbNewLine & ToysQuantity & " Toy(s) @ $15.00 = " & ToysTotal.ToString("C2") & vbNewLine & vbNewLine & "Subtotal = " & SubTotal.ToString("C2") & vbNewLine & vbNewLine & "Tax = " & Tax.ToString("C2") & vbNewLine & vbNewLine & "Total = " & TotalPrice.ToString("C2")
    End Sub

    Protected Sub Clear_Button_Click(sender As Object, e As EventArgs) Handles Clear_Button.Click
        Magazine_Txt_Box.Text = ""
        Books_Txt_Box.Text = ""
        Toys_Txt_Box.Text = ""
        Sub_Total_Txt_Box.Text = ""
        Tax_Txt_Box.Text = ""
        Total_Txt_Box.Text = ""
    End Sub
End Class