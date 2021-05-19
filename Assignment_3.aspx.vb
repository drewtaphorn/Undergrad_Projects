Public Class Assignment_3
    Inherits System.Web.UI.Page

    Public Shared gSubaruLegacyCounter, gSubaruForesterCounter, gSubaruOutbackCounter, gSubaruImprezaCounter, gSubaruBRZCounter, gTotalSaleQuotes, gTotalSales As Decimal
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Subaru_Image.ImageUrl = "Subaru.png"
    End Sub

    Protected Sub Total_Sales_Quote_Submit_Button_Click(sender As Object, e As EventArgs) Handles Total_Sales_Quote_Submit_Button.Click

        gTotalSaleQuotes = gSubaruBRZCounter + gSubaruForesterCounter + gSubaruImprezaCounter + gSubaruOutbackCounter + gSubaruOutbackCounter

        Total_Sales_Quote_Txt_Box.Text = "Subaru Legacy: " & gSubaruLegacyCounter & vbNewLine & "Subaru Forester: " & gSubaruForesterCounter & vbNewLine & "Subaru Outback: " & gSubaruOutbackCounter & vbNewLine & "Subaru Impreza: " & gSubaruImprezaCounter & vbNewLine & "Subaru BRZ: " & gSubaruBRZCounter & vbNewLine & "Total Sale Quotes for Today: " & gTotalSaleQuotes & vbNewLine & "Total Sales for Today: " & gTotalSales.ToString("C2")
    End Sub

    Protected Sub Submit_Sales_Quote_Button_Click(sender As Object, e As EventArgs) Handles Submit_Sales_Quote_Button.Click

        Dim BasePrice As Decimal
        Dim CommissionRate As Decimal
        Dim DiscountRate As Decimal
        Dim Accessories As Decimal
        Dim PaintJob As Decimal
        Dim StateTaxRate As Decimal
        Dim SubTotal, Discount, StateTax, Commission, Total As Decimal
        Dim Str As String = ""

        If Base_Price_List.SelectedIndex = 1 Then
            BasePrice = 22895
        ElseIf Base_Price_List.SelectedIndex = 2 Then
            BasePrice = 24795
        ElseIf Base_Price_List.SelectedIndex = 3 Then
            BasePrice = 26795
        ElseIf Base_Price_List.SelectedIndex = 4 Then
            BasePrice = 27495
        ElseIf Base_Price_List.SelectedIndex = 5 Then
            BasePrice = 28845
        End If

        For Each li As ListItem In Base_Price_List.Items
            If li.Selected Then
                Str &= "Car Price:" & vbNewLine
                Exit For
            End If
        Next

        For Each li As ListItem In Base_Price_List.Items
            If li.Selected Then
                BasePrice += li.Value
                Str &= li.Text & " @" & FormatCurrency(li.Value, 2) & vbNewLine
            End If
        Next

        If Base_Price_List.SelectedIndex = 1 Then
            gSubaruLegacyCounter = gSubaruLegacyCounter + 1
        ElseIf Base_Price_List.SelectedIndex = 2 Then
            gSubaruForesterCounter = gSubaruForesterCounter + 1
        ElseIf Base_Price_List.SelectedIndex = 3 Then
            gSubaruOutbackCounter = gSubaruOutbackCounter + 1
        ElseIf Base_Price_List.SelectedIndex = 4 Then
            gSubaruImprezaCounter = gSubaruImprezaCounter + 1
        ElseIf Base_Price_List.SelectedIndex = 5 Then
            gSubaruBRZCounter = gSubaruBRZCounter + 1
        End If

        If Yes_Checkbox_VIP.Checked = True Then
            CommissionRate = 0.015
        End If
        If No_Checkbox_VIP.Checked = True Then
            CommissionRate = 0.01
        End If
        If Yes_Checkbox_SalesPromo.Checked = True Then
            DiscountRate = 0.1
        End If
        If No_Checkbox_SalesPromo.Checked = True Then
            DiscountRate = 0
        End If

        For Each li As ListItem In Accessories_Checkbox_List.Items
            If li.Selected Then
                Str &= "Accessories:" & vbNewLine
                Exit For
            End If
        Next

        For Each li As ListItem In Accessories_Checkbox_List.Items
            If li.Selected Then
                Accessories += li.Value
                Str &= li.Text & " @" & FormatCurrency(li.Value, 2) & vbNewLine
            End If
        Next
        For Each li As ListItem In Paint_Color_List.Items
            If li.Selected Then
                Str &= "Paint Color: " & vbNewLine
                Exit For
            End If
        Next
        For Each li As ListItem In Paint_Color_List.Items
            If li.Selected Then
                Str &= li.Text & vbNewLine
                Exit For
            End If
        Next
        For Each li As ListItem In Paint_Job_List.Items
            If li.Selected Then
                Str &= "Paint Job:" & vbNewLine
                Exit For
            End If
        Next

        For Each li As ListItem In Paint_Job_List.Items
            If li.Selected Then
                PaintJob += li.Value
                Str &= li.Text & " @" & FormatCurrency(li.Value, 2) & vbNewLine
            End If
        Next
        For Each li As ListItem In State_Drop_List.Items
            If li.Selected Then
                Str &= "State of Purchase: " & li.Text & vbNewLine
                Exit For
            End If
        Next

        For Each li As ListItem In State_Drop_List.Items
            If li.Selected Then
                StateTaxRate += li.Value
            End If
        Next

        Discount = BasePrice * DiscountRate
        StateTax = SubTotal * StateTaxRate
        SubTotal = Base_Price_List.SelectedValue + Accessories + PaintJob - Discount
        StateTax = SubTotal * StateTaxRate
        Commission = SubTotal * CommissionRate

        Total = SubTotal + StateTax

        gTotalSales += Total

        Summary_Txt_Box.Text = Str & vbNewLine & "Sub Total: " & SubTotal.ToString("C2") & vbNewLine & "Tax: " & StateTax.ToString("C2") & vbNewLine & "Total: " & Total.ToString("C2")

    End Sub
    Protected Sub Clear_Button_Click(sender As Object, e As EventArgs) Handles Clear_Button.Click
        Base_Price_List.SelectedIndex = 0
        Paint_Color_List.SelectedIndex = 0
        Paint_Job_List.SelectedIndex = 0
        State_Drop_List.SelectedIndex = 0

        For Each li In Accessories_Checkbox_List.Items
            If li.selected = True Then li.selected = False
        Next
        If Yes_Checkbox_VIP.Checked = True Then
            Yes_Checkbox_VIP.Checked = False
        End If
        If No_Checkbox_VIP.Checked = True Then
            No_Checkbox_VIP.Checked = False
        End If
        If Yes_Checkbox_SalesPromo.Checked = True Then
            Yes_Checkbox_SalesPromo.Checked = False
        End If
        If No_Checkbox_SalesPromo.Checked = True Then
            No_Checkbox_SalesPromo.Checked = False
        End If
        Summary_Txt_Box.Text = ""
        Total_Sales_Quote_Txt_Box.Text = ""
    End Sub
End Class