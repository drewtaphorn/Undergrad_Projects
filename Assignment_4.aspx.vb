Public Class Assignment_3
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Public Shared TotalGrossPay, TotalTax As Decimal

    Protected Sub Tax_Total_Button_Click(sender As Object, e As EventArgs) Handles Tax_Total_Button.Click
        Tax_Total_Txt_Box.Text = "Tax Total for " & Location_Drop_Down_List.Text & ": " & TotalTax.ToString("C2")
    End Sub

    Protected Sub Gross_Total_Button_Click(sender As Object, e As EventArgs) Handles Gross_Total_Button.Click
        Gross_Total_Txt_Box.Text = "Gross Pay Total for " & Location_Drop_Down_List.Text & ": " & TotalGrossPay.ToString("C2")
    End Sub

    Protected Sub Calculate_Button_Click(sender As Object, e As EventArgs) Handles Calculate_Button.Click

        Dim Hours, OvertimeHours, WeekendPay, OvertimePay, HourlyPay As Decimal
        Dim HourRate, TaxRate, Tax, GrossPay, NetPay As Decimal
        Dim MondayHours, TuesdayHours, WednesdayHours, ThursdayHours, FridayHours, SaturdayHours As Decimal
        Dim Str As String = ""


        If IsNumeric(Monday_Hrs_Txt_Box.Text) Then
            MondayHours = Convert.ToDecimal(Monday_Hrs_Txt_Box.Text)
        Else
            Response.Write("Please type a number in the Monday Hours Text Box")
            Exit Sub
        End If
        If IsNumeric(Tuesday_Hrs_Txt_Box.Text) Then
            TuesdayHours = Convert.ToDecimal(Tuesday_Hrs_Txt_Box.Text)
        Else
            Response.Write("Please type a number in the Tuesday Hours Text Box")
            Exit Sub
        End If
        If IsNumeric(Wednesday_Hrs_Txt_Box.Text) Then
            WednesdayHours = Convert.ToDecimal(Wednesday_Hrs_Txt_Box.Text)
        Else
            Response.Write("Please type a number in the Wednesday Hours Text Box")
            Exit Sub
        End If
        If IsNumeric(Thursday_Hrs_Txt_Box.Text) Then
            ThursdayHours = Convert.ToDecimal(Thursday_Hrs_Txt_Box.Text)
        Else
            Response.Write("Please type a number in the Thursday Hours Text Box")
            Exit Sub
        End If
        If IsNumeric(Friday_Hrs_Txt_Box.Text) Then
            FridayHours = Convert.ToDecimal(Friday_Hrs_Txt_Box.Text)
        Else
            Response.Write("Please type a number in the Friday Hours Text Box")
            Exit Sub
        End If
        If IsNumeric(Saturday_Hrs_Txt_Box.Text) Then
            SaturdayHours = Convert.ToDecimal(Saturday_Hrs_Txt_Box.Text)
        Else
            Response.Write("Please type a number in the Saturday Hours Text Box")
            Exit Sub
        End If

        If IsNumeric(Saturday_Hrs_Txt_Box.Text) > 8 Then
            Response.Write("Saturday Hours cannot exceed over 8 hours.")
        End If

        For Each li As ListItem In Location_Drop_Down_List.Items
            If li.Selected Then
                HourRate += li.Value
                Exit For
            End If
        Next

        Hours = MondayHours + TuesdayHours + WednesdayHours + ThursdayHours + FridayHours
        If Hours > 40 Then
            OvertimeHours = Hours - 40
        End If

        HourlyPay = Hours * HourRate
        OvertimePay = OvertimeHours * HourRate * 1.5
        WeekendPay = SaturdayHours * HourRate * 2
        GrossPay = HourlyPay + OvertimePay + WeekendPay

        If Location_Drop_Down_List.SelectedIndex = 1 Then
            TaxRate = 0.0495 + 0.22 'State + Federal
        ElseIf Location_Drop_Down_List.SelectedIndex = 2 Then
            TaxRate = 0.0323 + 0.22 'State + Federal
        ElseIf Location_Drop_Down_List.SelectedIndex = 3 Then
            TaxRate = 0.054 + 0.22 'State + Federal
        End If

        Tax = GrossPay * TaxRate
        NetPay = GrossPay - Tax

        TotalTax += Tax
        TotalGrossPay += GrossPay

        Calculate_Txt_Box.Text = "Employee Name: " & Employee_Name_Txt_Box.Text & vbNewLine & "Hours worked: " & Hours & vbNewLine & "Overtime Hours: " & OvertimeHours & vbNewLine & "Weekend Hours: " & SaturdayHours & vbNewLine & "Gross Pay: " & GrossPay.ToString("C2") & vbNewLine & "Taxes: " & Tax.ToString("C2") & vbNewLine & "Net Pay: " & NetPay.ToString("C2")

    End Sub
    Protected Sub Clear_Button_Click(sender As Object, e As EventArgs) Handles Clear_Button.Click
        Location_Drop_Down_List.SelectedIndex = 0
        Employee_Name_Txt_Box.Text = ""
        Monday_Hrs_Txt_Box.Text = ""
        Tuesday_Hrs_Txt_Box.Text = ""
        Wednesday_Hrs_Txt_Box.Text = ""
        Thursday_Hrs_Txt_Box.Text = ""
        Friday_Hrs_Txt_Box.Text = ""
        Saturday_Hrs_Txt_Box.Text = ""
        Calculate_Txt_Box.Text = ""
    End Sub
End Class