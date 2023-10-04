# Movie-Rental-App
Visual Basic code for a Movie Rental App
  



Public Class Form1

    Private NumUsers As Integer = 0
    Dim RentalSum As Decimal

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles VidFormatGroupBox.Enter

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Const vhs = 10
        Const dvd = 15
        Const VhsNew = 12
        Const DvdNew = 18
        Const MemberDiscountRate = 0.1
        Dim TotalAmount As Decimal


        If VhsRadioButton.Checked = True Then
            If NewReleaseCheckbox.Checked = True Then
                If MembershipCheckbox.Checked = True Then
                    TotalAmount = VhsNew - (vhs * MemberDiscountRate)
                ElseIf MembershipCheckbox.Checked = False Then
                    TotalAmount = VhsNew
                End If

            ElseIf NewReleaseCheckbox.Checked = False Then
                If MembershipCheckbox.Checked = True Then
                    TotalAmount = vhs - (vhs * MemberDiscountRate)
                ElseIf MembershipCheckbox.Checked = False Then
                    TotalAmount = vhs
                End If
            End If


        ElseIf VhsRadioButton.Checked = False Then
            If NewReleaseCheckbox.Checked = True Then
                If MembershipCheckbox.Checked = True Then
                    TotalAmount = dvd - (dvd * MemberDiscountRate)
                ElseIf MembershipCheckbox.Checked = False Then
                    TotalAmount = DvdNew
                End If
            ElseIf NewReleaseCheckbox.Checked = False Then
                If MembershipCheckbox.Checked = True Then
                    TotalAmount = dvd - (dvd * MemberDiscountRate)
                ElseIf MembershipCheckbox.Checked = False Then
                    TotalAmount = dvd
                End If
            End If

        End If

        NumUsers = NumUsers + 1
        RentalSum += TotalAmount

        MessageBox.Show("Price of your order is " & TotalAmount.ToString("C2"))


    End Sub


    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        MovieTitleTextbox.Clear()
        MembershipCheckbox.Checked = False
        NewReleaseCheckbox.Checked = False
        VhsRadioButton.Checked = False
        DvdRadioButton.Checked = False
    End Sub

    Private Sub OrderCompleteButton_Click(sender As Object, e As EventArgs) Handles OrderCompleteButton.Click
        Dim completeforsure As String = "Do you want to complete your order?"
        Dim result As DialogResult = MessageBox.Show(completeforsure, "Quit?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = Windows.Forms.DialogResult.Yes Then
            MovieTitleTextbox.Clear()
            MembershipCheckbox.Checked = False
            NewReleaseCheckbox.Checked = False
            VhsRadioButton.Checked = False
            DvdRadioButton.Checked = False
        End If



    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        MessageBox.Show("Number of customers: " & NumUsers &
         ControlChars.NewLine & "Sum of rental amount is $" & RentalSum.ToString)

    End Sub
End Class

