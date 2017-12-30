Option Strict On

Public Class TipCalculator
    '=============================================================================================================
    ' Author:	Holly Eaton
    ' 
    ' Program:	Tip and Split Check Calculator
    ' 
    ' Dev Env:	Visual Studio
    ' 
    ' Description:
    '   Purpose:    Project that will determine how much a 15%, 20% or 25% tip is and calculate a new bill total.
    '               The calculator will also split the check between a specified number of diners.
    '
    ' 	Input:      The amount of the bill. The number of people paying the bill. The user must then click the
    '                desired calculate button for a "Standard Tip" 15%, "Great Tip" 20% or an "Amazing Tip" 25%. 
    '
    ' 	Process:    Validates the users input for numerical characters only; presents a message asking for numbers
    '               only when user enters something else.
    '               Calculates the desired tip, new total and splits the check among the specified number of diners.
    '
    ' 	Output:     Textual information for the user displaying the tip amount, total bill and split for each diner.
    ' 
    '===============================================================================================================
    '   Declared Constants:
    '	SngSTANDARD_TIP
    '   SngGREAT_TIP
    '   SngAMAZING_TIP
    '	
    '===============================================================================================================
    '   Variables for user entered data:
    '   sngBill
    '   sngNumPaying
    '	Note: (dim and cast both directly to single then no type casting necessary in calculations)
    '	Example: Dim sngBill As Single = Csng(txtBill.Text)
    '	
    '===============================================================================================================
    '	Variables for calculated values:
    '	sngTipAmount
    '	sngNewBillTotal
    '	sngEach
    '	
    '================================================================================================================
    '	Calculations in pseudocode
    '	Set option strict on at top 
    '	TipAmount = (Bill * STANDARD_TIP)
    '	NewBillTotal = (Bill * STANDARD_TIP) + (Bill)
    '   Each = (NewBillTotal / NumPaying)
    '  
    '	same code for great tip and amazing tip just different constant for calculation
    '
    '==================================================================================================================
    '==================================================================================================================
    '==================================================================================================================

    'Declared Constants:
    Const SngSTANDARD_TIP As Single = 0.15
    Const SngGREAT_TIP As Single = 0.2
    Const sngAMAZING_TIP As Single = 0.25

    Private Sub btnCalcStandTip_Click(sender As Object, e As EventArgs) Handles btnCalcStandTip.Click
        Dim sngTipAmount As Single
        Dim sngNewBillTotal As Single
        Dim sngEach As Single

        'Variables for user entered data:
        Try
            Dim sngBill As Single = CSng(txtBill.Text)
            Try
                Dim sngNumPaying As Single = CSng(txtNumPay.Text)

                'Calculations for Tip Amount and New Bill Total
                sngTipAmount = (sngBill * SngSTANDARD_TIP)
                sngNewBillTotal = (sngBill * SngSTANDARD_TIP) + (sngBill)
                sngEach = (sngNewBillTotal / sngNumPaying)

                'format and output the results
                lblTipAmount.Text = sngTipAmount.ToString("c")
                lblNewBillTotal.Text = sngNewBillTotal.ToString("c")
                lblPercent.Text = SngSTANDARD_TIP.ToString("p")
                lblEach.Text = sngEach.ToString("C")

            Catch ex As Exception
                'What to do if user data entered into txtNumPaying is invalid and cannot be cast to single.
                'Tell user what to enter
                MessageBox.Show("Please enter the number of people paying using only numerical characters.")

                'reset input area
                txtNumPay.Text = "0"

                'put insertion point inside of Bill textbox
                txtNumPay.Focus()
            End Try
            
        Catch ex As Exception
            'What to do if user data entered into txtBill is invalid and cannot be cast to single.
            'Tell user what to enter
            MessageBox.Show("Please enter the amount of the bill using only numerical characters.")

            'reset input area
            txtBill.Text = "0"

            'put insertion point inside of Bill textbox
            txtBill.Focus()
        End Try
       

    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        'Clear the textboxes and labels
        txtBill.Text = "0"
        lblTipAmount.Text = String.Empty
        lblNewBillTotal.Text = String.Empty
        txtNumPay.Text = "1"
        lblEach.Text = String.Empty

        'Give the focus to txtBill
        txtBill.Focus()

    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Display Messagebox
        MessageBox.Show("Thanks for using the Tip and Split Check Calculator")

        'Close the form
        Me.Close()
    End Sub

    Private Sub btnCalcGreatTip_Click(sender As Object, e As EventArgs) Handles btnCalcGreatTip.Click
        Dim sngTipAmount As Single
        Dim sngNewBillTotal As Single
        Dim sngEach As Single

        'Variables for user entered data:
        Try
            Dim sngBill As Single = CSng(txtBill.Text)
            Try
                Dim sngNumPaying As Single = CSng(txtNumPay.Text)

                'Calculations for Tip Amount and New Bill Total
                sngTipAmount = (sngBill * SngGREAT_TIP)
                sngNewBillTotal = (sngBill * SngGREAT_TIP) + (sngBill)
                sngEach = (sngNewBillTotal / sngNumPaying)

                'format and output the results
                lblTipAmount.Text = sngTipAmount.ToString("c")
                lblNewBillTotal.Text = sngNewBillTotal.ToString("c")
                lblPercent.Text = SngGREAT_TIP.ToString("p")
                lblEach.Text = sngEach.ToString("C")

            Catch ex As Exception
                'What to do if user data entered into txtNumPaying is invalid and cannot be cast to single.
                'Tell user what to enter
                MessageBox.Show("Please enter the number of people paying using only numerical characters.")

                'reset input area
                txtNumPay.Text = "0"

                'put insertion point inside of Bill textbox
                txtNumPay.Focus()
            End Try

        Catch ex As Exception
            'What to do if user data entered into txtBill is invalid and cannot be cast to single.
            'Tell user what to enter
            MessageBox.Show("Please enter the amount of the bill using only numerical characters.")

            'reset input area
            txtBill.Text = "0"

            'put insertion point inside of Bill textbox
            txtBill.Focus()
        End Try

    End Sub

    Private Sub btnCalcAmazingTip_Click(sender As Object, e As EventArgs) Handles btnCalcAmazingTip.Click

        Dim sngTipAmount As Single
        Dim sngNewBillTotal As Single
        Dim sngEach As Single

        'Variables for user entered data:
        Try
            Dim sngBill As Single = CSng(txtBill.Text)
            Try
                Dim sngNumPaying As Single = CSng(txtNumPay.Text)

                'Calculations for Tip Amount and New Bill Total
                sngTipAmount = (sngBill * sngAMAZING_TIP)
                sngNewBillTotal = (sngBill * sngAMAZING_TIP) + (sngBill)
                sngEach = (sngNewBillTotal / sngNumPaying)

                'format and output the results
                lblTipAmount.Text = sngTipAmount.ToString("c")
                lblNewBillTotal.Text = sngNewBillTotal.ToString("c")
                lblPercent.Text = sngAMAZING_TIP.ToString("p")
                lblEach.Text = sngEach.ToString("C")

            Catch ex As Exception
                'What to do if user data entered into txtNumPaying is invalid and cannot be cast to single.
                'Tell user what to enter
                MessageBox.Show("Please enter the number of people paying using only numerical characters.")

                'reset input area
                txtNumPay.Text = "0"

                'put insertion point inside of Bill textbox
                txtNumPay.Focus()
            End Try
        Catch ex As Exception
            'What to do if user data entered into txtBill is invalid and cannot be cast to single.
            'Tell user what to enter
            MessageBox.Show("Please enter the amount of the bill using only numerical characters.")

            'reset input area
            txtBill.Text = "0"

            'put insertion point inside of Bill textbox
            txtBill.Focus()
        End Try

    End Sub

End Class
