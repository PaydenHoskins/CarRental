'Payden Hoskins
'Rcet 2265
'Spring 2025
'Car Rental Form
'https://github.com/PaydenHoskins/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary

Public Class RentalForm
    'Calculates the total owed
    Sub CalculateCharge()
        Dim milesDriven As Integer
        Dim beginOdometer As Integer
        Dim endingOdometer As Integer
        Dim mileCharge As Double
        Dim discount As Double
        Dim daysRented As Integer
        Dim totalCharge As Double
        Dim totalDiscount As Double
        endingOdometer = CInt(EndOdometerTextBox.Text)
        beginOdometer = CInt(BeginOdometerTextBox.Text)
        milesDriven = (endingOdometer - beginOdometer)
        If AAAcheckbox.Checked = True And Seniorcheckbox.Checked = True Then
            discount = 0.08
        ElseIf AAAcheckbox.Checked = True And Seniorcheckbox.Checked = False Then
            discount = 0.05
        ElseIf AAAcheckbox.Checked = False And Seniorcheckbox.Checked = True Then
            discount = 0.03
        ElseIf AAAcheckbox.Checked = False And Seniorcheckbox.Checked = False Then
            discount = 0.00
        End If
        If MilesradioButton.Checked = True Then
            If milesDriven < 0 Then
                MsgBox("You cannot drive less than 0 miles.")
                BeginOdometerTextBox.Focus()
                EndOdometerTextBox.Focus()
            ElseIf milesDriven >= 0 Then
                TotalMilesTextBox.Text = ($"{CStr(milesDriven)}mi")
                Select Case milesDriven
                    Case 0 To 200
                        mileCharge = 0
                    Case 201 To 500
                        mileCharge = 0.12
                    Case >= 500
                        mileCharge = 0.1
                End Select
                MileageChargeTextBox.Text = (milesDriven * mileCharge).ToString("C")
                daysRented = CInt(DaysTextBox.Text)
                DayChargeTextBox.Text = (daysRented * 15).ToString("C")
                daysRented = CInt(DaysTextBox.Text)
                DayChargeTextBox.Text = (daysRented * 15).ToString("C")
                totalCharge = (CInt(DayChargeTextBox.Text) + CInt(MileageChargeTextBox.Text))
                totalDiscount = (totalCharge * discount)
                TotalDiscountTextBox.Text = ($"-{totalDiscount.ToString("C")}")
                TotalChargeTextBox.Text = (totalCharge - totalDiscount).ToString("C")
                TotalCharges(True, (totalCharge - totalDiscount))
                TotalMilesDriven(True, milesDriven)
            End If
        ElseIf MilesradioButton.Checked = False Then
            If milesDriven < 0 Then
                MsgBox("You cannot drive less than 0 miles.")
                BeginOdometerTextBox.Focus()
                EndOdometerTextBox.Focus()
            ElseIf milesDriven >= 0 Then
                TotalMilesTextBox.Text = ($"{CStr(milesDriven)}mi")
                Select Case milesDriven
                    Case 0 To 321
                        mileCharge = 0
                    Case 322 To 804
                        mileCharge = 0.12
                    Case >= 805
                        mileCharge = 0.1
                End Select
                MileageChargeTextBox.Text = (milesDriven * mileCharge).ToString("C")
                daysRented = CInt(DaysTextBox.Text)
                DayChargeTextBox.Text = (daysRented * 15).ToString("C")
                totalCharge = (CInt(DayChargeTextBox.Text) + CInt(MileageChargeTextBox.Text))
                totalDiscount = (totalCharge * discount)
                TotalDiscountTextBox.Text = ($"-{totalDiscount.ToString("C")}")
                TotalChargeTextBox.Text = (totalCharge - totalDiscount).ToString("C")
                TotalCharges(True, (totalCharge - totalDiscount))
                TotalMilesDriven(True, milesDriven)
            End If
        End If

    End Sub
    'total customer count
    Function CustomerCount(Optional yes As Boolean = False) As Integer
        Static customers As Integer = 0
        If yes Then
            customers += 1
        End If
        Return customers
    End Function
    'counter for total miles
    Function TotalMilesDriven(Optional yes As Boolean = False, Optional miles As Double = 0) As Double
        Static milesTotal As Double = 0
        If yes Then
            milesTotal += miles
        End If
        Return milesTotal
    End Function
    'counter for total amount made
    Function TotalCharges(Optional yes As Boolean = False, Optional charge As Double = 0) As Double
        Static totalCharge As Double = 0
        If yes Then
            totalCharge += charge
        End If
        Return totalCharge
    End Function
    'Event handlers ***************************************
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click, CalculateToolStripMenuItem.Click
        Dim addressValid As Boolean
        Dim cityValid As Boolean
        Dim nameValid As Boolean
        Dim stateValid As Boolean
        Dim zipcodeValid As Boolean
        Dim beginOdometer As Boolean
        Dim endingOdometer As Boolean
        Dim dayChargeValid As Boolean

        Dim message As String

        Try
            If NameTextBox.Text <> "" Then
                nameValid = True
            ElseIf NameTextBox.Text = "" Then
                nameValid = False
                NameTextBox.Focus()
                message &= "Please enter a valid Name." & vbNewLine
            End If
            If AddressTextBox.Text <> "" Then
                addressValid = True
            ElseIf AddressTextBox.Text = "" Then
                addressValid = False
                AddressTextBox.Focus()
                message &= "Please enter a valid Address." & vbNewLine
            End If
            If CityTextBox.Text <> "" Then
                cityValid = True
            ElseIf CityTextBox.Text = "" Then
                cityValid = False
                CityTextBox.Focus()
                message &= "Please enter a valid City." & vbNewLine
            End If
            If StateTextBox.Text <> "" Then
                stateValid = True
            ElseIf StateTextBox.Text = "" Then
                stateValid = False
                StateTextBox.Focus()
                message &= "Please enter a valid State." & vbNewLine
            End If
            If ZipCodeLabel.Text <> "" Then
                zipcodeValid = True
            ElseIf ZipCodeTextBox.Text = "" Then
                zipcodeValid = False
                ZipCodeLabel.Focus()
                message &= "Please enter a valid ZipCode." & vbNewLine
            End If
            If IsNumeric(BeginOdometerTextBox.Text) = True Then
                beginOdometer = True
            ElseIf IsNumeric(BeginOdometerTextBox.Text) = False Then
                beginOdometer = False
                BeginOdometerTextBox.Focus()
                message &= "Please enter a valid Beginning Odometer Reading." & vbNewLine
            End If
            If IsNumeric(EndOdometerTextBox.Text) = True Then
                endingOdometer = True
            ElseIf IsNumeric(EndOdometerTextBox.Text) = False Then
                endingOdometer = False
                EndOdometerTextBox.Focus()
                message &= "Please enter a valid Ending Odometer Reading." & vbNewLine
            End If
            If IsNumeric(DaysTextBox.Text) = True Then
                Select Case CInt(DaysTextBox.Text)
                    Case 1 To 45
                        dayChargeValid = True
                    Case Else
                        dayChargeValid = False
                        DaysTextBox.Focus()
                        DaysTextBox.Clear()
                        message &= "Please enter a valid number of days from 1 to 45." & vbNewLine
                End Select
            ElseIf IsNumeric(DaysTextBox.Text) = False Then
                dayChargeValid = False
                DaysTextBox.Focus()
                DaysTextBox.Clear()
                message &= "Please enter a valid number of days from 1 to 45." & vbNewLine
            End If
            If addressValid = False Or cityValid = False Or nameValid = False Or stateValid = False Or zipcodeValid = False Or beginOdometer = False Or endingOdometer = False Or dayChargeValid = False Then
                MsgBox(message, MsgBoxStyle.Exclamation)
            ElseIf addressValid = True And cityValid = True And nameValid = True And stateValid = True And zipcodeValid = True And beginOdometer = True And endingOdometer = True And dayChargeValid = True Then
                CalculateCharge()
                CustomerCount(True)

            End If
        Catch ex As Exception
            MsgBox("Something Just Broke!")
        End Try
    End Sub
    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click, ExitToolStripMenuItem1.Click
        Dim msg = "Are you sure you wanna leave???"
        Dim style = MsgBoxStyle.YesNo
        Dim title = "Leave"
        Dim messageBox = MsgBox(msg, style, title)

        If messageBox = MsgBoxResult.Yes Then
            Me.Close()
        ElseIf messageBox = MsgBoxResult.No Then
        End If
    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click, ClearToolStripMenuItem1.Click
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()
        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()
    End Sub
    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click, SummaryToolStripMenuItem1.Click
        Dim totalCustomers As Integer
        Dim totalMiles As Double
        Dim totalCharge As Double
        totalCustomers = CustomerCount()
        totalMiles = TotalMilesDriven()
        totalCharge = TotalCharges()
        MsgBox($"{totalCustomers}{vbNewLine} {totalMiles}{vbNewLine} {totalCharge.ToString("C")}{vbNewLine}")
    End Sub
End Class