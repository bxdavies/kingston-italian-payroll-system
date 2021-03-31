Imports System
Module Main
    Dim MenuOption, EmployeeName, EmployeeAddressLine2, EmployeeAddressTown, EmployeeAddressPostCode, EmployeePhoneNumber, windowtext As String
    Dim EmployeeAge, EmployeeAddressLine1, TotalHoursWorked, NormalHoursWorked, OvertimeHoursWorked, TaxRate As Integer
    Dim EmployeePay, NormalPay, OvertimePay, GrossPay, IncomeTax, NationalInsurance, NetPay As Decimal
    Dim EmployeeType As Boolean
    Sub Main()
        'Displays a menu to the user upon opening the program'
        windowtext = "Kingston Italian Payroll System"
        Console.Title = windowtext
        'Changes the background colour and text colour of the console window'
        While True
            Console.BackgroundColor = ConsoleColor.DarkBlue
            Console.ForegroundColor = ConsoleColor.White
            Console.Clear()
            Console.WriteLine("Kingston Italian Payroll System")
            Console.WriteLine("© Cameron Bools 2020")
            Console.WriteLine("-------------------------------")
            Console.WriteLine("1. Input Employee Details")
            Console.WriteLine("2. Generate Payslip")
            Console.WriteLine("3. Help ")
            Console.WriteLine("4. Exit")
            Console.WriteLine("-------------------------------")
            Console.WriteLine("")
            Console.Write("Enter Choice: ")
            MenuOption = Console.ReadLine()
            If IsNumeric(MenuOption) Then
                If MenuOption = "1" Then
                    'Displays menu to input details about the employee'
                    Console.Clear()
                    inputdetails()
                ElseIf MenuOption = "2" Then
                    'Displays the pay slip and works out the maths'
                    Console.Clear()
                    Calculations()
                    payslip()
                ElseIf MenuOption = "3" Then
                    'Displays a help menu'
                    Console.Clear()
                    help()
                ElseIf MenuOption = "4" Then
                    'Exits the program'
                    End
                Else
                    Console.WriteLine("Invalid Option!")
                End If
            Else
                Console.Clear()
                Console.WriteLine("Invalid Input!")
            End If
        End While
    End Sub
    Sub inputdetails()
        'Asks the user to enter the relevant information. If the input is invalid, an exception will be triggered'
        While True
            Console.WriteLine(“Enter employee name: ”)
            EmployeeName = Console.ReadLine()
            If IsNumeric(EmployeeName) Then
                Console.WriteLine("Invalid input!")
                Continue While
            End If
            Exit While
        End While
        While True
            Dim strEmployeeAddressLine1 As String
            Console.WriteLine(“Enter employee house number: ”)
            strEmployeeAddressLine1 = Console.ReadLine()
            If Integer.TryParse(strEmployeeAddressLine1, EmployeeAddressLine1) = True Then
                Exit While
            ElseIf Integer.TryParse(strEmployeeAddressLine1, EmployeeAddressLine1) = False Then
                Console.WriteLine("Invalid input!")
            End If
        End While

        While True
            Console.WriteLine(“Enter employee street name: “)
            EmployeeAddressLine2 = Console.ReadLine()
            If IsNumeric(EmployeeAddressLine2) Then
                Console.WriteLine("Invalid input!")
                Continue While
            End If
            Exit While
        End While

        While True
            Console.WriteLine(“Enter employee post code: “)
            EmployeeAddressPostCode = Console.ReadLine()
            Exit While
        End While

        While True
            Dim strEmployeeType As String
            Console.WriteLine(“Is employee an apprentice? (YES or NO): “)
            strEmployeeType = Console.ReadLine()
            strEmployeeType = strEmployeeType.ToUpper()
            If strEmployeeType = "YES" Then
                EmployeeType = True
                Exit While
            ElseIf strEmployeeType = "NO" Then
                EmployeeType = False
                Exit While
            Else
                Console.WriteLine("Invalid Input!")
                Continue While
            End If
        End While

        'Only shows the prompts if the employee has worked overtime hours'
        While EmployeeAge < 16
            Try
                Console.WriteLine(“Enter employee age: “)
                EmployeeAge = Console.ReadLine()
                If IsNumeric(EmployeeAge) Then
                    Exit While
                Else
                    Console.WriteLine("Invalid Input!")
                    Continue While
                End If
            Catch ex As Exception
                Console.WriteLine("Invalid Input!")
            End Try

        End While
        While True
            Try
                Console.WriteLine(“Enter the total hours the employee has worked: “)
                TotalHoursWorked = Console.ReadLine()

                If IsNumeric(TotalHoursWorked) Then
                    Exit While
                Else
                    Console.WriteLine("Invalid Input!")
                    Continue While
                End If
            Catch ex As Exception
                Console.WriteLine("Invalid Input!")
            End Try

        End While

        While True
            Console.WriteLine(“Enter employee phone number: “)
            EmployeePhoneNumber = Console.ReadLine()
            If IsNumeric(EmployeePhoneNumber) Then
                Exit While
            Else
                Console.WriteLine("Invalid Input!")
                Continue While
            End If
        End While
        Console.WriteLine("Details complete! Press any key to return to main menu...")
        Console.ReadLine()
        Console.Clear()
    End Sub
    Sub Calculations()
        'Calculations for working out the pay for the employee'
        'If the employee has worked for more than 224 hours, then overtime is counted. Else overtime is set to 0 by default'
        If TotalHoursWorked > 224 Then
            OvertimeHoursWorked = TotalHoursWorked - 224
        Else
            OvertimeHoursWorked = 0
        End If
        NormalHoursWorked = TotalHoursWorked - OvertimeHoursWorked
        NormalPay = NormalHoursWorked * Pay()
        OvertimePay = OvertimeHoursWorked * Pay()
        GrossPay = NormalPay + OvertimePay
        IncomeTax = GrossPay * Tax()
        NationalInsurance = GrossPay * 0.1
        NetPay = GrossPay - (IncomeTax + NationalInsurance)
    End Sub
    Sub payslip()
        'Outputs the payslip with all of the information entered in the input information stage'
        'Line 127 calls the pay() function to work out the pay for the employee'
        'Line 132 calls the Tax() function to work out the tax'
        Console.WriteLine("---------------------------------------------------")
        Console.WriteLine("Name: " & EmployeeName)
        Console.WriteLine("House Number: " & EmployeeAddressLine1)
        Console.WriteLine("Street: " & EmployeeAddressLine2)
        Console.WriteLine("Post Code: " & EmployeeAddressPostCode)
        Console.WriteLine("Age: " & EmployeeAge)
        Console.WriteLine("Total Hours Worked: " & TotalHoursWorked)
        Console.WriteLine("Employee Phone Number: " & EmployeePhoneNumber)
        Console.WriteLine("Is employee an Apprentice?: " & EmployeeType)
        Console.WriteLine("Hours Worked: " & NormalHoursWorked & " hours")
        Console.WriteLine("Overtime: " & OvertimeHoursWorked & " hours")
        Console.WriteLine("Pay Rate: " & "£" & Pay())
        Console.WriteLine("Normal Pay: " & "£" & NormalPay)
        Console.WriteLine("Overtime Pay: " & "£" & OvertimePay)
        Console.WriteLine("Gross Pay: " & "£" & GrossPay)
        Console.WriteLine("Tax Rate: " & "£" & Tax() * 100)
        Console.WriteLine("Taxed: " & "£" & IncomeTax)
        Console.WriteLine("National Insurance: " & "£" & NationalInsurance)
        Console.WriteLine("Net Pay: " & "£" & NetPay)
        Console.WriteLine("---------------------------------------------------")
        Console.WriteLine("Press any key to return to main menu")
        Console.ReadLine()
        Console.Clear()
    End Sub
    Sub help()
        Console.Clear()
        Console.WriteLine("Help Menu")
        Console.WriteLine("---------------------------------------------------")
        Console.WriteLine("Main Menu Option Definitions")
        Console.WriteLine("")
        Console.WriteLine("1 - Takes you to the input details form. Details can be entered to generate a pay slip")
        Console.WriteLine("2 - Generates a pay slip. The details entered in the first option will be displayed here along with the pay.")
        Console.WriteLine("3 - Opens this help menu")
        Console.WriteLine("4 - Exits the program")
        Console.WriteLine("Entering the corresponding number and pressing enter will open the menu")
        Console.WriteLine("---------------------------------------------------")
        Console.WriteLine("Input Details Form")
        Console.WriteLine("")
        Console.WriteLine("Upon entering the menu, you will be presented with a field to fill out. Once filled out, the next field will be displayed")
        Console.WriteLine("")
        Console.WriteLine("Once all the fileds have been filled out, the details will be temporarily stored, so a pay slip can be generated. Only one pay slip for one employee can be generated at a time.")
        Console.WriteLine("---------------------------------------------------")
        Console.WriteLine("Pay Slip")
        Console.WriteLine("")
        Console.WriteLine("Upon opening the pay slip menu, you will be displayed the pay slip with all the informaion you entered in the previous menu. Fields such as pay, national insurance, hours worked etc will be calculated by the program automatically.")
        Console.WriteLine("---------------------------------------------------")
        Console.WriteLine("Press any key to return to the main menu...")
        Console.ReadLine()
        Console.Clear()
    End Sub
    Function Pay()
        'This function will work out the pay for the employee'
        Dim EmployeePay As Decimal
        'If the employee is an apprentice'
        If EmployeeType = False Then
            If EmployeeAge >= 21 Then
                EmployeePay = 6.31
                Return EmployeePay
            ElseIf EmployeeAge >= 18 And EmployeeAge <= 20 Then
                EmployeePay = 5.03
                Return EmployeePay
            ElseIf EmployeeAge >= 16 And EmployeeAge < 18 Then
                EmployeePay = 3.72
                Return EmployeePay
            Else
                EmployeePay = 0.00
                Return EmployeePay
            End If
            'If the employee is not an apprentice and is a worker'
        ElseIf EmployeeType = True Then
            If EmployeeAge > 19 Then
                EmployeePay = 2.68
                Return EmployeePay
            ElseIf EmployeeAge <= 19 Then
                EmployeePay = 1.78
                Return EmployeePay
            Else
                EmployeePay = 0.00
                Return EmployeePay
            End If
        Else
            EmployeePay = 0.00
            Return EmployeePay
        End If
    End Function
    Function Tax()
        'Function to work out the tax'
        Dim TaxRate As Decimal
        'If the employee earns less than £2667.50, then 0.2% tax is charged'
        If GrossPay < 2667.5 Then
            TaxRate = 0.2
            Return TaxRate
        ElseIf GrossPay > 2667.5 And GrossPay < 12500 Then
            'If the employee earns in between £2667.50 and £12,500, then 0.4% tax is charged'
            TaxRate = 0.4
            Return TaxRate
        ElseIf GrossPay > 12500 Then
            'If the employee earns more than £12,500 then 0.45% tax is charged'
            TaxRate = 0.45
            Return TaxRate
        End If
        Return TaxRate
    End Function
End Module
