' Simple VBScript Calculator
' Save as calculator.vbs and run by double-clicking or: cscript //nologo calculator.vbs

Option Explicit

Dim choice, a, b, result

Do
    choice = InputBox( _
        "Simple VBScript Calculator" & vbCrLf & vbCrLf & _
        "1) Add" & vbCrLf & _
        "2) Subtract" & vbCrLf & _
        "3) Multiply" & vbCrLf & _
        "4) Divide" & vbCrLf & _
        "5) Exit", _
        "Calculator")

    If choice = "" Then Exit Do ' user cancelled

    Select Case Trim(choice)
        Case "1", "Add", "add", "+"
            a = GetNumber("Enter the first number:")
            If IsEmpty(a) Then Exit Do
            b = GetNumber("Enter the second number:")
            If IsEmpty(b) Then Exit Do
            result = a + b
            MsgBox a & " + " & b & " = " & result, vbInformation, "Result"
        Case "2", "Subtract", "subtract", "-"
            a = GetNumber("Enter the first number:")
            If IsEmpty(a) Then Exit Do
            b = GetNumber("Enter the second number:")
            If IsEmpty(b) Then Exit Do
            result = a - b
            MsgBox a & " - " & b & " = " & result, vbInformation, "Result"
        Case "3", "Multiply", "multiply", "*", "x", "X"
            a = GetNumber("Enter the first number:")
            If IsEmpty(a) Then Exit Do
            b = GetNumber("Enter the second number:")
            If IsEmpty(b) Then Exit Do
            result = a * b
            MsgBox a & " * " & b & " = " & result, vbInformation, "Result"
        Case "4", "Divide", "divide", "/"
            a = GetNumber("Enter the first number:")
            If IsEmpty(a) Then Exit Do
            b = GetNumber("Enter the second number:")
            If IsEmpty(b) Then Exit Do
            If b = 0 Then
                MsgBox "Error: Division by zero is not allowed.", vbCritical, "Error"
            Else
                result = a / b
                MsgBox a & " / " & b & " = " & result, vbInformation, "Result"
            End If
        Case "5", "Exit", "exit", "q", "Q"
            Exit Do
        Case Else
            MsgBox "Invalid choice. Enter 1-5 (or operation name).", vbExclamation, "Invalid"
    End Select
Loop

MsgBox "Goodbye!", vbInformation, "Calculator"


' Helper function to repeatedly prompt until numeric input or cancel
Function GetNumber(prompt)
    Dim s
    s = InputBox(prompt)
    If s = "" Then
        GetNumber = Empty ' propagate cancel
        Exit Function
    End If

    If IsNumeric(s) Then
        GetNumber = CDbl(s)
    Else
        MsgBox "That is not a number. Please try again or Cancel to quit.", vbExclamation, "Invalid input"
        GetNumber = GetNumber(prompt) ' retry
    End If
End Function