Attribute VB_Name = "Module1"
Sub New_Invoice_Sheet()
'
' New_Invoice_Sheet Button Macro
'
' Workout Protection: Worksheet.Protect "Password", UserInterfaceOnly := True
 
    Dim x As Integer
    Dim Waiting As Boolean
    Dim DueDate
    Dim Title
    Dim CONTRACTNO
    
    CONTRACTNO = Mid(Application.ActiveWorkbook.Path, InStr(Application.ActiveWorkbook.Path, "Contracts\") + 10, 5)
    Title = "New Payment Sheet"
    
' First Payment?

    If Range("H5") = "YY###" Then
        NewSheet.Show
    Range("H5") = CONTRACTNO
    Range("H7") = "=XLOOKUP(H5,Info!P4:P186,Info!Q4:Q186,""CONTRACT NOT FOUND"")"
    End If

' Troubleshoot Sheet
    
    
    If ActiveSheet.Name <> CStr(Format(Range("H13"), "000")) Then
        ActiveSheet.Name = Format(Range("H13"), "000")
        Application.Goto Range("H13"), Scroll:=True
        MsgBox "Warning Sheet Name Was Not Equal to Payment No.", vbCritical, Title
        Application.Goto Range("A1"), Scroll:=True
    End If
    
' Prepare New Sheet Name & Payment No.

    x = CStr(CInt(ActiveSheet.Name) + 1)
    ActiveSheet.Copy Before:=ActiveSheet
    On Error Resume Next
    ActiveSheet.Name = Format(x, "000")
    
    
' Update Amount(£), Payment No., Previously Certified, Payment Bought Forward and Date

    Range("E48").Formula = "=SUM(J22:J47)"
    
    Range("H13").Value = x

    Range("J62").Value = Range("J60").Value
    Range("J20").Value = Range("J49").Value
    
    Range("C58") = Date
    
'Automatic Item Number

    Worksheets("Info").Range("A22:A47").Copy
        ActiveWorkbook.ActiveSheet.Range("A22").PasteSpecial
    Range("F22:F47").ClearContents

' Input VAT, App No. and Invoice No.
    
    Dim answer As Integer
'    Dim InvoiceBook As Workbook
    
    Dim NextFriday As Date
    If Range(D2) = "F-122" Then
    Else
        answer = MsgBox("20% Reverse VAT Charge?", vbQuestion + vbYesNo + vbDefaultButton1, Title)

        If answer = vbYes Then
            Worksheets("Info").Range("H4:L8").Copy
            ActiveWorkbook.ActiveSheet.Range("F63").PasteSpecial
        ElseIf answer = vbNo Then
            Worksheets("Info").Range("B4:F8").Copy
            ActiveWorkbook.ActiveSheet.Range("F63").PasteSpecial
        End If
    End If
    
    NextFriday = Date + 8 - Weekday(Date, vbFriday)
    
    Range("C48").Value = InputBox("Input S/C Application No" & vbCrLf & vbCrLf & "(Previous: " & Range("C48") & ")", Title)
    Range("H48") = CDate(InputBox("Input Date of Invoice", Title))
    
    Waiting = True
    
    While Waiting
    
        DueDate = CDate(InputBox("Input Payment Due Date" & vbCrLf & vbCrLf & "(Nearest Friday: " & NextFriday & ")", Title))
    
        If DueDate < Date Then
        MsgBox "Date Must Be After Todays Date", , Title
    
        ElseIf Weekday(DueDate) <> 6 Then
        MsgBox "Date Must Be On A Friday", , Title
        
        Else
        Waiting = False
        
        End If
        
    Wend
    
    Range("C64") = DueDate
    
' Manual User Input Information & Troubleshooting

    MsgBox "Input Quantities and Amount(£) then Print", , Title
    
    ' Check Tax
    
    If Range("H15") = 0.2 Or Range("H15") = 0.3 Then
        If Range("H56") <> Range("J54") Then
            Application.Goto Range("H56"), Scroll:=True
            MsgBox "Warning Tax Auto Corrected Please Check", vbCritical, Title
            Application.Goto Range("A1"), Scroll:=True
            Range("H56") = Range("J54")
        End If
    End If
    
    

End Sub

