Option Explicit

Sub main()

Call fill_colors
Call fill_color_codes

End Sub
Sub fill_colors()
    'this fills 10 cells with randomly selected colors
    
    Dim i As Integer, t As Integer
    
    For t = 1 To 52
        For i = 1 To 10
            Range("A1").Value = ""
            Range("A2").Offset(i - 1, t - 1).Value = "Project " & CStr(i)
            Range("B2").Offset(i - 1, t - 1).Interior.Color = choose_color()
        Next i
    Next t
    
End Sub

Function choose_color()

Dim rand_num As Double

    rand_num = WorksheetFunction.RandBetween(1, 4)
    If rand_num = 1 Then
        choose_color = 255
        Exit Function
    ElseIf rand_num = 2 Then
        choose_color = 49407
        Exit Function
    ElseIf rand_num = 3 Then
        choose_color = 12611584
        Exit Function
    ElseIf rand_num = 4 Then
        choose_color = 5287936
        Exit Function
    End If
    
        
End Function

Sub fill_sheets()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim N As Integer, i As Integer, t_ As Integer, t As Integer
Dim wb As Workbook

t = 52
N = 100

For t_ = 1 To t

    For i = 1 To N
        
        'open new workbook
        Set wb = Workbooks.Add
        wb.Activate
    
        'fill colors
        Call fill_colors
        
        'save workbook
        wb.SaveAs ("C:\Users\n2\Documents\Code\VBA - Projects\project" & CStr(i) & "_t" & CStr(t_))
        
        'close workbook
        wb.Close
        Set wb = Nothing
        
    Next i

Next t_

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub fill_color_codes()

Dim rng As Range, icell As Range


Set rng = Application.InputBox("Select range", Type:=8)
For Each icell In rng
    icell.Value = icell.Interior.Color
Next

End Sub
