Option Explicit

Sub fill_colors()
    'this fills 10 cells with randomly selected colors
    
    Dim i As Integer
    
    For i = 1 To 10
        Range("A1").Offset(i - 1, 0).Value = "Project " & CStr(i)
        Range("B1").Offset(i - 1, 0).Interior.Color = choose_color()
    Next i
        
End Sub

Function choose_color()

Dim rand_num As Double
    'this chooses a random cell color
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
'this randomly generates 100 project workstreams' RAG statuses for 52 weeks
'we'll use these RAG statuses to demonstrate the project completion forecasting algorithm
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim N As Integer, i As Integer, t_ As Integer, T As Integer
Dim wb As Workbook

T = 52
N = 100

For t_ = 1 To T

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
