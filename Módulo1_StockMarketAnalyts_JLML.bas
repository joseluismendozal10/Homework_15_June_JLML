Attribute VB_Name = "Módulo1"
Sub StockMarketAnalyst()

    Call performanceOff
    Call allWorksheets
    
End Sub
Function allWorksheets()

    Dim f_activeSheet As Worksheet
    For Each f_activeSheet In Worksheets
    
    f_activeSheet.Select
    
    Call Ticker
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent change"
    Range("L1").Value = "Total Stock Volume"
    
    Next
      
End Function
Function Ticker()


    Call performanceOn
    Dim CurrentTicker1, CurrentTicker2 As String
    Dim CurrentVolume, TotalVolume As Double
    Dim CurrentRow As Integer
    Dim LoopCounter As Integer
    Dim Lastrow As Long
    
    TotalVolume = 0
    CurrentRow = 1
    LoopCounter = -1
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To Lastrow
        CurrentTicker1 = Cells(i, 1).Value2
        CurrentTicker2 = Cells(i + 1, 1).Value2
        CurrentVolume = Cells(i, 7).Value
        TotalVolume = TotalVolume + CurrentVolume
        LoopCounter = LoopCounter + 1
        
        
    If (CurrentTicker2 <> CurrentTicker1) Then
        CurrentRow = CurrentRow + 1
        Cells(CurrentRow, 9).Value = CurrentTicker1
        Cells(CurrentRow, 12).Value = TotalVolume
        
    Call StockVolume(CurrentRow, i, LoopCounter)
    
    TotalVolume = 0
    LoopCounter = -1
    
    End If
    Next i
    
    Call performanceOff
    
End Function
Function StockVolume(f_CurrentRow, f_i, F_LoopCounter)

    Dim OpenPrice, ClosePrice As Double
    
    OpenPrice = Cells(f_i - F_LoopCounter, 3).Value2
    ClosePrice = Cells(f_i, 6).Value2
    Cells(f_CurrentRow, 10).Value = ClosePrice - OpenPrice
    Cells(f_CurrentRow, 10).NumberFormat = "0.000000"
    
    If (Cells(f_CurrentRow, 10).Value) >= 0 Then
        Cells(f_CurrentRow, 10).Interior.ColorIndex = 4
    Else
        Cells(f_CurrentRow, 10).Interior.ColorIndex = 3
    End If
    
    If (OpenPrice = 0 Or ClosePrice = 0) Then
        Cells(f_CurrentRow, 11).Value = 0
    Else
        Cells(f_CurrentRow, 11).Value = ClosePrice / OpenPrice - 1
    End If
        Cells(f_CurrentRow, 11).NumberFormat = "0.00%"
        
End Function
Function performanceOn()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
End Function
Function performanceOff()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Function



