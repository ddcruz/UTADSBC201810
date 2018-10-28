Attribute VB_Name = "Module1"
Private Sub CalculateYearlyMetricsPerTicker()
    Dim CurrentTicker As String
    Dim i As Integer
    i = 1
    
    Dim FirstOpen As Double
    Dim LastClose As Double

    For x = 2 To Cells(Rows.Count, 1).End(xlUp).Row + 1
        If CurrentTicker <> Cells(x, 1).Value Then
            If IsNumeric(Cells(x - 1, 6).Value) Then
                LastClose = Cells(x - 1, 6).Value
                
                'Yearly Change
                Cells(i, 10).Value = LastClose - FirstOpen
                If Cells(i, 10).Value < 0 Then
                    Cells(i, 10).Interior.Color = vbRed
                ElseIf Cells(i, 10).Value > 0 Then
                    Cells(i, 10).Interior.Color = vbGreen
                End If
                
                'Percent Change
                With Cells(i, 11)
                    If FirstOpen <> 0 Then
                        .Value = (LastClose / FirstOpen) - 1
                    End If
                    .Style = "Percent"
                    .NumberFormat = "0.00%"
                End With
                
            End If
            CurrentTicker = Cells(x, 1).Value

            FirstOpen = Cells(x, 3).Value
            i = i + 1
            Cells(i, 9).Value = CurrentTicker
        End If
        
        If Not IsEmpty(Cells(x, 7).Value) Then
            Cells(i, 12).Value = Cells(i, 12).Value + Cells(x, 7).Value
        End If
        
    Next x
End Sub

Private Sub CalculateOutliers()
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As LongLong
    
    For x = 2 To Cells(Rows.Count, 9).End(xlUp).Row
        If GreatestPercentIncrease < Cells(x, 11).Value Then
            GreatestPercentIncrease = Cells(x, 11).Value
            Cells(2, 16).Value = Cells(x, 9).Value
            With Cells(2, 17)
                .Value = GreatestPercentIncrease
                .Style = "Percent"
                .NumberFormat = "0.00%"
            End With
        End If
        
        If GreatestPercentDecrease > Cells(x, 11).Value Then
            GreatestPercentDecrease = Cells(x, 11).Value
            Cells(3, 16).Value = Cells(x, 9).Value
            With Cells(3, 17)
                .Value = GreatestPercentDecrease
                .Style = "Percent"
                .NumberFormat = "0.00%"
            End With
        End If
        
        If GreatestTotalVolume < Cells(x, 12).Value Then
            GreatestTotalVolume = Cells(x, 12).Value
            Cells(4, 16).Value = Cells(x, 9).Value
            Cells(4, 17).Value = GreatestTotalVolume
        End If
    Next x
End Sub

Private Sub ClearCalcs()
    'Clear the outliers calc
    Range("P2:Q4").Value = ""
    
    'Clear the stock calcs
    Range("I2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Selection.ClearFormats
    Range("I2").Select
    
End Sub

Private Sub NewHeaders()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest total volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
End Sub

Public Sub RunCalcForAllSheets()
    For Each xlSheet In Worksheets
        xlSheet.Select
        NewHeaders
        CalculateYearlyMetricsPerTicker
        CalculateOutliers
    Next
End Sub

Private Sub ClearCalcForAllSheets()
'Change scope from Private to Public so you can run this to clear any previous execution of calculations
'Do not use before executing RunCalcForAllSheets, doing so will delete some data from <vol> column
    For Each xlSheet In Worksheets
        xlSheet.Select
        ClearCalcs
    Next
End Sub

