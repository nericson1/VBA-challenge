VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub run_all()

    For Each ws In Sheets
        ws.Activate

        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        summary_table ws
        conditional_formatting_yearly ws
        conditional_formatting_percent ws
        
        Dim greatest_tick_name As String
        Dim greatest_tick_value As Double
        greatest_values ws, greatest_tick_name, greatest_tick_value
        
        Dim smallest_tick_name As String
        Dim smallest_tick_value As Double
        smallest_values ws, smallest_tick_name, smallest_tick_value
        
        Dim greatest_volume_tick_name As String
        Dim greatest_volume_tick_value As Variant
        greatest_volume ws, greatest_volume_tick_name, greatest_volume_tick_value
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 14).Value = "WORKSHEET"
        ws.Cells(1, 15).Value = ws.Name
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = greatest_tick_name
        ws.Cells(2, 16).Value = greatest_tick_value
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = smallest_tick_name
        ws.Cells(3, 16).Value = smallest_tick_value
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = greatest_volume_tick_name
        ws.Cells(4, 16).Value = greatest_volume_tick_value

        If overall_greatest_increase < greatest_tick_value Then
        overall_greatest_increase = greatest_tick_value
        overall_greatest_tick_name = greatest_tick_name
        End If
        
        If overall_greatest_decrease > smallest_tick_value Then
        overall_greatest_decrease = smallest_tick_value
        overall_smallest_tick_name = smallest_tick_name
        End If
        
        If overall_greatest_volume_value < greatest_volume_tick_value Then
        overall_greatest_volume_value = greatest_volume_tick_value
        overall_greatest_volume_name = greatest_volume_tick_name
        End If
        
    ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
    ws.Range("P2:P3").NumberFormat = "0.00%"
        
    Next ws

End Sub

Sub summary_table(ByVal ws As Worksheet)

    Dim column As Integer
    column = 1
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim totalrow As Long
    totalrow = 2
    Dim volumetotal As Variant
    
    For i = 2 To lastrow
        'Add stock volume to total
        volumetotal = volumetotal + ws.Cells(i, 7)
        If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
        'Update summary table
        ws.Cells(totalrow, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(totalrow, 12).Value = volumetotal
        ws.Cells(totalrow, 10).Value = ws.Cells(i, column + 5).Value - ws.Cells(i - 250, column + 2).Value
        'Percent change = new value minus old, divide difference by old number and multiply by 100
        ws.Cells(totalrow, 11).Value = (ws.Cells(totalrow, 10).Value / ws.Cells(i - 250, column + 2).Value)
        'Move new ticket total down one row
        totalrow = totalrow + 1
        'Reset stock volume total for next ticker
        volumetotal = 0
        End If
    Next i
    
End Sub

Sub conditional_formatting_yearly(ByVal ws As Worksheet)
        ws.Activate
        ws.Range("J2").Select
        'Loop through each row within the worksheet and assign fill color
        Do Until IsEmpty(ActiveCell)
            If ActiveCell.Value > 0 Then
                ActiveCell.Interior.ColorIndex = 4
                Else:
                ActiveCell.Interior.ColorIndex = 3
                End If
                'Move the active cell down to next row
                ActiveCell.Offset(1, 0).Select
        Loop
End Sub

Sub conditional_formatting_percent(ByVal ws As Worksheet)
        ws.Activate
        ws.Range("K2").Select
        'Loop through each row within the worksheet and assign fill color
        Do Until IsEmpty(ActiveCell)
            If ActiveCell.Value > 0 Then
                ActiveCell.Interior.ColorIndex = 4
                Else:
                ActiveCell.Interior.ColorIndex = 3
                End If
                'Move the active cell down to next row
                ActiveCell.Offset(1, 0).Select
        Loop
End Sub

Function greatest_values(ByVal ws As Worksheet, ByRef tick_name As String, ByRef tick_val As Double)
        ws.Range("K2").Select
        
        'Declare var to store the highest value found
        Dim maxVal As Double
        maxVal = 0

        'Declare var to store the ticker # of the highest value found
        Dim maxValTicker As String
        maxValTicker = ""
        
        'Loop through each row within the worksheet until an empty row is hit
        Do Until IsEmpty(ActiveCell)
            If ActiveCell.Value > maxVal Then
                maxVal = ActiveCell.Value
                'The ticker name is two columns behind the greatest value
                maxValTicker = ActiveCell.Offset(0, -2)
                End If
                'Move the active cell down to next row
                ActiveCell.Offset(1, 0).Select
        Loop
    
        tick_name = maxValTicker
        tick_val = maxVal

End Function

Sub smallest_values(ByVal ws As Worksheet, ByRef tick_name As String, ByRef tick_val As Double)
    'Declare var to store the smallest value found
    Dim minVal As Double
    minVal = 0

    'Declare var to store the ticker # of the smallest value found
    Dim minValTicker As String
    minValTicker = ""

    ws.Activate
    ws.Range("K2").Select
    'Loop through each row within the worksheet until an empty row is hit
    Do Until IsEmpty(ActiveCell)
        If ActiveCell.Value < minVal Then
            minVal = ActiveCell.Value
            'The ticker name is two columns behind the smallest value
            minValTicker = ActiveCell.Offset(0, -2)
            End If
            'Move the active cell down to next row
            ActiveCell.Offset(1, 0).Select
    Loop

    tick_name = minValTicker
    tick_val = minVal

End Sub

Sub greatest_volume(ByVal ws As Worksheet, ByRef tick_name As String, ByRef tick_val As Variant)
    'Declare var to store the highest volume found
    Dim maxVal As Variant
    maxVal = 0

    'Declare var to store the ticker # of the highest volume found
    Dim maxValTicker As String
    maxValTicker = ""

    ws.Activate
    ws.Range("L2").Select
    'Loop through each row within the worksheet until an empty row is hit
    Do Until IsEmpty(ActiveCell)
        If ActiveCell.Value > maxVal Then
            maxVal = ActiveCell.Value
            'The ticker name is two columns behind the greatest value
            maxValTicker = ActiveCell.Offset(0, -3)
            End If
            'Move the active cell down to next row
            ActiveCell.Offset(1, 0).Select
    Loop
    
    tick_name = maxValTicker
    tick_val = maxVal

End Sub
