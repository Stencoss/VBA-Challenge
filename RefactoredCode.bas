Attribute VB_Name = "Loops"

' Calls the other modules to be run on all the worksheets in the work book
' https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
'****************************************************************************************
'*                              RUN ME                                                  *
'****************************************************************************************
Sub run_all_sheets():

    Dim next_worksheet As Worksheet         ' Create a worksheet object
    
    
    Application.ScreenUpdating = False
    ' Loop
    For Each next_worksheet In Worksheets   ' Loops through the different worksheets
        next_worksheet.Select               ' goes through worksheets
        Call pull_data
        Call find_greatest
        Call look_pretty
    Next
    
    Application.ScreenUpdating = True



End Sub



' Working loop - combine all data extration functions into loop.

Sub pull_data():

    Dim sht As Worksheet            ' Store active sheet for switching
    Dim LastRow As Long             ' Collect the last row of a column - (ticker)
    Dim ticker_tracker As String    ' Stores the ticker value
    Dim total_volume As Double      ' Store the total value of stock
    Dim next_line As Integer        ' Track lines for chart
    Dim open_price As Double        ' Open_price to get difference      open - close =
    Dim yearly_change As Double     ' Yearly change from open to close
    Dim pecent_growth As Double     ' Find the percent growth of the stock
    Dim vol_open                    ' Open volume to find growth
    
    ' Initializing Variables
    Set sht = ActiveSheet                                   ' Have to use set for an obeject
    LastRow = sht.Range("A1").CurrentRegion.Rows.Count      ' Puts the number of rows in this variable
    ticker_tracker = Cells(2, 1).Value                      ' Sets the ticker value to first ticker
    total_volume = 0                                        ' Sets total volume to a start of 0
    next_line = 2                                           ' Initial value to be 2 row below title
    open_price = Cells(2, "C").Value                        ' Initial open
    vol_open = Cells(3, "G").Value                          ' First value is 0 - skipped it
    
    'Loop through the whole list - currently testing less
    For i = 2 To LastRow
            
        total_volume = total_volume + Cells(i, "G").Value     ' Add volume to total volume
            
            
            
        ' Open_price is sometimes  0 which errors out - have to check for this
        If ticker_tracker <> Cells(i + 1, 1) And open_price = 0 Then        ' If current ticker != to next ticker and open price not 0
            Cells(next_line, "I") = ticker_tracker                          ' Print the current ticker_tracker to chart
            Cells(next_line, "L") = total_volume                            ' Print total volume to chart
            yearly_change = Cells(i, "F").Value - open_price                ' Get total change
            Cells(next_line, "J") = yearly_change                           ' Print yearly_change to the chart
            percent_growth = 0                                              ' Get the percent chnage
            Cells(next_line, "K") = FormatPercent(percent_growth, 2)        ' Print percent_growth to chart
            
            yearly_change = 0                                               ' Get ready for next stock
            open_price = Cells(i + 1, "C").Value                            ' Get new starting price from next stock
            next_line = next_line + 1                                       ' Add 1 to next line to not overwrite the previous
            ticker_tracker = Cells(i + 1, 1).Value                          ' Change the value of ticker_tracker to next stock
            total_volume = 0                                                ' Reset total_volume to 0 - new stock
        ' Normal check - no 0's
        ElseIf ticker_tracker <> Cells(i + 1, 1) Then
            Cells(next_line, "I") = ticker_tracker                          ' Print the current ticker_tracker to chart
            Cells(next_line, "L") = total_volume                            ' Print total volume to chart
            yearly_change = Cells(i, "F").Value - open_price                ' Get total change
            Cells(next_line, "J") = yearly_change                           ' Print yearly_change to the chart
            percent_growth = yearly_change / open_price                     ' Get the percent chnage
            Cells(next_line, "K") = FormatPercent(percent_growth, 2)        ' Print percent_growth to chart
            
            yearly_change = 0                                 ' Get ready for next stock
            open_price = Cells(i + 1, "C").Value              ' Get new starting price from next stock
            next_line = next_line + 1                         ' Add 1 to next line to not overwrite the previous
            ticker_tracker = Cells(i + 1, 1).Value            ' Change the value of ticker_tracker to next stock
            total_volume = 0                                  ' Reset total_volume to 0 - new stock
        
        End If
    
    Next i

End Sub


' Make headers for new columns and add conditional formatting
' Need to confirm if formatting is in vba or if we can do in excel
Sub look_pretty():

    Dim sht As Worksheet
    Dim LastRow As Long
    
    Set sht = ActiveSheet
    LastRow = sht.Range("J1").CurrentRegion.Rows.Count

    Range("I1:L1").Font.Bold = True
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    ' Chart for values
    Range("N3").Value = "Greatest % Increase"
    Range("N4").Value = "Greatest % Decrease"
    Range("N5").Value = "Greatest Total Volume"
    Range("O2").Value = "Ticker"
    Range("P2").Value = "Value"


    ' Playing with colors - positve is green and negative is red
    For i = 2 To LastRow
    
        If Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 11) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
End Sub

' This module will find the greatest percent change for increase and decrease
' and the greatest total volume.  Needs to output the values and a ticker of
' what company it belongs to.


Sub find_greatest()

    ' Create variables
    Dim g_inc_percent As Double         ' greatest increase by percentage
    Dim g_dec_percent As Double         ' greatest decrease by percentage
    Dim g_volume As Double              ' greatest total volume
    Dim sht As Worksheet
    Dim LastRow As Long
    
    Set sht = ActiveSheet
    LastRow = sht.Range("J1").CurrentRegion.Rows.Count
    
    ' Loop though column k and find the largest and smallest percent
    'Find and print the highest percent
    g_inc_percent = WorksheetFunction.Max(Range("K2:K3000"))
    Range("P3").Value = FormatPercent(g_inc_percent)
    
    'Find an print the lowest percent
    g_dec_percent = WorksheetFunction.Min(Range("K2:K3000"))
    Range("P4").Value = FormatPercent(g_dec_percent)
    
    'Find and print the highest volume
    g_volume = WorksheetFunction.Max(Range("L2:L3000"))
    Range("P5") = g_volume
    '
    For i = 2 To LastRow
    
        If Cells(i, 11) = g_inc_percent Then
            Range("O3").Value = (Cells(i, 9).Value)
        End If
    
        If Cells(i, 11) = g_dec_percent Then
            Range("O4").Value = (Cells(i, 9).Value)
        End If
        
        If Cells(i, 12) = g_volume Then
            Range("O5").Value = (Cells(i, 9).Value)
        End If
    
    Next i

End Sub



