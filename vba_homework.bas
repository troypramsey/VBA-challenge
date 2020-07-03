Attribute VB_Name = "Module1"
Sub VBAhw()



'-----------------
' SHEETS LOOP
'-----------------

For ws = 1 To Worksheets.Count
    If ActiveSheet.Index = Worksheets.Count Then
    Worksheets(1).Activate
Else
    ActiveSheet.Next.Activate
End If
    
    ' Credit to https://www.automateexcel.com/vba/activate-select-sheet/


    '--------------------
    ' DECLARING VARIABLES
    '--------------------
    
    ' Column Variables
    
    Dim ticker As String
    Dim y_change As Double
    y_change = 0
    Dim open_price As Double
    open_price = Range("C2").Value
    Dim close_price As Double
    close_price = 0
    Dim y_pct As Double
    Dim total_volume As Double
    total_volume = 0
    Dim summary_head As Long
    summary_head = 2
    
    ' Challenge Variables
    
    Dim max_increase As Double
    max_increase = 0
    Dim max_increase_ticker As String
    Dim max_decrease As Double
    max_decrease = 0
    Dim max_decrease_ticker As String
    Dim max_vol As Double
    max_vol = 0
    Dim max_vol_ticker As String
    
    ' End of rows variable
    
    Dim last_row As Long
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
    '----------------
    ' ADDING HEADERS
    '----------------
    
    ' Primary assignment headers
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Challenge column headers
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    
    ' Challenge row headers
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    '--------------------------
    ' TICKER LOOP
    '--------------------------
    
    For r = 2 To last_row
    
        ' Move to next company
        
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
            ticker = Cells(r, 1).Value
            close_price = Cells(r, 6).Value
            
        ' Calculate total volume
        
            total_volume = total_volume + Cells(r, 7).Value
        
        ' Incrementing max volume challenge variable
            
            If total_volume > max_vol Then
                max_vol = total_volume
                max_vol_ticker = ticker
            End If
            
            
        ' Calculate Yearly Change
        
            y_change = close_price - open_price
        
        ' Using the open price and Yearly Change to assign the Percentage Change
            
            If open_price = 0 Then
                y_pct = 0
            Else
                y_pct = y_change / open_price
            End If
            
        ' Incrementing greatest % increase and decrease challenge variables
        
            If y_pct > max_increase Then
                max_increase = y_pct
                max_increase_ticker = ticker
            End If
            
            If y_pct < max_decrease Then
                max_decrease = y_pct
                max_decrease_ticker = ticker
            End If
            
            
    '-----------------------------
    ' FILLING THE REQUIRED COLUMNS
    '-----------------------------
            
            ' Filling the Ticker column
            Range("I" & summary_head).Value = ticker
            
            ' Filling and formatting the Yearly Change column
            Range("J" & summary_head).Value = y_change
                If y_change > 0 Then
                    Range("J" & summary_head).Interior.Color = RGB(0, 255, 0)
                Else
                    Range("J" & summary_head).Interior.Color = RGB(255, 0, 0)
                End If
                
            ' Filling the Percent Change column
            Range("K" & summary_head).Value = y_pct
            
            ' Filling the Total Stock Volume column
            Range("L" & summary_head).Value = total_volume
            
            
            
    '--------------------------
    ' RESETTING VARIABLES
    '--------------------------
            
            ' Move the summary table head
            summary_head = summary_head + 1
            
            ' Reset Yearly Change
            y_change = 0
            
            ' Next company's open price
            open_price = Cells(r + 1, 3).Value
            
            ' Reset Percentage change
            y_pct = 0
            
            ' Reset the Total Stock Volume variable
            total_volume = 0
            
            
    ' End of company selection IF condition
        Else
        
    '-------------------------------------------
    ' PERFORMING OPERATIONS ON CURRENT COMPANY
    '-------------------------------------------
            
            ' Updating the Total Stock Volume
            total_volume = total_volume + Cells(r, 7).Value
            

    
            
    ' End of entire company selection IF statement
        End If
        
        
    Next r
    
'-----------------------------
' FILLING CHALLENGE COLUMNS
'-----------------------------
        
     Range("Q2").Value = max_increase
     Range("P2").Value = max_increase_ticker
 
     Range("Q3").Value = max_decrease
     Range("P3").Value = max_decrease_ticker

     Range("Q4").Value = max_vol
     Range("P4").Value = max_vol_ticker
            
        
'-----------------------------
' FORMATTING AND CLEANUP
'-----------------------------

' Setting the style on the Percent Change column
Range("K2:K" & last_row).Style = "Percent"
Range("K2:K" & last_row).NumberFormat = "0.00%"

' Setting the style on the max increase and max decrease challenge variables
Range("Q2:Q3").Style = "Percent"
Range("Q2:Q3").NumberFormat = "0.00%"

' Autofitting column width
ThisWorkbook.ActiveSheet.Cells.EntireColumn.AutoFit



Next

End Sub
Sub ClearSheet()
'
' USED TO CLEAR SHEETS DURING TESTING
'

'

For ws = 1 To Worksheets.Count
    If ActiveSheet.Index = Worksheets.Count Then
    Worksheets(1).Activate
Else
    ActiveSheet.Next.Activate
End If

    Range("I:I,J:J,K:K,L:L,O:O,P:P,Q:Q").Select
    Range("Q1").Activate
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
Next
End Sub
