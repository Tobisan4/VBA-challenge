Attribute VB_Name = "Module1"
Sub Stock_analyzer():

'-------------------------------------------------
' Variables
'-------------------------------------------------

' Set variable for holding the stock ticker name
Dim Ticker_Name As String

' Set variable for holding the total volume per ticker name
Dim Ticker_Total_Volume As Double
Ticker_Total_Volume = 0

' Set variable for holding the opening price of the year
Dim Ticker_Opening_Price As Double
Ticker_Opening_Price = 0

' Set variable for holding the closing price of the year
Dim Ticker_Closing_Price As Double
Ticker_Closing_Price = 0

' Set variable for holding the price change
Dim Price_Change As Double
Price_Change = 0

' Set variable for holding the percent change
Dim Percent_Change As Double
Percent_Change = 0

' Set variable to keep track of the location of each ticker name in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Set variable for holding greatest %change increase data
Dim Great_Percent_Increase As Double
Great_Percent_Increase = 0

' Set variable for holding greatest %change decrease data
Dim Great_Percent_Decrease As Double
Great_Percent_Decrease = 0

' Set variable for holding greatest total volume data
Dim Great_Total_Volume As Double
Great_Total_Volume = 0

' Set variable for holding greatest %change increase ticker name
Dim Great_Percent_Increase_Ticker As String

' Set variable for holding greatest %change decrease ticker name
Dim Great_Percent_Decrease_Ticker As String

' Set variable for holding greatest total volume ticker name
Dim Great_Total_Volume_Ticker As String

Dim movedate As Integer


'-------------------------------------------------
' Looping through all sheets
'-------------------------------------------------

For Each ws In Worksheets

'-------------------------------------------------
' Creating summary table headers
'-------------------------------------------------

' Create the Ticker header of the summary table
ws.Range("I1").Value = "Ticker"

' Create the Yearly Change header of the summary table
ws.Range("J1").Value = "Yearly Change"

' Create the Percent Change header of the summary table
ws.Range("K1").Value = "Percent Change"

' Create the Total Stock Volume header of the summary table
ws.Range("L1").Value = "Total Stock Volume"

' Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
'-------------------------------------------------
' Looping through all the stocks on the current sheet to calculate Price Change,
' % Change & Total Stock Volume, also populate the summary table
'-------------------------------------------------

' Loop through all the stock entrees
For i = 2 To LastRow

    ' Check if we are still within the same stock ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
          
        ' Set the closing price of the year
        Ticker_Closing_Price = ws.Cells(i, 6).Value
          
        ' Set the stock ticker name
        Ticker_Name = ws.Cells(i, 1).Value
        
        ' Add the stock ticker name to the summary table
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

        ' Calculate the yearly price change
        Price_Change = (Ticker_Closing_Price - Ticker_Opening_Price)

        ' Add the yearly price change to the summary table
        ws.Range("J" & Summary_Table_Row).Value = Price_Change
        
        ' Setting the number format to a number rounded to 2 decimal places
        ws.Range("J" & Summary_Table_Row).NumberFormat = "#0.00"
          
        ' Check if there is a price change, if there is no change ...
        If ws.Range("J" & Summary_Table_Row).Value = 0 Then
                
            'Set cell color fill to no fill
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
                    
        ' If there is a positive price change
        ElseIf ws.Range("J" & Summary_Table_Row).Value > 0 Then
          
            'Set cell color fill to green
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
        ' If there is a negative price change
        ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
          
            'Set cell color fill to red
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
        End If
                  
        ' Check if there is a price change, if there is no change ...
        If Price_Change = 0 Then
            
            ' Set percentage to 0
            ws.Range("K" & Summary_Table_Row).Value = 0
            
            ' Setting the number format to percentage
                ws.Range("K" & Summary_Table_Row).NumberFormat = "#0.00%"
          
        ' Check if there is a price change, if there is a change ...
        Else
            
            ' Calculate the percent change
            Percent_Change = (Ticker_Closing_Price - Ticker_Opening_Price) / Ticker_Opening_Price
            
            ' Add the percent change to the summary table
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
            
            ' Setting the number format to percentage
            ws.Range("K" & Summary_Table_Row).NumberFormat = "#0.00%"
                
        End If
           
        '    Add to the stock total volume
        Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
          
        ' Add the total stock volume to the summary table
        ws.Range("L" & Summary_Table_Row).Value = Ticker_Total_Volume
          
        ' Increment the summary table row by 1
        Summary_Table_Row = Summary_Table_Row + 1
          
        ' Reset the stock total volume
        Ticker_Total_Volume = 0
          
        ' Reset the stock opening price
        Ticker_Opening_Price = 0
          
        ' Reset the stock closing price
        Ticker_Closing_Price = 0
          
    ' If the cell immediately following a row is the same stock ticker...
    Else
          
        ' Check if row is holding data for the opening day of the year
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
              
            'Check if ws.Cells(i,3).Value = 0 AND ws.Cells(i,7).Value = 0, if it is TRUE...
            If ws.Cells(i, 3).Value = 0 And ws.Cells(i, 7).Value = 0 Then
                
                ' Set movedate counter to 1
                movedate = 1
                    
                ' Loop until you found a non zero opening day cell price AND non zero volume
                Do Until ws.Cells(i + movedate, 3).Value > 0 Or ws.Cells(i + movedate, 1) <> ws.Cells(i + movedate + 1, 1)
                                    
                    movedate = movedate + 1
                        
                Loop
                    
                'Set the opening price of the year
                Ticker_Opening_Price = ws.Cells(i + movedate, 3).Value
                
            Else
                
                'Set the opening price of the year
                Ticker_Opening_Price = ws.Cells(i, 3).Value
                
            End If
                
        ' If the row is not holding the opening day information
        Else
           
        ' Add to the stock total volume
        Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value
                
        End If
    
    End If
    
    Next i
       
'-------------------------------------------------
' Identifying the Greatest % increase",
' Greatest % decrease and Greatest total volume
'-------------------------------------------------
    
    ' Determine the Last Row of the summary table
    Summary_LastRow = Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Setting the starting values of Great_Percent_Increase, Great_Percent_Increase_Ticker,Great_Percent_Decrease, Great_Percent_Decrease_Ticker, Great_Total_Volume, Great_Total_Volume_Ticker
    Great_Percent_Increase = ws.Cells(2, 11).Value
    Great_Percent_Increase_Ticker = ws.Cells(2, 9).Value
    Great_Percent_Decrease = ws.Cells(2, 11).Value
    Great_Percent_Decrease_Ticker = ws.Cells(2, 9).Value
    Great_Total_Volume = ws.Cells(2, 12).Value
    Great_Total_Volume_Ticker = ws.Cells(2, 9).Value


    ' Loop through all the summary table entrees
    For si = 3 To Summary_LastRow
    
        ' Compare Great_Percent_Increase with the current cell, if the current cell is higher ...
        If Great_Percent_Increase < ws.Cells(si, 11).Value Then
    
            Great_Percent_Increase = ws.Cells(si, 11).Value
            Great_Percent_Increase_Ticker = ws.Cells(si, 9).Value
            
        ' Compare Great_Percent_Increase with the current cell, if the current cell is lower ...
        Else
        
        End If
        
        ' Compare Great_Percent_Decrease with the current cell, if the current cell is lower ...
        If Great_Percent_Decrease > ws.Cells(si, 11).Value Then
    
            Great_Percent_Decrease = ws.Cells(si, 11).Value
            Great_Percent_Decrease_Ticker = ws.Cells(si, 9).Value

        ' Compare Great_Percent_Decrease with the current cell, if the current cell is higher ...
        Else
        
        End If
 
        ' Compare the total stock volume of the current cell and the next cell, if the current cell is higher ...
        If Great_Total_Volume < ws.Cells(si, 12).Value Then
    
            Great_Total_Volume = ws.Cells(si, 12).Value
            Great_Total_Volume_Ticker = ws.Cells(si, 9).Value
        
        Else
                  
        End If
        
    Next si
    
    ' Create the greatest percent increase label
    ws.Range("O2").Value = "Greatest % Increase"
    
    ' Create the greatest percent decrease label
    ws.Range("O3").Value = "Greatest % Decrease"
    
    ' Create the greatest total volume label
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ' Create the ticker label
    ws.Range("P1").Value = "Ticker"
    ' Create the volume label
    ws.Range("Q1").Value = "Volume"
    
    ' Populate greatest %increase ticker
    ws.Range("P2").Value = Great_Percent_Increase_Ticker
    
    ' Populate greatest %decrease ticker
    ws.Range("P3").Value = Great_Percent_Decrease_Ticker
    
    ' Populate greatest total volume ticker
    ws.Range("P4").Value = Great_Total_Volume_Ticker
    
    ' Populate greatest %increase data
    ws.Range("Q2").Value = Great_Percent_Increase
    
    ' Setting the number format to percentage
    ws.Range("Q2").NumberFormat = "#0.00%"
    
    ' Populate greatest %decrease data
    ws.Range("Q3").Value = Great_Percent_Decrease
    
    ' Setting the number format to percentage
    ws.Range("Q3").NumberFormat = "#0.00%"
    
    'Populate greatest total volume data
    ws.Range("Q4").Value = Great_Total_Volume
        
    ' Setting the number format to a number without decimal places
    ws.Range("Q4").NumberFormat = "#0"
    
    ' Setting the summary table row value back to 2
    Summary_Table_Row = 2
    
    ' Setting Great_Percent_Increase, Great_Percent_Decrease, Great_Total_Volume back to 0
    Great_Percent_Increase = 0
    Great_Percent_Decrease = 0
    Great_Total_Volume = 0

    ' Autofit data
    ws.Columns("I:Q").AutoFit
         
Next ws
     
End Sub

