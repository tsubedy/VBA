Sub Stock_analysis()
    
    ' Setting initial variable for holding the ticker name
    Dim Ticker As String
    Ticker = " "
    
    ' Setting an initial variable for total stock volume for each ticker
    Dim Ticker_Vol As Double
    Ticker_Vol = 0
    
    ' Settin other variables
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Price_Change As Double
    Dim Percent_Change As Double
    
    ' Initializing variable values
    Open_Price = 0
    Close_Price = 0
    Price_Change = 0
    Percent_Change = 0
    
    ' Tracking the location for each ticker name in the summary table
    
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    ' Setting initial row count for the current worksheet
    Dim Lrow As Long
    Dim i As Long
    
    Lrow = Cells(rows.Count, 1).End(xlUp).Row
    'MsgBox (Lrow)
    
   ' Setting Headers for the Summary Table
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    ' Setting Open Price for the first Ticker.
    ' Rest of the ticker's open price are initialized inside the for loop
    
    Open_Price = Cells(2, 3).Value
        
    ' Looping for rows beginning from row2
    
    For i = 2 To Lrow
      
    ' Checking if  within the same ticker name,
    ' if not - print the results in the summary table
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Setting the ticker name
            Ticker = Cells(i, 1).Value
                
            ' Calculating Price_Change and Percent_Change
                Close_Price = Cells(i, 6).Value
                Price_Change = Close_Price - Open_Price
            
            If Open_Price <> 0 Then
                Percent_Change = (Price_Change / Open_Price) * 100
            
            Else
            
            ' Checking if any cells have a zero value for open_price
                MsgBox ("For " & Ticker & ", Row " & CStr(i) & ": <Open> is " & Open_Price & ". Please, Fix and proceed.")
            
            End If
                
            ' Adding Total Volume to the Ticker
            Ticker_Vol = Ticker_Vol + Cells(i, 7).Value
              
            ' Printing the Ticker Name in the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker
            
            ' Printing the price change for the Ticker Name in the Summary Table
            Range("J" & Summary_Table_Row).Value = Price_Change
            
            ' Filling "Yearly Change", i.e. Price_Change with Green and Red colors
            If (Price_Change > 0) Then  'Fill column with GREEN color - positive
                Range("J" & Summary_Table_Row).Interior.Color = vbGreen
            
            Else 'If (Price_Change <= 0) then Filling column with RED color - Negative
                Range("J" & Summary_Table_Row).Interior.Color = vbRed
            
            End If
                
            ' Printing the Percentage Change in the Summary Table
            Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
            
            ' Printing the Total Ticker Volume in the Summary Table
            Range("L" & Summary_Table_Row).Value = Ticker_Vol
                
            
            ' Going to the next row in the summary table
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reseting Price_Change and Percent_Change in order to go next Ticker
            Price_Change = 0
            Close_Price = 0
            
            Open_Price = Cells(i + 1, 3).Value
            Percent_Change = 0
            Ticker_Vol = 0
                
        'Else - If the immediate next row has the same ticker name just add to Total Ticker Volume
        
        Else
        
        ' Adding the Total Ticker Volume with the previous value
            Ticker_Vol = Ticker_Vol + Cells(i, 7).Value
        
        End If
             
    Next i
        
End Sub
