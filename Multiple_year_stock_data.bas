Attribute VB_Name = "Module1"

Sub StocksData()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each WS In Worksheets
        
        ' Define Ticker variable
    
        Dim Ticker As String
        
        ' Define Year Open variable
    
        Dim Year_Open As Double
    
        ' Define Year Close variable
    
        Dim Year_Close As Double
        
        ' Define Total Volume variable to hold total stocks per ticker
    
        Dim Total_Volume As Double
    
        ' Set the greatest incraese of data variable
        Dim Greatest_Increase As Double
        
        ' Set the greatest decrease of data variable
        Dim Greatest_Decrease As Double
        
        ' Set the headers of the Stock Sumary Data
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Set the headers of the Greatest Changes Table
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        
        Year_Open_First = 2
        Row_Current = 2
            
        ' Determine the Last Row of the data
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        Year_Open = WS.Cells(Year_Open_First, 3).Value
        
        ' Loop through all the rows of data
        For i = 2 To LastRow
                                
           ' After reaching the last row of the current ticker symbol data
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                
                Ticker = WS.Cells(i, 1).Value
                WS.Cells(Row_Current, 9).Value = Ticker
                
                ' Set year's close price
                Year_Close = WS.Cells(i, 6).Value
                
                ' Calculating the yearly change between the year close and year open
                Yearly_Change = Year_Close - Year_Open
                WS.Cells(Row_Current, 10).Value = Yearly_Change
                
                ' Calcuating the percent change between the year close and year open
                Percent_Change = ((Yearly_Change) / Year_Open)
                WS.Cells(Row_Current, 11).Value = Percent_Change
                WS.Cells(Row_Current, 11).Style = "Percent"
                
                ' Add volume of row to total volume
                Total_Volume = Total_Volume + WS.Cells(i, 7).Value
                WS.Cells(Row_Current, 12).Value = Total_Volume
                                            
                ' Increase the row of the summary table for next ticker information
                Row_Current = Row_Current + 1
            
                ' Reset total volume for the next ticker data
                Total_Volume = 0
                
                Yearly_Change = 0
                
                Year_Open = WS.Cells(i + 1, 3).Value
            
            ' While looping through data of the current ticker symbol
            Else
                ' Add volume of row to total volume
                Total_Volume = Total_Volume + WS.Cells(i, 7).Value
            
            End If
              
       Next i
       

       Greatest_Increase = 0
       Greatest_Decrease = Cells(2, 11)
       Greatest_Total_Volume = 0
       Last_Row2 = WS.Cells(Rows.Count, 9).End(xlUp).Row
       
       ' Loop through the Yearly change data and color the cells based on + / -
       For j = 2 To Last_Row2
       
        If WS.Cells(j, 10) >= 0 Then
            WS.Cells(j, 10).Interior.ColorIndex = 4
        
        Else
            WS.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
        
        If WS.Cells(j, 11) > Greatest_Increase Then
            Greatest_Increase = WS.Cells(j, 11).Value
                  
            WS.Cells(2, 17).Value = Greatest_Increase
            WS.Cells(2, 17).Style = "Percent"
            WS.Cells(2, 16).Value = Cells(j, 9).Value
        
        ElseIf WS.Cells(j, 11) < Greatest_Decrease Then
            Greatest_Decrease = WS.Cells(j, 11).Value
        
            WS.Cells(3, 17).Value = Greatest_Decrease
            WS.Cells(3, 17).Style = "Percent"
            WS.Cells(3, 16).Value = WS.Cells(j, 9).Value
        
        End If
        
        If WS.Cells(j, 12).Value > Greatest_Total_Volume Then
            Greatest_Total_Volume = WS.Cells(j, 12).Value
        
            WS.Cells(4, 17).Value = Greatest_Total_Volume
            WS.Cells(4, 16).Value = Cells(j, 9).Value
        
        End If
       
       Next j
       
    Next WS

End Sub
