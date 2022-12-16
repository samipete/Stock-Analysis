Attribute VB_Name = "Module1"
'create a script that loops through "all the stocks for one year" and outputs:
    'ticker symbol
    'yearly change from opening price to closing of a given year
    'percent change from opening price
    'total stock volume of stock
    
Sub stockanalysis():

' to loop through all worksheets in workbook
   Dim Current As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each Current In Worksheets
           
'define variables
   Dim openPrice, closePrice, yearlyChange, percentChange As Double
   Dim tickname As String
'set vol count at zero
    tickvol = 0
     
'set initial opening price that can be changed in the loop
     openPrice = Current.Cells(2, 3)
     
 'label summary table
        Current.Cells(1, 9).Value = "Ticker"
        Current.Cells(1, 10).Value = "Yearly Change"
        Current.Cells(1, 11).Value = "Percent Change"
        Current.Cells(1, 12).Value = "Total Stock Volume"
        
'Find last row in first column
    Dim I, lastRow, tickersumRow, Table As Integer
    lastRow = Current.Cells(Rows.Count, 1).End(xlUp).Row
    
    'keep track of what row we want summary table to start (similar to i)
    tickersumRow = 2
    
'loop through ticker names in first column
        For I = 2 To lastRow
        
         'add ticker name to loop
               tickname = Current.Cells(I, 1).Value
            
            'add ticker volume to loop
               tickvol = tickvol + Current.Cells(I, 7).Value
            
        
        ' use conditional statement  to search for when ticker name changes
            If Current.Cells(I + 1, "A") <> Current.Cells(I, 1).Value Then
             
         'this is where i get confused...
            
            'print ticker name in summary table
               Current.Cells(tickersumRow, "I").Value = tickname
                
            'print ticker volume
                Current.Cells(tickersumRow, "L").Value = tickvol
            
            'collect closeprice info
                closePrice = Current.Cells(I, 6).Value
            
            'find the yearly change
                yearlyChange = (closePrice - openPrice)
                
            'print yearly change in summary table
                Current.Cells(tickersumRow, "J").Value = yearlyChange
                
            'conditional formatting green+ red -
            If yearlyChange >= 0 Then
                Current.Cells(tickersumRow, "J").Interior.ColorIndex = 10
                
            Else
                Current.Cells(tickersumRow, "J").Interior.ColorIndex = 3
                
            End If
            
            'print yearly percent change
                percentChange = yearlyChange / openPrice
                ' percentChange = Format(percentChange, "0.00%")
                
                Current.Cells(tickersumRow, "K").Value = percentChange
                Current.Cells(tickersumRow, "K").NumberFormat = "0.00%"
                
        
            'reset openPrice
                openPrice = Current.Cells(I + 1, "C").Value
                
      
            'add one to tickersumRow
            tickersumRow = tickersumRow + 1
            
            'reset tickvol to 0
            tickvol = 0
            
        
            
        End If
        
        Next I
        
        Next
      
End Sub










