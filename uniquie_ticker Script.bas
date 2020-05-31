Attribute VB_Name = "Unique_Ticker"
Sub Unique_Ticker():

Dim ticker_row As Integer

ticker_row = 2

stock_volume = 0

Dim ticker As String

For i = 2 To 797711

         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         
        ticker = Cells(i, 1).Value
        
       stock_volume = stock_volume + Cells(i, 7).Value
       
       Range("H" & ticker_row).Value = ticker
       
       Range("K" & ticker_row).Value = stock_volume
       
       ticker_row = ticker_row + 1
       
       stock_volume = 0
            
       Else: stock_volume = stock_volume + Cells(i, 7).Value
           
      End If
      
Next i
    
End Sub
