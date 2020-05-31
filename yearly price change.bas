Attribute VB_Name = "yearly_price_chg"
Sub yearly_price_chg()

Dim mydate, mymonth, myday As String
Dim price_row As Long

price_row = 2

For i = 2 To 3000

    
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
        Do While Cells(i, 2).vlaue <> " "
        mydate = CStr(Cells(i, 2))
            If Len(mydate) = 10 Then
                mymonth = Right(mydate, 2, 2)
                myday = Mid(mydate, 3, 2)
                
            End If
            
                If myday = 1 And mymonth = 1 Then
                jan_price = Cells(i, 3).Value
                
                If myday = 30 And mymont = 12 Then
                dec_price = Cells(i, 3).Value
                
                End If
            Next i
        
        Range("I" & price_row).Value = (jan_price - dec_price)
        
        price_row = price_row + 1
        
    Next i
    
        
End Sub
