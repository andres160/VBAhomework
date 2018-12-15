Sub total_volume()

    total_vol = 0
    summary_ref = 2
    
    For i = 2 To 10000
    
        current1 = Cells(i, 1).Value
        next1 = Cells(i + 1, 1).Value
        current_vol = Cells(i, 7).Value
        

        If current1 = next1 Then
            total_vol = total_vol + current_vol
            
        ElseIf current1 <> next1 Then
       
            Cells(summary_ref, "J").Value = total_vol
            
            Cells(summary_ref, "I").Value = current1

            Total = 0
            
            summary_ref = summary_ref + 1
       
        End If
        
    Next i
    
End Sub
