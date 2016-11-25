Attribute VB_Name = "MakeList"
Sub MakeList()
    
    'getting values from input area
    '-------------------------------
    Dim values As Variant
    values = Range("B2:I4").Value
    
    'making a list
    '-------------------------------
    Dim counts(1 To 8) As Integer
    
    For num = 1 To (3 ^ 8)
        'making a row
        Cells(num + 7, 1).Value = num
        For i = 1 To 8
            Cells(num + 7, i + 1).Value = values(counts(i) + 1, i)
        Next i
        
        'increment counters
        For i = 8 To 1 Step -1
            If counts(i) < 2 Then
                counts(i) = counts(i) + 1
                Exit For
            Else
                counts(i) = 0
            End If
        Next i
    Next num

End Sub

Sub ClearList()
    Range("A8:I6568").Clear
End Sub
