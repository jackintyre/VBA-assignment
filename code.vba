Attribute VB_Name = "Module1"
Sub stock():

Dim ticker As String
Dim i As Long
Dim j As Long
Dim volume As Variant
volume = 0

Dim startvalue As Double

startvalue = 0
Dim ws As Worksheet
Set ws = ActiveSheet
For Each Current In Worksheets:

startvalue = Current.Cells(2, 3).Value

Current.Cells(1, 10).Value = "Ticker"
Current.Cells(1, 11).Value = "Yearly Change"
Current.Cells(1, 12).Value = "Percent Change"
Current.Cells(1, 13).Value = "Total Volume"

Current.Cells(1, 16).Value = "Ticker"
Current.Cells(1, 17).Value = "Value"

Current.Cells(2, 15).Value = "Greatest % increase"
Current.Cells(3, 15).Value = "Greatest % decrease"
Current.Cells(4, 15).Value = "Greatest Total Volume"



Dim lastrow As Long
j = 2
Current.Cells(j, 10) = Current.Cells(2, 1).Value
'ReDim ticker(j)

'ReDim volume(j)
Dim percentchange As Double

lastrow = Current.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To (lastrow + 1)

    
    If (Current.Cells(i, 1).Value <> Current.Cells(j, 10)) Then
       ' ReDim Preserve ticker(j - 2)
         Current.Cells(j, 11).Value = Current.Cells((i - 1), 6).Value - startvalue
       
        If ((Current.Cells((i - 1), 6).Value - startvalue) <> 0) And (startvalue <> 0) Then

              percentchange = ((Current.Cells((i - 1), 6).Value / startvalue) - 1) * 100
                Current.Cells(j, 12).Value = percentchange
         
        ElseIf ((Current.Cells((i - 1), 6).Value) = 0) And (startvalue <> 0) Then
         
            percentchange = -100
            Current.Cells(j, 12).Value = percentchange
            
        ElseIf ((Current.Cells((i - 1), 6).Value) = 0) And (startvalue = 0) Then
          
        
             percentchange = 0
            Current.Cells(j, 12).Value = percentchange
        End If
        
         
       
        
        ticker = Current.Cells(i, 1).Value
       
        startvalue = Current.Cells(i, 3).Value
       
        
        
        j = j + 1
        Current.Cells(j, 10).Value = ticker
          volume = Current.Cells(i, 7).Value
      
       
     Current.Cells(j, 13).Value = volume
      
        
    Else
        If (startvalue = 0) Then
            startvalue = Current.Cells((i - 1), 6).Value
        End If
        
        volume = volume + Current.Cells(i, 7).Value
        Current.Cells(j, 13).Value = volume
        If (volume > Current.Cells(4, 17).Value) Then
        
            Current.Cells(4, 16) = ticker
            Current.Cells(4, 17) = volume
    
        End If
        
    End If
    
   
        
Next i



lastrow = Current.Cells(Rows.Count, 10).End(xlUp).Row
Dim c As Integer

For c = 2 To lastrow
percentchange = Current.Cells(c, 12).Value
    
    If (Current.Cells(c, 11).Value >= 0) Then
        
        Current.Cells(c, 11).Interior.ColorIndex = 4
        
    Else
        
        Current.Cells(c, 11).Interior.ColorIndex = 3
        
    End If
     If (percentchange > Current.Cells(2, 17).Value) Then
        ticker = Current.Cells(c, 10).Value
           Current.Cells(2, 16) = ticker
            Current.Cells(2, 17) = percentchange
    
        End If
          If (percentchange < Current.Cells(3, 17).Value) Then
        ticker = Current.Cells(c, 10).Value
            Current.Cells(3, 16) = ticker
            Current.Cells(3, 17) = percentchange
            
    
        End If
    
Next c

        
    
Next


End Sub




