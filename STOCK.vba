Sub Stock_Market_Analysis()

'Check through each worksheet'
Dim wkst As Worksheet
    For Each wkst In ActiveWorkbook.Worksheets
    wkst.Activate
    
        'Find the lastrow'
        lastrow = wkst.Cells(Rows.Count, "A").End(xlUp).Row
        'Rows.Count is the number of rows in a worksheet, just over one million'

        'Create Variable to hold Value
        Dim Yearly_Open As Double
        Dim Yearly_Close As Double
        Dim Yearly As Double
        Dim Ticker_Value As String
        Dim Percent As Double
        
        Dim Volume As Double
        Volume = 0

        Dim DataPoints As Double
        DataPoints = 2
        Dim Category As Integer
        Category = 1

        Dim i As Long
        
        'Set first row to the title
        Cells(1, "J").Value = "Yearly "
        Cells(1, "K").Value = "Percent"
        Cells(1, "I").Value = "Ticker Value"
        Cells(1, "L").Value = "Volume of Stock"
        
        'Set Initial Yearly Open
        Yearly_Open = Cells(2, Category + 2).Value
         ' Loop through all ticker symbol
        
        For i = 2 To lastrow
         'Check and see if the index is in the same ticker cell'
            If Cells(i + 1, Category).Value <> Cells(i, Category).Value Then
            
                ' Sets stock names to Ticker value'
                Ticker_Value = Cells(i, Category).Value
                'Set Yearly Close
                Yearly_Close = Cells(i, Category + 5).Value
                'Add Yearly Change between each stock value'
                Yearly = Yearly_Close - Yearly_Open
                ' This adds the total value to volume from G'
                
                'Puts ticker value in new column'
                Cells(DataPoints, Category + 8).Value = Ticker_Value
                'Adds yearly value to column J'
                Cells(DataPoints, Category + 9).Value = Yearly

                If (Yearly_Open = 0 And Yearly_Close <> 0) Then
                    Percent = 1
                'Set percent to zero if open&close are zero'
                
                ElseIf (Yearly_Open = 0 And Yearly_Close = 0) Then
                    Percent = 0
                Else
                    'Divide yearly value by open value to get percentage'
                    Percent = Yearly / Yearly_Open
                    'Set equal to column K'
                    Cells(DataPoints, Category + 10).Value = Percent
                    'Set value to percentage format '
                    Cells(DataPoints, Category + 10).NumberFormat = "0.00%"
                    
                End If
                
                Volume = Volume + Cells(i, Category + 6).Value
                'Reset the opened yearly value'
                'Adds volume value in column L'
                Cells(DataPoints, Category + 11).Value = Volume
                Yearly_Open = Cells(i + 1, Category + 2)
                'Add one to the iterated total table row
                DataPoints = DataPoints + 1
                'Reset the Volume Total'
    
                Volume = 0
                'Set percent equal to one if Open equals zero & close does not equal zero'
                
            Else
            
                'Stock Volume is added to its appropiate cell '
                Volume = Volume + Cells(i, Category + 6).Value
                
            End If
            
        Next i
        
'----------------------------------------------------------------'
        
        ' This is a second for loop for lastrow'
        TwoLastRow = wkst.Cells(Rows.Count, Category + 8).End(xlUp).Row
        ' Color positive, negative, and null cells
        For j = 2 To TwoLastRow
        
        'Check if values in "J" are less than zero'
            If Cells(j, Category + 9).Value < 0 Then
            'Color the inside red'
                Cells(j, Category + 9).Interior.ColorIndex = 3
                
               'Check if values in "J" are less than zero'
            ElseIf (Cells(j, Category + 9).Value > 0 Or Cells(j, Category + 9).Value = 0) Then
            'Color the inside green'
                Cells(j, Category + 9).Interior.ColorIndex = 4
                
            End If
            
        Next j
        
    Next wkst
    
End Sub
