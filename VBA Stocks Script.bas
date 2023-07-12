Attribute VB_Name = "Module1"
Sub stocks():
    
    For Each ws In Worksheets
    
        Dim ticker As String
        Dim first As Double
        Dim last As Double
        Dim change As Double
        Dim percent As Double
        Dim volume As Double
        Dim length As Double
        Dim srow As Integer
        srow = 2
        length = 800000
        For I = 2 To length
            'find beginning
            If ws.Cells(I - 1, 1).Value <> ws.Cells(I, 1).Value Then
                'record first price
                first = ws.Cells(I, 3).Value
                volume = volume + ws.Cells(I, 7).Value
            ElseIf ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                ' record ticker
                ticker = ws.Cells(I, 1).Value
                ' record last price
                last = ws.Cells(I, 6).Value
                volume = volume + ws.Cells(I, 7)
                change = last - first
                percent = change / first
                ws.Range("I" & srow).Value = ticker
                ws.Range("J" & srow).Value = change
                ws.Range("K" & srow).Value = FormatPercent(percent)
                ws.Range("L" & srow).Value = volume
                srow = srow + 1
                volume = 0
                
            Else
                volume = volume + ws.Cells(I, 7)
                
            End If
        
        Next I
        
        Dim inc As Double
        Dim dec As Double
        Dim vol As LongLong
        Dim it As String
        Dim dt As String
        Dim vt As String
        
        inc = 0
        dec = 0
        vol = 0
        
        For I = 2 To srow - 1
            If ws.Cells(I, 10).Value > 0 Then
                ws.Cells(I, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(I, 10).Value < 0 Then
                ws.Cells(I, 10).Interior.ColorIndex = 3
            Else
            End If
            
            If ws.Cells(I, 11).Value > inc Then
                inc = ws.Cells(I, 11).Value
                it = ws.Cells(I, 9).Value
            ElseIf ws.Cells(I, 11).Value < dec Then
                dec = ws.Cells(I, 11).Value
                dt = ws.Cells(I, 9).Value
            End If
            
            If ws.Cells(I, 12).Value > vol Then
                vol = ws.Cells(I, 12).Value
                vt = ws.Cells(I, 9).Value
            End If
        Next I
            
        ws.Range("P2").Value = it
        ws.Range("p3").Value = dt
        ws.Range("p4").Value = vt
    
        ws.Range("q2").Value = FormatPercent(inc)
        ws.Range("q3").Value = FormatPercent(dec)
        ws.Range("q4").Value = vol
    Next ws

End Sub

