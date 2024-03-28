Sub Alphabet_Soup()

' ============================                              ' a - data row start number
' = Testing on the Alphabets =                              ' b - data row end number
' ============================                              ' cope - first open "C"
' I started out being weird about                           ' d - summary row start number
' this alphabet stuff then figured                          ' e - summary row end number
' I'd keep it b/c you see a lot of                          ' fast - last closed in "F"
' these... robot or otherwise...                            ' g - volume sum in "G"
' and figured this would either                             ' hop - activeworkbook (if needed)
' make you laugh or iritate you...                          ' inc - increase
' either way, it'll be memorable <3                         ' j - start of data summary ***REPLACE WITH 2***
                                                            ' kick - ticker for max
Dim a, b, d, e, g, j As Integer                             ' lick - ticker for min
Dim cope, fast, inc, min, n, q As Double                    ' min - decrease
Dim kick, lick, pick, tick As String                        ' o -
'Dim hop As Workbook                                        ' pick - ticker for max volue
Dim skip As Worksheet                                       ' q - "percent" or "quotient"
                                                            ' r -
' =======================                                   ' skip - activeworksheet
' = Sheet Loop Hop.Skip =                                   ' tick - ticker
' =======================
                                                            
For Each skip In ActiveWorkbook.Worksheets
    
    ' =======================
    ' = Loop Summary Detail =
    ' =======================
    
    skip.Range("I1").Value = "Ticker"                       ' Setting headers for width protocol
    skip.Range("J1").Value = "Yearly Change"
    skip.Range("K1").Value = "Percent Change"
    skip.Range("L1").Value = "Total Stock Volume"
    skip.Range("J:K").ColumnWidth = 15
    skip.Range("L:L").ColumnWidth = 20
    
    b = skip.Cells(Rows.Count, 1).End(xlUp).Row             ' Establishing first loop pre-values
    g = 0
    j = 2
    cope = skip.Cells(2, 3).Value
    'MsgBox (b)
        
        For a = 2 To b
              
                If skip.Cells(a + 1, 1).Value = skip.Cells(a, 1).Value Then
                
                    g = g + skip.Cells(a, 7).Value
                    'MsgBox (g)
            
                    Else
                    
                    g = g + skip.Cells(a, 7).Value
                    tick = skip.Cells(a, 1).Value
                    fast = skip.Cells(a, 6).Value
                    q = (fast / cope - 1)
                    'MsgBox (tick & " - " & cope & " - " & fast & " - " & g)
                    'MsgBox (FormatPercent(q, 2))
                    skip.Cells(j, 9).Value = tick
                    skip.Cells(j, 10).Value = fast - cope                   'Yearly Change
                    skip.Cells(j, 11).Value = FormatPercent(q, 2)           'Percent Change
                    skip.Cells(j, 12).Value = FormatNumber(g, 0)
                    '==============================================================================
                    'Please don't dock me for improving readability, I'm electing to use the number
                    'format over text to avoid Scientific output - but can conver to text by using:
                    ' .Value = "'" & g (source: I'm familiar with the aprotrophe; use excel often)
                    '==============================================================================
                    
                            If q < 0 Then                                   'Conditional Logic
                                skip.Cells(j, 10).Interior.ColorIndex = 3
                                Else
                                skip.Cells(j, 10).Interior.ColorIndex = 4
                            End If
                    
                    g = 0
                    cope = skip.Cells(a + 1, 3).Value
                    j = j + 1
                
                End If
            
        Next a
        
        ' =======================
        ' = Loop Total Summary =
        ' =======================
                
        skip.Range("O2").Value = "Greatest % Increase"
        skip.Range("O3").Value = "Greatest % Decrease"
        skip.Range("O4").Value = "Greatest Total Volume"
        skip.Range("P1").Value = "Ticker"
        skip.Range("Q1").Value = "Value"
        
        
        e = skip.Cells(Rows.Count, 9).End(xlUp).Row
        inc = 0
        min = 0
        n = 0
        'MsgBox (e)
        
                For d = 2 To e
                    If skip.Cells(d, 11).Value > inc Then
                        inc = skip.Cells(d, 11).Value
                        kick = skip.Cells(d, 9).Value
                    Else: End If
                Next d
        
        'MsgBox (kick & " - " & inc)
        skip.Range("P2").Value = kick
        skip.Range("Q2").Value = FormatPercent(inc, 2)
        
                For d = 2 To e
                    If skip.Cells(d, 11).Value < min Then
                        min = skip.Cells(d, 11).Value
                        lick = skip.Cells(d, 9).Value
                    Else: End If
                Next d
                
        'MsgBox (lick & " - " & min)
        skip.Range("P3").Value = lick
        skip.Range("Q3").Value = FormatPercent(min, 2)
        
                For d = 2 To e
                    If skip.Cells(d, 12).Value > n Then
                        n = skip.Cells(d, 12).Value
                        pick = skip.Cells(d, 9).Value
                    Else: End If
                Next d
        
        'MsgBox (pick & " - " & FormatNumber(n, 0))
        skip.Range("P4").Value = pick
        skip.Range("Q4").Value = FormatNumber(n, 0)
        skip.Range("O:O").ColumnWidth = Len(skip.Range("O4"))
        skip.Range("Q:Q").ColumnWidth = Len(skip.Range("Q4")) + (Len(skip.Range("Q4")) / 3)
        
        ' =================================================================
        ' See? I made it harder for myself to get a relative width. ^ This
        ' divides the lenght of the cell by 3 to add in the qty of commmas
        ' =================================================================
        
        'MsgBox (Len(Range("Q4")) + (Len(Range("Q4")) / 3))

Next skip

End Sub


