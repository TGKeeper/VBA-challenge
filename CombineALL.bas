Attribute VB_Name = "CombineALL"
Sub CombineALL()

Dim wks As Worksheet
For Each wks In Worksheets

lastrow = wks.Cells(Rows.Count, 1).End(xlUp).Row

wks.Range("I1").Value = "TICKER"
wks.Range("J1").Value = "YEARLY CHANGE"
wks.Range("K1").Value = "PERCENT CHANGE"
wks.Range("L1").Value = "TOTAL STOCK VOLUME"

wks.Range("N2").Value = "GREATEST % INCREASE"
wks.Range("N3").Value = "GREATEST % DECREASE"
wks.Range("N4").Value = "GREATEST TOTAL VOLUME"

Dim tickerUNIQ As String

Dim tickerTOTAL As Double
tickerTOTAL = 0

Dim summaryrow As Integer
summaryrow = 2

Dim summaryrow2 As Integer
summaryrow2 = 2

Dim summaryrow3 As Integer
summaryrow3 = 2

Dim openvalue As Double
Dim closevalue As Double
Dim yrlychng As Double
Dim percchange As Double

Dim jcolor As Long
jcolor = 2

Dim kcolor As Long
kcolor = 2

For s = 2 To lastrow

   
    If wks.Cells(s + 1, 1).Value <> wks.Cells(s, 1).Value Then
        
        tickerUNIQ = wks.Cells(s, 1).Value
        

        tickerTOTAL = tickerTOTAL + wks.Cells(s, 7).Value
    
        wks.Range("I" & summaryrow).Value = tickerUNIQ
    
        wks.Range("L" & summaryrow).Value = tickerTOTAL
           
        summaryrow = summaryrow + 1
    
        tickerTOTAL = 0
        
    
    Else
    
        tickerTOTAL = tickerTOTAL + wks.Cells(s, 7).Value
        
    End If
    
    
    If Right(wks.Cells(s, 2).Value, 4) = "0102" Then
        openvalue = wks.Cells(s, 3)
        
        

    ElseIf Right(wks.Cells(s, 2).Value, 4) = "1231" Then 'removed range ref to only store open/closevalue
        closevalue = wks.Cells(s, 6)
    
    
    wks.Range("J" & summaryrow2).Value = closevalue - openvalue
    
    
    yrlychng = closevalue - openvalue
    
    wks.Range("K" & summaryrow3).Value = yrlychng / openvalue
    
    percchange = yrlychng / openvalue
 
    summaryrow2 = summaryrow2 + 1
    summaryrow3 = summaryrow3 + 1
    
End If
                     
    
Next s

    
Dim lastSUMMARYrow As Integer
lastSUMMARYrow = wks.Cells(Rows.Count, 9).End(xlUp).Row

For J = 2 To lastSUMMARYrow

If wks.Cells(J, 10).Value > 0 Then
    wks.Cells(J, 10).Interior.ColorIndex = 4
    
    
    ElseIf wks.Cells(J, 10) < 0 Then
    wks.Cells(J, 10).Interior.ColorIndex = 3
    
    
    ElseIf wks.Cells(J, 10) = 0 Then
    wks.Cells(J, 10).Interior.ColorIndex = 0
    
    
End If


Next J

For K = 2 To lastSUMMARYrow

If wks.Cells(K, 11).Value > 0 Then
    wks.Cells(K, 11).Interior.ColorIndex = 4
    
    
    ElseIf wks.Cells(K, 11) < 0 Then
    wks.Cells(K, 11).Interior.ColorIndex = 3
    
    
    ElseIf wks.Cells(K, 11) = 0 Then
    wks.Cells(K, 11).Interior.ColorIndex = 0
    
End If
    
Next K
    
    
Dim Krng As Range
    Set Krng = wks.Range("K2:K500")
    
    Dim MaxK As Double
    MaxK = WorksheetFunction.Max(Krng)
    
    Dim MinK As Double
    MinK = WorksheetFunction.Min(Krng)
    
    Dim MostVol As Double
    Set V = wks.Range("L2:L200")
    MostVol = WorksheetFunction.Max(V)

wks.Range("O2") = MaxK
wks.Range("O2").NumberFormat = "0.00%"
wks.Range("O3") = MinK
wks.Range("O3").NumberFormat = "0.00%"
wks.Range("O4") = MostVol

Krng.NumberFormat = "0.00%"
wks.Columns("I:O").AutoFit

Next wks
End Sub

