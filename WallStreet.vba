Sub VBA()
  Dim earliest As Object, rowNum As Long
  Set dict = CreateObject("Scripting.Dictionary")
  
  For rowNum = 2 To UBound(Columns("A").Value)
    Dim ticker As String, theDate As Long, openAmt As Double, closeAmt As Double, stockVolume As Double
    ticker = Columns("A").Rows(rowNum).Value
    theDate = Columns("B").Rows(rowNum).Value
    openAmt = Columns("C").Rows(rowNum).Value
    closeAmt = Columns("F").Rows(rowNum).Value
    stockVol = Columns("G").Rows(rowNum).Value
    
    Dim perTickerDict As Object
    
    If dict.Exists(ticker) Then
      Set perTickerDict = dict(ticker)
      
      If theDate < perTickerDict("earliestDate") Then
        perTickerDict.Remove "earliestDate"
        perTickerDict.Remove "earliestOpen"
        perTickerDict.Add "earliestDate", theDate
        perTickerDict.Add "earliestOpen", openAmt
      End If
      
      If theDate > perTickerDict("latestDate") Then
        perTickerDict.Remove "latestDate"
        perTickerDict.Remove "latestClose"
        perTickerDict.Add "latestDate", theDate
        perTickerDict.Add "latestClose", closeAmt
      End If
      
      Dim totStockVol As Double
      totStockVol = perTickerDict("totStockVol") + stockVol
      perTickerDict.Remove "totStockVol"
      perTickerDict.Add "totStockVol", totStockVol
    Else
      Set perTickerDict = CreateObject("Scripting.Dictionary")
      
      perTickerDict.Add "earliestDate", theDate
      perTickerDict.Add "latestDate", theDate
      perTickerDict.Add "earliestOpen", openAmt
      perTickerDict.Add "latestClose", closeAmt
      perTickerDict.Add "totStockVol", stockVol
      
      dict.Add ticker, perTickerDict
    End If
  Next
  
  Columns("I").Rows(1).Value = "ticker"
  Columns("J").Rows(1).Value = "yearlyChange"
  Columns("K").Rows(1).Value = "percentChange"
  Columns("L").Rows(1).Value = "totStockVol"
  
  Dim startIndex As Long
  startIndex = 2
  
  For tickerIndex = LBound(dict.Keys()) To UBound(dict.Keys())
    ticker = dict.Keys()(tickerIndex)
    
    rowNum = startIndex + tickerIndex
    
    openAmt = dict(ticker)("earliestOpen")
    closeAmt = dict(ticker)("latestClose")
    totStockVol = dict(ticker)("totStockVol")
    Dim yearlyChange As Double
    yearlyChange = openAmt - closeAmt
    Dim percentChange As Double
    If openAmt = 0 Then
        percentChange = 1
    Else
        percentChange = yearlyChange / openAmt
    End If
    
    Columns("I").Rows(rowNum).Value = ticker
    Columns("J").Rows(rowNum).Value = yearlyChange
    Columns("K").Rows(rowNum).Value = percentChange
    Columns("L").Rows(rowNum).Value = totStockVol
    
  Next
End Sub
