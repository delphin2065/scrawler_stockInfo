Sub inputData()
    Application.DisplayAlerts = False
   

    Call clearWorksheet
    
    Dim fileName As String
    Dim filePath As String
    Dim wb As Workbook
    Dim rng As Range
    Dim dateOrd As Long
    Dim cumPLOrd As Long
    Dim wrksOrd As Long
    Dim nRow As Long
    
    Dim entryPriceOrd As Long
    Dim entryShareOrd As Long
    
    
    
    filePath = ThisWorkbook.Path
    fileName = filePath & "\" & ThisWorkbook.Worksheets(1).Range("B1").Value
    wrksOrd = ThisWorkbook.Worksheets(1).Range("B2").Value
    dateOrd = ThisWorkbook.Worksheets(1).Range("B3").Value
    cumPLOrd = ThisWorkbook.Worksheets(1).Range("B4").Value
    
    entryPriceOrd = ThisWorkbook.Worksheets(1).Range("B5").Value
    entryShareOrd = ThisWorkbook.Worksheets(1).Range("B6").Value

  
    Set wb = GetObject(fileName)
    
    
    With wb.Worksheets(wrksOrd).Range("A1").CurrentRegion
        nRow = .Rows.Count
        ThisWorkbook.Worksheets(2).Range("A1").Resize(nRow, 4).Value = .Columns(dateOrd).Value
        ThisWorkbook.Worksheets(2).Range("B1").Resize(nRow, 1).Value = .Columns(cumPLOrd).Value
        ThisWorkbook.Worksheets(2).Range("C1").Resize(nRow, 1).Value = .Columns(entryPriceOrd).Value
        ThisWorkbook.Worksheets(2).Range("D1").Resize(nRow, 1).Value = .Columns(entryShareOrd).Value
    
    
    End With
    
    Set rng = Nothing
    Set wb = Nothing
    
    Call closeOtherWorkbook
    Call calInICap
    Call calPL
    Call calReturnYrs
    
'    Call calDrowDown
    Call calProfitIndicator
    
    Application.DisplayAlerts = True
End Sub

Sub calPL()
    Dim rng As Range
    Dim cell As Range
    Dim nCol As Long
    Dim nCol2 As Long
    
    Dim nRow As Long
    
    Set rng = ThisWorkbook.Worksheets(2).Range("A1").CurrentRegion
        nCol = rng.End(xlToRight).Column + 1
        nCol2 = rng.End(xlToRight).Column + 2
        nRow = rng.End(xlDown).Row
        rng.Cells(1, nCol) = "PL"
        rng.Cells(2, nCol).Formula = "=B2"
        rng.Cells(3, nCol).Formula = "=B3-B2"
        rng.Cells(3, nCol).Copy
        rng.Range(rng.Cells(3, nCol), rng.Cells(nRow, nCol)).PasteSpecial xlPasteAll
        
        rng.Cells(1, nCol2) = "countWinLoss"
        rng.Cells(2, nCol2).Formula = "=IF(F2>=0,1,-1)"
        rng.Cells(2, nCol2).Copy
        rng.Range(rng.Cells(2, nCol2), rng.Cells(nRow, nCol2)).PasteSpecial xlPasteAll
        
        
    
    Set rng = Nothing
    
    Set rng = ThisWorkbook.Worksheets(2).Range("A1").CurrentRegion
        rng.Copy
        rng.PasteSpecial xlPasteValues
    Set rng = Nothing
End Sub


Sub calDrowDown()
    Dim rng As Range
    Dim cell As Range
    Dim nCol As Long
    Dim nRow As Long
    
    Application.Calculation = xlCalculationManual
    Set rng = ThisWorkbook.Worksheets(2).Range("A1").CurrentRegion
        nCol = rng.End(xlToRight).Column + 1
        nRow = rng.End(xlDown).Row
        rng.Cells(1, nCol) = "DrowDown"
        rng.Cells(2, nCol).Formula = "=B2-MAX(B$2:B2)"
        rng.Cells(2, nCol).Copy
        rng.Range(rng.Cells(2, nCol), rng.Cells(nRow, nCol)).PasteSpecial xlPasteAll
    Set rng = Nothing
    Application.Calculation = xlCalculationAutomatic
    
    Set rng = ThisWorkbook.Worksheets(2).Range("A1").CurrentRegion
        rng.Copy
        rng.PasteSpecial xlPasteValues
    Set rng = Nothing
    

End Sub

Sub calInICap()
    Dim rng As Range
    Dim cell As Range
    Dim nCol As Long
    Dim nRow As Long
    
    Application.Calculation = xlCalculationManual
    Set rng = ThisWorkbook.Worksheets(2).Range("A1").CurrentRegion
        nCol = rng.End(xlToRight).Column + 1
        nRow = rng.End(xlDown).Row
        rng.Cells(1, nCol) = "iniCap"
        rng.Cells(2, nCol).Formula = "=C2*ABS(D2)*1000"
        rng.Cells(2, nCol).Copy
        rng.Range(rng.Cells(2, nCol), rng.Cells(nRow, nCol)).PasteSpecial xlPasteAll
    Set rng = Nothing
    Application.Calculation = xlCalculationAutomatic
    
    Set rng = ThisWorkbook.Worksheets(2).Range("A1").CurrentRegion
        rng.Copy
        rng.PasteSpecial xlPasteValues
    Set rng = Nothing
    

End Sub
Sub calProfitIndicator()
    Dim nRow As Long
    Dim startYrs As Long
    Dim endYrs As Long
    
    
    ' calculate WinLoss ratio
    ThisWorkbook.Worksheets(1).Range("B11").Formula = "=COUNTIFS(result!G:G,1)/COUNT(result!G:G)"
    
    
    ' calculate MDD
    ThisWorkbook.Worksheets(1).Range("B12").Formula = "=MIN(Est_iniCap!E:E)"
    
    
    ' calculate netProfit
    nRow = ThisWorkbook.Worksheets(3).Range("A1").CurrentRegion.Rows.Count
    ThisWorkbook.Worksheets(1).Range("B13") = ThisWorkbook.Worksheets(3).Range("D" & nRow)


    ' calculate riskReturn ratio
    nRow = ThisWorkbook.Worksheets(3).Range("A1").CurrentRegion.Rows.Count
    ThisWorkbook.Worksheets(1).Range("B14") = ThisWorkbook.Worksheets(1).Range("B13") / Abs(ThisWorkbook.Worksheets(1).Range("B12").Value)


    '   calculate PL ratio
    ThisWorkbook.Worksheets(1).Range("B15").Formula = "=ABS((SUMIFS(result!E:E,result!F:F,"">=0"")/COUNTIFS(result!F:F,"">=0""))/(SUMIFS(result!E:E,result!F:F,""<0"")/COUNTIFS(result!F:F,""<0"")))"

    '   iniCap
    ThisWorkbook.Worksheets(1).Range("B16").Formula = "=MAX(Est_iniCap!B:B)"

    ' returnRate
    ThisWorkbook.Worksheets(1).Range("B17").Formula = ThisWorkbook.Worksheets(1).Range("B13") / ThisWorkbook.Worksheets(1).Range("B16")

End Sub


Sub clearWorksheet()
    ThisWorkbook.Worksheets(2).Cells.Delete
    ThisWorkbook.Worksheets(3).Cells.Delete
    
    ThisWorkbook.Worksheets(1).Range("B11", "B20").ClearContents
    
    Call deleteCharts

End Sub


Sub closeOtherWorkbook()
    Dim wb As Workbook
    
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            wb.Close SaveChanges:=False
        
        End If
    Next wb

End Sub



Sub drawPicture()
    Dim ws As Worksheet
    Dim wsChart As Worksheet
    Dim chartObj As ChartObject
    Dim rng As Range
    Dim nRow As Long

    Set wsChart = ThisWorkbook.Worksheets(1)
        Set ws = ThisWorkbook.Worksheets(3)
            nRow = ws.Range("A1").CurrentRegion.Rows.Count
            Set rng = ws.Range("A2", "G" & nRow)
                Set chartObj = wsChart.ChartObjects.Add(Left:=wsChart.Range("H2").Left, Top:=wsChart.Range("H2").Top, Width:=375, Height:=225)
                    With chartObj.Chart
                        .ChartType = xlLine ' 折線圖
                        With .SeriesCollection.NewSeries
                            .Name = "cumPL" ' 數據系列名稱
                            .XValues = rng.Columns(1) ' 設定 X 軸數據
                            .Values = rng.Columns(4) ' 設定 Y 軸數據
                        End With
                        
                        
                        With .SeriesCollection.NewSeries
                            .Name = "DrawDown" ' 區域圖系列名稱
                            .XValues = rng.Columns(1) ' 使用相同的 X 軸數據
                            .Values = rng.Columns(5) ' 使用相同的 Y 軸數據
                            .ChartType = xlArea ' 設定為區域圖
                        End With
                                        
                        
                        ' 設置圖表標題
                        .HasTitle = True
                        .ChartTitle.Text = "損益折線圖"
                        
                        ' 設置 X 軸和 Y 軸標題
                        .Axes(xlCategory, xlPrimary).HasTitle = False
'                        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Date"
                        .Axes(xlCategory, xlPrimary).TickLabelPosition = xlLow
                        
                        .Axes(xlValue, xlPrimary).HasTitle = False
'                        .Axes(xlValue, xlPrimary).AxisTitle.Text = "cumPL"
                        
    '                    設置圖例
                        .HasLegend = True
                        .Legend.Position = xlLegendPositionTop
                        
                    End With
                Set chartObj = Nothing
            Set rng = Nothing
        Set ws = Nothing
    Set wsChart = Nothing
End Sub


Sub deleteCharts()
    Dim wsChart As Worksheet
    Dim i As Integer

    
    Set wsChart = ThisWorkbook.Worksheets(1)

    For i = wsChart.ChartObjects.Count To 1 Step -1
        wsChart.ChartObjects(i).Delete
    Next i
    Set wsChart = Nothing
End Sub


Sub calReturnYrs()
    Dim rng As Range
    Dim nRow As Long
    Dim nCol As Long
    
    Set rng = ThisWorkbook.Worksheets(2).Range("A1").CurrentRegion.Columns(1)
        rng.Copy
        ThisWorkbook.Worksheets(3).Range("A1").PasteSpecial xlPasteValues
    Set rng = Nothing
    
    
    Set rng = ThisWorkbook.Worksheets(3).Range("A1").CurrentRegion
        rng.RemoveDuplicates Columns:=1, Header:=xlYes
    Set rng = Nothing
    
'    計算每日需使用資金
    Set rng = ThisWorkbook.Worksheets(3).Range("A1").CurrentRegion
        nRow = rng.Rows.Count
        nCol = rng.Columns.Count + 1
        rng.Cells(1, nCol) = "使用資金"
        rng.Cells(2, nCol).Formula = "=SUMIFS(result!E:E,result!A:A,Est_iniCap!A2)"
        rng.Cells(2, nCol).Copy Destination:=rng.Range(rng.Cells(2, nCol), rng.Cells(nRow, nCol))
        
        
        
    Set rng = Nothing
    
'    計算每日損益 + 累積損益 + DD
    Set rng = ThisWorkbook.Worksheets(3).Range("A1").CurrentRegion
        nRow = rng.Rows.Count
        nCol = rng.Columns.Count + 1
        rng.Cells(1, nCol) = "損益"
        rng.Cells(2, nCol).Formula = "=SUMIFS(result!F:F,result!A:A,Est_iniCap!A2)"
        rng.Cells(2, nCol).Copy Destination:=rng.Range(rng.Cells(2, nCol), rng.Cells(nRow, nCol))
        
        nCol = rng.Columns.Count + 2
        rng.Cells(1, nCol) = "累積損益"
        rng.Cells(2, nCol).Formula = "=" & rng.Cells(2, nCol - 1).Address(False, False)
        rng.Cells(3, nCol).Formula = "=" & rng.Cells(3, nCol - 1).Address(False, False) & "+" & rng.Cells(2, nCol).Address(False, False)
        rng.Cells(3, nCol).Copy Destination:=rng.Range(rng.Cells(3, nCol), rng.Cells(nRow, nCol))
    
        nCol = rng.Columns.Count + 3
        rng.Cells(1, nCol) = "最大拉回"
        rng.Cells(2, nCol).Formula = "=" & rng.Cells(2, nCol - 1).Address(False, False) & "-" & "MAX(" & rng.Cells(2, nCol - 1).Address(False, False) & ":" & rng.Cells(2, nCol - 1).Address(True, True) & ")"
        rng.Cells(2, nCol).Copy Destination:=rng.Range(rng.Cells(2, nCol), rng.Cells(nRow, nCol))
    
    Set rng = Nothing
    
'    計算每日累績報酬率
    Set rng = ThisWorkbook.Worksheets(3).Range("A1").CurrentRegion
        Dim iniCap As Double
        
        nRow = rng.Rows.Count
        nCol = rng.Columns.Count + 1
        iniCap = WorksheetFunction.Max(rng.Columns(nCol - 4))
        
        rng.Cells(1, nCol) = "報酬率(%)"
        rng.Cells(2, nCol).Formula = "=" & rng.Cells(2, nCol - 2).Address(False, False) & "/" & iniCap
        rng.Cells(2, nCol).Copy Destination:=rng.Range(rng.Cells(2, nCol), rng.Cells(nRow, nCol))
            
    Set rng = Nothing
    
'    計算DD(%)
    Set rng = ThisWorkbook.Worksheets(3).Range("A1").CurrentRegion
        
        nRow = rng.Rows.Count
        nCol = rng.Columns.Count + 1
        
        rng.Cells(1, nCol) = "最大拉回(%)"
        rng.Cells(2, nCol).Formula = "=" & rng.Cells(2, nCol - 1).Address(False, False) & "-" & "MAX(" & rng.Cells(2, nCol - 1).Address(False, False) & ":" & rng.Cells(2, nCol - 1).Address(True, True) & ")"
        rng.Cells(2, nCol).Copy Destination:=rng.Range(rng.Cells(2, nCol), rng.Cells(nRow, nCol))
        
    Set rng = Nothing
    
    Set rng = ThisWorkbook.Worksheets(3).Range("A1").CurrentRegion
        rng.Copy
        rng.Range("A1").PasteSpecial xlPasteValues
    Set rng = Nothing
End Sub
