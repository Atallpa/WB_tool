Attribute VB_Name = "CL_MRG"
Public Const iter = 1000
Public Const n = 1
Public Const years = 50
Public Data(n, years + 1, iter) As Single
Public GDP(years, iter) As Single

Sub MRG_CL()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'Minimum Revenue Guarantee - Monte Carlo Simulation, David Duarte dduarte@worldbank.org (2018)
    
    Dim bias   As Single
    bias = Sheets("MRG").Range("bias")
    
    'Get correlations
    Dim corr(1 To n) As Single
    For i = 1 To n
        corr(i) = Sheets("MRG").Cells(15 * i, 3).Value
    Next i
    
    'Get standard deviations
    Dim stand(1 To n) As Single
    For i = 1 To n
        stand(i) = Sheets("MRG").Cells(14 * i, 3).Value
    Next i
    
    'GDP data
    Dim Ge(1 To years) As Single
    Dim Gf(1 To years) As Single
    Dim Gr(1 To years) As Variant
    Dim Gs(1 To years) As Single
    Dim SD_GDP As Single
    '    SD_GDP = Sheets("GDP").Range("C7").Value
    '
    '    For j = 1 To years
    '        Ge(j) = Sheets("GDP").Cells(10, 3 + j).Value
    '        Gf(j) = Sheets("GDP").Cells(11, 3 + j).Value
    '        Gr(j) = Sheets("GDP").Cells(12, 3 + j).Value
    '    Next j
    '
    'Simulate revenues and calculate payments
    
    Dim Xg(n, years) As Single                    'guaranteed amount
    Dim Xf(n, years) As Single                    'forecasted amount
    Dim Xe(n, years) As Single                    'effective amount
    Dim Xs(n, years) As Single                    'simulated amount
    Dim Mu(n, years) As Variant
    
    Dim Ps(n, years) As Single                    'simulated payment
    'Dim Data(N, years, iter) As Single
    
    For i = 1 To n
        For j = 1 To years
            Xg(i, j) = Sheets("MRG").Cells(21 * i, 3 + j).Value
            Xf(i, j) = Sheets("MRG").Cells(19 * i, 3 + j).Value
            Xe(i, j) = Sheets("MRG").Cells(18 * i, 3 + j).Value
            Mu(i, j) = Sheets("MRG").Cells(20 * i, 3 + j).Value
        Next j
    Next i
    
    For h = 1 To iter
        
        'Random correlated shocks
        
        Dim GDP_Shocks(years) As Single
        Dim C_Shocks(n, years) As Single
        
        For i = 1 To years
            Randomize
            GDP_Shocks(i) = Rnd
            Randomize
            For j = 1 To n
                C_Shocks(j, i) = corr(j) * GDP_Shocks(i) + Sqr(1 - (corr(j) ^ 2)) * Rnd
            Next j
        Next i
        
        'GDP simulation
        '        For j = 1 To years
        '            If Ge(j) > 0 Then
        '                Gs(j) = Ge(j)
        '            Else
        '                If Gf(j) > 0 Then
        '                    Gs(j) = WorksheetFunction.Norm_Inv(GDP_Shocks(j), Gf(j), SD_GDP * Gf(j))
        '
        '                Else
        '                    Gs(j) = Gs(j - 1) * Exp(Gr(j) + (SD_GDP ^ 2) / 2 + SD_GDP * GDP_Shocks(j))
        '                End If
        '            End If
        '            GDP(j, h) = Gs(j)
        '        Next j
        
        For i = 1 To n
            For j = 1 To years
                
                ' Simulated revenue
                Randomize
                
                If Xf(i, j) > 0 Then
                    If Xe(i, j) > 0 Then
                        Xs(i, j) = Xe(i, j)
                    Else
                        If j = 1 Then
                            Xs(i, j) = WorksheetFunction.Norm_Inv(Rnd, Xf(i, j) * (1 - bias), stand(i) * Xf(i, j))
                        Else
                            If Xf(i, j - 1) = 0 Then
                                Xs(i, j) = WorksheetFunction.Norm_Inv(Rnd, Xf(i, j) * (1 - bias), stand(i) * j * Xf(i, j))
                            Else
                                Xs(i, j) = Xs(i, j - 1) * Exp(Mu(i, j) - ((stand(i) ^ 2) / 2) + stand(i) * C_Shocks(i, j))
                            End If
                        End If
                    End If
                End If
                
                ' Simulated payments
                
                If Xg(i, j) = 0 Then
                    Ps(i, j) = 0
                Else
                    Ps(i, j) = WorksheetFunction.Max(0, Xg(i, j) - Xs(i, j))
                End If
                
                'save data in the bix matrix
                Data(i, j, h) = Ps(i, j)
                
            Next j
        Next i
        
        Application.StatusBar = "Progress: " & h & " of iter: " & Format(h / iter, "0%")
    Next h
    
    ' present value
    Dim rf     As Single
    rf = Sheets("MRG").Range("Rf").Value
    For i = 1 To n
        For h = 1 To iter
            For j = 1 To years
                Data(i, years + 1, h) = Data(i, years + 1, h) + Data(i, j, h) / ((1 + rf) ^ j)
            Next j
        Next h
    Next i
    
    'Means and standard deviation
    Application.StatusBar = "Updating means and standard deviations"
    Dim m(n, years) As Single
    Dim SD(n, years) As Single
    Dim aux    As Single
    
    For i = 1 To n
        For j = 1 To years
            For h = 1 To iter
                aux = m(i, j)
                m(i, j) = m(i, j) + Data(i, j, h)
                SD(i, j) = SD(i, j) + ((h * Data(i, j, h) - aux) ^ 2) / (h * (h + 1))
            Next h
            m(i, j) = m(i, j) / iter
            SD(i, j) = SD(i, j) / iter
            SD(i, j) = Sqr(SD(i, j))
        Next j
    Next i
    
    For i = 1 To n
        
        For j = 1 To years
            If Xg(i, j) = 0 Then
                Sheets("MRG").Cells(25 * i, 3 + j).Value = ""
                Sheets("MRG").Cells(26 * i, 3 + j).Value = ""
            Else
                Sheets("MRG").Cells(25 * i, 3 + j).Value = m(i, j)
                Sheets("MRG").Cells(26 * i, 3 + j).Value = SD(i, j)
            End If
        Next j
    Next i
    
    'GDP mean and standard deviation
    
    '    Dim GDP_M(years) As Single
    '    Dim GDP_S(years) As Single
    '
    '    For j = 1 To years
    '        For h = 1 To iter
    '            aux = GDP_M(j)
    '            GDP_M(j) = GDP_M(j) + GDP(j, h)
    '            GDP_S(j) = GDP_S(j) + ((h * GDP(j, h) - aux) ^ 2) / (h * (h + 1))
    '        Next h
    '        GDP_M(j) = GDP_M(j) / iter
    '        GDP_S(j) = GDP_S(j) / iter
    '        GDP_S(j) = Sqr(GDP_S(j))
    '    Next j
    '
    '
    '    For j = 1 To years
    '        Sheets("GDP").Cells(15, 3 + j).Value = GDP_M(j)
    '        Sheets("GDP").Cells(16, 3 + j).Value = GDP_S(j)
    '    Next j
    
    
    
    Application.Calculation = xlCalculationAutomatic
    
    Call Histograms(1, 51)
    Call DropDown2_Change
    
    FanChart
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub FanChart()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Sorting (1)
    
    Dim arr1(), arr2(), arr3(), arr4(), arr5(), arr6(), arr7(), arr8(), arr9(), arr10(), arr11(), _
        varMid(), varMax(), varMin(), var10(), var20(), var30(), var40(), var50(), var60(), var70(), var80(), var90() As Variant
    
    For i = 1 To years
        
        ReDim Preserve arr1(i), arr2(i), arr3(i), arr4(i), arr5(i), arr6(i), arr7(i), arr8(i), arr9(i), arr10(i), arr11(i), _
            varMid(i), varMax(i), varMin(i), var10(i), var20(i), var30(i), var40(i), var50(i), var60(i), var70(i), var80(i), var90(i)
        
        arr1(i) = Data(1, i, 0)
        arr2(i) = Data(1, i, Round(0.1 * iter))
        arr3(i) = Data(1, i, Round(0.2 * iter))
        arr4(i) = Data(1, i, Round(0.3 * iter))
        arr5(i) = Data(1, i, Round(0.4 * iter))
        arr6(i) = Data(1, i, Round(0.5 * iter))
        arr7(i) = Data(1, i, Round(0.6 * iter))
        arr8(i) = Data(1, i, Round(0.7 * iter))
        arr9(i) = Data(1, i, Round(0.8 * iter))
        arr10(i) = Data(1, i, Round(0.9 * iter))
        arr11(i) = Data(1, i, iter)
        var50(i) = arr6(i)
        
        If arr6(i) = "" Then varMax(i) = "" Else:  varMax(i) = arr11(i) - arr10(i)
        If arr6(i) = "" Then var90(i) = "" Else:  var90(i) = arr10(i) - arr9(i)
        If arr6(i) = "" Then var80(i) = "" Else:  var80(i) = arr9(i) - arr8(i)
        If arr6(i) = "" Then var70(i) = "" Else:  var70(i) = arr8(i) - arr7(i)
        If arr6(i) = "" Then var60(i) = "" Else:  var60(i) = arr7(i) - arr6(i)
        If arr6(i) = "" Then var50(i) = "" Else:  var50(i) = arr6(i) - arr5(i)
        If arr6(i) = "" Then var40(i) = "" Else:  var40(i) = arr5(i) - arr4(i)
        If arr6(i) = "" Then var30(i) = "" Else:  var30(i) = arr4(i) - arr3(i)
        If arr6(i) = "" Then var20(i) = "" Else:  var20(i) = arr3(i) - arr2(i)
        If arr6(i) = "" Then var10(i) = "" Else:  var10(i) = arr2(i) - arr1(i)
        If arr6(i) = "" Then varMin(i) = "" Else:  varMin(i) = arr1(i)
        
    Next i
    
    Set DataRange = Worksheets("MRG").Range("D25:BA25")
    
    Dim coll   As New Collection
    Dim itm
    
    On Error Resume Next
    For Each cell In DataRange.Cells
        If cell.Value <> "" Then coll.Add cell.Offset(-1, 0).Value, CStr(cell.Value)
    Next cell
    
    Dim x      As Long, xArr() As Variant
    For Each itm In coll
        ReDim Preserve xArr(x)
        xArr(x) = itm
        x = x + 1
    Next
    
    Dim myChart As Object
    
    Set myChart = ActiveSheet.ChartObjects("fan_chart")
    
    myChart.Activate
    
    With ActiveChart
        .FullSeriesCollection(1).XValues = xArr()
        .FullSeriesCollection(1).Values = var50()
        .FullSeriesCollection(2).Values = var10()
        .FullSeriesCollection(3).Values = var20()
        .FullSeriesCollection(4).Values = var30()
        .FullSeriesCollection(5).Values = var40()
        .FullSeriesCollection(6).Values = var50()
        .FullSeriesCollection(7).Values = var60()
        .FullSeriesCollection(8).Values = var70()
        .FullSeriesCollection(9).Values = var80()
        .FullSeriesCollection(10).Values = var90()
        For x = 1 To 50
            .ChartGroups(1).FullCategoryCollection(x).IsFiltered = True
        Next x
        For x = 1 To coll.Count
            .ChartGroups(1).FullCategoryCollection(x).IsFiltered = False
        Next x
    End With
    
    Range("A1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub DropDown2_Change()

    Sorting (1)
    Dim x As Single
    x = Sheets("MRG").Cells(29, 5)
    Dim U(years) As Single
    Dim L(years) As Single
    
    For i = 1 To years
        U(i) = Data(1, i, Round(((x + 1) / 2) * iter))
        L(i) = Data(1, i, Round(((1 - x) / 2) * iter))
        
        If Sheets("MRG").Cells(25, 3 + i).Value = "" Then
            Sheets("MRG").Cells(30, 3 + i).Value = ""
            Sheets("MRG").Cells(31, 3 + i).Value = ""
        Else
            Sheets("MRG").Cells(30, 3 + i).Value = U(i)
            Sheets("MRG").Cells(31, 3 + i).Value = L(i)
        End If
    Next i
    
End Sub


