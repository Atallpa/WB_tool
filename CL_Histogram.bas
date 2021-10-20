Attribute VB_Name = "CL_Histogram"
Function Histograms(project As Single, year As Single)

    Dim Minval As Single
    Dim Maxval As Single
    Dim Step   As Single
    Dim hist(20, 3) As Single

    For h = 1 To iter
        If Data(project, year, h) < Minval Then
            Minval = Data(project, year, h)
        ElseIf Data(project, year, h) > Maxval Then
            Maxval = Data(project, year, h)
        End If
    Next h
    
    Step = (Maxval - Minval) / 18                 'number of bins minus 2 (the extremes)
    
    For h = 1 To 20
        hist(h, 1) = Minval + Step * (h - 1)
        hist(h, 2) = hist(h, 1) + Step
    Next h
    
    For h = 1 To iter
        For f = 1 To 20
            If Data(project, year, h) < hist(f, 2) Then
                hist(f, 3) = hist(f, 3) + 1
                GoTo continue2
            Else
                GoTo continue1
            End If
continue1:
        Next f
continue2:
    Next h
    
    
    Dim LBin(), HBin(), Freq(), Prob(), CProb() As Variant
    
    ReDim LBin(1 To 20), HBin(1 To 20), Freq(1 To 20), Prob(1 To 20), CProb(1 To 20)
    
    For h = 1 To 20
        LBin(h) = Format(hist(h, 1), "0,00")      'L bin
        HBin(h) = Format(hist(h, 2), "0,00")      'H bin
        Freq(h) = Format(hist(h, 3), "0,00")      'Freq
        Prob(h) = WorksheetFunction.IfError((Freq(h) / 1000), 0)
        If h = 1 Then CProb(1) = Prob(h) Else: CProb(h) = WorksheetFunction.IfError((Prob(h) + CProb(h - 1)), 0)
    Next h
    
    Dim myChart As Object: Set myChart = Sheets("ReportTemplate").ChartObjects("histo_chart")
    
    myChart.Activate
    
    On Error Resume Next
    With ActiveChart
        .FullSeriesCollection(2).Values = CProb()
        .FullSeriesCollection(1).Values = Prob()
        .FullSeriesCollection(1).XValues = LBin()
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0%"
        .FullSeriesCollection(1).Name = "=""Probability"""
        .FullSeriesCollection(2).Name = "=""C Probability"""
    End With
    
End Function

Function histogramsG(year As Single)

    Dim Minval As Single
    Dim Maxval As Single
    Dim Step   As Single
    Dim hist(20, 3) As Single

    For h = 1 To iter
        If GDP(year, h) < Minval Then
            Minval = GDP(year, h)
        ElseIf GDP(year, h) > Maxval Then
            Maxval = GDP(year, h)
        End If
    Next h
    
    Step = (Maxval - Minval) / 18                 'number of bins minus 2 (the extremes)
    
    For h = 1 To 20
        hist(h, 1) = Minval + Step * (h - 1)
        hist(h, 2) = hist(h, 1) + Step
    Next h
    
    For h = 1 To iter
        For f = 1 To 20
            If GDP(year, h) < hist(f, 2) Then
                hist(f, 3) = hist(f, 3) + 1
                GoTo continue2
            Else
                GoTo continue1
            End If
continue1:
        Next f
continue2:
    Next h
    
    For h = 1 To 20
        Sheets("GDP").Cells(22 + h, 24).Value = hist(h, 1)
        Sheets("GDP").Cells(22 + h, 25).Value = hist(h, 2)
        Sheets("GDP").Cells(22 + h, 26).Value = hist(h, 3)
    Next h
    
End Function
Public Function Sorting(project As Single)
    Dim aux(iter) As Single

    For j = 1 To years
        
        For i = 1 To iter
            aux(i) = Data(project, j, i)
        Next i
        
        Call QuickSort(aux, LBound(aux), UBound(aux))
        
        For i = 1 To iter
            Data(project, j, i) = aux(i)
        Next i
        
    Next j
    
End Function

Public Function SortingG()
    Dim aux(iter) As Single

    For j = 1 To years
        
        For i = 1 To iter
            aux(i) = GDP(j, i)
        Next i
        
        Call QuickSort(aux, LBound(aux), UBound(aux))
        
        For i = 1 To iter
            GDP(j, i) = aux(i)
        Next i
        
    Next j
    
End Function

Sub QuickSort(arr, Lo As Long, Hi As Long)

    Dim varPivot As Variant
    Dim varTmp As Variant
    Dim tmpLow As Long
    Dim tmpHi  As Long
    tmpLow = Lo
    tmpHi = Hi
    varPivot = arr((Lo + Hi) \ 2)
    
    Do While tmpLow <= tmpHi
        
        Do While arr(tmpLow) < varPivot And tmpLow < Hi
            tmpLow = tmpLow + 1
        Loop
        
        Do While varPivot < arr(tmpHi) And tmpHi > Lo
            tmpHi = tmpHi - 1
        Loop
        
        If tmpLow <= tmpHi Then
            varTmp = arr(tmpLow)
            arr(tmpLow) = arr(tmpHi)
            arr(tmpHi) = varTmp
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
        
    Loop
    
    If Lo < tmpHi Then QuickSort arr, Lo, tmpHi
    If tmpLow < Hi Then QuickSort arr, tmpLow, Hi
    
End Sub


Sub DropDown1_Change()
    Dim year As Single
    year = Sheets("MRG").Cells(35, 19).Value
    Call Histograms(1, year)
End Sub
