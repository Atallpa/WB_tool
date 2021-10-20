Attribute VB_Name = "transferData"
Sub transferDataToPFRAM()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim i, j, k, L As Single
    Dim dataArr, rngArr As Variant
    Dim str, rngStr    As String
    
    
    'AUTO CALCULATING REVENUES
    
    If Sheets("Pivot_Data").Range("Stor_Howtheprojectisfunded") = "Tax payers--public funded" And Sheets("Pivot_Data").Range("Stor_PROJECTIONOFREVENUES") = "Auto" Then Calc_Row1
    
    'OPERATION AND MAINTENANCE COSTS
    
    rngStr = "Stor_Maintenance;Stor_Operation;Stor_Userfeesforgovernment;Stor_Royalties;Stor_Otherpaymentstogovernment;Stor_Othercosts;Stor_Revenueguarantees" 'String with range names
    rngArr = Split(rngStr, ";")                   'Array with ranges
    
    For j = LBound(rngArr) To UBound(rngArr)
        str = Range(rngArr(j))                    'String with data
        dataArr = Split(str, ";")                 'Array with ranges and data
        For i = LBound(dataArr) To UBound(dataArr)
            
            If j = UBound(rngArr) Then
                Worksheets("Template").Cells(82, 3 + i) = dataArr(i)
            Else
                Worksheets("Template").Cells(71 + j, 3 + i) = dataArr(i)
            End If
            
        Next i
        PctDone = j / 33
        ProgressBarForm.Label1.Caption = "Operation And Maintenance Costs"
        Call UpdateProgress(PctDone)
    Next j
    
    'OTHER ADJUSTMENT FACTOR PROJECTION OF USER FEES - UNITARY PRICE
    
    Dim oafArr(1 To 20) As Variant
    Dim oafStr As String
    
    For k = 1 To 20
        If k > 10 Then
            oafArr(k) = "Stor_Demand" & k - 10 & "UserOtherAdjustmentfactor"
        Else
            oafArr(k) = "Stor_Price" & k & "UserOtherAdjustmentfactor" 'Array with ranges
        End If
        PctDone = 33 + (k / 66)
        ProgressBarForm.Label1.Caption = "Other Adjustment Factor Projection Of User Fees"
        Call UpdateProgress(PctDone)
        
    Next k
    
    For j = LBound(oafArr) To UBound(oafArr)
        oafStr = Range(oafArr(j))                 'String with data
        dataArr = Split(oafStr, ";")              'Array with ranges and data
        
        For i = LBound(dataArr) To UBound(dataArr)
            If j > 11 Then
                Worksheets("Template").Cells(390 + ((j - 11) * 6), 4 + i) = dataArr(i)
            Else
                Worksheets("Template").Cells(310 + ((j - 1) * 8), 4 + i) = dataArr(i)
            End If
        Next i
        PctDone = 0.18 + (j / 100)
        Call UpdateProgress(PctDone)
    Next j
    
    'OTHER ADJUSTMENT FACTOR PROJECTION OF GOVERNMENT - UNITARY PRICE
    
    
    For k = 1 To 20
        If k > 10 Then
            oafArr(k) = "Stor_Demand" & k - 10 & "GovOtherAdjustmentfactor"
        Else
            oafArr(k) = "Stor_Price" & k & "GovOtherAdjustmentfactor" 'Array with ranges
        End If
        PctDone = (0.38) + (k / 150)
        Call UpdateProgress(PctDone)
    Next k
    
    For j = LBound(oafArr) To UBound(oafArr)
        oafStr = Range(oafArr(j))                 'String with data
        dataArr = Split(oafStr, ";")              'Array with ranges and data
        
        For i = LBound(dataArr) To UBound(dataArr)
            If j > 11 Then
                Worksheets("Template").Cells(549 + ((j - 11) * 6), 4 + i) = dataArr(i)
            Else
                Worksheets("Template").Cells(469 + ((j - 1) * 8), 4 + i) = dataArr(i)
            End If
        Next i
        PctDone = (0.51) + (j / 41)
        ProgressBarForm.Label1.Caption = "Other Adjustment Factor Projection Of Government"
        Call UpdateProgress(PctDone)
    Next j
    
    EIR_SeekTemplate
    
    Unload ProgressBarForm
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub Calc_Row1()
    Worksheets("Template").Range("C78").FormulaR1C1 = "=R[540]C[1]"
    Worksheets("Template").Range("C78").AutoFill Destination:=Worksheets("Template").Range("C78:AZ78"), Type:=xlFillDefault
End Sub

Sub EIR_SeekTemplate()
    On Error Resume Next
    If ThisWorkbook.Worksheets("Template").Range("T37").Value = 0 Then
        ThisWorkbook.Worksheets("Template").Range("T38").Copy
        ThisWorkbook.Worksheets("Template").Range("WACC").PasteSpecial xlPasteValues
        Application.CutCopyMode = False
    Else
        ThisWorkbook.Worksheets("Template").Range("WACC") = ""
        ThisWorkbook.Worksheets("Template").Range("T36").GoalSeek Goal:=Range("T37").Value, ChangingCell:=Range("WACC")
    End If
   
End Sub

