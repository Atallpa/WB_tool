Attribute VB_Name = "Module4"
Sub Menuvfm1_Click()
'LIGHT---------------
    For i = 1 To 4
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
'DARK-----------------
    
    ActiveSheet.Shapes.Range(Array("Menuvfm1")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menuvfm1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    VfM_P1
    
End Sub
Sub Menuvfm2_Click()
'LIGHT---------------
    For i = 1 To 4
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
'DARK-----------------
    
    ActiveSheet.Shapes.Range(Array("Menuvfm2")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menuvfm2")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    VfM_P8
    
End Sub

Sub Menuvfm3_Click()
'LIGHT---------------
    For i = 1 To 4
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
'DARK-----------------
    
    ActiveSheet.Shapes.Range(Array("Menuvfm3")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menuvfm3")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    VfM_P10
    
End Sub

Sub Menuvfm4_Click()
'LIGHT---------------
    For i = 1 To 4
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
'DARK-----------------
    
    ActiveSheet.Shapes.Range(Array("Menuvfm4")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menuvfm4")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    VfM_P12
    
End Sub
