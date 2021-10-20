Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
    
        
    Application.ScreenUpdating = False
    Dim cell, rng As Range
    Dim i As Integer
    
    Set rng = Selection
    i = 1
    
    For Each cell In rng
        If cell.Font.Size = 24 And cell <> "" Then
            ActiveSheet.Shapes.Range(Array("Arr_Down")).Select
            Selection.Copy
            cell.Offset(0, -1).Select
            ActiveSheet.Paste
            With Selection
                .Left = ActiveCell.Offset(0, 1).Left - (Selection.Width / 2)
                .Top = ((ActiveCell.Height / 2) + ActiveCell.Top) - Selection.Height / 2
                .Name = "Arr_Down" & i
                i = i + 1
            End With
        End If
    Next
    
    i = 1
    
    For Each cell In rng
        If cell.Font.Size = 14 And cell <> "" Then
            ActiveSheet.Shapes.Range(Array("Arr_Left")).Select
            Selection.Copy
            cell.Offset(0, -1).Select
            ActiveSheet.Paste
            With Selection
                .Left = ActiveCell.Offset(0, 1).Left - (Selection.Width / 2)
                .Top = ((ActiveCell.Height / 2) + ActiveCell.Top) - Selection.Height / 2
                .Name = "Arr_Left" & i
                i = i + 1
            End With
        End If
    Next
    Application.ScreenUpdating = True
    
End Sub

Sub addChkBox()

    Dim chk As Object
    Dim j, rng As Range
    Dim i As Integer
    
    Set rng = Selection
            
    Dim h As Double
    
    h = ActiveCell.Height
    
    i = 1
    For Each j In rng
        j.RowHeight = h
        Set chk = ActiveSheet.CheckBoxes.Add(354.5, 3710, 72, 72) 'ActiveSheet.CheckBoxes.Add(354.5, 3710, 72, 72).Select
        With chk
            .Name = "vfm_RPChk_" & i
            .Caption = ""
            .Height = h
            .Width = 20
            .Left = j.Left + (j.Width / 2) - (chk.Width / 2)
            .Top = j.Top + (h / 2) - (chk.Height / 2)
            '.LinkedCell = j.Address
            .Value = True
            i = i + 1
        End With
    Next j
    
End Sub
