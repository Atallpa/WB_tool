Attribute VB_Name = "MacroRuns"
Option Explicit

Public ProjectOptionSelected
Public ProjectListed
Public ProjectDelete
Public ProjectCreate_Edit

Dim i          As Long
Dim cell       As Range

Private Sub ob_YesShare_Click()

    Application.ScreenUpdating = False
    ActiveSheet.Shapes.Range(Array("ob_NoShare")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("ob_YesShare")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("ob_YesShare")).Fill.Visible = msoFalse
        Rows("170:172").EntireRow.Hidden = True
    Else
        ActiveSheet.Shapes.Range(Array("ob_YesShare")).Fill.Visible = msoTrue
        Rows("170:172").EntireRow.Hidden = False
    End If
    Range("Coll_Governmentshareholding").Value = "=E170"
    Application.ScreenUpdating = True
    
End Sub

Private Sub ob_NoShare_Click()

    Application.ScreenUpdating = False
    ActiveSheet.Shapes.Range(Array("ob_YesShare")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("ob_NoShare")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("ob_NoShare")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("ob_NoShare")).Fill.Visible = msoTrue
        Rows("170:172").EntireRow.Hidden = True
    End If
    
    Range("E170").Value = 0
    Range("AA1").Select
    Application.ScreenUpdating = True
    
End Sub
Private Sub ob_YesOtherPymnt_Click()
    Application.ScreenUpdating = False
    
    If Range("Coll_OtherGovPaymnt").Value = "YES" Then Exit Sub
    Range("Coll_OtherGovPaymnt").Value = "YES"
    
    ActiveSheet.Shapes.Range(Array("ob_NoOtherPymnt")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("ob_YesOtherPymnt")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("ob_YesOtherPymnt")).Fill.Visible = msoFalse
        Rows("418:422").EntireRow.Hidden = True
    Else
        ActiveSheet.Shapes.Range(Array("ob_YesOtherPymnt")).Fill.Visible = msoTrue
        Rows("418:422").EntireRow.Hidden = False
        Range("yes_no_OP") = "True"
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub ob_NoOtherPymnt_Click()
    Application.ScreenUpdating = False
    
    If Range("Coll_OtherGovPaymnt").Value = "NO" Then Exit Sub
    Range("Coll_OtherGovPaymnt").Value = "NO"
    
    ActiveSheet.Shapes.Range(Array("ob_YesOtherPymnt")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("ob_NoOtherPymnt")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("ob_NoOtherPymnt")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("ob_NoOtherPymnt")).Fill.Visible = msoTrue
        Rows("418:422").EntireRow.Hidden = True
        Range("yes_no_OP") = "False"
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub OB_DB_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoTrue
    End If
    
    If Range("Coll_Typeofproject").Value = "DB" Then Exit Sub
    Range("Coll_Typeofproject").Value = "DB"
    
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub


Private Sub OB_DBFO_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If Range("Coll_Typeofproject").Value = "DBFO" Then Exit Sub
    Range("Coll_Typeofproject").Value = "DBFO"
    
    ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoTrue
    End If
    
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub


Private Sub OB_BOT_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If Range("Coll_Typeofproject").Value = "BOT" Then Exit Sub
    Range("Coll_Typeofproject").Value = "BOT"
    
    ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoTrue
    End If
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub


Private Sub OB_BBO_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If Range("Coll_Typeofproject").Value = "BBO" Then Exit Sub
    Range("Coll_Typeofproject").Value = "BBO"
    
    ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoTrue
    End If
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub


Private Sub OB_BOO_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If Range("Coll_Typeofproject").Value = "BOO" Then Exit Sub
    Range("Coll_Typeofproject").Value = "BOO"
    
    ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoTrue
    End If
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub


Private Sub SpinnerAssetsNum_Change()

    Application.ScreenUpdating = False
    
    Dim rng    As Long
    
    SPAsset1_Click
    
    UnhideAll
    
    ActiveSheet.Shapes.Range(Array("SPAsset1")).Visible = True
    
    For i = 2 To 10
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Visible = False
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Left = ActiveSheet.Shapes.Range(Array("SPAsset" & i - 1)).Left + ActiveSheet.Shapes.Range(Array("SPAsset" & i - 1)).Width
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Height = ActiveSheet.Shapes.Range(Array("SPAsset" & i - 1)).Height
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Top = ActiveSheet.Shapes.Range(Array("SPAsset" & i - 1)).Top
    Next i
    
    rng = Range("D205").Value
    
    For i = 1 To (0 + rng)
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Visible = True
    Next i
    
    Range(Cells(209, 69), Cells(213, 69 - (9 - rng))).ClearContents
    
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    Range("Currency").EntireRow.Hidden = True
    Range("General_Info").EntireRow.Hidden = True
    Range("Project_Financed").EntireRow.Hidden = True
    Range("Step7").EntireRow.Hidden = True
    Range("endLeft").EntireColumn.Hidden = True
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub SpinnerServiceNum_Change()

    Dim rng    As Long

    ActiveSheet.Shapes.Range(Array("SPSer1")).Visible = True
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    For i = 1 To 10
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Visible = False
    Next i
    
    rng = Range("Amount_of_services").Value
    
    For i = 1 To (0 + rng)
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Visible = True
    Next i
    
    On Error Resume Next
    If rng <> 10 Then
        Range(Cells(312, 150), Cells(312 - (9 - rng), 156)).ClearContents
        Range(Cells(324, 151), Cells(324 - (9 - rng), 153)).ClearContents
        Range(Cells(324, 155), Cells(324 - (9 - rng), 156)).ClearContents
        Range(Cells(334, 151), Cells(334 - (9 - rng), 153)).ClearContents
        Range(Cells(334, 155), Cells(334 - (9 - rng), 156)).ClearContents
    End If
    
End Sub

Private Sub SPAsset1_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset1")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 1
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    
    Application.ScreenUpdating = True
End Sub
Private Sub SPAsset2_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset2")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 2
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub SPAsset3_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset3")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 3
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    Application.ScreenUpdating = True
    
End Sub
Private Sub SPAsset4_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset4")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 4
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    Application.ScreenUpdating = True
    
End Sub
Private Sub SPAsset5_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset5")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 5
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    Application.ScreenUpdating = True
    
End Sub
Private Sub SPAsset6_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset6")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 6
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    Application.ScreenUpdating = True
    
End Sub

Private Sub SPAsset7_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset7")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 7
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    Application.ScreenUpdating = True
    
End Sub
Private Sub SPAsset8_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset8")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 8
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    Application.ScreenUpdating = True
    
End Sub
Private Sub SPAsset9_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset9")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 9
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    Application.ScreenUpdating = True
    
End Sub
Private Sub SPAsset10_Click()

    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 1 To 10
        
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Select
        
        With Selection
            .Height = Cells(208, 1).Height - 2
            .Top = Cells(208, 1).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPAsset170")).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    
    Dim myArray As Variant
    Dim myString As String
    Dim cell   As Range
    Dim x      As Long
    
    myString = "D211;D215;D219;D224;D228"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        Cells(209 + x, 59).Offset(0, Range("CurrentAsset")) = Range(myArray(x)).Value
    Next x
    
    Range("CurrentAsset") = 10
    
    For x = LBound(myArray) To UBound(myArray)
        Range(myArray(x)).Value = Cells(209 + x, 59).Offset(0, Range("CurrentAsset"))
    Next x
    
    Range("D211").Select
    
    '***********
    Application.ScreenUpdating = True
    
End Sub

Private Sub Gov_funded_Click()

    Application.ScreenUpdating = False
    
    If Range("Coll_Howtheprojectisfunded").Value = "Tax payers--public funded" Then Exit Sub
    
    ActiveSheet.Shapes.Range(Array("Use_FundIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    ActiveSheet.Shapes.Range(Array("CombIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    ActiveSheet.Shapes.Range(Array("Gov_FundIco")).Select
    With Selection.ShapeRange.Fill
        If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Else
            .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
        End If
    End With
    
    Range("Coll_Howtheprojectisfunded").Value = "Tax payers--public funded"
    
    User_fund_Click
    For i = 2 To 6
        ActiveSheet.Shapes.Range(Array("CB_" & i)).Select: Selection.Value = xlOff
    Next
    User_fund_Click
    
    ActiveSheet.Shapes.Range(Array("CB_4")).Select: Selection.Value = xlOn
    ActiveSheet.Shapes.Range(Array("CB_6")).Select: Selection.Value = xlOn
    
    Range("A348,A354").EntireRow.Hidden = False
    Range("A347,A353").EntireRow.Hidden = True
    Range("OF_PriceUF").ClearContents
    Range("OF_DemandUF").ClearContents
    Range("OF_PriceGov").ClearContents
    Range("OF_DemandGov").ClearContents
    Range("BN327:BO328,BN332:BO333,BU327:BV328,BU332:BV333").ClearContents
    
    Range("AA1").Select
    
    ActiveSheet.Shapes.Range(Array("Gov_fund")).Visible = False
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub User_funded_Click()

    Application.ScreenUpdating = False
    
    If Range("Coll_Howtheprojectisfunded").Value = "Users of the services--user funded" Then Exit Sub
    
    ActiveSheet.Shapes.Range(Array("Gov_FundIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    ActiveSheet.Shapes.Range(Array("CombIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    ActiveSheet.Shapes.Range(Array("Use_FundIco")).Select
    
    With Selection.ShapeRange.Fill
        
        If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Else
            .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
        End If
    End With
    
    Range("Coll_Howtheprojectisfunded").Value = "Users of the services--user funded"
    
    Range("A348,A354").EntireRow.Hidden = True
    Range("A347,A353").EntireRow.Hidden = False
    Range("OF_PriceGov").ClearContents
    Range("OF_DemandGov").ClearContents
    Range("BN327:BO328,BN332:BO333,BU327:BV328,BU332:BV333").ClearContents
    
    Gov_fund_Click
    For i = 2 To 6
        ActiveSheet.Shapes.Range(Array("CB_" & i)).Select: Selection.Value = xlOff
    Next
    User_fund_Click
    
    ActiveSheet.Shapes.Range(Array("CB_4")).Select: Selection.Value = xlOn
    ActiveSheet.Shapes.Range(Array("CB_6")).Select: Selection.Value = xlOn
    
    Range("AA1").Select
    
    ActiveSheet.Shapes.Range(Array("Gov_fund")).Visible = False
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub Comb_funded_Click()

    Application.ScreenUpdating = False
    
    If Range("Coll_Howtheprojectisfunded").Value = "Combined" Then Exit Sub
    
    ActiveSheet.Shapes.Range(Array("Gov_FundIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    ActiveSheet.Shapes.Range(Array("Use_FundIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    ActiveSheet.Shapes.Range(Array("CombIco")).Select
    
    With Selection.ShapeRange.Fill
        
        If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Else
            .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
        End If
    End With
    
    Range("Coll_Howtheprojectisfunded").Value = "Combined"
    
    Range("A348,A354").EntireRow.Hidden = False
    Range("A347,A353").EntireRow.Hidden = False
    
    Range("OF_PriceUF").ClearContents
    Range("OF_DemandUF").ClearContents
    Range("OF_PriceGov").ClearContents
    Range("OF_DemandGov").ClearContents
    Range("BN327:BO328,BN332:BO333,BU327:BV328,BU332:BV333").ClearContents
    
    Range("AA1").Select
    
    ActiveSheet.Shapes.Range(Array("Gov_fund")).Visible = True
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub Loc_Currency_Click()

    Dim answer As Integer
    
    If Range("Currency_Portfolio") <> "Dom" Then
        answer = MsgBox("Foreign currency is used to represent macroeconomic parameters. For the tool to continue working properly, you have to convert all the quantitative information into foreign currency.", vbCritical + vbOKOnly, "Currency")
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
        
    ActiveSheet.Shapes.Range(Array("Fx_CurIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    ActiveSheet.Shapes.Range(Array("Local_CurIco")).Select
    
    With Selection.ShapeRange.Fill
        
        If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Else
            .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
        End If
    End With
    
    Range("Coll_Currency").Value = "Dom"
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub FX_Currency_Click()

    Dim answer As Integer
    
    If Range("Currency_Portfolio") = "Dom" Then
        answer = MsgBox("Local currency is used to represent macroeconomic parameters. For the tool to continue working properly, you have to convert all the quantitative information into local currency.", vbCritical + vbOKOnly, "Currency")
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
        
    ActiveSheet.Shapes.Range(Array("Local_CurIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    ActiveSheet.Shapes.Range(Array("Fx_CurIco")).Select
    
    With Selection.ShapeRange.Fill
        
        If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Else
            .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
        End If
    End With
    Range("Coll_Currency").Value = "FX"
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub EnterManual_Click()

    Application.ScreenUpdating = False
    
    If Range("Coll_CalculateRevenue").Value = "Manual" Then Exit Sub
    
    Range("Coll_CalculateRevenue").Value = "Manual"
    
    ActiveSheet.Shapes.Range(Array("CalcuIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    ActiveSheet.Shapes.Range(Array("EntManIco")).Select
    
    With Selection.ShapeRange.Fill
        
        If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Else
            .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
        End If
    End With
    
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub CB_4_Click()

    Dim Uf, Gf, co As String

    Uf = "Users of the services--user funded"
    Gf = "Tax payers--public funded"
    co = "Combined"
    
    Application.ScreenUpdating = False
    '
    If Range("A348").EntireRow.Hidden = True And Range("ES337") = "Users funded" Then Range("A345:A349").EntireRow.Hidden = True
    If Range("A345:A349").EntireRow.Hidden = True And Range("ES337") = "Users funded" Then Range("A345,A346,A347,A349").EntireRow.Hidden = False
    
    'UNITARY PRICE FOR SERVICE
    
    ' Government funded
    If Range("Serv_Funded") = Gf Then
        If Range("Serv_PriceOtherGov") = True Then
            Range("Serv_PriceOtherGov") = Range("Price_OFactor")
            Range("OF_PriceGov").ClearContents
            Range("345:349").EntireRow.Hidden = True
            
        ElseIf Range("Serv_PriceOtherGov") = False Then
            Range("Serv_PriceOtherGov") = Range("Price_OFactor")
            Range("345:349").EntireRow.Hidden = False
            Range("347:347").EntireRow.Hidden = True ' Opposite
        End If
    End If
    
    ' Users funded
    If Range("Serv_Funded") = Uf Then
        If Range("Serv_PriceOther") = True Then
            Range("Serv_PriceOther") = Range("Price_OFactor")
            Range("OF_PriceUF").ClearContents
            Range("345:349").EntireRow.Hidden = True
            
        ElseIf Range("Serv_PriceOther") = False Then
            Range("Serv_PriceOther") = Range("Price_OFactor")
            Range("345:349").EntireRow.Hidden = False
            Range("348:348").EntireRow.Hidden = True ' Opposite
        End If
    End If
    
    ' Combined funded
    If Range("Serv_Funded") = co Then
        If Range("UF") = True Then
            If Range("Serv_PriceOther") = True Then
                Range("Serv_PriceOther") = Range("Price_OFactor")
                Range("OF_PriceUF").ClearContents
                Range("345:349").EntireRow.Hidden = True
                
            ElseIf Range("Serv_PriceOther") = False Then
                Range("Serv_PriceOther") = Range("Price_OFactor")
                Range("345:349").EntireRow.Hidden = False
                Range("348:348").EntireRow.Hidden = True ' Opposite
            End If
        End If
        
        If Range("GF") = True Then
            If Range("Serv_PriceOtherGov") = True Then
                Range("Serv_PriceOtherGov") = Range("Price_OFactor")
                Range("OF_PriceGov").ClearContents
                Range("345:349").EntireRow.Hidden = True
                
            ElseIf Range("Serv_PriceOtherGov") = False Then
                Range("Serv_PriceOtherGov") = Range("Price_OFactor")
                Range("345:349").EntireRow.Hidden = False
                Range("347:347").EntireRow.Hidden = True ' Opposite
            End If
        End If
    End If
    
    Range("Ser_Name").Select
    
    Application.ScreenUpdating = True
    
End Sub
Private Sub CB_6_Click()


    Dim Uf, Gf, co As String

    Uf = "Users of the services--user funded"
    Gf = "Tax payers--public funded"
    co = "Combined"
    
    Application.ScreenUpdating = False
    '
    If Range("A354").EntireRow.Hidden = True And Range("ES337") = "Users funded" Then Range("A351:A355").EntireRow.Hidden = True
    If Range("A351:A355").EntireRow.Hidden = True And Range("ES337") = "Users funded" Then Range("A351:A355").EntireRow.Hidden = False
    
    'UNITARY Demand FOR SERVICE
    
    ' Government funded
    If Range("Serv_Funded") = Gf Then
        If Range("Serv_DemaOtherGov") = True Then
            Range("Serv_DemaOtherGov") = Range("Demand_OFactor")
            Range("OF_DemandGov").ClearContents
            Range("351:355").EntireRow.Hidden = True
            
        ElseIf Range("Serv_DemaOtherGov") = False Then
            Range("Serv_DemaOtherGov") = Range("Demand_OFactor")
            Range("351:355").EntireRow.Hidden = False
            Range("353:353").EntireRow.Hidden = True ' Opposite
        End If
    End If
    
    ' Users funded
    If Range("Serv_Funded") = Uf Then
        If Range("Serv_DemaOther") = True Then
            Range("Serv_DemaOther") = Range("Demand_OFactor")
            Range("OF_DemandUF").ClearContents
            Range("351:355").EntireRow.Hidden = True
            
        ElseIf Range("Serv_DemaOther") = False Then
            Range("Serv_DemaOther") = Range("Demand_OFactor")
            Range("351:355").EntireRow.Hidden = False
            Range("354:354").EntireRow.Hidden = True ' Opposite
        End If
    End If
    
    ' Combined funded
    If Range("Serv_Funded") = co Then
        If Range("UF") = True Then
            If Range("Serv_DemaOther") = True Then
                Range("Serv_DemaOther") = Range("Demand_OFactor")
                Range("OF_DemandUF").ClearContents
                Range("351:355").EntireRow.Hidden = True
                
            ElseIf Range("Serv_DemaOther") = False Then
                Range("Serv_DemaOther") = Range("Demand_OFactor")
                Range("351:355").EntireRow.Hidden = False
                Range("354:354").EntireRow.Hidden = True ' Opposite
            End If
        End If
        
        If Range("GF") = True Then
            If Range("Serv_DemaOtherGov") = True Then
                Range("Serv_DemaOtherGov") = Range("Demand_OFactor")
                Range("OF_DemandGov").ClearContents
                Range("351:355").EntireRow.Hidden = True
                
            ElseIf Range("Serv_DemaOtherGov") = False Then
                Range("Serv_DemaOtherGov") = Range("Demand_OFactor")
                Range("351:355").EntireRow.Hidden = False
                Range("353:353").EntireRow.Hidden = True ' Opposite
            End If
        End If
    End If
    
    Range("Ser_Name").Select
    
    Application.ScreenUpdating = True
    
End Sub


Private Sub CalculPFRAM_Click()

    Application.ScreenUpdating = False
    
    
    If Range("Coll_CalculateRevenue").Value = "Auto" Then Exit Sub
    
    Range("Coll_CalculateRevenue").Value = "Auto"
    
    ActiveSheet.Shapes.Range(Array("EntManIco")).Select
    Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    ActiveSheet.Shapes.Range(Array("CalcuIco")).Select
    
    With Selection.ShapeRange.Fill
        
        If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
            .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Else
            .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
        End If
    End With
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    
End Sub
'    ***************************************************
'    * ROWS AND COLUMN HIDDEN WHILE COMPLETTING SURVEY *
'    ***************************************************
Public Sub StartingSurvey()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    UnhideAll
    Worksheets("Project_Data").Select
    Range("Header").EntireRow.Hidden = True
    Range("Step1").EntireRow.Hidden = True
    Range("endLeft").EntireColumn.Hidden = True
    Range("Coll_Description").Select
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu1")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu1B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu1B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Range("AA1").Select
    ScrollToTop
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub Cont_ProjJust_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    
    Range("Header").EntireRow.Hidden = True
    Range("Step2").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("endLeft").EntireColumn.Hidden = True
    Range("AA1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub Cont_ProjFund_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("Step3").EntireRow.Hidden = True
    Range("endLeft").EntireColumn.Hidden = True
    Range("AA1").Select
    ScrollToTop
    
    '=========================
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Private Sub Back_ProjJust_Click()
    Cont_ProjJust_Click
End Sub
Private Sub Cont_ProjCurr_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    '=========================
    
    UnhideAll
    
    
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu1")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu1B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu1B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Range("Step4").EntireRow.Hidden = True
    Range("endLeft").EntireColumn.Hidden = True
    Range("AA1").Select
    ScrollToTop
    '=========================
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Private Sub Back_ProjFund_Click()
    Cont_ProjFund_Click
End Sub
Private Sub Cont_ProjGenInfo_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    '=========================
    
    UnhideAll
    
    
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    Range("Currency").EntireRow.Hidden = True
    
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu2")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu2")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu2B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu2B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Range("Step5_Menu").EntireRow.Hidden = True
    
    Range("Step5").EntireRow.Hidden = True
    
    Range("endLeft").EntireColumn.Hidden = True
    Range("AA1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub Back_ProjCurr_Click()
    Cont_ProjCurr_Click
End Sub
Private Sub Cont_ProjFinan_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    '=========================
    
    UnhideAll
    
    
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    Range("Currency").EntireRow.Hidden = True
    Range("General_Info").EntireRow.Hidden = True
    Range("Step6").EntireRow.Hidden = True
    Range("endLeft").EntireColumn.Hidden = True
    
    Range("AA1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Private Sub Back_ProjGenInfo_Click()
    Cont_ProjGenInfo_Click
End Sub
Private Sub Cont_AssetCharact_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    '=========================
    
    UnhideAll
    
    
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    Range("Currency").EntireRow.Hidden = True
    Range("General_Info").EntireRow.Hidden = True
    Range("Project_Financed").EntireRow.Hidden = True
    Range("Step7").EntireRow.Hidden = True
    Range("endLeft").EntireColumn.Hidden = True
    Range("AA1").Select
    '=========================
    
    ActiveSheet.Shapes.Range(Array("SpinnerAssetsNum")).Select
    With Selection
        .Value = 1
    End With
    
    For i = 1 To 10
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Visible = False
    Next i
    
    For i = 1 To Range("Coll_AmountOfAssets")
        ActiveSheet.Shapes.Range(Array("SPAsset" & i)).Visible = True
    Next i
    
    SPAsset1_Click
    ScrollToTop
    
    '=========================
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Private Sub Back_ProjFinan_Click()
    Cont_ProjFinan_Click
End Sub
Private Sub Cont_CalcFee_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    '=========================
    
    UnhideAll
    
    
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    Range("Currency").EntireRow.Hidden = True
    Range("General_Info").EntireRow.Hidden = True
    Range("Project_Financed").EntireRow.Hidden = True
    Range("Characteristic_Asset").EntireRow.Hidden = True
    
    Range("Step8_Menu").EntireRow.Hidden = True
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu3")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu3")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu3B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu3B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Range("Step8").EntireRow.Hidden = True
    
    Range("endLeft").EntireColumn.Hidden = True
    
    SPAsset1_Click
    YoungestProyect
    
    Range("AA1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Private Sub Back_AssetCharact_Click()
    Cont_AssetCharact_Click
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu2")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu2")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu2B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu2B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
End Sub

Private Sub Cont_AnnualRev_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    '=========================
    
    If Range("Coll_CalculateRevenue").Value = "Manual" Then
        
        UnhideAll
        
        Range("Header").EntireRow.Hidden = True
        Range("Description").EntireRow.Hidden = True
        Range("Justification").EntireRow.Hidden = True
        Range("project_funded").EntireRow.Hidden = True
        Range("Currency").EntireRow.Hidden = True
        Range("General_Info").EntireRow.Hidden = True
        Range("Project_Financed").EntireRow.Hidden = True
        Range("Characteristic_Asset").EntireRow.Hidden = True
        Range("Service_fees").EntireRow.Hidden = True
        
        
        Range("Step9").EntireRow.Hidden = True
        Rows("270:272").EntireRow.Hidden = True
        Range("endLeft2").EntireColumn.Hidden = True
        Range("endRight").EntireColumn.Hidden = True
        Range("AA1").Select
        SpinnerRevYear_Change
        
        'HIDE COMBINED ROW
        If Range("Coll_Howtheprojectisfunded") <> "Combined" Then
            Range("A281").EntireRow.Hidden = True
            Range("Annual_Revenue").ClearContents
        Else
            Range("A281").EntireRow.Hidden = False
        End If
        
    Else
        
        
        UnhideAll
        BorrarBarraService
        
        For i = 1 To 10
            ActiveSheet.Shapes.Range(Array("SPSer" & i)).Visible = False
        Next i
        
        For i = 1 To Range("Amount_of_services")
            ActiveSheet.Shapes.Range(Array("SPSer" & i)).Visible = True
        Next i
        
        SPSer1_Click
        
        AlignServiceAssetTab
        
        Range("Header").EntireRow.Hidden = True
        Range("Description").EntireRow.Hidden = True
        Range("Justification").EntireRow.Hidden = True
        Range("project_funded").EntireRow.Hidden = True
        Range("Currency").EntireRow.Hidden = True
        Range("General_Info").EntireRow.Hidden = True
        Range("Project_Financed").EntireRow.Hidden = True
        Range("Characteristic_Asset").EntireRow.Hidden = True
        Range("Service_fees").EntireRow.Hidden = True
        Range("Annual_revenue_projection").EntireRow.Hidden = True
        
        Range("Step10").EntireRow.Hidden = True
        
        Range("endLeft2").EntireColumn.Hidden = True
        Range("endRight").EntireColumn.Hidden = True
        Range("AA1").Select
    End If
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    
    ScrollToTop
    ActiveSheet.Shapes.Range(Array("Menu3")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu3")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu3B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu3B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Public Sub RiskMatrix()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu6")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu6")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu6B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu6B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Sheets("Risk_Matrix").Activate
    
    Matrix.Show
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub Back_CalcFee_Click()

    Sheets("Project_Data").Range("$B$5") = Sheets("Project_Data").Range("$EA$168")
    
    SPSer1_Click
    User_fund_Click
    Cont_CalcFee_Click
    
End Sub
Private Sub Cont_OandM_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    '=========================
    
    UnhideAll
    
    BorrarBarraOandM
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    Range("Currency").EntireRow.Hidden = True
    Range("General_Info").EntireRow.Hidden = True
    Range("Project_Financed").EntireRow.Hidden = True
    Range("Characteristic_Asset").EntireRow.Hidden = True
    Range("Service_fees").EntireRow.Hidden = True
    
    Range("Annual_revenue_projection").EntireRow.Hidden = True
    Range("Service_revenue").EntireRow.Hidden = True
    Range("MRG").EntireRow.Hidden = True
    Range("Debt_Guarantee").EntireRow.Hidden = True
    
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu4")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu4")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu4B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu4B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    
    If Range("OandM_M") = True Then Rows("455").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("455").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
    If Range("OandM_O") = True Then Rows("456").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("456").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
    If Range("OandM__UF") = True Then Rows("457").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("457").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
    If Range("OandM_R") = True Then Rows("458").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("458").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
    If Range("OandM__OP") = True Then Rows("459").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("459").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
    If Range("OandM_OC") = True Then Rows("460").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("460").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
    
    Sheets("Project_Data").Range("$B$5") = Sheets("Project_Data").Range("$EA$168")
    
    Range("Step13").EntireRow.Hidden = True
    Range("endLeft2").EntireColumn.Hidden = True
    Range("endRight").EntireColumn.Hidden = True
    
    Range("AA1").Select
    ScrollToTop
    '=========================
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub Back_OandM_Click()
    Cont_OandM_Click
End Sub

Private Sub Cont_RevGuar_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    '=========================
    
    UnhideAll
    
    
    BorrarBarraGuarantees
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    Range("Currency").EntireRow.Hidden = True
    Range("General_Info").EntireRow.Hidden = True
    Range("Project_Financed").EntireRow.Hidden = True
    Range("Characteristic_Asset").EntireRow.Hidden = True
    Range("Service_fees").EntireRow.Hidden = True
    
    Range("Annual_revenue_projection").EntireRow.Hidden = True
    Range("Service_revenue").EntireRow.Hidden = True
    
    Range("Step11").EntireRow.Hidden = True
    Range("endLeft2").EntireColumn.Hidden = True
    Range("endRight").EntireColumn.Hidden = True
    
    'LIGHT
    For i = 1 To 7
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu5")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu5")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveSheet.Shapes.Range(Array("Menu5B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu5B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Range("AA1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub Back_MRG_Click()
    Cont_RevGuar_Click
End Sub

Private Sub Cont_DebtGuar_Click()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    '=========================
    
    UnhideAll
    
    BorrarBarraOtherPayment
    Range("Header").EntireRow.Hidden = True
    Range("Description").EntireRow.Hidden = True
    Range("Justification").EntireRow.Hidden = True
    Range("project_funded").EntireRow.Hidden = True
    Range("Currency").EntireRow.Hidden = True
    Range("General_Info").EntireRow.Hidden = True
    Range("Project_Financed").EntireRow.Hidden = True
    Range("Characteristic_Asset").EntireRow.Hidden = True
    Range("Service_fees").EntireRow.Hidden = True
    
    Range("Annual_revenue_projection").EntireRow.Hidden = True
    Range("Service_revenue").EntireRow.Hidden = True
    Range("MRG").EntireRow.Hidden = True
    
    Range("Step12").EntireRow.Hidden = True
    Range("endLeft2").EntireColumn.Hidden = True
    Range("endRight").EntireColumn.Hidden = True
    
    If Range("yes_no_OP") = "True" Then Rows("418:422").EntireRow.Hidden = False Else: Rows("418:422").EntireRow.Hidden = True
    
    Range("AA1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub Cont_OandMAuto_Click()
    SPSer1_Click
    User_fund_Click
    Cont_OandM_Click
End Sub

Private Sub Back_CalcRev_Click()
    Cont_AnnualRev_Click
End Sub

Public Sub UnhideAll()
    Cells.Select
    Selection.EntireRow.Hidden = False
    Selection.EntireColumn.Hidden = False
End Sub

'**********************************
'    SERVICE SPIN-BUTTONS         *
'**********************************

Private Sub SPSer1_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 1
    
    On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub SPSer2_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 2
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Private Sub SPSer3_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 3
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Private Sub SPSer4_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 4
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub SPSer5_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 5
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub SPSer6_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 6
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub SPSer7_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 7
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub SPSer8_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 8
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub SPSer9_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 9
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub SPSer10_Click()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim a      As Integer: a = 10
    
    'On Error Resume Next
    For i = 1 To Range("Amount_of_services")
        
        ActiveSheet.Shapes.Range(Array("SPSer" & i)).Select
        
        With Selection
            .Height = Cells(307, 61).Height - 2
            .Top = Cells(307, 61).Top + 2
        End With
        
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        
    Next
    
    ' NEXT
    
    ActiveSheet.Shapes.Range(Array("SPSer" & a)).Select
    
    With Selection
        .ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ShapeRange.Fill.ForeColor.Brightness = -0.5
    End With
    
    Selection.ShapeRange.TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
    
    '***********
    Range("CurrentSer") = a
    '***********
    
    '******************************************************************
    '******************************************************************
    
    Dim KeyCells As Range
    
    Range("Ser_Name") = Range("Coll_Service" & Range("CurrentSer") & "Nameofservice") 'Service name
    
    Range("Ser_Unit") = Range("Coll_Service" & Range("CurrentSer") & "Unit") 'What unit of measure does this service use?
    
    Range("Ser_YearOPer") = Range("Coll_Service" & Range("CurrentSer") & "Startyear") 'When is the first year of operation of the service?  (This year should be from or after the construction of the asset is finished)
    
    'REVENUES SERVICE
    If Range("Serv_Funded") = "Tax payers--public funded" Then 'Gov. Funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
    ElseIf Range("Serv_Funded") = "Users of the services--user funded" Then 'Users Funded
        
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    ElseIf Range("Serv_Funded") = "Combined" Then 'Combined
        
        Range("GF_Price") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialPrice") 'Unit Price Government funded
        
        Range("GF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "GovInitialDemand") 'Demand Government funded
        
        Range("UF_Price") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialPrice") 'Unit Price Users funded
        
        Range("UF_Demand") = Range("Coll_Service" & Range("CurrentSer") & "UFInitialDemand") 'Demand Users funded
        
    End If
    
    User_fund_Click
    
    '================================================================
    
    Range("Amount_of_services").Select
    
    '******************************************************************
    '******************************************************************
    
    Saves_OtherAdj_Factor
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub SpinnerRevYear_Change()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim Y, x, n, j As Long
    Dim DataRange As Range
    
    Y = Range("Coll_Lengthofcontract").Value
    n = Range("Coll_Startyear")
    x = Range("Sp_number").Value
    
    ActiveSheet.Shapes.Range(Array("SpinnerRevYear")).Select
    With Selection
        .Min = 0
        .Max = Y
        .SmallChange = 1
        .LinkedCell = "$BJ$275"
    End With
    
    If x = Y + 2 Then x = 1
    
    '====
    
    Set DataRange = ActiveSheet.Range(Cells(178, 66 + x), Cells(179, 115))
    DataRange.Copy
    
    Cells(268, 66).PasteSpecial Paste:=xlPasteValues
    
    Range("Sp_numberLast").Value = x
    
    '====
    Range("revenues_to_copy").Select
    Selection.Copy
    
    Dim Anc2   As Range
    
    Set Anc2 = Sheets("Project_Data").Range("$BI$279")
    
    Anc2.Select
    ActiveSheet.Paste
    
    'Annual revenues
    
    If x <> 0 Then
        Range(Cells(279, 116 - x), Cells(283, 115)).Select
        Selection.Delete Shift:=xlToLeft
        '        Range(Cells(Range("AncorCell_Rev").Row - 1, Range("AncorCell_Rev").Column + 3 + (y - x + 2)), Cells(Range("AncorCell_Rev").Row + 3, Range("AncorCell_Rev").Column + 49 + 6)).Select
        '        Selection.Delete Shift:=xlToLeft
    End If
    
    Range("AnnualRev_Formulas").Select
    Selection.Copy
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    Range("AnnualRev_Formulas2").Select
    Selection.Copy
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    ActiveWorkbook.Names("Projected_Annual_Rev").RefersToR1C1 = "=Project_Data!R280C66:R282C115"
    ActiveWorkbook.Names("Annual_Revenue").RefersToR1C1 = "=Project_Data!R281C66:R281C115"
    ActiveWorkbook.Names("Annual_Revenue2").RefersToR1C1 = "=Project_Data!R282C66:R282C115"
    
    Range("YrRevenueStart").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub YoungestProyect()

    Dim serArr() As Variant
    Dim p
    
    If Cells(209, i) <> "" Then
        
        For i = 1 To 10
            If Cells(209, 59 + i) <> "" Then
                ReDim Preserve serArr(i)
                serArr(i) = Cells(210, 59 + i) + Cells(211, 59 + i)
            End If
        Next i
        On Error GoTo ErrorHandler
        p = Application.Match(Application.Min(serArr), serArr, 0) - 1
        
        Cells(215, 60) = Cells(209, 59 + p)
        Exit Sub
    Else
        Sheets("Project_Data").Range("$BH$215") = Sheets("Project_Data").Range("$BH$209")
    End If
    
ErrorHandler:
    Cells(215, 60) = "Empty"
    Exit Sub
End Sub

Private Sub OtherFactorPrice()

    Dim unitArr() As Variant
    Dim i      As Long
    Dim unitRng As Range: Set unitRng = Range("OtherFactorPrice")

    For Each cell In unitRng
        ReDim Preserve unitArr(i)
        If cell.Value <> "" Then
            unitArr(i) = cell.Value
            i = i + 1
        End If
    Next cell
    
    For i = LBound(unitArr) To UBound(unitArr)
        Cells(305 + Range("CurrentSer").Value, 99 + i) = unitArr(i)
    Next i
    
End Sub

Private Sub OtherFactorDemand()

    Dim unitArr() As Variant
    Dim i      As Long
    Dim unitRng As Range: Set unitRng = Range("OtherFactorDemand")

    For Each cell In unitRng
        ReDim Preserve unitArr(i)
        If cell.Value <> "" Then
            unitArr(i) = cell.Value
            i = i + 1
        End If
    Next cell
    
    For i = LBound(unitArr) To UBound(unitArr)
        Cells(315 + Range("CurrentSer").Value, 99 + i) = unitArr(i)
    Next i
    
End Sub

Function CONCATENATEMULTIPLE(Ref As Range, Separator As String) As String
    Dim cell   As Range
    Dim result As String

    For Each cell In Ref
        If cell = "" Then result = result & 0 & Separator
        If cell <> "" Then result = result & cell.Value & Separator
    Next cell
    
    CONCATENATEMULTIPLE = Left(result, Len(result) - 1)
End Function

Private Sub BorrarBarraService()

    Application.ScreenUpdating = False
    
    If ProjectOptionSelected = "Edit" Then
        Dim arr(0 To 3) As Variant
        arr(0) = Range("$EY$315")
        arr(1) = Range("$EZ$315")
        arr(2) = Range("$EY$325")
        arr(3) = Range("$EZ$325")
    End If
    
    Range("OF_DemandValue").Copy
    Range("AncorCell_OFP_StoreData").PasteSpecial xlPasteValues
    Range("OF_PriceValue").Copy
    Range("AncorCell_OFD_StoreData").PasteSpecial xlPasteValues
    
    Dim Y, x, n As Long
    
    Y = Range("Coll_Lengthofcontract").Value
    n = Range("Coll_Startyear")
    
    Range("TemplateToCopy_PriceAndDemand").Select
    Selection.Copy
    Range("AncorCell_Serv").Select
    ActiveSheet.Paste
    
    'Annual service other adjustment factor
    
    Range(Cells(Range("AncorCell_Serv").Row, Range("AncorCell_Serv").Column + 4 + (Y + 2)), Cells(Range("AncorCell_Serv").Row + 10, Range("AncorCell_Serv").Column + 50 + 4)).Select
    Selection.Delete Shift:=xlToLeft
    
    ActiveWorkbook.Names("OF_PriceUF").RefersToR1C1 = "=Project_Data!R347C66:R347C115"
    ActiveWorkbook.Names("OF_PriceGov").RefersToR1C1 = "=Project_Data!R348C66:R348C115"
    
    ActiveWorkbook.Names("OF_DemandUF").RefersToR1C1 = "=Project_Data!R353C66:R353C115"
    ActiveWorkbook.Names("OF_DemandGov").RefersToR1C1 = "=Project_Data!R354C66:R354C115"
    
    ActiveWorkbook.Names.Add Name:="OF_RangeValue", RefersToR1C1:="=Project_Data!R346C66:R354C115"
    
    Dim rng    As Range: Set rng = Range("BN346:DK346")
    For Each cell In rng
        If cell.HasFormula Then
            cell.Value = cell.Value
        End If
    Next cell
    
    Set rng = Range("BN352:DK352")
    For Each cell In rng
        If cell.HasFormula Then
            cell.Value = cell.Value
        End If
    Next cell
    
    If ProjectOptionSelected = "Edit" Then
        Range("$EY$315") = arr(0)
        Range("$EZ$315") = arr(1)
        Range("$EY$325") = arr(2)
        Range("$EZ$325") = arr(3)
    End If
    
    Range("Ser_Name").Select
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub BorrarBarraGuarantees()
    '
    '    Application.ScreenUpdating = False
    '    Application.Calculation = xlManual
    '    Application.DisplayAlerts = False
    '    Application.EnableEvents = False

    Range("GuaranteeValues").Copy
    Range("AncorCell_RG_StoreData").PasteSpecial xlPasteValues
    
    Dim Y, x, n As Long
    
    Y = Range("Coll_Lengthofcontract").Value
    n = Range("Coll_Startyear")
    
    Range("TemplateToCopy_Guarantee").Select
    Selection.Copy
    Range("AncorCell_Guar").Select
    ActiveSheet.Paste
    
    'Annual service other adjustment factor
    
    Range(Cells(Range("AncorCell_Guar").Row, Range("AncorCell_Guar").Column + 4 + (Y + 2)), Cells(Range("AncorCell_Guar").Row + 3, Range("AncorCell_Guar").Column + 50 + 4)).Select
    Selection.Delete Shift:=xlToLeft
    
    ActiveWorkbook.Names("Revenue_Guarantee").RefersToR1C1 = "=Project_Data!R387C66:R387C115"
    ActiveWorkbook.Names("RG_RangeValue").RefersToR1C1 = "=Project_Data!R386C66:R387C115"
    
    Dim rng    As Range: Set rng = Range("BN386:DK386")
    For Each cell In rng
        If cell.HasFormula Then
            cell.Value = cell.Value
        End If
    Next cell
    
    Range("Ser_Name").Select
    
    '    Application.ScreenUpdating = True
    '    Application.Calculation = xlAutomatic
    '    Application.DisplayAlerts = True
    '    Application.EnableEvents = True
End Sub

Private Sub BorrarBarraOtherPayment()
    '
    '    Application.ScreenUpdating = False
    '    Application.Calculation = xlManual
    '    Application.DisplayAlerts = False
    '    Application.EnableEvents = False

    Range("OGP_Values").Copy
    Range("AncorCell_OGP_StoreData").PasteSpecial xlPasteValues
    
    Dim Y, x, n As Long
    
    Y = Range("Coll_Lengthofcontract").Value
    n = Range("Coll_Startyear")
    
    Range("TemplateToCopy_OtherPaymnt").Select
    Selection.Copy
    Range("AncorCell_OtherPaymnt").Select
    ActiveSheet.Paste
    
    'Annual service other payment of governments
    
    Range(Cells(Range("AncorCell_OtherPaymnt").Row, Range("AncorCell_OtherPaymnt").Column + 4 + (Y + 2)), Cells(Range("AncorCell_OtherPaymnt").Row + 3, Range("AncorCell_OtherPaymnt").Column + 50 + 4)).Select
    Selection.Delete Shift:=xlToLeft
    
    ActiveWorkbook.Names("OGP").RefersToR1C1 = "=Project_Data!R421C66:R421C115"
    ActiveWorkbook.Names("OGP_RangeValue").RefersToR1C1 = "=Project_Data!R420C66:R421C115"
    
    Dim rng    As Range: Set rng = Range("BN420:DK420")
    For Each cell In rng
        If cell.HasFormula Then
            cell.Value = cell.Value
        End If
    Next cell
    
    Range("AA1").Select
    '
    '    Application.ScreenUpdating = True
    '    Application.Calculation = xlAutomatic
    '    Application.DisplayAlerts = True
    '    Application.EnableEvents = True
End Sub

Private Sub BorrarBarraOandM()

    '    Application.ScreenUpdating = False
    '    Application.Calculation = xlManual
    '    Application.DisplayAlerts = False
    '    Application.EnableEvents = False

    Range("OandM_StoreData").Copy
    Range("AncorCell_OandM_StoreData").PasteSpecial xlPasteValues
    
    Dim Y, x, n As Long
    
    Y = Range("Coll_Lengthofcontract").Value
    n = Range("Coll_Startyear")
    
    Range("TemplateToCopy_OandM").Select
    Selection.Copy
    Range("AncorCell_OandM").Select
    ActiveSheet.Paste
    
    'Annual service other adjustment factor
    
    Range(Cells(Range("AncorCell_OandM").Row, Range("AncorCell_OandM").Column + 2 + (Y + 4)), Cells(Range("AncorCell_OandM").Row + 8, Range("AncorCell_OandM").Column + 50 + 4)).Select
    Selection.Delete Shift:=xlToLeft
    
    '==========
    Dim myArray As Variant
    Dim myString As String
    
    myString = "BI445;BI447;BI449;BP445;BP447;BP449"
    
    myArray = Split(myString, ";")
    
    For x = LBound(myArray) To UBound(myArray)
        ActiveSheet.Shapes.Range(Array("CB_" & x + 7)).Left = Range(myArray(x)).Left + Range(myArray(x)).Width
        ActiveSheet.Shapes.Range(Array("CB_" & x + 7)).Top = Range(myArray(x)).Top + 1.5
    Next x
    
    '    ActiveSheet.Shapes.Range(Array("CB_7")).Select
    '    Selection.Value = xlOn
    '    ActiveSheet.Shapes.Range(Array("CB_8")).Select
    '    Selection.Value = xlOn
    
    Range("452:456,461:461").EntireRow.Hidden = False
    
    ActiveWorkbook.Names("OandM_Maint").RefersToR1C1 = "=Project_Data!R455C66:R455C115"
    ActiveWorkbook.Names("OandM_Oper").RefersToR1C1 = "=Project_Data!R456C66:R456C115"
    ActiveWorkbook.Names("OandM_OthCost").RefersToR1C1 = "=Project_Data!R460C66:R460C115"
    ActiveWorkbook.Names("OandM_OthPayToGov").RefersToR1C1 = "=Project_Data!R459C66:R459C115"
    ActiveWorkbook.Names("OandM_Roya").RefersToR1C1 = "=Project_Data!R458C66:R458C115"
    ActiveWorkbook.Names("OandM_UserF").RefersToR1C1 = "=Project_Data!R457C66:R457C115"
    ActiveWorkbook.Names("OandM_AnnualAmounts").RefersToR1C1 = "=Project_Data!R454C66:R460C115"
    '=========
    
    Dim rng    As Range: Set rng = Range("BN454:DK454")
    For Each cell In rng
        If cell.HasFormula Then
            cell.Value = cell.Value
        End If
    Next cell
    
    Range("BI452").Select
    '
    '    Application.ScreenUpdating = True
    '    Application.Calculation = xlAutomatic
    '    Application.DisplayAlerts = True
    '    Application.EnableEvents = True
End Sub

'============================
'    O&M OPTIONS            =
'============================
Private Sub CB_7_Click()
    If Range("OandM_M") = True Then Rows("455").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("455").EntireRow.Hidden = True: Range("BI451").Select
    '    If Range("OandM_O") = True Then Rows("456").EntireRow.Hidden = False
    '    If Range("OandM__UF") = True Then Rows("457").EntireRow.Hidden = False
    '    If Range("OandM_R") = True Then Rows("458").EntireRow.Hidden = False
    '    If Range("OandM__OP") = True Then Rows("459").EntireRow.Hidden = False
    '    If Range("OandM_OC") = True Then Rows("460").EntireRow.Hidden = False
    
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    'On Error Resume Next
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
End Sub
Private Sub CB_8_Click()
    If Range("OandM_O") = True Then Rows("456").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("456").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    'On Error Resume Next
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
End Sub
Private Sub CB_9_Click()
    If Range("OandM__UF") = True Then Rows("457").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("457").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    'On Error Resume Next
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
End Sub
Private Sub CB_10_Click()
    If Range("OandM_R") = True Then Rows("458").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("458").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    'On Error Resume Next
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
End Sub
Private Sub CB_11_Click()
    If Range("OandM__OP") = True Then Rows("459").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("459").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    'On Error Resume Next
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
End Sub
Private Sub CB_12_Click()
    If Range("OandM_OC") = True Then Rows("460").EntireRow.Hidden = False: Range("BI451").Select Else: Rows("460").EntireRow.Hidden = True: Range("BI451").Select
    If Range("allOff") = 0 Then Rows("452:461").EntireRow.Hidden = True
    'On Error Resume Next
    If Range("allOff") = 1 Then Range("452:454,461:461").EntireRow.Hidden = False
End Sub
Private Sub ResizeMenu()

    '    Application.ScreenUpdating = False
    '    Application.Calculation = xlManual
    '    Application.DisplayAlerts = False
    '    Application.EnableEvents = False


    UnhideAll
    
    
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("Menu" & i)).Select
        With Selection
            .Height = Range("Menu_Opt1").Height * 2
            .Width = Range("B10:f10").Width
            .Top = Range("Menu_Opt1").Top
            .Left = Range("Menu_Opt" & i).Left
        End With
        ActiveSheet.Shapes.Range(Array("Main_Menu")).Select
        With Selection
            .Width = Range("Main_MenuRange").Height * 2
            .Height = Range("Main_MenuRange").Height * 2
            .Top = Range("Main_MenuRange").Top + 1
            .Left = ((Range("C9:D9").Width / 2) + Range("a9").Width * 1.5) - 3
        End With
    Next i
    
    For i = 1 To 5
        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Select
        With Selection
            .Height = ActiveSheet.Shapes.Range(Array("Menu1")).Height
            .Width = ActiveSheet.Shapes.Range(Array("Menu1")).Width
            .Top = ActiveSheet.Shapes.Range(Array("Menu1")).Top
            If i = 1 Then .Left = Range("BG10").Left + ActiveSheet.Shapes.Range(Array("Menu1")).Left Else: .Left = ActiveSheet.Shapes.Range(Array("Menu" & i - 1 & "B")).Left + ActiveSheet.Shapes.Range(Array("Menu" & i - 1 & "B")).Width
        End With
        ActiveSheet.Shapes.Range(Array("Main_MenuB")).Select
        With Selection
            .Height = ActiveSheet.Shapes.Range(Array("Main_Menu")).Height
            .Width = ActiveSheet.Shapes.Range(Array("Main_Menu")).Width
            .Top = ActiveSheet.Shapes.Range(Array("Main_Menu")).Top
            .Left = Range("BG10").Left + ActiveSheet.Shapes.Range(Array("Main_Menu")).Left
        End With
    Next i
    
    Menu1_Click
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

'Private Sub ()
'    Dim s      As Shape
'    ActiveSheet.Unprotect Password:="1230"
'    For Each s In ActiveSheet.Shapes
'        s.Locked = True
'    Next
'    ActiveSheet.Protect Password:="1230"
'End Sub
'
''un-protect all shapes in active sht
'Private Sub ()
'    Dim s      As Shape
'    ActiveSheet.Unprotect Password:="1230"
'    For Each s In ActiveSheet.Shapes
''        s.Locked = False
''    Next
'End Sub

'************************************
'          MENU LINKS
'************************************

Public Sub Menu1_Click()
    StartingSurvey
End Sub

Private Sub Menu2_Click()
    Cont_ProjGenInfo_Click
End Sub
Private Sub Menu3_Click()
    Cont_CalcFee_Click
End Sub
Private Sub Menu4_Click()
    Cont_OandM_Click
End Sub
Private Sub Menu5_Click()
    Cont_RevGuar_Click
End Sub
Private Sub Menu1B_Click()
    StartingSurvey
End Sub

Private Sub Menu2B_Click()
    Cont_ProjGenInfo_Click
End Sub
Private Sub Menu3B_Click()
    Cont_CalcFee_Click
End Sub
Private Sub Menu4B_Click()
    Cont_OandM_Click
End Sub
Private Sub Menu5B_Click()
    Cont_RevGuar_Click
End Sub
Private Sub GoTo1_Click()
    StartingSurvey
End Sub

Private Sub GoTo2_Click()
    Cont_ProjJust_Click
End Sub
Private Sub GoTo3_Click()
    Cont_ProjFund_Click
End Sub
Private Sub GoTo4_Click()
    Cont_ProjCurr_Click
End Sub
Private Sub GoTo17_Click()
    Cont_ProjGenInfo_Click
End Sub
Private Sub GoTo18_Click()
    Cont_ProjFinan_Click
End Sub
Private Sub GoTo19_Click()
    Cont_AssetCharact_Click
End Sub
Private Sub GoTo25_Click()
    Cont_CalcFee_Click
End Sub
Private Sub GoTo26_Click()
    Cont_AnnualRev_Click
End Sub
Private Sub GoTo27_Click()
    Cont_OandMAuto_Click
End Sub
Private Sub GoTo28_Click()
    Cont_RevGuar_Click
End Sub
Private Sub GoTo29_Click()
    Cont_DebtGuar_Click
End Sub
Public Sub ScrollToTop()
    ActiveWindow.ScrollRow = 1
End Sub


'***************************
'    REVENUES Buttons
'***************************

Private Sub User_fund_Click()


    Range("UF") = True
    Range("GF") = False
    
    Dim Uf, Gf, co As String
    
    Uf = "Users of the services--user funded"
    Gf = "Tax payers--public funded"
    co = "Combined"
    
    '***************************************************************************************************************************************************************************************************************************************************************
    '***************************************************************************************************************************************************************************************************************************************************************
    
    Dim UserValue, GovValue As Boolean
    Dim arr    As Variant
    
    'INFLATION
    
    If Range("Serv_Funded") = co Or Range("Serv_Funded") = Uf Then
        If Range("Coll_Price" & Range("CurrentSer") & "UserDomesticinflationindexed") = "No" Then UserValue = False Else: UserValue = True
        Range("Serv_Inflation") = UserValue
        Range("INFLATION") = UserValue
    End If
    
    If Range("Serv_Funded") = Gf Then
        If Range("Coll_Price" & Range("CurrentSer") & "GovDomesticinflationindexed") = "No" Then GovValue = False Else: GovValue = True
        Range("Serv_Inflation") = GovValue
        Range("INFLATION") = GovValue
        
    End If
    
    'NER
    
    If Range("Serv_Funded") = co Or Range("Serv_Funded") = Uf Then
        If Range("Coll_Price" & Range("CurrentSer") & "UserNERindexed") = "No" Then UserValue = False Else: UserValue = True
        Range("Serv_Ner") = UserValue
        Range("NER") = UserValue
        
    End If
    
    If Range("Serv_Funded") = Gf Then
        If Range("Coll_Price" & Range("CurrentSer") & "GovNERindexed") = "No" Then GovValue = False Else: GovValue = True
        Range("Serv_Ner") = GovValue
        Range("NER") = GovValue
        
    End If
    
    'GDP
    
    If Range("Serv_Funded") = co Or Range("Serv_Funded") = Uf Then
        If Range("Coll_Demand" & Range("CurrentSer") & "UserLinktoGDP") = "No" Then UserValue = False Else: UserValue = True
        Range("Serv_GDP") = UserValue
        Range("GDP") = UserValue
        
    End If
    
    If Range("Serv_Funded") = Gf Then
        If Range("Coll_Demand" & Range("CurrentSer") & "GovLinktoGDP") = "No" Then GovValue = False Else: GovValue = True
        Range("Serv_GDP") = GovValue
        Range("GDP") = GovValue
        
    End If
    
    'OTHER FACTOR PRICE
    
    '=========== Price =========================
    
    'Users funded
    If Range("Serv_Funded") = Uf Then
        If Range("Coll_Price" & Range("CurrentSer") & "UserOtherAdjustmentfactor") = "" Then UserValue = False Else: UserValue = True
        
        Range("Serv_PriceOther") = UserValue
        Range("Price_OFactor") = UserValue
        
        If UserValue = False Then
            Range("$345:$349").EntireRow.Hidden = True
        ElseIf UserValue = True Then
            Range("$345:$349").EntireRow.Hidden = False
            Range("$348:$348").EntireRow.Hidden = True
            
            arr = Split(Range("Coll_Price" & Range("CurrentSer") & "UserOtherAdjustmentfactor"), ";")
            
            For i = LBound(arr) To UBound(arr)
                Cells(347, 66 + i) = arr(i)
            Next i
            
        End If
        
    End If
    
    'Gov. funded
    If Range("Serv_Funded") = Gf Then
        If Range("Coll_Price" & Range("CurrentSer") & "GovOtherAdjustmentfactor") = "" Then GovValue = False Else: GovValue = True
        
        Range("Serv_PriceOtherGov") = GovValue
        Range("Price_OFactor") = GovValue
        
        If GovValue = False Then
            Range("$345:$349").EntireRow.Hidden = True
        ElseIf GovValue = True Then
            Range("$345:$349").EntireRow.Hidden = False
            Range("$347:$347").EntireRow.Hidden = True
            
            arr = Split(Range("Coll_Price" & Range("CurrentSer") & "GovOtherAdjustmentfactor"), ";")
            
            For i = LBound(arr) To UBound(arr)
                Cells(348, 66 + i) = arr(i)
            Next i
            
        End If
    End If
    
    'Combined funded
    If Range("Serv_Funded") = co Then
        
        If Range("UF") = True Then
            If Range("Coll_Price" & Range("CurrentSer") & "UserOtherAdjustmentfactor") = "" Then UserValue = False Else: UserValue = True
            
            Range("Serv_PriceOther") = UserValue
            Range("Price_OFactor") = UserValue
            
            If UserValue = False Then
                Range("$345:$349").EntireRow.Hidden = True
            ElseIf UserValue = True Then
                Range("$345:$349").EntireRow.Hidden = False
                Range("$348:$348").EntireRow.Hidden = True
                
                arr = Split(Range("Coll_Price" & Range("CurrentSer") & "UserOtherAdjustmentfactor"), ";")
                
                For i = LBound(arr) To UBound(arr)
                    Cells(347, 66 + i) = arr(i)
                Next i
            End If
        End If
        If Range("GF") = True Then
            If Range("Coll_Price" & Range("CurrentSer") & "GovOtherAdjustmentfactor") = "" Then GovValue = False Else: GovValue = True
            
            Range("Serv_PriceOtherGov") = GovValue
            Range("Price_OFactor") = GovValue
            
            If GovValue = False Then
                Range("$345:$349").EntireRow.Hidden = True
            ElseIf GovValue = True Then
                Range("$345:$349").EntireRow.Hidden = False
                Range("$347:$347").EntireRow.Hidden = True
                
                arr = Split(Range("Coll_Price" & Range("CurrentSer") & "GovOtherAdjustmentfactor"), ";")
                
                For i = LBound(arr) To UBound(arr)
                    Cells(348, 66 + i) = arr(i)
                Next i
                
            End If
        End If
    End If
    
    'OTHER FACTOR DEMAND
    
    '=========== Demand =========================
    
    'Users funded
    If Range("Serv_Funded") = Uf Then
        If Range("Coll_Demand" & Range("CurrentSer") & "UserOtherAdjustmentfactor") = "" Then UserValue = False Else: UserValue = True
        
        Range("Serv_DemaOther") = UserValue
        Range("Demand_OFactor") = UserValue
        
        If UserValue = False Then
            Range("$351:$355").EntireRow.Hidden = True
        ElseIf UserValue = True Then
            Range("$351:$355").EntireRow.Hidden = False
            Range("$354:$354").EntireRow.Hidden = True
            
            arr = Split(Range("Coll_Demand" & Range("CurrentSer") & "UserOtherAdjustmentfactor"), ";")
            
            For i = LBound(arr) To UBound(arr)
                Cells(353, 66 + i) = arr(i)
            Next i
            
        End If
        
    End If
    
    'Gov. funded
    If Range("Serv_Funded") = Gf Then
        If Range("Coll_Demand" & Range("CurrentSer") & "GovOtherAdjustmentfactor") = "" Then GovValue = False Else: GovValue = True
        
        Range("Serv_DemaOtherGov") = GovValue
        Range("Demand_OFactor") = GovValue
        
        If GovValue = False Then
            Range("$351:$355").EntireRow.Hidden = True
        ElseIf GovValue = True Then
            Range("$351:$355").EntireRow.Hidden = False
            Range("$353:$353").EntireRow.Hidden = True
            
            arr = Split(Range("Coll_Demand" & Range("CurrentSer") & "GovOtherAdjustmentfactor"), ";")
            
            For i = LBound(arr) To UBound(arr)
                Cells(354, 66 + i) = arr(i)
            Next i
            
        End If
    End If
    
    'Combined funded
    If Range("Serv_Funded") = co Then
        
        If Range("UF") = True Then
            If Range("Coll_Demand" & Range("CurrentSer") & "UserOtherAdjustmentfactor") = "" Then UserValue = False Else: UserValue = True
            
            Range("Serv_DemaOther") = UserValue
            Range("Demand_OFactor") = UserValue
            
            If UserValue = False Then
                Range("$351:$355").EntireRow.Hidden = True
            ElseIf UserValue = True Then
                Range("$351:$355").EntireRow.Hidden = False
                Range("$354:$354").EntireRow.Hidden = True
                
                arr = Split(Range("Coll_Demand" & Range("CurrentSer") & "UserOtherAdjustmentfactor"), ";")
                
                For i = LBound(arr) To UBound(arr)
                    Cells(353, 66 + i) = arr(i)
                Next i
            End If
        End If
        If Range("GF") = True Then
            If Range("Coll_Demand" & Range("CurrentSer") & "GovOtherAdjustmentfactor") = "" Then GovValue = False Else: GovValue = True
            
            Range("Serv_DemaOtherGov") = GovValue
            Range("Demand_OFactor") = GovValue
            
            If GovValue = False Then
                Range("$351:$355").EntireRow.Hidden = True
            ElseIf GovValue = True Then
                Range("$351:$355").EntireRow.Hidden = False
                Range("$353:$353").EntireRow.Hidden = True
                
                arr = Split(Range("Coll_Demand" & Range("CurrentSer") & "GovOtherAdjustmentfactor"), ";")
                
                For i = LBound(arr) To UBound(arr)
                    Cells(354, 66 + i) = arr(i)
                Next i
                
            End If
        End If
    End If
    
    '***************************************************************************************************************************************************************************************************************************************************************
    '***************************************************************************************************************************************************************************************************************************************************************
    
    Dim onn, off As Variant
    
    onn = "User_fund"
    off = "Gov_fund"
    With ActiveSheet.Shapes.Range(Array(onn))
        .Fill.ForeColor.RGB = RGB(59, 56, 56)
        .Line.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .Shadow.Visible = msoFalse
        .TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .TextFrame2.TextRange.Font.Bold = msoTrue
    End With
    
    With ActiveSheet.Shapes.Range(Array(off))
        .Fill.ForeColor.RGB = RGB(217, 217, 217)
        .Shadow.Visible = msoFalse
        .TextFrame2.TextRange.Font.Bold = msoFalse
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    Range("Ser_Name").Select
    
End Sub
Private Sub Gov_fund_Click()
    Dim Uf, Gf, co As String


    Range("UF") = False
    Range("GF") = True
    
    Uf = "Users of the services--user funded"
    Gf = "Tax payers--public funded"
    co = "Combined"
    
    '***************************************************************************************************************************************************************************************************************************************************************
    '***************************************************************************************************************************************************************************************************************************************************************
    
    Dim UserValue, GovValue As Boolean
    Dim arr    As Variant
    
    'INFLATION
    If Range("Coll_Price" & Range("CurrentSer") & "GovDomesticinflationindexed") = "No" Then GovValue = False Else: GovValue = True
    Range("Serv_InflationGov") = GovValue
    Range("INFLATION") = GovValue
    
    'NER
    If Range("Coll_Price" & Range("CurrentSer") & "GovNERindexed") = "No" Then GovValue = False Else: GovValue = True
    Range("Serv_NerGov") = GovValue
    Range("NER") = GovValue
    
    'GDP
    If Range("Coll_Demand" & Range("CurrentSer") & "GovLinktoGDP") = "No" Then GovValue = False Else: GovValue = True
    Range("Serv_GDPGov") = GovValue
    Range("GDP") = GovValue
    
    'OTHER FACTOR PRICE User/Combined Funded
    
    '=========== Price =========================
    
    If Range("Coll_Price" & Range("CurrentSer") & "GovOtherAdjustmentfactor") = "" Then GovValue = False Else: GovValue = True
    
    Range("Serv_PriceOtherGov") = GovValue
    Range("Price_OFactor") = GovValue
    
    If GovValue = False Then
        Range("$345:$349").EntireRow.Hidden = True
    ElseIf GovValue = True Then
        Range("$345:$349").EntireRow.Hidden = False
        Range("$347:$347").EntireRow.Hidden = True
        
        arr = Split(Range("Coll_Price" & Range("CurrentSer") & "GovOtherAdjustmentfactor"), ";")
        
        For i = LBound(arr) To UBound(arr)
            Cells(348, 66 + i) = arr(i)
        Next i
    End If
    
    'OTHER FACTOR DEMAND
    
    '=========== Demand =========================
    
    If Range("Coll_Demand" & Range("CurrentSer") & "GovOtherAdjustmentfactor") = "" Then GovValue = False Else: GovValue = True
    
    Range("Serv_DemaOtherGov") = GovValue
    Range("Demand_OFactor") = GovValue
    
    If GovValue = False Then
        Range("$351:$355").EntireRow.Hidden = True
    ElseIf GovValue = True Then
        Range("$351:$355").EntireRow.Hidden = False
        Range("$353:$353").EntireRow.Hidden = True
        
        arr = Split(Range("Coll_Demand" & Range("CurrentSer") & "GovOtherAdjustmentfactor"), ";")
        
        For i = LBound(arr) To UBound(arr)
            Cells(354, 66 + i) = arr(i)
        Next i
        
    End If
    
    'Combined funded
    If Range("Serv_Funded") = co Then
        If Range("Coll_Demand" & Range("CurrentSer") & "GovOtherAdjustmentfactor") = "" Then GovValue = False Else: GovValue = True
        
        Range("Serv_DemaOtherGov") = GovValue
        Range("Demand_OFactor") = GovValue
        
        If GovValue = False Then
            Range("$351:$355").EntireRow.Hidden = True
        ElseIf GovValue = True Then
            Range("$351:$355").EntireRow.Hidden = False
            Range("$353:$353").EntireRow.Hidden = True
            
            arr = Split(Range("Coll_Demand" & Range("CurrentSer") & "GovOtherAdjustmentfactor"), ";")
            
            For i = LBound(arr) To UBound(arr)
                Cells(354, 66 + i) = arr(i)
            Next i
        End If
    End If
    
    '***************************************************************************************************************************************************************************************************************************************************************
    '***************************************************************************************************************************************************************************************************************************************************************
    
    Dim onn, off As Variant
    
    off = "User_fund"
    onn = "Gov_fund"
    With ActiveSheet.Shapes.Range(Array(onn))
        .Fill.ForeColor.RGB = RGB(59, 56, 56)
        .Line.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .Shadow.Visible = msoFalse
        .TextFrame2.TextRange.Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .TextFrame2.TextRange.Font.Bold = msoTrue
    End With
    
    With ActiveSheet.Shapes.Range(Array(off))
        .Fill.ForeColor.RGB = RGB(217, 217, 217)
        .Shadow.Visible = msoFalse
        .TextFrame2.TextRange.Font.Bold = msoFalse
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    End With
    
    Range("Ser_Name").Select
    
End Sub
Private Sub transferRevenues()

    Dim myArray, arr As Variant
    Dim myString, str As String
    Dim DataRange, rng As Range
    Dim cell   As Range
    Dim x      As Long
    Dim ancorCell As Range: Set ancorCell = Range("CurrentSer")

    Set DataRange = ActiveSheet.Range("ResultRevenues")
    
    For Each cell In DataRange.Cells
        myString = myString & ";|;" & cell.Value
    Next cell
    
    myArray = Split(myString, ";|;")
    
    
    For x = LBound(myArray) + 1 To UBound(myArray)
        Cells(302 + ancorCell, 152 + x) = myArray(x)
    Next x
    
    Set rng = ActiveSheet.Range("ResultRevenue1")
    
    For Each cell In rng.Cells
        str = str & ";|;" & cell.Value
    Next cell
    
    arr = Split(str, ";|;")
    
    
    For x = LBound(arr) + 1 To UBound(arr)
        Cells(302 + ancorCell, 149 + x) = arr(x)
    Next x
End Sub
Private Sub transfer_INF_NER_GDP()

    Dim myArray, arr As Variant
    Dim myString, strg As String
    Dim DataRange, rng1, rng2 As Range
    Dim cell   As Range
    Dim x      As Long
    Dim c, r   As Long
    Dim ancorCell As Range: Set ancorCell = Range("CurrentSer")

    c = 150
    
    '    If Range("Serv_Funded") <> "Combined" Then
    '        If Range("Serv_Funded") = "Tax payers--public funded" Then r = 314 Else: r = 324
    '
    '        Set DataRange = ActiveSheet.Range("$EU$336:$EW$336")
    '
    '        For Each cell In DataRange.Cells
    '            myString = myString & ";|;" & cell.Value
    '        Next cell
    '
    '        myArray = Split(myString, ";|;")
    '
    '
    '        For x = LBound(myArray) + 1 To UBound(myArray)
    '            Cells(r + ancorCell, c + x) = myArray(x)
    '        Next x
    '    Else
    Set rng1 = ActiveSheet.Range("$EU$336:$EW$336")
    Set rng2 = ActiveSheet.Range("$EU$340:$EW$340")
    
    r = 324
    c = 150
    
    For Each cell In rng1.Cells
        myString = myString & ";|;" & cell.Value
    Next cell
    
    myArray = Split(myString, ";|;")
    
    
    For x = LBound(myArray) + 1 To UBound(myArray)
        Cells(r + ancorCell, c + x) = myArray(x)
    Next x
    
    r = 314
    c = 150
    
    For Each cell In rng2.Cells
        strg = strg & ";|;" & cell.Value
    Next cell
    
    arr = Split(strg, ";|;")
    
    
    For x = LBound(arr) + 1 To UBound(arr)
        Cells(r + ancorCell, c + x) = arr(x)
    Next x
    '    End If
    
    
End Sub

Private Sub transfer_OtherFactor()

    Dim myArray, arr As Variant
    Dim myString, strg As String
    Dim DataRange, rng1, rng2 As Range
    Dim cell   As Range
    Dim x      As Long
    Dim c, r   As Long
    Dim ancorCell As Range: Set ancorCell = Range("CurrentSer")

    c = 154
    
    If Range("Serv_Funded") <> "Combined" Then
        If Range("Serv_Funded") = "Tax payers--public funded" Then r = 314 Else: r = 324
        
        Set DataRange = ActiveSheet.Range("Result_OtherFactorUF")
        
        For Each cell In DataRange.Cells
            myString = myString & ";|;" & cell.Value
        Next cell
        
        myArray = Split(myString, ";|;")
        
        
        For x = LBound(myArray) + 1 To UBound(myArray)
            Cells(r + ancorCell, c + x) = myArray(x)
        Next x
    Else
        Set rng1 = ActiveSheet.Range("Result_OtherFactorUF")
        Set rng2 = ActiveSheet.Range("Result_OtherFactorGF")
        
        r = 324
        c = 154
        
        For Each cell In rng1.Cells
            myString = myString & ";|;" & cell.Value
        Next cell
        
        myArray = Split(myString, ";|;")
        
        
        For x = LBound(myArray) + 1 To UBound(myArray)
            Cells(r + ancorCell, c + x) = myArray(x)
        Next x
        
        r = 314
        c = 154
        
        For Each cell In rng2.Cells
            strg = strg & ";|;" & cell.Value
        Next cell
        
        arr = Split(strg, ";|;")
        
        
        For x = LBound(arr) + 1 To UBound(arr)
            Cells(r + ancorCell, c + x) = arr(x)
        Next x
    End If
    
    
End Sub

Public Sub Menu_Macro1_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    'LIGHT
    For i = 1 To 2
        ActiveSheet.Shapes.Range(Array("Menu_Macro" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu_Macro" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu_Macro1")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu_Macro1")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Range("Macro_Header").EntireRow.Hidden = True
    Range("Country_Param").EntireRow.Hidden = False
    Range("Fiscal_Rule").EntireRow.Hidden = True
    
    Range("PortfolioYrBegin").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Private Sub Menu_Macro2_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    frmFiscal_Ceiling.Show
    
    Sheets("Macro_Data").Activate
    
    'LIGHT
    For i = 1 To 2
        ActiveSheet.Shapes.Range(Array("Menu_Macro" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menu_Macro" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'DARK
    ActiveSheet.Shapes.Range(Array("Menu_Macro2")).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menu_Macro2")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    Range("Macro_Header").EntireRow.Hidden = True
    Range("Country_Param").EntireRow.Hidden = True
    Range("Fiscal_Rule").EntireRow.Hidden = False
    
    Cells(43, 8).Select
    
    '    ActiveSheet.Shapes.Range(Array("Fiscal_CeilingChart")).Visible = True
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub valuesOLD()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wb     As Workbook: Set wb = ThisWorkbook
    Dim ws     As Worksheet: Set ws = wb.Worksheets("Macro_Data")
    Dim valstr As String
    
    For i = 0 To 49
        valstr = valstr & ";" & Application.Sum(ws.Range(Cells(25, 8 + i), Cells(33, 8 + i)))
    Next i
    
    ws.Cells(2, 1) = valstr
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub Main_MenuB_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    If Range("Difference_Macro").Value = True Then
        Dim answer As Integer
        
        answer = MsgBox("We have detected that there were changes in macroeconomic data, do you want to reflect these changes for the next reports?", vbInformation + vbYesNo, "Update Macroparameters")
        
        Select Case answer
            Case vbYes
                valuesOLD
            Case vbNo
                
        End Select
    End If
    Worksheets("Main_Menu").Activate
    Cells(13, 3).Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub Main_Menu_Macroeconomic_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Worksheets("Macro_Data").Activate
    Menu_Macro1_Click
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Private Sub Main_Menu_Comb_funded_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    UnhideAll
    
    Range("Main_Menu").EntireRow.Hidden = True
    Range("Proj_Data_Menu").EntireRow.Hidden = False
    Range("Proj_Data_SubMenu").EntireRow.Hidden = True
    
    Range("C44").Select
    Worksheets("Main_Menu").ComboBox1.Visible = False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
Public Sub VfM_Main_Menu_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Sheets("Project_Data").Activate
    UserForm3.Show
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Public Sub Project_Data_Main_Menu_Click()
Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ThisWorkbook.Worksheets("Template").Activate
    Call EIR_SeekTemplate
    Call Proj_Summary
    Call ExpensesCashFlow
    Call Gov_St
    Worksheets("Main_Menu").Activate
    Call Back_ProjData_Click
    Cells(13, 3).Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub Back_ProjData_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    UnhideAll
    
    Range("Main_Menu").EntireRow.Hidden = False
    Range("Proj_Data_Menu").EntireRow.Hidden = True
    Range("Proj_Data_SubMenu").EntireRow.Hidden = True
    Worksheets("Main_Menu").ComboBox1.Visible = False
    Cells(13, 3).Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub Edit_Project_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ActiveSheet.Shapes.Range(Array("Cont_ProjSelection")).Visible = True
    
    ProjectOptionSelected = "Edit"
    
    UnhideAll
    
    Range("Main_Menu").EntireRow.Hidden = True
    Range("Proj_Data_Menu").EntireRow.Hidden = True
    Range("Proj_Data_SubMenu").EntireRow.Hidden = False
    Worksheets("Main_Menu").ComboBox1.Visible = True
    Cells(56, 2).Value = "Edit or view a project from portfolio"
    Range("B56").Font.Color = -1003520
    ListCombobox1
    Worksheets("Main_Menu").ComboBox1.Value = "<< Select project from list to edit >>"
    Cells(59, 2).Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub Back_ProjSelection_Click()
    Main_Menu_Comb_funded_Click
End Sub

Private Sub Delete_Project_Click()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    
    UnhideAll
    
    ActiveSheet.Shapes.Range(Array("Cont_ProjSelection")).Visible = False
    
    Range("Main_Menu").EntireRow.Hidden = True
    Range("Proj_Data_Menu").EntireRow.Hidden = True
    Range("Proj_Data_SubMenu").EntireRow.Hidden = False
    Worksheets("Main_Menu").ComboBox1.Visible = True
    Cells(56, 2).Value = "Delete a project from portfolio"
    Range("B56").Font.Color = -16777024
    ListCombobox1
    Worksheets("Main_Menu").ComboBox1.Value = "<< Select project from list to delete>>"
    Cells(59, 2).Select
    
    ProjectOptionSelected = "Delete"
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub


Private Sub Create_Project_Click()
    ActiveSheet.Shapes.Range(Array("Cont_ProjSelection")).Visible = True
    ProjectOptionSelected = "New"
    NewProjectEmpty
End Sub


'    ***************************************************
'    * COMBOBOX1 MAIN MENU ACTION *
'    ***************************************************

Sub DeleteProject()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim lCol   As Long: lCol = Worksheets("Store_Data").Cells(9, Columns.Count).End(xlToLeft).Column
    Dim lColMatrix   As Long: lColMatrix = Worksheets("Risk_Matrix").Cells(17, Columns.Count).End(xlToLeft).Column
    
    If lCol - 4 = 0 Then
        Dim answer As Integer
        answer = MsgBox("This action cannot be executed because the portfolio is empty of projects.", vbInformation + vbOKOnly, "Empty portfolio")
        Exit Sub
    End If
    
    Dim wss    As Sheets: Set wss = ThisWorkbook.Worksheets
    Dim ws     As Worksheet: Set ws = wss("Store_Data")
    Dim wsMatrix     As Worksheet: Set wsMatrix = wss("Risk_Matrix")
    Dim rng    As Range: Set rng = ws.Range(ws.Cells(9, 5), ws.Cells(9, lCol))
    Dim rngMatrix    As Range: Set rngMatrix = wsMatrix.Range(wsMatrix.Cells(17, 25), wsMatrix.Cells(17, lColMatrix))
    Dim cell   As Range
    Dim ProjSelect As Long
    Dim arr()  As Variant
    Dim nameproj As Variant
    
    ProjSelect = ThisWorkbook.Sheets("Main_Menu").ComboBox1.ListIndex
    nameproj = Worksheets("Main_Menu").ComboBox1.Value
    
    answer = MsgBox("You are about to delete " & nameproj & " project. Would you like to continue?", vbCritical + vbYesNo, "Project delete")
    
    Select Case answer
        Case vbYes
            
            ws.Columns(ProjSelect + 5).Delete
            wsMatrix.Columns(ProjSelect + 25).Delete
            
            On Error Resume Next
            
            i = 1
            For Each cell In rng
                cell.Value = "P" & i
                i = i + 1
                
            Next cell
            
            i = 1
            For Each cell In rngMatrix
                cell.Value = "P" & i
                i = i + 1
                
            Next cell
            
            ListCombobox1
            
        Case vbNo
            Exit Sub
            
    End Select
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub ListCombobox1()
    Dim lCol   As Long: lCol = Worksheets("Store_Data").Cells(9, Columns.Count).End(xlToLeft).Column
    Dim rng    As Range
    Dim wss    As Sheets: Set wss = ThisWorkbook.Worksheets
    Dim ws     As Worksheet: Set ws = wss("Store_Data")
    Dim cell   As Range
    Dim x      As Long

    Set rng = ws.Range(ws.Cells(17, 5), ws.Cells(17, lCol))
    
    Dim arr()  As Variant
    For Each cell In rng
        ReDim Preserve arr(x)
        If cell <> "" Then arr(x) = cell.Value
        x = x + 1
    Next cell
    
    Worksheets("Main_Menu").ComboBox1.Clear
    Worksheets("Main_Menu").ComboBox1.List = Application.Transpose(arr())
    
End Sub

Private Sub AlignServiceAssetTab()

    Application.ScreenUpdating = False
    
    Dim w      As Long, h As Long
    Dim TopPosition As Long, LeftPosition As Long
    Dim chtObj As ChartObject
    Dim i      As Long, NumCols As Long
    Dim cellRef As Range
    
    
    'ASSET ALIGNMENT
    ActiveSheet.Shapes.Range(Array("SPAsset1")).Visible = True
    ActiveSheet.Shapes.Range(Array("SPAsset1")).Select
    Set cellRef = Sheets("Project_Data").Range("$D$208")
    
    'Get size of active chart
    w = 48.24
    h = 12.24
    NumCols = 10
    
    'Change starting positions, if necessary
    TopPosition = (cellRef.Top + cellRef.Height) - h
    LeftPosition = cellRef.Left
    
    For i = 1 To 10
        With ActiveSheet.Shapes.Range(Array("SPAsset" & i))
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 1) Mod NumCols) * w
            .Top = TopPosition + Int((i - 1) / NumCols) * h
        End With
    Next i
    
    'SERVICE ALIGNMENT
    
    ActiveSheet.Shapes.Range(Array("SPAsset1")).Select
    Set cellRef = Sheets("Project_Data").Range("$BI$307")
    
    'Get size of active chart
    w = 48.24
    h = 12.24
    NumCols = 10
    
    'Change starting positions, if necessary
    TopPosition = (cellRef.Top + cellRef.Height) - h
    LeftPosition = cellRef.Left
    
    For i = 1 To 10
        With ActiveSheet.Shapes.Range(Array("SPSer" & i))
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 1) Mod NumCols) * w
            .Top = TopPosition + Int((i - 1) / NumCols) * h
        End With
    Next i
    
    Application.ScreenUpdating = True
End Sub

Private Sub loadDataofProjEdit()

    frmWait.Show
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    NewProjectEmpty
    
    Dim wss    As Sheets: Set wss = ThisWorkbook.Worksheets
    Dim ws     As Worksheet: Set ws = wss("Store_Data")
    Dim FormulaInCell, arr As Variant
    Dim myArray As Variant
    Dim myString As String
    
    Dim x, Y, j, n As Long
    x = 1
    Y = 1
    j = (ThisWorkbook.Sheets("Project_Data").Range("current_project") + 5)
    
    Dim strArr1, strArr2 As String
    Dim rng1, rng2, cell As Range
    Dim arrCell1, arrCell2 As Variant
    
    Set rng1 = ThisWorkbook.Worksheets("TStore_Data").Range("RowsStore1")
    Set rng2 = ThisWorkbook.Worksheets("TStore_Data").Range("RowsStore2")
    
    For Each cell In rng1.Cells
        strArr1 = cell.Value & ";" & strArr1
    Next cell
    
    arrCell1 = Split(strArr1, ";")
    
    For Each cell In rng2.Cells
        strArr2 = cell.Value & ";" & strArr2
    Next cell
    
    arrCell2 = Split(strArr2, ";")
    
    For x = LBound(arrCell1) To UBound(arrCell1) - 1
        With ws.Cells(arrCell1(x), j)
            wss("Project_Data").Range(arrCell2(x)) = .Value
        End With
    Next x
    
    For i = 161 To 166
        arr = Split(ws.Cells(i, j).Value, ";")
        For x = LBound(arr) To UBound(arr)
            wss("Project_Data").Cells(455 + (i - 161), 66 + x).Value = arr(x)
        Next x
    Next i
    
    If i = 159 Then
        arr = Split(ws.Cells(i, j).Value, ";")
        For x = LBound(arr) To UBound(arr)
            wss("Project_Data").Cells(282, 66 + x).Value = arr(x)
        Next x
    End If
    
    'SERVICES LOAD DATA
    
    myString = "Asset_name;YrsOfConstruction;YrAssetBegin;Useful_life;Constr_cost;Land_private"
    myArray = Split(myString, ";")
    
    For i = 0 To 5
        wss("Project_Data").Range(myArray(i)).Value = wss("Project_Data").Cells(209 + i, 60)
    Next i
    
    'FINANCE LOAD DATA
    
    myString = "170;177;181;185;189"
    myArray = Split(myString, ";")
    
    For i = 0 To 4
        If i = 0 Then
            Cells(myArray(i), j) = wss("Project_Data").Cells(1 + i, 27).Value
        Else
            Cells(myArray(i), 4) = wss("Project_Data").Cells(1 + i, 27).Value
        End If
        
    Next i
    
    'REVENUES GUARANTEES
    
    myArray = Split(ws.Cells(275, j).Value, ";")
    
    For i = LBound(myArray) To UBound(myArray)
        wss("Project_Data").Cells(387, 66 + i) = myArray(i)
    Next i
    
    'OTHER PAYMENT OF THE GOVERNMENT
    
    If Range("Coll_Typeofproject") = "YES" Then ob_YesOtherPymnt_Click
    If Range("Coll_Typeofproject") = "NO" Then ob_NoOtherPymnt_Click
    
    
    myArray = Split(ws.Cells(278, j).Value, ";")
    
    For i = LBound(myArray) To UBound(myArray)
        wss("Project_Data").Cells(421, 66 + i) = myArray(i)
    Next i
    
    
    
    'HOW THE PROJECT IS FUNDED
    If Range("Coll_Howtheprojectisfunded") = "Tax payers--public funded" Then Gov_funded_Click
    If Range("Coll_Howtheprojectisfunded") = "Users of the services--user funded" Then User_funded_Click
    If Range("Coll_Howtheprojectisfunded") = "Combined" Then Comb_funded_Click
    
    'TYPE OF PROJECT
    If Range("Coll_Typeofproject") = "DB" Then OB_DB_Click
    If Range("Coll_Typeofproject") = "DBFO" Then OB_DBFO_Click
    If Range("Coll_Typeofproject") = "BOT" Then OB_BOT_Click
    If Range("Coll_Typeofproject") = "BBO" Then OB_BBO_Click
    If Range("Coll_Typeofproject") = "BOO" Then OB_BOO_Click
    
    'CURRENCY
    If Range("Coll_Currency") = "DOM" Then Loc_Currency_Click
    If Range("Coll_Currency") = "FX" Then FX_Currency_Click
    
    'REVENUES CALCULATION
    If Range("Coll_CalculateRevenue") = "Manual" Then EnterManual_Click
    If Range("Coll_CalculateRevenue") = "Auto" Then CalculPFRAM_Click
    
    AlignServiceAssetTab
    LoadImages
    
    For i = 1 To 10
        wss("Project_Data").Shapes.Range(Array("SPAsset" & i)).Visible = False
        wss("Project_Data").Shapes.Range(Array("SPSer" & i)).Visible = False
    Next i
    
    'MANUAL ANNUAL REVENUES ====================================================
    
    'IS IT MORE THAN 0?
    Dim L      As Long
    arr = Split(Worksheets("Store_Data").Cells(159, 2).Value, ";")
    For x = LBound(arr) To UBound(arr)
        arr(x) = arr(x) + L
        L = arr(x)
    Next x
    
    If i <> 0 Then
        arr = Split(ws.Cells(159, 2).Value, ";")
        For x = LBound(arr) To UBound(arr)
            wss("Project_Data").Cells(164, 66 + x).Value = arr(x)
        Next x
    End If
    
    If i <> 0 Then
        arr = Split(ws.Cells(160, 2).Value, ";")
        For x = LBound(arr) To UBound(arr)
            wss("Project_Data").Cells(163, 66 + x).Value = arr(x)
        Next x
    End If
    
    Sheets("Project_Data").Range("sp_number") = Sheets("Project_Data").Range("sp_numb_load") - 1
    
    SpinnerRevYear_Change
    
    Dim n1     As Long
    
    n = Range("sp_numb_load").Value - 1
    n1 = Range("Coll_Lengthofcontract").Value - n + 1
    
    
    Sheets("Project_Data").Range(Cells(163, 66 + n), Cells(164, 66 + n1)).Copy
    
    Sheets("Project_Data").Range("BN281").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '====================================================
    
    For i = 22 To 26
        FormulaInCell = Mid(ws.Cells(i, 2).Formula, 2, Len(ws.Cells(i, 2).Formula))
        With ws.Cells(i, j)
            wss("Project_Data").Range(FormulaInCell) = .Value
        End With
    Next i
    
    'O&M OPTIONS TRUE/FALSE
    
    For i = 1 To 6
        wss("Project_Data").Cells(445 + i, 56) = wss("Project_Data").Cells(445 + i, 57)
        
    Next i
    
    Load_Data_RiskM
    
    Worksheets("Project_Data").Select
    
    Unload frmWait
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Private Sub NewProjectEmpty()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' Put in array all the name of projects
    Dim cell
    Dim arr() As Variant
    
    For Each cell In Sheets("Store_Data").Range("$17:$17")
        If cell <> "" Then
            ReDim Preserve arr(i)
            arr(i) = cell.Value
            i = i + 1
        End If
    Next cell
    
    Dim wss    As Sheets: Set wss = ThisWorkbook.Worksheets
    Dim ws     As Worksheet: Set ws = wss("Store_Data")
    Dim FormulaInCell As Variant
    Dim x, Y, j  As Long
    x = 1
    Y = 1
    j = (ThisWorkbook.Sheets("Project_Data").Range("current_project") + 5)
    
    'Enter name when new
    If ProjectOptionSelected = "New" Then
        Dim projName As Variant
        projName = InputBox("Name of the project", "Project name")
        If StrPtr(projName) = 0 Then
            Exit Sub
        ElseIf projName = vbNullString Then
            Exit Sub
        Else
            For i = LBound(arr) To UBound(arr)
                If LCase(projName) = LCase(arr(i)) Then
                    Dim answer As Integer
                    answer = MsgBox("This projet already exist in your portfolio. Please select other name for this new project.", vbCritical + vbOKOnly, "Project name")
                    Exit Sub
                End If
            Next i
            'Nothing
        End If
    End If
    
    wss("Project_Data").Activate
    
    
    For i = 11 To 276
        FormulaInCell = Mid(ws.Cells(i, 2).Formula, 2, Len(ws.Cells(i, 2).Formula))
        
        If ws.Cells(i, j) = "" Then GoTo 1
        If i > 160 And i < 167 Then GoTo 1
        
        
        If i = 275 Or i = 159 Then GoTo 1
        
        If i = 30 Or i = 36 Or i = 42 Or i = 48 Or i = 54 Or i = 60 Or i = 66 Or i = 72 Or i = 78 Or i = 84 Then wss("Project_Data").Range("Coll_Asset" & x & "Yearconstructionbegins").Value = "": x = x + 1: GoTo 1
        
        If i = 91 Or i = 98 Or i = 105 Or i = 112 Or i = 119 Or i = 126 Or i = 133 Or i = 140 Or i = 147 Or i = 154 Then wss("Project_Data").Range("Coll_Service" & Y & "Startyear").Value = "": Y = Y + 1: GoTo 1
        
        With ws.Cells(i, j)
            wss("Project_Data").Range(FormulaInCell) = ""
        End With
1                 Next i
        
        
        wss("Project_Data").Range("CellsToDelete").Value = ""
        wss("Project_Data").Range("AssetService_Delete").Value = ""
        wss("Project_Data").Range("OandM_Delete").Value = ""
        wss("Project_Data").Range("Coll_AmountOfAssets").Value = 1
        wss("Project_Data").Range("Amount_of_Services").Value = 1
        SpinnerAssetsNum_Change
        SpinnerServiceNum_Change
        
        
        '--------------------------------------------------------------------- OPTION AND IMAGE BUTTONS
        Dim myArray As Variant
        Dim myString As String
        
        myString = "Gov_funded;User_funded;Comb_funded;Local_CurIco;Fx_CurIco;EnterManual;CalculPFRAM"
        
        myArray = Split(myString, ";")
        
        For x = LBound(myArray) To UBound(myArray)
            ActiveSheet.Shapes.Range(Array(myArray(x))).Select
            Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        Next x
        
        myString = "OB_DB;OB_DBFO;OB_BOT;OB_BBO;OB_BOO"
        
        myArray = Split(myString, ";")
        
        For x = LBound(myArray) To UBound(myArray)
            ActiveSheet.Shapes.Range(Array(myArray(x))).Fill.Visible = msoFalse
        Next x
        '---------------------------------------------------------------------
        
        '---------------------------------------------------------------------DELETE BOXES
        ActiveWorkbook.Names.Add Name:="Box_delete", RefersToR1C1:= _
        "=Project_Data!R237C66:R242 C115,Project_Data!R247C66:R247C115,Project_Data!R252C66:R252C115,Project_Data!R257C66:R258C115,Project_Data!R263C66:R264C115,Project_Data!R269C66:R269C115,Project_Data!R282C66:R282C114,Project_Data!R347C66:R348C114,Project_Data!R353C66:R354C114,Project_Data!R387C66:R387C114,Project_Data!R421C66:R421C114,Project_Data!R455C66:R460C114"
        
        wss("Project_Data").Range("Box_delete").Value = ""
        '---------------------------------------------------------------------
        
        UnhideAll
        AlignServiceAssetTab
        
        
        For i = 2 To 12
            ActiveSheet.Shapes.Range(Array("CB_" & i)).Select
            Selection.Value = xlOff
        Next i
        
        'Name saved when new project

        If ProjectOptionSelected = "New" Then
            NewProject
            Dim lCol   As Long
            lCol = Worksheets("Store_Data").Cells(11, Columns.Count).End(xlToLeft).Column
            
            Range("Coll_Nameoftheproject") = projName
            ws.Cells(17, lCol) = projName
        End If
        
        New_Data_RiskM                            'ADD PROJECT TO RISK MATRIX
        
        ws.Activate
        
        Menu1_Click
        
        If ProjectOptionSelected = "New" Then Call SaveBT_Click

        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        
End Sub

Private Sub NewProject()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim setPnt As Range: Set setPnt = Worksheets("Store_Data").Range("E11")
    Dim rngPivot As Range: Set rngPivot = Worksheets("Store_Data").Range("B11:B276")
    Dim lCol   As Long
    
    lCol = Worksheets("Store_Data").Cells(9, 500).End(xlToLeft).Column
    
    If setPnt = "" Then                           'Project 1
        rngPivot.Copy
        setPnt.Offset(-2, 0).Value = "P1"
        setPnt.Offset(-2, 0).Interior.ThemeColor = xlThemeColorDark2
        setPnt.PasteSpecial xlPasteValues
    Else
        rngPivot.Copy
        Worksheets("Store_Data").Cells(9, lCol + 1).Value = "P" & (lCol - 3)
        Worksheets("Store_Data").Cells(9, lCol + 1).Interior.ThemeColor = xlThemeColorDark2
        Worksheets("Store_Data").Cells(11, lCol + 1).PasteSpecial xlPasteValues
    End If
    '
    ListCombobox1                                 'Add to list with new project
    
    ThisWorkbook.Sheets("Project_Data").Range("current_project") = lCol - 4
    
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Private Sub SaveProgress()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim setPnt As Range: Set setPnt = Worksheets("Store_Data").Range("E11")
    Dim rngPivot As Range: Set rngPivot = Worksheets("Store_Data").Range("B11:B276")
    Dim lCol   As Long
    
    lCol = Worksheets("Store_Data").Cells(11, Columns.Count).End(xlToLeft).Column
    
    If setPnt = "" Then                           'Project 1
        rngPivot.Copy
        setPnt.Offset(-2, 0).Value = "P1"
        setPnt.Offset(-2, 0).Interior.ThemeColor = xlThemeColorDark2
        setPnt.PasteSpecial xlPasteValues
    Else
        rngPivot.Copy
        Worksheets("Store_Data").Cells(11, lCol).PasteSpecial xlPasteValues
    End If
    
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

'**********************************************************
'*      LOAD IMAGES
'**********************************************************

Private Sub LoadImages()

    ' GOV FUNDED IMAGE
    If Range("Coll_Howtheprojectisfunded").Value = "Tax payers--public funded" Then
        
        ActiveSheet.Shapes.Range(Array("Use_FundIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        ActiveSheet.Shapes.Range(Array("CombIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        
        ActiveSheet.Shapes.Range(Array("Gov_FundIco")).Select
        With Selection.ShapeRange.Fill
            If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            Else
                .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
            End If
        End With
        
        Range("A348,A354").EntireRow.Hidden = False
        Range("A347,A353").EntireRow.Hidden = True
        
        Range("AA1").Select
        
        ActiveSheet.Shapes.Range(Array("Gov_fund")).Visible = False
    End If
    
    ' USER FUNDED IMAGE
    If Range("Coll_Howtheprojectisfunded").Value = "Users of the services--user funded" Then
        
        ActiveSheet.Shapes.Range(Array("Gov_FundIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        ActiveSheet.Shapes.Range(Array("CombIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        
        ActiveSheet.Shapes.Range(Array("Use_FundIco")).Select
        
        With Selection.ShapeRange.Fill
            
            If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            Else
                .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
            End If
        End With
        
        Range("A348,A354").EntireRow.Hidden = True
        Range("A347,A353").EntireRow.Hidden = False
        
        Range("AA1").Select
        
        ActiveSheet.Shapes.Range(Array("Gov_fund")).Visible = False
    End If
    
    ' COMBINED FUNDED IMAGE
    If Range("Coll_Howtheprojectisfunded").Value = "Combined" Then
        
        ActiveSheet.Shapes.Range(Array("Gov_FundIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        ActiveSheet.Shapes.Range(Array("Use_FundIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        
        ActiveSheet.Shapes.Range(Array("CombIco")).Select
        
        With Selection.ShapeRange.Fill
            
            If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            Else
                .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
            End If
        End With
        
        Range("A348,A354").EntireRow.Hidden = False
        Range("A347,A353").EntireRow.Hidden = False
        
        Range("AA1").Select
        
        ActiveSheet.Shapes.Range(Array("Gov_fund")).Visible = True
    End If
    
    ' LOCAL CURRENCY IMAGE
    If Range("Coll_Currency").Value = "Dom" Then
        
        ActiveSheet.Shapes.Range(Array("Fx_CurIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        
        ActiveSheet.Shapes.Range(Array("Local_CurIco")).Select
        
        With Selection.ShapeRange.Fill
            
            If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            Else
                .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
            End If
        End With
        
        Range("AA1").Select
    End If
    
    ' FX CURRENCY IMAGE
    If Range("Coll_Currency").Value = "FX" Then
        
        ActiveSheet.Shapes.Range(Array("Local_CurIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        
        ActiveSheet.Shapes.Range(Array("Fx_CurIco")).Select
        
        With Selection.ShapeRange.Fill
            
            If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            Else
                .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
            End If
        End With
        Range("AA1").Select
    End If
    
    ' MANUAL CALCULATION IMAGE
    If Range("Coll_CalculateRevenue").Value = "Manual" Then
        
        Range("Coll_CalculateRevenue").Value = "Manual"
        
        ActiveSheet.Shapes.Range(Array("CalcuIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        
        ActiveSheet.Shapes.Range(Array("EntManIco")).Select
        
        With Selection.ShapeRange.Fill
            
            If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            Else
                .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
            End If
        End With
        
        Range("AA1").Select
    End If
    
    ' PFRAM CALCULATION IMAGE
    If Range("Coll_CalculateRevenue").Value = "Auto" Then
        
        Range("Coll_CalculateRevenue").Value = "Auto"
        
        ActiveSheet.Shapes.Range(Array("EntManIco")).Select
        Selection.ShapeRange.Fill.ForeColor.ObjectThemeColor = msoThemeColorBackground1
        
        ActiveSheet.Shapes.Range(Array("CalcuIco")).Select
        
        With Selection.ShapeRange.Fill
            
            If .ForeColor.ObjectThemeColor = msoThemeColorAccent6 And .ForeColor.Brightness = 0.400000006 Then
                .ForeColor.ObjectThemeColor = msoThemeColorBackground1
            Else
                .ForeColor.ObjectThemeColor = msoThemeColorAccent6: .ForeColor.Brightness = 0.400000006
            End If
        End With
        Range("AA1").Select
    End If
    'ob_YesShare
    
    If Range("Coll_Governmentshareholding").Value <> 0 Then
        ActiveSheet.Shapes.Range(Array("ob_NoShare")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("ob_YesShare")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("ob_YesShare")).Fill.Visible = msoFalse
            Rows("170:172").EntireRow.Hidden = True
        Else
            ActiveSheet.Shapes.Range(Array("ob_YesShare")).Fill.Visible = msoTrue
            Rows("170:172").EntireRow.Hidden = False
        End If
    End If
    
    'ob_NoShare
    
    If Range("Coll_Governmentshareholding").Value = 0 Then
        ActiveSheet.Shapes.Range(Array("ob_YesShare")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("ob_NoShare")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("ob_NoShare")).Fill.Visible = msoFalse
        Else
            ActiveSheet.Shapes.Range(Array("ob_NoShare")).Fill.Visible = msoTrue
            Rows("170:172").EntireRow.Hidden = True
        End If
    End If
    
    'ob_YesOtherPymnt
    
    If Range("Coll_OtherGovPaymnt") = "YES" Then
        ActiveSheet.Shapes.Range(Array("ob_NoOtherPymnt")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("ob_YesOtherPymnt")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("ob_YesOtherPymnt")).Fill.Visible = msoFalse
            Rows("418:422").EntireRow.Hidden = True
        Else
            ActiveSheet.Shapes.Range(Array("ob_YesOtherPymnt")).Fill.Visible = msoTrue
            Rows("418:422").EntireRow.Hidden = False
            Range("yes_no_OP") = "True"
        End If
    End If
    
    
    'ob_NoOtherPymnt
    If Range("Coll_OtherGovPaymnt") = "NO" Then
        ActiveSheet.Shapes.Range(Array("ob_YesOtherPymnt")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("ob_NoOtherPymnt")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("ob_NoOtherPymnt")).Fill.Visible = msoFalse
        Else
            ActiveSheet.Shapes.Range(Array("ob_NoOtherPymnt")).Fill.Visible = msoTrue
            Rows("418:422").EntireRow.Hidden = True
            Range("yes_no_OP") = "False"
        End If
    End If
    
    'OB_DB
    
    If Range("Coll_Typeofproject").Value = "DB" Then
        ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
        Else
            ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoTrue
        End If
    End If
    
    'OB_DBFO
    
    If Range("Coll_Typeofproject").Value = "DBFO" Then
        
        ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
        Else
            ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoTrue
        End If
    End If
    
    'OB_BOT
    
    If Range("Coll_Typeofproject").Value = "BOT" Then
        Range("Coll_Typeofproject").Value = "BOT"
        
        ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
        Else
            ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoTrue
        End If
    End If
    
    'OB_BBO
    
    If Range("Coll_Typeofproject").Value = "BBO" Then
        Range("Coll_Typeofproject").Value = "BBO"
        
        ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
        Else
            ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoTrue
        End If
    End If
    
    'OB_BOO
    
    If Range("Coll_Typeofproject").Value = "BOO" Then
        Range("Coll_Typeofproject").Value = "BOO"
        
        ActiveSheet.Shapes.Range(Array("OB_DBFO")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BOT")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_BBO")).Fill.Visible = msoFalse
        ActiveSheet.Shapes.Range(Array("OB_DB")).Fill.Visible = msoFalse
        
        If ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoTrue Then
            ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoFalse
        Else
            ActiveSheet.Shapes.Range(Array("OB_BOO")).Fill.Visible = msoTrue
        End If
    End If
    
End Sub

Private Sub CB_2_Click()

    Dim Uf, Gf, co As String
    Dim InflationValue As String

    Uf = "Users of the services--user funded"
    Gf = "Tax payers--public funded"
    co = "Combined"
    
    If Range("INFLATION") = False Then InflationValue = "No" Else: InflationValue = "Yes"
    
    If Range("Serv_Funded") = Uf Then             ' USERS FUNDED
        Range("Coll_Price" & Range("CurrentSer") & "UserDomesticinflationindexed") = InflationValue
        Range("Serv_Inflation") = Range("Inflation")
    End If
    
    
    If Range("Serv_Funded") = Gf Then             ' GOV. FUNDED
        Range("Coll_Price" & Range("CurrentSer") & "GovDomesticinflationindexed") = InflationValue
        Range("Serv_Inflation") = Range("Inflation")
    End If
    
    If Range("Serv_Funded") = co Then             ' COMBINED FUNDED
        If Range("UF") = True Then
            Range("Coll_Price" & Range("CurrentSer") & "UserDomesticinflationindexed") = InflationValue
            Range("Serv_Inflation") = Range("Inflation")
        ElseIf Range("GF") = True Then
            Range("Coll_Price" & Range("CurrentSer") & "GovDomesticinflationindexed") = InflationValue
            Range("Serv_InflationGov") = Range("Inflation")
        End If
    End If
    
End Sub
Private Sub CB_3_Click()


    Dim Uf, Gf, co As String
    Dim NerValue As String

    Uf = "Users of the services--user funded"
    Gf = "Tax payers--public funded"
    co = "Combined"
    
    If Range("NER") = False Then NerValue = "No" Else: NerValue = "Yes"
    
    If Range("Serv_Funded") = Uf Then             ' USERS FUNDED
        Range("Coll_Price" & Range("CurrentSer") & "UserNERindexed") = NerValue
        Range("Serv_NER") = Range("NER")
    End If
    If Range("Serv_Funded") = Gf Then             ' GOV. FUNDED
        Range("Coll_Price" & Range("CurrentSer") & "GovNERindexed") = NerValue
        Range("Serv_NER") = Range("NER")
    End If
    
    If Range("Serv_Funded") = co Then             ' COMBINED FUNDED
        If Range("UF") = True Then
            Range("Coll_Price" & Range("CurrentSer") & "UserNERindexed") = NerValue
            Range("Serv_NER") = Range("NER")
        ElseIf Range("GF") = True Then
            Range("Coll_Price" & Range("CurrentSer") & "GovNERindexed") = NerValue
            Range("Serv_NERGov") = Range("NER")
        End If
    End If
    
End Sub

Private Sub CB_5_Click()

    Dim Uf, Gf, co As String
    Dim GDPValue As String

    Uf = "Users of the services--user funded"
    Gf = "Tax payers--public funded"
    co = "Combined"
    
    If Range("GDP") = False Then GDPValue = "No" Else: GDPValue = "Yes"
    
    If Range("Serv_Funded") = Uf Then             ' USERS FUNDED
        Range("Coll_Demand" & Range("CurrentSer") & "UserLinktoGDP") = GDPValue
        Range("Serv_GDP") = Range("GDP")
    End If
    If Range("Serv_Funded") = Gf Then             ' GOV. FUNDED
        Range("Coll_Demand" & Range("CurrentSer") & "GovLinktoGDP") = GDPValue
        Range("Serv_GDP") = Range("GDP")
    End If
    
    If Range("Serv_Funded") = co Then             ' COMBINED FUNDED
        If Range("UF") = True Then
            Range("Coll_Demand" & Range("CurrentSer") & "UserLinktoGDP") = GDPValue
            Range("Serv_GDP") = Range("GDP")
        ElseIf Range("GF") = True Then
            Range("Coll_Demand" & Range("CurrentSer") & "GovLinktoGDP") = GDPValue
            Range("Serv_GDPGov") = Range("GDP")
        End If
    End If
    
End Sub

Private Sub Cont_ProjSelection_Click()

    Dim lCol   As Long

    If Range("Ser_Name") = Worksheets("Main_Menu").ComboBox1.Value Then
        Dim answer As Integer
        
        answer = MsgBox("This project is already loaded. Do you want to continue with this version or load the project from the last version saved?", vbInformation + vbYesNo, "Option")
        
        Select Case answer
            Case vbYes
                ThisWorkbook.Worksheets("Project_Data").Activate
                Exit Sub
            Case vbNo
                'Do Nothing
        End Select
    End If
    
    ThisWorkbook.Sheets("Project_Data").Range("current_project") = Worksheets("Main_Menu").ComboBox1.ListIndex
    
    If Worksheets("Main_Menu").ComboBox1.Value = "<< Select project from list to edit >>" Then Exit Sub
    
    If ProjectOptionSelected = "Edit" Then loadDataofProjEdit
    If ProjectOptionSelected = "New" Then NewProjectEmpty
    lCol = Worksheets("Store_Data").Cells(11, Columns.Count).End(xlToLeft).Column
    'ThisWorkbook.Sheets("Project_Data").Range("current_project") = lCol - 4
    'Back_ProjData_Click
    EIR_SeekTemplate
    
End Sub

Private Sub OB_Dom()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ActiveSheet.Shapes.Range(Array("OB_Dom")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_FX")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("OB_Dom")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("OB_Dom")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("OB_Dom")).Fill.Visible = msoTrue
    End If
    
    If Range("Currency_Portfolio").Value = "Dom" Then Exit Sub
    Range("Currency_Portfolio").Value = "Dom"
    
    Range("AA1").Select
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
End Sub
Private Sub OB_FX()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ActiveSheet.Shapes.Range(Array("OB_FX")).Fill.Visible = msoFalse
    ActiveSheet.Shapes.Range(Array("OB_Dom")).Fill.Visible = msoFalse
    
    If ActiveSheet.Shapes.Range(Array("OB_FX")).Fill.Visible = msoTrue Then
        ActiveSheet.Shapes.Range(Array("OB_FX")).Fill.Visible = msoFalse
    Else
        ActiveSheet.Shapes.Range(Array("OB_FX")).Fill.Visible = msoTrue
    End If
    
    If Range("Currency_Portfolio").Value = "FX" Then Exit Sub
    Range("Currency_Portfolio").Value = "FX"
    
    Range("AA1").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
End Sub

Public Sub Print_Report_Click()
    Dim answer As Integer

    answer = MsgBox("The active project " & Range("Coll_Nameoftheproject") & " is ready to print. If you want to select another project, please do so from the Project Data menu.", vbQuestion + vbOKOnly, "Print report")
    
    Print_ReportMenu.Show
End Sub

Private Sub Cont_TBD_Click()
    EIR_SeekTemplate
    Risk_Menu.Show
End Sub

Private Sub RevenueManual()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim i, j   As Single
    Dim dataArr, rngArr As Variant
    Dim str, rngStr    As String
    
    rngStr = "Stor_Enterrevenueforcasted"
    
    rngArr = Split(rngStr, ";")
    
    For j = LBound(rngArr) To UBound(rngArr)
        str = Range(rngArr(j))
        dataArr = Split(str, ";")
        For i = LBound(dataArr) To UBound(dataArr)
            Worksheets("Template").Cells(460, 4 + i) = dataArr(i)
        Next i
    Next j
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

'Private Sub LoadRevenuesManual()
'    Application.ScreenUpdating = False
'    Application.Calculation = xlManual
'    Application.DisplayAlerts = False
'    Application.EnableEvents = False
'
'    Dim i, j  As Single
'    Dim dataArr, rngArr As Variant
'    Dim str, rngStr    As String
'
'    Range(Sp_number).Value = 0
'
'    rngStr = "Stor_Enterrevenueforcasted"
'
'    rngArr = Split(rngStr, ";")
'
'    For j = LBound(rngArr) To UBound(rngArr)
'        str = Range(rngArr(j))
'        dataArr = Split(str, ";")
'        For i = LBound(dataArr) To UBound(dataArr)
'            Worksheets("Template").Cells(460, 4 + i) = dataArr(i)
'        Next i
'    Next j
'
'    rngStr = "Stor_Enterrevenueforcasted2"
'
'    rngArr = Split(rngStr, ";")
'
'    For j = LBound(rngArr) To UBound(rngArr)
'        str = Range(rngArr(j))
'        dataArr = Split(str, ";")
'        For i = LBound(dataArr) To UBound(dataArr)
'            Worksheets("Template").Cells(460, 4 + i) = dataArr(i)
'        Next i
'    Next j
'
'    Application.ScreenUpdating = True
'    Application.Calculation = xlAutomatic
'    Application.DisplayAlerts = True
'    Application.EnableEvents = True
'
'End Sub

Private Sub SaveBT_Click()

    Dim wb     As Workbook: Set wb = ThisWorkbook
    Dim ws     As Worksheet: Set ws = wb.Worksheets("Store_Data")

    Sheets("Project_Data").Range("$B$5") = Sheets("Project_Data").Range("$EA$169") 'Manual revenues
    Sheets("Project_Data").Range("$B$6") = Sheets("Project_Data").Range("$EA$168") 'Manual revenues1

    With ws
        .Range("B11:B1000").Copy
        .Cells(11, (ThisWorkbook.Sheets("Project_Data").Range("current_project") + 5)).PasteSpecial Paste:=xlPasteValues
    End With
    
End Sub
Private Sub Saves_OtherAdj_Factor()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim ws     As Worksheet: Set ws = ThisWorkbook.Worksheets("Project_Data")
    Dim arr    As Variant
    Dim x      As Long
    
    'Gov. Funded
    For i = 0 To 9                                'Price
        If Sheets("Project_Data").Cells(315 + i, 157) = "False" Then Range(Cells(303 + i, 167), Cells(303 + i, 216)).ClearContents: GoTo 1
        arr = Split(ws.Cells(315 + i, 155).Value, ";")
        For x = LBound(arr) To UBound(arr)
            ws.Cells(303 + i, 167 + x).Value = arr(x)
        Next x
1                 Next i
        
        For i = 0 To 9                            ' Demand
            If Sheets("Project_Data").Cells(315 + i, 158) = "False" Then Range(Cells(313 + i, 167), Cells(313 + i, 216)).ClearContents: GoTo 2
            arr = Split(ws.Cells(315 + i, 156).Value, ";")
            For x = LBound(arr) To UBound(arr)
                ws.Cells(313 + i, 167 + x).Value = arr(x)
            Next x
2                     Next i
            
            'User funded
            For i = 0 To 9                        'Price
                If Sheets("Project_Data").Cells(325 + i, 157) = "False" Then Range(Cells(324 + i, 167), Cells(324 + i, 216)).ClearContents: GoTo 3
                arr = Split(ws.Cells(325 + i, 155).Value, ";")
                For x = LBound(arr) To UBound(arr)
                    ws.Cells(324 + i, 167 + x).Value = arr(x)
                Next x
3                         Next i
                For i = 0 To 9                    'Demand
                    If Sheets("Project_Data").Cells(325 + i, 158) = "False" Then Range(Cells(334 + i, 167), Cells(334 + i, 216)).ClearContents: GoTo 4
                    arr = Split(ws.Cells(325 + i, 156).Value, ";")
                    For x = LBound(arr) To UBound(arr)
                        ws.Cells(334 + i, 167 + x).Value = arr(x)
                    Next x
4                             Next i
                    
                    'Other option on/off
                    Sheets("Project_Data").Range("$FC$341") = Sheets("Project_Data").Range("$FC$342")
                    Sheets("Project_Data").Range("$FD$341") = Sheets("Project_Data").Range("$FD$342")
                    
                    Application.ScreenUpdating = True
                    Application.Calculation = xlAutomatic
                    Application.DisplayAlerts = True
                    Application.EnableEvents = True
                    
End Sub

Public Sub Run_Shock()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Sheets("SensitivityAnalysis").Activate
    
    Call Baseline_Shock
    Range("shocks_range").Select
    Selection.Copy
    Range("X57").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("GDP_shock,start_year_GDP_shock,end_year_GDP_shock,er_shock,start_year_er_shock,end_year_er_shock,inflation_shock,start_year_inflation_shock,end_year_inflation_shock,adjustment_inflation___er_shock").Select
    Selection.ClearContents
    
    Call Shock_Row
    Call Baseline_Shock
    
    Range("X57:AB63").Select
    Selection.Copy
    Range("GDP_shock").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("shocks_rangeOLD").Select
    Selection.ClearContents
    
    Range("BD206").Select
    Selection.Copy
    Range("BD205").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("X57:AE63").ClearContents
    
    Call Macro_Shock_Row
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Public Sub Clear_Cont()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim a      As Long
    
    Sheets("SensitivityAnalysis").Activate
    
    Range("GDP_shock,start_year_GDP_shock,end_year_GDP_shock,er_shock,start_year_er_shock,end_year_er_shock,inflation_shock,start_year_inflation_shock,end_year_inflation_shock,adjustment_inflation___er_shock").Select
    Selection.ClearContents
    Range("C250:AZ289").Select
    Selection.ClearContents
    Range("BD205").ClearContents
    Range("BD208:BD237").Value = 0
    Call Shock_Row
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Private Sub Baseline_Shock()
    Range("C208:AZ247").Select
    Selection.Copy
    Range("C250").Select
    Selection.PasteSpecial Paste:=xlPasteValues
End Sub

Private Sub Shock_Row()
    ActiveWindow.ScrollRow = 53
    Range("A53").Select
End Sub

Private Sub Macro_Shock_Row()
    ActiveWindow.ScrollRow = 73
    Range("A73").Select
End Sub

Private Sub Proj_Shock_Row()
    ActiveWindow.ScrollRow = 127
    Range("A127").Select
End Sub

'=================================REPORT GENERATOR

Private Sub ResetTitleCharts()

    Dim arr(1) As Variant

    For i = 1 To 22
        ActiveSheet.ChartObjects("Chart " & i).Activate
        With ActiveChart
            arr(1) = .ChartTitle.Text
            .ChartTitle.Delete
            .HasTitle = True
            .ChartTitle.Text = arr(1)
            .ChartTitle.Font.Name = "Bahnschrift Light SemiCondensed"
            .ChartTitle.Font.Size = 12
            .ChartTitle.Font.Bold = False
        End With
    Next i
    
End Sub

Public Sub Proj_Summary()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim rngCopy As Range
    Dim i      As Long
    
    Set rngCopy = Range("template_projectSUM")
    i = Range("current_project").Value
    rngCopy.Copy
    
    Sheets("SummaryProjects").Select
    Sheets("SummaryProjects").Cells(16 * (i + 1), 1).PasteSpecial _
                                       Paste:=xlPasteValuesAndNumberFormats
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub Gov_St()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim rngCopy As Range
    Dim i      As Long
    
    Set rngCopy = Range("Gov_St")
    i = Range("current_project").Value
    rngCopy.Copy
    Sheets("GovernmentStatements").Cells(73 * (i + 1), 1).PasteSpecial _
                                            Paste:=xlPasteValuesAndNumberFormats
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub

Public Sub ChartsAligmentOnPage()
    Dim w, h, TopPosition, LeftPosition, i, NumCols As Long

    Dim MS, FS, AR, CL, CT1, CT2, CT3, CLR, CS As Long

    MS = 351                                      '    MACROECONOMIC SENSITIVITY
    FS = 406                                      '    FISCAL SENSITIVITY
    AR = 274                                      '    ACCOUNTING_REPORTING
    CL = 317                                      '    CONT.LIAB
    CT1 = 480                                     '    CONTRACT TERMINATION1
    CT2 = 507                                     '    CONTRACT TERMINATION 2
    CT3 = 535                                     '    CONTRACT TERMINATION 3
    CLR = 575                                     '    CONTINGENT LIABILITIES REVENUES
    CS = 251                                      '    CASH FLOW
    Worksheets("ReportTemplate").Activate
    'MACROECONOMIC SENSITIVITY
    NumCols = 2
    LeftPosition = 2
    w = (Worksheets("ReportTemplate").Range(Cells(363, 1), Cells(364, 17)).Width - 8) / 2
    h = (Worksheets("ReportTemplate").Range(Cells(3364, 1), Cells(3402, 17)).Height) / 2
    
    TopPosition = Cells(MS, 1).Top
    LeftPosition = Cells(364, 1).Left + 2
    For i = 7 To 10
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 1) Mod NumCols) * w
            If i = 7 Or i = 8 Then .Top = TopPosition Else: .Top = TopPosition + Int((i - 7) / NumCols) * h
        End With
    Next i
    
    'FISCAL SENSITIVITY
    TopPosition = Cells(FS, 1).Top
    LeftPosition = 2
    For i = 11 To 16
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 1) Mod NumCols) * w
            .Top = TopPosition + Int((i - 11) / NumCols) * h 'If i = 11 Or i = 12 Then .Top = TopPosition Else: .Top = TopPosition + Int((i - 11) / NumCols) * h
        End With
    Next i
    
    'ACCOUNTING & REPORTING
    TopPosition = Cells(AR, 1).Top
    LeftPosition = 2
    For i = 2 To 4
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 2) Mod NumCols) * w
            .Top = TopPosition + Int((i - 2) / NumCols) * h
        End With
    Next i
    
    'CONT. LIAB
    TopPosition = Cells(CL, 1).Top
    LeftPosition = 2
    For i = 5 To 6
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            .Height = h
            If i = 2 Then .Left = 2 Else: .Left = LeftPosition + ((i - 1) Mod NumCols) * w
            .Top = TopPosition
        End With
    Next i
    
    'CONTRACT TERMINATION 1
    TopPosition = Cells(CT1, 1).Top
    LeftPosition = 2
    For i = 17 To 18
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 17) Mod NumCols) * w
            .Top = TopPosition + Int((i - 17) / NumCols) * h
        End With
    Next i
    
    'CONTRACT TERMINATION 2
    TopPosition = Cells(CT2, 1).Top
    LeftPosition = 2
    For i = 19 To 20
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 19) Mod NumCols) * w
            .Top = TopPosition + Int((i - 19) / NumCols) * h
        End With
    Next i
    
    'CONTRACT TERMINATION 3
    TopPosition = Cells(CT3, 1).Top
    LeftPosition = 2
    For i = 21 To 22
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 21) Mod NumCols) * w
            .Top = TopPosition + Int((i - 21) / NumCols) * h
        End With
    Next i
    
    'CONTINGENT LIABILITIES REVENUES
    TopPosition = Cells(CLR, 1).Top
    LeftPosition = 2
    For i = 23 To 25
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            If i = 25 Then .Width = w * 2
            .Height = h
            .Left = LeftPosition + ((i - 23) Mod NumCols) * w
            .Top = TopPosition + Int((i - 23) / NumCols) * h
        End With
    Next i
    
    'CASH FLOW
    NumCols = 1
    LeftPosition = 2
    w = (Range(Cells(364, 1), Cells(364, 16)).Width - 8)
    TopPosition = Cells(CS, 1).Top
    For i = 1 To 1
        With ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i)
            .Width = w
            .Height = h
            .Left = LeftPosition + ((i - 1) Mod NumCols) * w
            .Top = TopPosition + Int((i - 1) / NumCols) * h
        End With
    Next i
    
End Sub


Public Sub ChartAjustYearHorizon()

    Dim c, i, x As Long

    c = Sheets("TStore_Data").Range("$B$20").Value
    
    For i = 1 To 25
        On Error Resume Next
        ThisWorkbook.Worksheets("ReportTemplate").ChartObjects("Chart " & i).Activate
        For x = 1 To 50
            ActiveChart.ChartGroups(1).FullCategoryCollection(x).IsFiltered = True
        Next x
        
        For x = 1 To c
            ActiveChart.ChartGroups(1).FullCategoryCollection(x).IsFiltered = False
        Next x
        
    Next i
    
End Sub

'RISK MATRIX ==============================================================================================================================

Private Sub New_Data_RiskM()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim myRowArr, myAsseArr As Variant
    Dim myString, myRow As String
    Dim x, i    As Long
    Dim DataRange, cell As Range
    Dim wb     As Workbook: Set wb = ThisWorkbook
    Dim ws     As Worksheet
    Set ws = wb.Sheets("Risk_Matrix")
    
    'LOAD PROJECT DATA
    ws.Select
    
    ws.Range(Cells(18, (24 + Range("current_project") + 1)), Cells(318, (24 + Range("current_project") + 1))).Select
    Selection.Copy
    ws.Cells(18, 24).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    
    Set DataRange = ws.Range("T18:T304")
    
    For Each cell In DataRange.Cells
        If cell.Value = "" Then GoTo 1
        myRow = myRow & ";" & cell.Value
1                 Next cell
        
        myString = myRow
        myRowArr = Split(myString, ";")
        
        For x = LBound(myRowArr) + 1 To UBound(myRowArr)
            ws.Cells(Cells(myRowArr(x), 20), 6).ClearContents
            ws.Cells(Cells(myRowArr(x), 20), 7).ClearContents
            ws.Cells(Cells(myRowArr(x), 20), 9).ClearContents
        Next x
        
        If Cells(17, 25) = "" Then
            Cells(17, 25) = "P1"
            Range("X17:X318").Copy
            Cells(17, 24).End(xlToRight).Select
            Selection.PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        Else
            Cells(17, 24).End(xlToRight).Select
            Selection.AutoFill Destination:=Range(Cells(17, ActiveCell.Column), Cells(17, ActiveCell.Column + 1)), Type:=xlFillDefault
            Range("X18:X318").Copy
            Cells(17, 24).End(xlToRight).Select
            Selection.PasteSpecial Paste:=xlPasteFormats
            Application.CutCopyMode = False
        End If
        
        Load_Data_RiskM
        
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
End Sub
Public Sub Load_Data_RiskM()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim myRowArr, myAsseArr As Variant
    Dim myString, myRow As String
    Dim x, i    As Long
    Dim DataRange, cell As Range
    Dim wb     As Workbook: Set wb = ThisWorkbook
    Dim ws     As Worksheet
    Set ws = wb.Sheets("Risk_Matrix")
    
    'LOAD PROJECT DATA
    ws.Select
    
    ws.Range(Cells(18, (24 + Range("current_project") + 1)), Cells(318, (24 + Range("current_project") + 1))).Select
    Selection.Copy
    ws.Cells(18, 24).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    
    Set DataRange = ws.Range("T18:T304")
    
    For Each cell In DataRange.Cells
        If cell.Value = "" Then GoTo 1
        myRow = myRow & ";" & cell.Value
1                 Next cell
        
        myString = myRow
        myRowArr = Split(myString, ";")
        
        For x = LBound(myRowArr) + 1 To UBound(myRowArr)
            ws.Cells(Cells(myRowArr(x), 20), 6).ClearContents
            ws.Cells(Cells(myRowArr(x), 20), 7).ClearContents
            ws.Cells(Cells(myRowArr(x), 20), 9).ClearContents
        Next x
        
        If ws.Range("empty_risk").Value = 0 Then GoTo 2
        
        On Error Resume Next
        For x = LBound(myRowArr) + 1 To UBound(myRowArr)
            myAsseArr = Split((Cells(myRowArr(x), 24).Value), ";")
            ws.Cells(Cells(myRowArr(x), 20), 6) = myAsseArr(0)
            ws.Cells(Cells(myRowArr(x), 20), 7) = myAsseArr(1)
            ws.Cells(Cells(myRowArr(x), 20), 9) = myAsseArr(2)
        Next x
        
        For i = 0 To 10
            ws.Cells(308 + i, 19) = ws.Cells(308 + i, 24)
        Next i
2
        Application.ScreenUpdating = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        
End Sub

Public Sub Save_Data_RiskM()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim wb     As Workbook: Set wb = ThisWorkbook
    Dim ws     As Worksheet
    Set ws = wb.Sheets("Risk_Matrix")
    
    Application.GoTo Reference:="Risk_Assess"
    Selection.Copy
    
    ws.Cells(17, (24 + Range("current_project") + 1)) = "P" & Range("current_project") + 1
    ws.Cells(18, (24 + Range("current_project") + 1)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub ExportAsPDF()


    Dim FolderPath As String
    FolderPath = ActiveWorkbook.Path
    
    On Error GoTo ErrorHandler
    Sheets(Array("ReportTemplate", "Report_Annex", "Risk_Matrix")).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=FolderPath & "\FCCL Report PDF", _
                                    openafterpublish:=True, ignoreprintareas:=False
    
    MsgBox "The FCCL report have been successfully exported to: " & FolderPath
    
    ThisWorkbook.Worksheets("Main_Menu").Select
    
    Exit Sub
    
ErrorHandler:
    Dim answer As Integer
    
    answer = MsgBox("It seems that you have a PDF report that is open or active. Please close it and try again.", vbExclamation + vbOKOnly, "Error PDF")
    Exit Sub
    
    
End Sub

Sub ExpensesCashFlow()


    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Dim str    As String
    Dim arr    As Variant
    Dim lCol, fCol, x, n, i, k As Integer
    
    lCol = Sheets("Store_Data").Cells(277, Columns.Count).End(xlToLeft).Column
    fCol = 4
    
    
    For i = fCol To lCol
        arr = Split(Sheets("Store_Data").Cells(277, i).Value, ";")
        n = 1
        For x = LBound(arr) To UBound(arr)
            If Sheets("Store_Data").Cells(15, i) = "YES" Then
                Sheets("Store_Data").Cells(1000 + k, n) = arr(x)
                n = n + 1
            End If
        Next x
        k = k + 1
    Next i
    
    For i = fCol To lCol
        For x = 1 To 50
            Sheets("Store_Data").Select
            Sheets("Macro_Data").Cells(2, 7 + x) = Application.Sum(Sheets("Store_Data").Range(Cells(1000, x), Cells(1000 + k - 1, x)))
        Next x
    Next i
    
    Sheets("Store_Data").Select
    
    Sheets("Store_Data").Range("A980").Select
    Sheets("Store_Data").Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Sheets("Store_Data").Range("A1").Select
    Sheets("Main_Menu").Select
    
    For i = fCol To lCol
        arr = Split(Sheets("Store_Data").Cells(277, i).Value, ";")
        n = 1
        For x = LBound(arr) To UBound(arr)
            If Sheets("Store_Data").Cells(15, i) = "NO" Then
                Sheets("Store_Data").Cells(1000 + k, n) = arr(x)
                n = n + 1
            End If
        Next x
        k = k + 1
    Next i
    
    For i = fCol To lCol
        For x = 1 To 50
            Sheets("Store_Data").Select
            Sheets("Macro_Data").Cells(3, 7 + x) = Application.Sum(Sheets("Store_Data").Range(Cells(1000, x), Cells(1000 + k - 1, x)))
        Next x
    Next i
    
    Sheets("Store_Data").Select
    
    Sheets("Store_Data").Range("A980").Select
    Sheets("Store_Data").Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    Sheets("Store_Data").Range("A1").Select
    Sheets("Main_Menu").Select
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub Menu7_Click()
    UserForm3.Show
     Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
'    'LIGHT
'
'    For i = 1 To 7
'        ActiveSheet.Shapes.Range(Array("Menu" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
'        ActiveSheet.Shapes.Range(Array("Menu" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
'        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).Fill.ForeColor.RGB = RGB(217, 217, 217)
'        ActiveSheet.Shapes.Range(Array("Menu" & i & "B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
'    Next
'
'   'DARK
'        ActiveSheet.Shapes.Range(Array("Menu7")).Fill.ForeColor.RGB = RGB(59, 56, 56)
'        ActiveSheet.Shapes.Range(Array("Menu7")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
'        ActiveSheet.Shapes.Range(Array("Menu7B")).Fill.ForeColor.RGB = RGB(59, 56, 56)
'        ActiveSheet.Shapes.Range(Array("Menu7B")).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
'
'        Application.ScreenUpdating = True
'        Application.Calculation = xlAutomatic
'        Application.DisplayAlerts = True
'        Application.EnableEvents = True
End Sub

sub testing()
    MsgBox "testing github"
End Sub