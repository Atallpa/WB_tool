Attribute VB_Name = "Module8"
Option Explicit
Public NewVfmProject As Boolean

Sub LoadVfMProjectData()
    
    'ELIMINAR CUANDO SE CREE LA OPCION DE NUEVO PROYECTO!!!!!!!!***************************************************************
    NewVfmProject = False
    'ELIMINAR CUANDO SE CREE LA OPCION DE NUEVO PROYECTO!!!!!!!!***************************************************************
    
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    Dim ws As Worksheet: Set ws = Sheets("Value for Money")
    
    Dim i, n, ProjActive As Integer:
    
    'Set project selected from column
    If NewVfmProject = True Then
         ProjActive = 78 ' <<<Es la columna BZ que esta vacia
    Else
        ProjActive = ws.Range("B2") + 80
    End If
    
    '==============================================================================================================================================================================
    '****ONLY  THOSE WITH ONE COLUMN DATA ****
    '==============================================================================================================================================================================
    
    'Determine row numbers
    Dim rowStr As String
    rowStr = "18;19;23;24;26;49;75;97;123;125;128;130;132;163;168;169;170;171;172;175;176;177;178;179;182;183;184;185;186;189;279;280;281;300;301;302;303;305;319;321;135"
    Dim rowArr As Variant: rowArr = Split(rowStr, ";")
    
    'Populate data loaded from column of project selected
    For i = LBound(rowArr) To UBound(rowArr)
        If i >= 0 And i < 6 Then n = 6
        If i = 6 Then n = 7
        If i = 7 Then n = 6
        If i >= 8 And i < 30 Then n = 7
        If i > 29 Then n = 6
        If i > 32 Then n = 9
        If i = 40 Then n = 8
        ws.Cells(rowArr(i), n).Value2 = ws.Cells(rowArr(i), ProjActive).Value2
    Next i
    
    
    
    '==============================================================================================================================================================================
    '****ONLY  THOSE WITH MORE THAN ONE COLUMN DATA ****
    '==============================================================================================================================================================================
    Dim CelStr As String
    Dim fila As Variant
    Dim rowArr2 As Variant
    Dim x As Integer
    
    '==============================================================================================================================================================================
    '****General and construction****
    '==============================================================================================================================================================================
    
    rowStr = "29"
    rowArr = Split(rowStr, ";")

        For i = LBound(rowArr) To UBound(rowArr)
        fila = rowArr(i)
        CelStr = ws.Cells(fila, ProjActive)
        If CelStr = "" Then CelStr = "0;0;0;0;0;0;0;0;0"
        rowArr2 = Split(CelStr, ";")
        
        For x = LBound(rowArr2) To UBound(rowArr2)
            For n = 5 To 13
                ws.Cells(fila, n).Value2 = rowArr2(n - 5)
            Next n
        Next x
    Next i
    
    
    '==============================================================================================================================================================================
    '****Operation and maintenance ****
    '==============================================================================================================================================================================
    
    'Concessionaire costs:
    fila = 46
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then CelStr = "0;0;0"
    rowArr = Split(CelStr, ";")
    ws.Cells(fila, 6).Value2 = rowArr(0)
    ws.Cells(fila, 9).Value2 = rowArr(1)
    ws.Cells(fila, 10).Value2 = rowArr(2)
    'Operation costs:
    fila = 47
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then CelStr = "0;0;0"
    rowArr = Split(CelStr, ";")
    ws.Cells(fila, 6).Value2 = rowArr(0)
    ws.Cells(fila, 9).Value2 = rowArr(1)
    ws.Cells(fila, 10).Value2 = rowArr(2)
    'Maintenance costs as % construction cost:
    fila = 52
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then CelStr = "0;0;0"
    rowArr = Split(CelStr, ";")
    ws.Cells(fila, 6).Value2 = rowArr(0)
    ws.Cells(fila, 9).Value2 = rowArr(1)
    ws.Cells(fila, 10).Value2 = rowArr(2)
    'Heavy maintenance 1:
    fila = 53
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then CelStr = "0;0"
    rowArr = Split(CelStr, ";")
    ws.Cells(fila, 6).Value2 = rowArr(0)
    ws.Cells(fila, 10).Value2 = rowArr(1)
    'Heavy maintenance 2:
    fila = 54
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then CelStr = "0;0"
    rowArr = Split(CelStr, ";")
    ws.Cells(fila, 6).Value2 = rowArr(0)
    ws.Cells(fila, 10).Value2 = rowArr(1)
    'Heavy maintenance 3:
    fila = 55
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then CelStr = "0;0"
    rowArr = Split(CelStr, ";")
    ws.Cells(fila, 6).Value2 = rowArr(0)
    ws.Cells(fila, 10).Value2 = rowArr(1)
    'Contract supervision cost:
    fila = 57
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then CelStr = "0;0"
    rowArr = Split(CelStr, ";")
    ws.Cells(fila, 6).Value2 = rowArr(0)
    ws.Cells(fila, 10).Value2 = rowArr(1)
    
    '==============================================================================================================================================================================
    '****Other parameters for simulation ****
    '==============================================================================================================================================================================
    
    rowStr = "101;102;103;104;105"
    rowArr = Split(rowStr, ";")
    For i = LBound(rowArr) To UBound(rowArr)
        fila = rowArr(i)
        CelStr = ws.Cells(fila, ProjActive)
        If CelStr = "" Then CelStr = "0;0;0"
        rowArr2 = Split(CelStr, ";")
        For x = LBound(rowArr2) To UBound(rowArr2)
            ws.Cells(fila, 9).Value2 = rowArr2(0)
            ws.Cells(fila, 11).Value2 = rowArr2(1)
            ws.Cells(fila, 13).Value2 = rowArr2(2)
        Next x
    Next i
    
    '==============================================================================================================================================================================
    '****Demand & revenues ****
    '==============================================================================================================================================================================
    
    'Period: (row 136)
    rowStr = "136"
    rowArr = Split(rowStr, ";")

        For i = LBound(rowArr) To UBound(rowArr)
        fila = rowArr(i)
        CelStr = ws.Cells(fila, ProjActive)
        If CelStr = "" Then CelStr = "0;0;0;0;0;0;0;0"
        rowArr2 = Split(CelStr, ";")
        
        For x = LBound(rowArr2) To UBound(rowArr2)
            For n = 8 To 15
                ws.Cells(fila, n).Value2 = rowArr2(n - 8)
            Next n
        Next x
    Next i
    
    'Annual growth rate of demand (%)
    n = 0
    For i = 140 To 147
        ReDim Preserve rowArr(n)
        rowArr(n) = i
        n = n + 1
    Next i
    
    For i = LBound(rowArr) To UBound(rowArr)
        fila = rowArr(i)
        CelStr = ws.Cells(fila, ProjActive)
        If CelStr = "" Then CelStr = "0;0;0;0;0;0;0;0;0;0;0;0;0;0"
        rowArr2 = Split(CelStr, ";")
        For x = LBound(rowArr2) To UBound(rowArr2)
            For n = 7 To 20
                ws.Cells(fila, n).Value2 = rowArr2(n - 7)
            Next n
        Next x
    Next i
    
    
    '==============================================================================================================================================================================
    '****Other payments from the government ****
    '==============================================================================================================================================================================
    
    rowStr = "207"
    rowArr = Split(rowStr, ";")

        For i = LBound(rowArr) To UBound(rowArr)
        fila = rowArr(i)
        CelStr = ws.Cells(fila, ProjActive)
        If CelStr = "" Then CelStr = "0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
        rowArr2 = Split(CelStr, ";")
        
        For x = LBound(rowArr2) To UBound(rowArr2)
            For n = 5 To 54
                ws.Cells(fila, n).Value2 = rowArr2(n - 5)
            Next n
        Next x
    Next i
    
    '==============================================================================================================================================================================
    '****Risk Parameter ****
    '==============================================================================================================================================================================
   
     'Risk Parameter array <<<<<<<<<<<<<==================
    
    rowStr = "219;220;221;222;226;227;228;229;233;237"
    rowArr = Split(rowStr, ";")

        For i = LBound(rowArr) To UBound(rowArr)
        fila = rowArr(i)
        CelStr = ws.Cells(fila, ProjActive)
        If CelStr = "" Then CelStr = "0;0"
        rowArr2 = Split(CelStr, ";")
        
        For x = LBound(rowArr2) To UBound(rowArr2)
            For n = 7 To 8
                ws.Cells(fila, n).Value2 = rowArr2(n - 7)
            Next n
        Next x
    Next i
    

    '==============================================================================================================================================================================
    '****Risk Allocation****
    '==============================================================================================================================================================================
    
    rowStr = "249;250;251;252;255;256;257;258;261"
    rowArr = Split(rowStr, ";")
    
    For i = LBound(rowArr) To UBound(rowArr)
        fila = rowArr(i)
        CelStr = ws.Cells(fila, ProjActive)
        If CelStr = "" Then CelStr = ";;"
        rowArr2 = Split(CelStr, ";")
        For x = LBound(rowArr2) To UBound(rowArr2)
            ws.Cells(fila, 8).Value2 = rowArr2(0)
            ws.Cells(fila, 11).Value2 = rowArr2(1)
            ws.Cells(fila, 14).Value2 = rowArr2(2)
        Next x
    Next i
    

    '==============================================================================================================================================================================
    '**** Country data****
    '==============================================================================================================================================================================
    
    'Country Data array <<<<<<<<<<<<<==================
    
    rowStr = "274;275"
    rowArr = Split(rowStr, ";")
    
    For i = LBound(rowArr) To UBound(rowArr)
        fila = rowArr(i)
        CelStr = ws.Cells(fila, ProjActive)
        If CelStr = "" Then CelStr = "0;0;0"
        rowArr2 = Split(CelStr, ";")
        For x = LBound(rowArr2) To UBound(rowArr2)
            ws.Cells(fila, 8).Value2 = rowArr2(0)
            ws.Cells(fila, 10).Value2 = rowArr2(1)
            ws.Cells(fila, 12).Value2 = rowArr2(2)
        Next x
    Next i
    
    '***********************************************************************************
    '***OBJECTS***
    '***********************************************************************************
    
    'Funding of the project
    fila = 76
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then
        CelStr = "TRUE;FALSE;FALSE"
    End If
        
    rowArr = Split(CelStr, ";")
    ws.OLEObjects("OptionButton1").Object.Value = rowArr(0)
    ws.OLEObjects("OptionButton2").Object.Value = rowArr(1)
    ws.OLEObjects("OptionButton3").Object.Value = rowArr(2)
    
    'Risk parameters

    fila = "223"
    CelStr = ws.Cells(fila, ProjActive)
    If CelStr = "" Then
        CelStr = "FALSE;FALSE;FALSE;FALSE;FALSE;FALSE;FALSE;FALSE;FALSE;FALSE"
    End If

    rowArr = Split(CelStr, ";")
    If rowArr(0) = "TRUE" Then ws.Shapes("vfm_RPChk_1").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_1").OLEFormat.Object.Value = False
    If rowArr(1) = "TRUE" Then ws.Shapes("vfm_RPChk_2").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_2").OLEFormat.Object.Value = False
    If rowArr(2) = "TRUE" Then ws.Shapes("vfm_RPChk_3").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_3").OLEFormat.Object.Value = False
    If rowArr(3) = "TRUE" Then ws.Shapes("vfm_RPChk_4").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_4").OLEFormat.Object.Value = False
    If rowArr(4) = "TRUE" Then ws.Shapes("vfm_RPChk_5").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_5").OLEFormat.Object.Value = False
    If rowArr(5) = "TRUE" Then ws.Shapes("vfm_RPChk_6").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_6").OLEFormat.Object.Value = False
    If rowArr(6) = "TRUE" Then ws.Shapes("vfm_RPChk_7").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_7").OLEFormat.Object.Value = False
    If rowArr(7) = "TRUE" Then ws.Shapes("vfm_RPChk_8").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_8").OLEFormat.Object.Value = False
    If rowArr(8) = "TRUE" Then ws.Shapes("vfm_RPChk_9").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_9").OLEFormat.Object.Value = False
    If rowArr(9) = "TRUE" Then ws.Shapes("vfm_RPChk_10").OLEFormat.Object.Value = True Else ws.Shapes("vfm_RPChk_10").OLEFormat.Object.Value = False

    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
 
Sub SaveVfMProjectData()
    
    Application.ScreenUpdating = False
    
    Dim rngOrigen, rngFinal As Range
    Dim ProjActive As Integer
    Dim ws As Worksheet: Set ws = Sheets("Value for Money")
    
    Set rngOrigen = ws.Range("CA:CA")
    ProjActive = ws.Range("B2") + 80
    Set rngFinal = ws.Cells(1, ProjActive)
    
    rngOrigen.Copy
    rngFinal.PasteSpecial Paste:=xlPasteValues
    
    Application.ScreenUpdating = True
    
End Sub

Sub StringWithData()

    Dim cell, rng As Range: Set rng = Range("CE2:CE321")
    Dim str As String
    
    For Each cell In rng
        If cell <> "" Then
            str = str & cell.Value & ";"
        End If
    Next cell
    
    Debug.Print str
    
End Sub

Sub ASDSA()
UnhideAll
End Sub
