Attribute VB_Name = "Module1"
Option Explicit
Public colorSelect As Integer

Sub Spinner2_Change()
        Range("CC_AInc") = Range("CC_AInc_Fix") / 100
End Sub

Sub Spinner3_Change()
    Range("OC_AInc") = Range("OC_AInc_Fix") / 100
End Sub

Sub Spinner4_Change()
    Range("MC_AInc") = Range("MC_AInc_Fix") / 100
End Sub

Sub VfM_P1()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_step1").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub VfM_P2()
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
  
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_step2").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub VfM_P3()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_step3").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub VfM_P4()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_step4").EntireRow.Hidden = True
    
    Range("A1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
''  -------- Table boolean of D-Other Parameter for simulation
    Dim rng1, rng2, rng3  As Range

    Set rng1 = Sheets("Project Information").Range("userfees")
    Set rng2 = Sheets("Project Information").Range("av_payment")
    Set rng3 = Sheets("Project Information").Range("combined")

    Range("vfm_parameters").EntireRow.Hidden = False

    If rng1.Value = True Then
         Range("vfm_Term_Max").EntireRow.Hidden = True
         Range("vfm_Term_Min").EntireRow.Hidden = True
         Range("vfm_Ava_Pymnt").EntireRow.Hidden = True
    End If
    If rng2.Value = True Then
         Range("vfm_Tar_base").EntireRow.Hidden = True
         Range("vfm_Term_Min").EntireRow.Hidden = True
         Range("vfm_Ava_Pymnt").EntireRow.Hidden = True
    End If
    If rng3.Value = True Then
         Range("vfm_Term_Min").EntireRow.Hidden = True
         Range("vfm_Equity").EntireRow.Hidden = True
    End If
    
End Sub

Sub VfM_P5()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_step5").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop

    '=========================
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub VfM_P6()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_p5_rng").EntireRow.Hidden = True
    Range("vfm_step6").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub

Sub VfM_P7()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_p5_rng").EntireRow.Hidden = True
    Range("vfm_p6_rng").EntireRow.Hidden = True
    Range("vfm_step7").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    colorSelect = 1
    Call switch_color
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Sub VfM_P8()

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_p5_rng").EntireRow.Hidden = True
    Range("vfm_p6_rng").EntireRow.Hidden = True
    Range("vfm_p7_rng").EntireRow.Hidden = True
    Range("vfm_step8").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    colorSelect = 2
    Call switch_color
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Sub VfM_P9()


    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_p5_rng").EntireRow.Hidden = True
    Range("vfm_p6_rng").EntireRow.Hidden = True
    Range("vfm_p7_rng").EntireRow.Hidden = True
    Range("vfm_p8_rng").EntireRow.Hidden = True
    Range("vfm_step9").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    colorSelect = 2
    Call switch_color
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Sub VfM_P10()


    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_p5_rng").EntireRow.Hidden = True
    Range("vfm_p6_rng").EntireRow.Hidden = True
    Range("vfm_p7_rng").EntireRow.Hidden = True
    Range("vfm_p8_rng").EntireRow.Hidden = True
    Range("vfm_p9_rng").EntireRow.Hidden = True
    Range("vfm_step10").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    colorSelect = 3
    Call switch_color
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Sub VfM_P11()


    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_p5_rng").EntireRow.Hidden = True
    Range("vfm_p6_rng").EntireRow.Hidden = True
    Range("vfm_p7_rng").EntireRow.Hidden = True
    Range("vfm_p8_rng").EntireRow.Hidden = True
    Range("vfm_p9_rng").EntireRow.Hidden = True
    Range("vfm_p10_rng").EntireRow.Hidden = True
    Range("vfm_step11").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Sub VfM_P11_backbtn()


    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_p5_rng").EntireRow.Hidden = True
    Range("vfm_p6_rng").EntireRow.Hidden = True
    Range("vfm_p7_rng").EntireRow.Hidden = True
    Range("vfm_p8_rng").EntireRow.Hidden = True
    Range("vfm_p9_rng").EntireRow.Hidden = True
    Range("vfm_p10_rng").EntireRow.Hidden = True
    Range("vfm_step11").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    colorSelect = 3
    Call switch_color
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Sub VfM_P12()


    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    '=========================
    
    UnhideAll
    
    Range("Header_VfM").EntireRow.Hidden = True
    Range("vfm_p1_rng").EntireRow.Hidden = True
    Range("vfm_p2_rng").EntireRow.Hidden = True
    Range("vfm_p3_rng").EntireRow.Hidden = True
    Range("vfm_p4_rng").EntireRow.Hidden = True
    Range("vfm_p5_rng").EntireRow.Hidden = True
    Range("vfm_p6_rng").EntireRow.Hidden = True
    Range("vfm_p7_rng").EntireRow.Hidden = True
    Range("vfm_p8_rng").EntireRow.Hidden = True
    Range("vfm_p9_rng").EntireRow.Hidden = True
    Range("vfm_p10_rng").EntireRow.Hidden = True
    Range("vfm_p11_rng").EntireRow.Hidden = True
    Range("vfm_step12").EntireRow.Hidden = True
    Range("A1").Select
    ScrollToTop
    '=========================
    
    colorSelect = 4
    Call switch_color
    
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
End Sub
Sub switch_color()
    
    Dim i As Byte
    For i = 1 To 4 'Color Default
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).Fill.ForeColor.RGB = RGB(217, 217, 217)
        ActiveSheet.Shapes.Range(Array("Menuvfm" & i)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    Next
    
    'Color selected
    ActiveSheet.Shapes.Range(Array("Menuvfm" & colorSelect)).Fill.ForeColor.RGB = RGB(59, 56, 56)
    ActiveSheet.Shapes.Range(Array("Menuvfm" & colorSelect)).TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
End Sub
