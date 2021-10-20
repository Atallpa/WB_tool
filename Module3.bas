Attribute VB_Name = "Module3"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("H225").Select
    ActiveSheet.CheckBoxes.Add(354.5, 3710, 72, 72).Select
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

    ActiveSheet.Shapes.Range(Array("test")).Select
    Selection.Delete
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ActiveSheet.Shapes.Range(Array("vfm_RPChk_1")).Select
    With Selection
        .Value = xlOff
        .LinkedCell = "$H$218"
        .Display3DShading = False
    End With
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"

    If ActiveSheet.Shapes.Range(Array("vfm_RPChk_1")).Value = True Then
        MsgBox "Hi"
    End If
   
End Sub
