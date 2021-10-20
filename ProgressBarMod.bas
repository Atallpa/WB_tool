Attribute VB_Name = "ProgressBarMod"
Option Explicit
Sub UpdateProgress(Pct)
    With ProgressBarForm
        .FrameProgress.Caption = Format(Pct, "0%")
        .LabelProgress.Width = Pct * (.FrameProgress.Width - 10)
        .Repaint
    End With
End Sub

Sub ShowUserForm()
    With ProgressBarForm
        'Use a color from the workbook's theme
        .LabelProgress.BackColor = ActiveWorkbook.Theme. _
            ThemeColorScheme.Colors(msoThemeAccent1)
        .LabelProgress.Width = 0
        .Show
    End With
End Sub
