Private Sub Workbook_Open()

On Error Resume Next

CommandBars.ExecuteMso "MinimizeRibbon"

With Application
    .DisplayFormulaBar = False
    .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",false)"
    .CommandBars("Worksheet Menu Bar").Enabled = False
    .DisplayStatusBar = False
    .DisplayScrollBars = False
    .Width = 1400
    .Height = 690
End With


With Windows("theAmazingSpinningWheel.xlsm")
    .DisplayWorkbookTabs = False
    .DisplayHorizontalScrollBar = False
    .DisplayVerticalScrollBar = False
End With


End Sub
