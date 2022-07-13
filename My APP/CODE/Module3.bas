Attribute VB_Name = "Module3"
Sub Sheet_Protection()
Attribute Sheet_Protection.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Sheet_Protection Macro
'

'
    Sheets("FORM").Select
    ActiveSheet.Unprotect
    Sheets("FORM").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub
Sub Print_ICA_Reports()
Attribute Print_ICA_Reports.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Print_Reports Macro
'

'
    Call PT_Refresh
    Sheets("ICA").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    
    
End Sub
Sub Print_IFI_Reports()
'
' Print_Reports Macro
'

'
    Call PT_Refresh
    Sheets("IFI").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    
    
End Sub
Sub Print_ISR_Reports()
'
' Print_Reports Macro
'

'
    Call PT_Refresh
    Sheets("ISR").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    
    
End Sub
Sub Print_VD_Reports()
'
' Print_Reports Macro
'

'
    Call PT_Refresh
    Sheets("VD").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    
    
End Sub

Sub Print_VML_Reports()
'
' Print_Reports Macro
'

'
    Call PT_Refresh
    Sheets("VML").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    
    
End Sub

Sub PT_Refresh()
Attribute PT_Refresh.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PT_Refresh Macro
'

'
    Range("A7").Select
    ActiveWorkbook.RefreshAll
End Sub
