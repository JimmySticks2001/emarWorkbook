Attribute VB_Name = "Module1"
Sub Phase1()
Attribute Phase1.VB_Description = "This sets the workbook for Phase 1 homework entry by customer."
Attribute Phase1.VB_ProcData.VB_Invoke_Func = "B\n14"
'
' Phase1 Macro
' This sets the workbook for Phase 1 homework entry by customer.
'
' Keyboard Shortcut: Ctrl+Shift+B
'
    Sheets("Instr Phase 2").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Instr Phase 3").Select
    ActiveWindow.SelectedSheets.Visible = False
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets("Diet-Rest").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Diet-Supp").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("eMAR Types Proc").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("eMAR Events").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("eMAR Reg").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("ORDER GROUPS").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("OE Roles").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("OE Security").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("REASON FOR CHANGE").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("NOTE CATEGORY").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Pre-Authorizations").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Override-Basic Duplicate").Select
    ActiveWindow.SelectedSheets.Visible = False
       ActiveWindow.SelectedSheets.Visible = False
    ActiveWindow.ScrollWorkbookTabs Sheets:=-8
    Sheets("ORDER TYPE").Select
    ActiveWindow.SmallScroll ToRight:=11
    ActiveSheet.Unprotect
    Columns("P:R").Select
    Selection.EntireColumn.Hidden = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("FREQUENCY").Select
    ActiveWindow.SmallScroll ToRight:=8
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    Sheets("ORDER CODE").Select
    ActiveSheet.Unprotect
    Columns("F:K").Select
    Selection.EntireColumn.Hidden = True
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    Range("B9").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("Resume Reason").Select
End Sub
Sub AfterPhase1()
Attribute AfterPhase1.VB_Description = "Run this to restore Workbook after Phase 1"
Attribute AfterPhase1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AfterPhase1 Macro
' Run this to restore Workbook after Phase 1
'

'
    Sheets("Instr Phase 2").Visible = True
    Sheets("Diet-Rest").Visible = True
    Sheets("Diet-Supp").Visible = True
    Sheets("eMAR Types Proc").Visible = True
    Sheets("eMAR Events").Visible = True
    Sheets("eMAR Reg").Visible = True
    Sheets("Instr Phase 3").Visible = True
    ActiveWindow.ScrollWorkbookTabs Sheets:=-2
    Sheets("Instr Phase 1").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Instr Phase 3").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("OE SOURCE").Select
    ActiveSheet.Unprotect
    Sheets("ORDER TYPE").Select
    ActiveSheet.Unprotect
    Sheets("FREQUENCY").Select
    ActiveSheet.Unprotect
    Sheets("REASON").Select
    ActiveSheet.Unprotect
    Sheets("ORDER CODE").Select
    ActiveSheet.Unprotect
    Sheets("DC Reason").Select
    ActiveSheet.Unprotect
    Sheets("Resume Reason").Select
    ActiveSheet.Unprotect
    Sheets("Diet-Rest").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-9
    Sheets("ORDER TYPE").Select
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.SmallScroll ToRight:=-2
    Sheets("FREQUENCY").Select
    ActiveWindow.SmallScroll ToRight:=5
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    Sheets("ORDER CODE").Select
    Columns("E:M").Select
    Selection.EntireColumn.Hidden = False
    ActiveWindow.SmallScroll ToRight:=1
    Sheets("ORDER CODE").Select
    With ActiveWorkbook.Sheets("ORDER CODE").Tab
        .Color = 49407
        .TintAndShade = 0
    End With
    Sheets("Resume Reason").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-3
End Sub
Sub Phase2()
Attribute Phase2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Phase2 Macro
'

'
    Sheets("eMAR Reg").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Resume Reason").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Hold Reason").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, _
        AllowFiltering:=True
    Sheets("DC Reason").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True _
        , AllowFiltering:=True
    Sheets("ORDER CODE").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("REASON").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("FREQUENCY").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("ORDER TYPE").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("OE SOURCE").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("Instr Phase 2").Select
End Sub
Sub After2()
Attribute After2.VB_Description = "Run this to review what the customer did with phase 2 homework."
Attribute After2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' After2 Macro
' Run this to review what the customer did with phase 2 homework.
'

'
    Sheets("Instr Phase 3").Visible = True
    Sheets("ORDER GROUPS").Visible = True
    Sheets("OE Roles").Visible = True
    Sheets("OE Security").Visible = True
    Sheets("REASON FOR CHANGE").Visible = True
    Sheets("NOTE CATEGORY").Visible = True
    Sheets("Pre-Authorizations").Visible = True
    Sheets("Override-Basic Duplicate").Visible = True
    ActiveSheet.Unprotect
    Sheets("Pre-Authorizations").Select
    ActiveSheet.Unprotect
    Sheets("NOTE CATEGORY").Select
    ActiveSheet.Unprotect
    Sheets("REASON FOR CHANGE").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("OE Roles").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("ORDER GROUPS").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("eMAR Reg").Select
    ActiveSheet.Unprotect
    Sheets("eMAR Events").Select
    ActiveSheet.Unprotect
    Sheets("eMAR Types Proc").Select
    ActiveSheet.Unprotect
    Sheets("Diet-Supp").Select
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Diet-Rest").Select
    ActiveSheet.Unprotect
    Sheets("Resume Reason").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Hold Reason").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("ORDER CODE").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("REASON").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("FREQUENCY").Select
    ActiveSheet.Unprotect
    Sheets("ORDER TYPE").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("OE SOURCE").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("Instr Phase 2").Select
    ActiveWindow.SelectedSheets.Visible = False
End Sub
Sub Phase3()
Attribute Phase3.VB_Description = "Set for customer to do phase 3 homework."
Attribute Phase3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Phase3 Macro
' Set for customer to do phase 3 homework.
'

'
    Sheets("OE SOURCE").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("ORDER TYPE").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("FREQUENCY").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("REASON").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("ORDER CODE").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("Hold Reason").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, _
        AllowFiltering:=True
    Sheets("Resume Reason").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("Diet-Rest").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("Diet-Supp").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("eMAR Types Proc").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, AllowSorting:= _
        True
    Sheets("eMAR Events").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, AllowSorting:= _
        True, AllowFiltering:=True
    Sheets("eMAR Reg").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:= _
        True
    Sheets("ORDER GROUPS").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("OE Roles").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("OE Security").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("REASON FOR CHANGE").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("NOTE CATEGORY").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("Pre-Authorizations").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Sheets("Override-Basic Duplicate").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFormattingCells:=True, AllowFormattingRows:=True, _
        AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
End Sub
Sub After3()
Attribute After3.VB_Description = "For Reviewing Phase 3 homework with the customer."
Attribute After3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' After3 Macro
' For Reviewing Phase 3 homework with the customer.
'

'
    ActiveSheet.Unprotect
    Sheets("Pre-Authorizations").Select
    ActiveSheet.Unprotect
    Sheets("NOTE CATEGORY").Select
    ActiveSheet.Unprotect
    Sheets("REASON FOR CHANGE").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("OE Security").Select
    ActiveSheet.Unprotect
    Sheets("OE Roles").Select
    ActiveSheet.Unprotect
    Sheets("ORDER GROUPS").Select
    ActiveSheet.Unprotect
    Sheets("eMAR Reg").Select
    ActiveSheet.Unprotect
    Sheets("eMAR Events").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-5
    Sheets("eMAR Types Proc").Select
    ActiveSheet.Unprotect
    Sheets("Diet-Supp").Select
    ActiveSheet.Unprotect
    Sheets("Diet-Rest").Select
    ActiveSheet.Unprotect
    Sheets("Resume Reason").Select
    ActiveSheet.Unprotect
    Sheets("Hold Reason").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-7
    Sheets("DC Reason").Select
    ActiveSheet.Unprotect
    Sheets("ORDER CODE").Select
    ActiveSheet.Unprotect
    Sheets("REASON").Select
    ActiveSheet.Unprotect
    Sheets("FREQUENCY").Select
    ActiveSheet.Unprotect
    Sheets("ORDER TYPE").Select
    ActiveSheet.Unprotect
    Sheets("OE SOURCE").Select
    ActiveSheet.Unprotect
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
End Sub
Sub RESETOE()
Attribute RESETOE.VB_Description = "RESETS workbook so macros will fire with no error.  After Reset, start again at Phase 1 and go in order until you get to the state you want."
Attribute RESETOE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RESETOE Macro
' RESETS workbook so macros will fire with no error.  After Reset, start again at Phase 1 and go in order until you get to the state you want.
'

'
    Sheets("Instr Phase 1").Visible = True
    Sheets("Instr Phase 2").Visible = True
End Sub
