Attribute VB_Name = "Module1"
'Takes in a string representing the cell which should hold the "Current" status in the NTST MACROS tab
Function updateCurrent(Cell As String)
    'Place "Current above the control button to signify the currect status of the workbook
    Sheets("NTST MACROS").Select
    Range("A1:G1").Clear
    Range(Cell).Value = "Current"
    Range(Cell).Interior.Color = RGB(67, 172, 106)
    Range(Cell).Font.Color = RGB(255, 255, 255)
    Range(Cell).Font.Bold = True
    Range(Cell).HorizontalAlignment = xlCenter
End Function

Sub Phase1()
'
' Phase1 Macro
' This sets the workbook for Phase 1 homework entry by customer.
'
' Keyboard Shortcut: Ctrl+Shift+B

    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect before starting work.
    ThisWorkbook.Unprotect

    Sheets("Instr Phase 2").Visible = False
    Sheets("Instr Phase 3").Visible = False
    Sheets("Diet-Rest").Visible = False
    Sheets("Diet-Supp").Visible = False
    Sheets("eMAR Types Proc").Visible = False
    Sheets("eMAR Events").Visible = False
    Sheets("eMAR Reg").Visible = False
    Sheets("ORDER GROUPS").Visible = False
    Sheets("OE Roles").Visible = False
    Sheets("OE Security").Visible = False
    Sheets("REASON FOR CHANGE").Visible = False
    Sheets("NOTE CATEGORY").Visible = False
    Sheets("Pre-Authorizations").Visible = False
    Sheets("Override-Basic Duplicate").Visible = False
    Sheets("NTST ONLY").Visible = False
    Sheets("Insulin").Visible = False
    'Sheets("NTST MACROS").Visible = False
    
    Sheets("ORDER TYPE").Select
    
    If (ActiveSheet.Columns("P:R").Hidden = False) Then
        Columns("P:R").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    'Hide phase 2 columns in order code
    Sheets("ORDER CODE").Select
    Columns("F:K").Select
    Selection.EntireColumn.Hidden = True

    'Place "Current above the control button to signify the currect status of the workbook
    updateCurrent ("B1")
    
    Sheets("Instr Phase 1").Protect
    
    ThisWorkbook.Protect
    
End Sub

Sub AfterPhase1()
'
' AfterPhase1 Macro
' Run this to restore Workbook after Phase 1
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect before starting work.
    ThisWorkbook.Unprotect
    
    'Unhide the columns in ORDER TYPE
    Sheets("ORDER TYPE").Select
    
    If (ActiveSheet.Columns("P:R").Hidden = True) Then
        Columns("P:R").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER TYPE tab color to orange
    Sheets("ORDER CODE").Tab.Color = RGB(255, 192, 0)
    
    'Set the sheets visiblity
    Sheets("Instr Phase 1").Visible = False
    Sheets("Instr Phase 2").Visible = True
    Sheets("Instr Phase 3").Visible = False
    Sheets("Diet-Rest").Visible = True
    Sheets("Diet-Supp").Visible = True
    Sheets("Insulin").Visible = True
    Sheets("eMAR Types Proc").Visible = True
    Sheets("eMAR Events").Visible = True
    Sheets("eMAR Reg").Visible = True
    Sheets("ORDER GROUPS").Visible = False
    Sheets("OE Roles").Visible = False
    Sheets("OE Security").Visible = False
    Sheets("REASON FOR CHANGE").Visible = False
    Sheets("NOTE CATEGORY").Visible = False
    Sheets("Pre-Authorizations").Visible = False
    Sheets("Override-Basic Duplicate").Visible = False
    Sheets("NTST ONLY").Visible = False
    
    'Place "Current above the control button to signify the currect status of the workbook
    updateCurrent ("C1")
    
    
    
    'Sheets("Instr Phase 3").Visible = True
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-2
    'Sheets("Instr Phase 1").Select
    'ActiveWindow.SelectedSheets.Visible = False
    'Sheets("Instr Phase 3").Select
    'ActiveWindow.SelectedSheets.Visible = False
    'Sheets("OE SOURCE").Select
    'ActiveSheet.Unprotect
    'Sheets("ORDER TYPE").Select
    'ActiveSheet.Unprotect
    'Sheets("FREQUENCY").Select
    'ActiveSheet.Unprotect
    'Sheets("REASON").Select
    'ActiveSheet.Unprotect
    'Sheets("ORDER CODE").Select
    'ActiveSheet.Unprotect
    'Sheets("DC Reason").Select
    'ActiveSheet.Unprotect
    'Sheets("Resume Reason").Select
    'ActiveSheet.Unprotect
    'Sheets("Diet-Rest").Select
    'ActiveWindow.ScrollWorkbookTabs Sheets:=1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-9
    'Sheets("ORDER TYPE").Select
    'ActiveWindow.LargeScroll ToRight:=-1
    'ActiveWindow.SmallScroll ToRight:=-2
    'Sheets("FREQUENCY").Select
    'ActiveWindow.SmallScroll ToRight:=5
    'ActiveWindow.ScrollColumn = 7
    'ActiveWindow.ScrollColumn = 6
    'ActiveWindow.ScrollColumn = 5
    'ActiveWindow.ScrollColumn = 4
    'ActiveWindow.ScrollColumn = 3
    'Sheets("ORDER CODE").Select
    'Columns("E:M").Select
    'Selection.EntireColumn.Hidden = False
    'ActiveWindow.SmallScroll ToRight:=1
    'Sheets("ORDER CODE").Select
    'With ActiveWorkbook.Sheets("ORDER CODE").Tab
    '    .Color = 49407
    '    .TintAndShade = 0
    'End With
    'Sheets("Resume Reason").Select
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-3
    
End Sub

Sub Phase2()
'
' Phase2 Macro
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect before starting work.
    ThisWorkbook.Unprotect
    
    'Unhide the columns in ORDER TYPE
    Sheets("ORDER TYPE").Select
    
    If (ActiveSheet.Columns("P:R").Hidden = True) Then
        Columns("P:R").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER TYPE tab color to orange
    Sheets("ORDER CODE").Tab.Color = RGB(255, 192, 0)
    
    'Set the sheets visiblity
    Sheets("Instr Phase 1").Visible = False
    Sheets("Instr Phase 2").Visible = True
    Sheets("Instr Phase 3").Visible = False
    Sheets("Diet-Rest").Visible = True
    Sheets("Diet-Supp").Visible = True
    Sheets("Insulin").Visible = True
    Sheets("eMAR Types Proc").Visible = True
    Sheets("eMAR Events").Visible = True
    Sheets("eMAR Reg").Visible = True
    Sheets("ORDER GROUPS").Visible = False
    Sheets("OE Roles").Visible = False
    Sheets("OE Security").Visible = False
    Sheets("REASON FOR CHANGE").Visible = False
    Sheets("NOTE CATEGORY").Visible = False
    Sheets("Pre-Authorizations").Visible = False
    Sheets("Override-Basic Duplicate").Visible = False
    Sheets("NTST ONLY").Visible = False
    
    'Place "Current above the control button to signify the currect status of the workbook
    updateCurrent ("D1")
    
    Sheets("Instr Phase 2").Protect


    'Sheets("eMAR Reg").Select
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    
    'Sheets("Resume Reason").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("Hold Reason").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("DC Reason").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("ORDER CODE").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("REASON").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("FREQUENCY").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("ORDER TYPE").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("OE SOURCE").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("Instr Phase 2").Select
    
End Sub

Sub After2()
'
' After2 Macro
' Run this to review what the customer did with phase 2 homework.
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect before starting work.
    ThisWorkbook.Unprotect
    
    'Set the sheets visiblity
    Sheets("Instr Phase 1").Visible = False
    Sheets("Instr Phase 2").Visible = False
    Sheets("Instr Phase 3").Visible = True
    Sheets("Diet-Rest").Visible = True
    Sheets("Diet-Supp").Visible = True
    Sheets("Insulin").Visible = True
    Sheets("eMAR Types Proc").Visible = True
    Sheets("eMAR Events").Visible = True
    Sheets("eMAR Reg").Visible = True
    Sheets("ORDER GROUPS").Visible = True
    Sheets("OE Roles").Visible = True
    Sheets("OE Security").Visible = True
    Sheets("REASON FOR CHANGE").Visible = True
    Sheets("NOTE CATEGORY").Visible = True
    Sheets("Pre-Authorizations").Visible = True
    Sheets("Override-Basic Duplicate").Visible = True
    Sheets("NTST ONLY").Visible = False

    'Place "Current above the control button to signify the currect status of the workbook
    updateCurrent ("E1")




    'Sheets("Instr Phase 3").Visible = True
    'Sheets("ORDER GROUPS").Visible = True
    'Sheets("OE Roles").Visible = True
    'Sheets("OE Security").Visible = True
    'Sheets("REASON FOR CHANGE").Visible = True
    'Sheets("NOTE CATEGORY").Visible = True
    'Sheets("Pre-Authorizations").Visible = True
    'Sheets("Override-Basic Duplicate").Visible = True
    'ActiveSheet.Unprotect
    'Sheets("Pre-Authorizations").Select
    'ActiveSheet.Unprotect
    'Sheets("NOTE CATEGORY").Select
    'ActiveSheet.Unprotect
    'Sheets("REASON FOR CHANGE").Select
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("OE Roles").Select
    'ActiveSheet.Unprotect
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("ORDER GROUPS").Select
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("eMAR Reg").Select
    'ActiveSheet.Unprotect
    'Sheets("eMAR Events").Select
    'ActiveSheet.Unprotect
    'Sheets("eMAR Types Proc").Select
    'ActiveSheet.Unprotect
    'Sheets("Diet-Supp").Select
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("Diet-Rest").Select
    'ActiveSheet.Unprotect
    'Sheets("Resume Reason").Select
    'ActiveSheet.Unprotect
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("Hold Reason").Select
    'ActiveSheet.Unprotect
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("ORDER CODE").Select
    'ActiveSheet.Unprotect
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("REASON").Select
    'ActiveSheet.Unprotect
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("FREQUENCY").Select
    'ActiveSheet.Unprotect
    'Sheets("ORDER TYPE").Select
    'ActiveSheet.Unprotect
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("OE SOURCE").Select
    'ActiveSheet.Unprotect
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("Instr Phase 2").Select
    'ActiveWindow.SelectedSheets.Visible = False
    
End Sub

Sub Phase3()
'
' Phase3 Macro
' Set for customer to do phase 3 homework.
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect before starting work.
    ThisWorkbook.Unprotect
    
    'Set the sheets visiblity
    Sheets("Instr Phase 1").Visible = False
    Sheets("Instr Phase 2").Visible = False
    Sheets("Instr Phase 3").Visible = True
    Sheets("Diet-Rest").Visible = True
    Sheets("Diet-Supp").Visible = True
    Sheets("Insulin").Visible = True
    Sheets("eMAR Types Proc").Visible = True
    Sheets("eMAR Events").Visible = True
    Sheets("eMAR Reg").Visible = True
    Sheets("ORDER GROUPS").Visible = True
    Sheets("OE Roles").Visible = True
    Sheets("OE Security").Visible = True
    Sheets("REASON FOR CHANGE").Visible = True
    Sheets("NOTE CATEGORY").Visible = True
    Sheets("Pre-Authorizations").Visible = True
    Sheets("Override-Basic Duplicate").Visible = True
    Sheets("NTST ONLY").Visible = False

    'Place "Current above the control button to signify the currect status of the workbook
    updateCurrent ("F1")

    'Protect regions of sheets which are visible to the clients
    
    'Hide NTST MACROS




    'Sheets("OE SOURCE").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("ORDER TYPE").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("FREQUENCY").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("REASON").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("ORDER CODE").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("Hold Reason").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("Resume Reason").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("Diet-Rest").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("Diet-Supp").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("eMAR Types Proc").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowSorting:=True
    
    'Sheets("eMAR Events").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("eMAR Reg").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowFiltering:=True
    
    'Sheets("ORDER GROUPS").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("OE Roles").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("OE Security").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("REASON FOR CHANGE").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("NOTE CATEGORY").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("Pre-Authorizations").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    
    'Sheets("Override-Basic Duplicate").Select
    'ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True

End Sub

Sub After3()
'
' After3 Macro
' For Reviewing Phase 3 homework with the customer.
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Disable protection on the workbook so we can alter the visibility of the sheets and do what
    ' we need to do.
    ThisWorkbook.Unprotect

    'Loop through all of the sheets in this workbook. Unprotect each sheet. Make every sheet, column,
    ' and row visible
    For Each loopedSheet In ThisWorkbook.Worksheets
        'Show this sheet
        loopedSheet.Visible = True
        'Activate this sheet so we can loop through it's contents
        loopedSheet.Select
        'Unprotect this sheet
        ActiveSheet.Unprotect
        'Loop through each column in the sheet, unhiding any that are hidden
        For Each sheetRange In ActiveSheet.UsedRange.Columns
            If sheetRange.Hidden = True Then
                ActiveSheet.Cells.EntireColumn.Hidden = False
            End If
        Next sheetRange 'next column
    Next loopedSheet 'next worksheet

    'Place "Current" above the control button to highlight the currect phase of the workbook
    updateCurrent ("G1")
   
End Sub

Sub RESETOE()
'
' RESETOE Macro
' RESETS workbook so macros will fire with no error.  After Reset, start again at Phase 1 and go in order until you get to the state you want.
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Disable protection on the workbook so we can alter the visibility of the sheets and do what
    ' we need to do.
    ThisWorkbook.Unprotect

    'Loop through all of the sheets in this workbook. Unprotect each sheet. Make every sheet, column,
    ' and row visible
    For Each loopedSheet In ThisWorkbook.Worksheets
        'Show this sheet
        loopedSheet.Visible = True
        'Activate this sheet so we can loop through it's contents
        loopedSheet.Select
        'Unprotect this sheet
        ActiveSheet.Unprotect
        'Loop through each column in the sheet, unhiding any that are hidden
        For Each sheetRange In ActiveSheet.UsedRange.Columns
            If sheetRange.Hidden = True Then
                ActiveSheet.Cells.EntireColumn.Hidden = False
            End If
        Next sheetRange 'next column
    Next loopedSheet 'next worksheet
    
    'Return ORDER CODE to it's default color
    Sheets("ORDER CODE").Tab.Color = RGB(0, 176, 240)

    'Place "Current" above the control button to highlight the currect phase of the workbook
    updateCurrent ("A1")
    
End Sub

