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

'This function calls the .Protect method on all sheets.
Function protectAll()
    For Each loopedSheet In ThisWorkbook.Worksheets
        loopedSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, AllowSorting:=True, AllowFiltering:=True
    Next loopedSheet 'next worksheet
End Function

'This function calls the .Unprotect method on all sheets.
Function unprotectAll()
    For Each loopedSheet In ThisWorkbook.Worksheets
        loopedSheet.Unprotect
    Next loopedSheet 'next worksheet
End Function

Sub Phase1()
'
' Phase1 Macro
' This sets the workbook for Phase 1 homework entry by customer.
'
' Keyboard Shortcut: Ctrl+Shift+B

    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance.
    Application.ScreenUpdating = False
    
    'Unprotect all of the sheets
    unprotectAll

    Sheets("Instr Phase 1").Visible = True
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
    
    'Hide the columns in ORDER TYPE
    Sheets("ORDER TYPE").Select
    
    If (ActiveSheet.Columns("P:R").Hidden = False) Then
        Columns("P:R").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    'Change ORDER TYPE tab color to blue
    Sheets("ORDER TYPE").Tab.Color = RGB(0, 176, 240)
    
    'Hide the columns in ORDER CODE
    Sheets("ORDER CODE").Select
    
    If (ActiveSheet.Columns("F:K").Hidden = False) Then
        Columns("F:K").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    'Change ORDER CODE tab color to blue
    Sheets("ORDER CODE").Tab.Color = RGB(0, 176, 240)

    'Place "Current above the control button to signify the currect status of the workbook
    updateCurrent ("B1")
    Sheets("NTST MACROS").Visible = False
    
    Sheets("Instr Phase 1").Select
    
    'Call the protectAll function to protect all of the sheets in this workbook
    protectAll
    
End Sub

Sub AfterPhase1()
'
' AfterPhase1 Macro
' Run this to restore Workbook after Phase 1
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect all of the sheets
    unprotectAll
    
    'Unhide the columns in ORDER TYPE
    Sheets("ORDER TYPE").Select
    
    If (ActiveSheet.Columns("P:R").Hidden = True) Then
        Columns("P:R").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER TYPE tab color to orange
    Sheets("ORDER TYPE").Tab.Color = RGB(255, 192, 0)
    
    'Unhide the columns in ORDER CODE
    Sheets("ORDER CODE").Select
    
    If (ActiveSheet.Columns("F:K").Hidden = True) Then
        Columns("F:K").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER CODE tab color to orange
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
    Sheets("NTST MACROS").Visible = True
    
    'Place "Current above the control button to signify the currect status of the workbook
    updateCurrent ("C1")
    
End Sub

Sub Phase2()
'
' Phase2 Macro
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect all of the sheets
    unprotectAll
    
    'Unhide the columns in ORDER TYPE
    Sheets("ORDER TYPE").Select
    
    If (ActiveSheet.Columns("P:R").Hidden = True) Then
        Columns("P:R").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER TYPE tab color to orange
    Sheets("ORDER TYPE").Tab.Color = RGB(255, 192, 0)
    
    'Unhide the columns in ORDER CODE
    Sheets("ORDER CODE").Select
    
    If (ActiveSheet.Columns("F:K").Hidden = True) Then
        Columns("F:K").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER CODE tab color to orange
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
    Sheets("NTST MACROS").Visible = False
    
    Sheets("Instr Phase 2").Select
    
    'Call the protectVisible function to protect all of the sheets which were made visible above
    protectAll

End Sub

Sub After2()
'
' After2 Macro
' Run this to review what the customer did with phase 2 homework.
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect all of the sheets
    unprotectAll
    
    'Unhide the columns in ORDER TYPE
    Sheets("ORDER TYPE").Select
    
    If (ActiveSheet.Columns("P:R").Hidden = True) Then
        Columns("P:R").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER TYPE tab color to orange
    Sheets("ORDER TYPE").Tab.Color = RGB(255, 192, 0)
    
    'Unhide the columns in ORDER CODE
    Sheets("ORDER CODE").Select
    
    If (ActiveSheet.Columns("F:K").Hidden = True) Then
        Columns("F:K").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER CODE tab color to orange
    Sheets("ORDER CODE").Tab.Color = RGB(255, 192, 0)
    
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
    Sheets("NTST MACROS").Visible = True

    'Place "Current above the control button to signify the currect status of the workbook
    updateCurrent ("E1")
    
End Sub

Sub Phase3()
'
' Phase3 Macro
' Set for customer to do phase 3 homework.
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect all of the sheets
    unprotectAll
    
    'Unhide the columns in ORDER TYPE
    Sheets("ORDER TYPE").Select
    
    If (ActiveSheet.Columns("P:R").Hidden = True) Then
        Columns("P:R").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER TYPE tab color to orange
    Sheets("ORDER TYPE").Tab.Color = RGB(255, 192, 0)
    
    'Unhide the columns in ORDER CODE
    Sheets("ORDER CODE").Select
    
    If (ActiveSheet.Columns("F:K").Hidden = True) Then
        Columns("F:K").Select
        Selection.EntireColumn.Hidden = False
    End If
    
    'Change ORDER CODE tab color to orange
    Sheets("ORDER CODE").Tab.Color = RGB(255, 192, 0)
    
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
    Sheets("NTST MACROS").Visible = False
    
    Sheets("Instr Phase 3").Select

    'Call the protectVisible function to protect all of the sheets which were made visible above
    protectAll

End Sub

Sub After3()
'
' After3 Macro
' For Reviewing Phase 3 homework with the customer.
'
    'Turn screen updating off. This prevents Excel from displaying all of the actions taken by
    ' this script. This has a huge impact on macro performance
    Application.ScreenUpdating = False
    
    'Unprotect all of the sheets
    unprotectAll

    'Loop through all of the sheets in this workbook. Make every sheet, column, and row visible
    For Each loopedSheet In ThisWorkbook.Worksheets
        'Show this sheet
        loopedSheet.Visible = True
        'Activate this sheet
        loopedSheet.Select
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
    
    'Unprotect all of the sheets
    unprotectAll

    'Loop through all of the sheets in this workbook. Make every sheet, column, and row visible.
    For Each loopedSheet In ThisWorkbook.Worksheets
        'Show this sheet
        loopedSheet.Visible = True
        'Activate this sheet
        loopedSheet.Select
        'Loop through each column in the sheet, unhiding any that are hidden
        For Each sheetRange In ActiveSheet.UsedRange.Columns
            If sheetRange.Hidden = True Then
                ActiveSheet.Cells.EntireColumn.Hidden = False
            End If
        Next sheetRange 'next column
    Next loopedSheet 'next worksheet
    
    'Return ORDER CODE to it's default color
    Sheets("ORDER CODE").Tab.Color = RGB(0, 176, 240)
    
    'Change ORDER TYPE to it's default color
    Sheets("ORDER TYPE").Tab.Color = RGB(0, 176, 240)

    'Place "Current" above the control button to highlight the currect phase of the workbook
    updateCurrent ("A1")
    
End Sub
