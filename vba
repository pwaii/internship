Option Explicit
' Excel Pivot Builder - text-only VBA version
' Paste this whole file into a standard VBA module.
'
' Main macros:
'   SetupPivotBuilderSheet
'   ChooseSourceWorkbook
'   BuildPivotTablesFromSetup
'
' The setup sheet acts as the user interface, so this works on machines where
' Python apps or custom executables are blocked.
Private Const SETUP_SHEET As String = "PivotBuilder_Setup"
Private Const HELPER_PREFIX As String = "__PivotSource"
Private Const FIELD_PICKER_SHEET As String = "PivotBuilder_FieldPicker"
Private Const CONDITION_BUILDER_SHEET As String = "PivotBuilder_Conditions"
Private Const VALUE_LIST_SHEET As String = "PivotBuilder_ValueLists"
Private mNextAutoFieldRefresh As Date
Private mLastAutoFieldKey As String
Private mSetupCellCache As Object
Private mApplyingDropdownAppend As Boolean
Public Sub SetupPivotBuilderSheet()
    Dim ws As Worksheet
    Dim defaultDataSheet As String
    Dim savedSourceWorkbook As Variant
    Dim savedDataSheet As Variant
    Dim savedTemplate As Variant
    Dim savedRows As Variant
    Dim savedIsNewLayout As Boolean
    Dim savedIsOldEnabledLayout As Boolean
    Dim savedHasNextRightColumn As Boolean
    Dim hadExistingSetup As Boolean
    Dim dataRow As Long
    Dim colIndex As Long
    
    defaultDataSheet = FirstVisibleDataSheetName(ThisWorkbook)
    Set ws = GetOrCreateSheet(ThisWorkbook, SETUP_SHEET)
    
    hadExistingSetup = Application.WorksheetFunction.CountA(ws.Cells) > 0
    If hadExistingSetup Then
        If LCase$(Trim$(CStr(ws.Range("A3").Value))) = "source workbook" Then
            savedSourceWorkbook = ws.Range("B3").Value
            savedDataSheet = ws.Range("B4").Value
            savedTemplate = ws.Range("B5").Value
        Else
            savedTemplate = ws.Range("B3").Value
            savedDataSheet = ws.Range("B4").Value
            savedSourceWorkbook = ws.Range("B5").Value
        End If
        savedIsNewLayout = LCase$(Trim$(CStr(ws.Range("A8").Value))) = "template"
        savedIsOldEnabledLayout = LCase$(Trim$(CStr(ws.Range("A8").Value))) = "enabled"
        savedHasNextRightColumn = LCase$(Trim$(CStr(ws.Range("M8").Value))) = "next pivot right"
        If savedIsNewLayout Or savedIsOldEnabledLayout Then
            If savedHasNextRightColumn Then
                savedRows = ws.Range("A9:N200").Value
            Else
                savedRows = ws.Range("A9:M200").Value
            End If
        Else
            savedRows = ws.Range("A9:K200").Value
        End If
    End If
    
    ws.Cells.Clear
    ws.Cells.UnMerge
    ws.Cells.Font.Name = "Aptos"
    ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = RGB(242, 244, 247)
    ws.Range("A1").Value = "Excel Pivot Builder"
    ws.Range("A1:N1").Merge
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").Interior.Color = RGB(141, 2, 31)
    ws.Rows("1:1").RowHeight = 28
    ws.Rows("3:6").RowHeight = 24
    ws.Range("A3").Value = "Source workbook"
    ws.Range("B3").Value = IIf(hadExistingSetup, savedSourceWorkbook, "")
    ws.Range("B3:D3").Merge
    ws.Range("B3:D3").HorizontalAlignment = xlLeft
    ws.Range("A4").Value = "Data sheet"
    ws.Range("B4").Value = IIf(hadExistingSetup And Trim$(CStr(savedDataSheet)) <> "", savedDataSheet, defaultDataSheet)
    
    ws.Range("A5").Value = "Selected template"
    ws.Range("B5").Value = IIf(hadExistingSetup And Trim$(CStr(savedTemplate)) <> "", savedTemplate, "All")
    ws.Range("A6").Value = "Fill one row per PivotTable. Blank Pivot Name rows are skipped."
    ws.Range("A6:D6").Merge
    ws.Range("A6").Font.Italic = True
    ws.Range("A3:A5").Font.Bold = True
    ws.Range("A3:D5").Interior.Color = RGB(255, 255, 255)
    ws.Range("A3:D5").Borders.Color = RGB(255, 255, 255)
    ws.Range("A8:N8").Value = Array( _
        "Template", _
