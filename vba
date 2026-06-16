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
        If savedIsNewLayout Or savedIsOldEnabledLayout Then
            savedRows = ws.Range("A9:M200").Value
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
    ws.Range("A1:M1").Merge
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
    ws.Range("B5").Value = IIf(hadExistingSetup And Trim$(CStr(savedTemplate)) <> "", savedTemplate, "Default")

    ws.Range("A6").Value = "Fill one row per PivotTable. Blank Pivot Name rows are skipped."
    ws.Range("A6:D6").Merge
    ws.Range("A6").Font.Italic = True
    ws.Range("A3:A5").Font.Bold = True
    ws.Range("A3:D5").Interior.Color = RGB(255, 255, 255)
    ws.Range("A3:D5").Borders.Color = RGB(255, 255, 255)

    ws.Range("A8:M8").Value = Array( _
        "Template", _
        "Pivot Name", _
        "Output Sheet", _
        "Save Behavior", _
        "Rows", _
        "Row Group Name", _
        "Group Rules", _
        "Display Values", _
        "Values To Count/Sum", _
        "Filters", _
        "Conditions", _
        "CSV File Name", _
        "Notes" _
    )
    ws.Range("A8:M8").Font.Bold = True
    ws.Range("A8:M8").Font.Color = RGB(255, 255, 255)
    ws.Range("A8:M8").Interior.Color = RGB(0, 0, 0)
    ws.Rows("8:8").RowHeight = 22

    If hadExistingSetup Then
        If savedIsNewLayout Then
            ws.Range("A9:M200").Value = savedRows
        ElseIf savedIsOldEnabledLayout Then
            For dataRow = 1 To UBound(savedRows, 1)
                If UCase$(Trim$(CStr(savedRows(dataRow, 1)))) = "YES" Or Trim$(CStr(savedRows(dataRow, 3))) <> "" Then
                    ws.Cells(8 + dataRow, "A").Value = savedRows(dataRow, 2)
                    ws.Cells(8 + dataRow, "B").Value = savedRows(dataRow, 3)
                    ws.Cells(8 + dataRow, "C").Value = savedRows(dataRow, 4)
                    ws.Cells(8 + dataRow, "D").Value = savedRows(dataRow, 5)
                    ws.Cells(8 + dataRow, "E").Value = savedRows(dataRow, 6)
                    ws.Cells(8 + dataRow, "F").Value = savedRows(dataRow, 7)
                    ws.Cells(8 + dataRow, "G").Value = ""
                    ws.Cells(8 + dataRow, "H").Value = savedRows(dataRow, 8)
                    ws.Cells(8 + dataRow, "I").Value = savedRows(dataRow, 9)
                    ws.Cells(8 + dataRow, "J").Value = savedRows(dataRow, 10)
                    ws.Cells(8 + dataRow, "K").Value = savedRows(dataRow, 11)
                    ws.Cells(8 + dataRow, "L").Value = savedRows(dataRow, 12)
                    ws.Cells(8 + dataRow, "M").Value = savedRows(dataRow, 13)
                End If
            Next dataRow
        Else
            For dataRow = 1 To UBound(savedRows, 1)
                For colIndex = 1 To UBound(savedRows, 2)
                    ws.Cells(8 + dataRow, colIndex).Value = savedRows(dataRow, colIndex)
                Next colIndex
            Next dataRow
        End If
    Else
        ws.Range("A9:M9").Value = Array( _
            "Default", _
            "Pivot_1", _
            "Pivot_Output", _
            "Save to input file", _
            "01", _
            "", _
            "", _
            "", _
            "05,06", _
            "", _
            "Status=Open; Region=East", _
            "", _
            "Conditions are optional. Use Field=Value; Field2=Value2" _
        )
    End If

    ws.Range("A9:M200").Interior.Color = RGB(255, 255, 255)
    ws.Range("A8:M200").Borders.Color = RGB(200, 205, 212)
    ws.Range("N1").Value = "Field Suggestions"
    ws.Range("N1:Q1").Merge
    ws.Range("N1").Font.Bold = True
    ws.Range("N1").Font.Color = RGB(255, 255, 255)
    ws.Range("N1").Interior.Color = RGB(141, 2, 31)
    ws.Range("N3").Value = "Available Fields"
    ws.Range("N3:O3").Merge
    ws.Range("N3").Font.Bold = True
    ws.Range("N3").Font.Color = RGB(255, 255, 255)
    ws.Range("N3").Interior.Color = RGB(0, 0, 0)
    ws.Range("N4:N500").Interior.Color = RGB(255, 255, 255)

    ws.Columns("A").ColumnWidth = 18
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 22
    ws.Columns("G").ColumnWidth = 42
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 16
    ws.Columns("J").ColumnWidth = 16
    ws.Columns("K").ColumnWidth = 26
    ws.Columns("L").ColumnWidth = 18
    ws.Columns("M").ColumnWidth = 38
    ws.Columns("N:Q").ColumnWidth = 22
    ws.Range("A1:M200").VerticalAlignment = xlTop
    ws.Range("A8:M200").WrapText = False
    ws.Range("E9:G200").WrapText = True
    ws.Range("G9:G200").WrapText = True
    ws.Range("K9:K200").WrapText = True
    ws.Rows("9:200").RowHeight = 54
    AddDataSheetDropdown ws, "B4"
    AddTemplateDropdown ws
    AddFieldValidations ws
    AddSetupButtons ws
    AddSetupComments ws
    ApplyBrandLayout ws
    If Trim$(CStr(ws.Range("B3").Value)) <> "" Then
        TryRefreshFieldSuggestions ws, False
    End If
    AddSaveBehaviorValidation ws
    ApplySaveBehaviorDisplay ws

    ws.Activate
    StartAutoFieldRefresh ws
    ws.Range("B3").Select

    MsgBox "Setup sheet refreshed. Existing template rows were preserved.", vbInformation
End Sub

Public Sub AddPivotBuilderButtons()
    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then
        MsgBox "Setup sheet not found. Run SetupPivotBuilderSheet first.", vbExclamation
        Exit Sub
    End If
    AddSetupButtons ThisWorkbook.Worksheets(SETUP_SHEET)
    MsgBox "Buttons added to " & SETUP_SHEET & ".", vbInformation
End Sub

Public Sub SaveCurrentTemplate()
    Dim setupWs As Worksheet
    Dim savePath As Variant

    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then
        MsgBox "Setup sheet not found. Run SetupPivotBuilderSheet first.", vbExclamation
        Exit Sub
    End If

    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    ApplyDropdownAppend setupWs
    ApplySaveBehaviorDisplay setupWs
    AddTemplateDropdown setupWs
    CaptureSetupCellCache setupWs

    If ThisWorkbook.Path = "" Then
        savePath = Application.GetSaveAsFilename( _
            InitialFileName:="PivotBuilder_Template.xlsm", _
            FileFilter:="Excel Macro-Enabled Workbook (*.xlsm),*.xlsm", _
            Title:="Save Pivot Builder template workbook" _
        )
        If savePath = False Then Exit Sub
        ThisWorkbook.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Else
        ThisWorkbook.Save
    End If

    MsgBox "Template saved in this tool workbook.", vbInformation
End Sub

Public Sub NewTemplateFromCurrent()
    Dim setupWs As Worksheet
    Dim currentTemplate As String
    Dim newTemplate As String
    Dim lastSetupRow As Long
    Dim rowIndex As Long
    Dim nextRow As Long
    Dim copiedCount As Long

    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then
        MsgBox "Setup sheet not found. Run SetupPivotBuilderSheet first.", vbExclamation
        Exit Sub
    End If

    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    currentTemplate = Trim$(CStr(setupWs.Range("B5").Value))
    If currentTemplate = "" Then
        MsgBox "Choose a selected template in B5 first.", vbExclamation
        Exit Sub
    End If

    newTemplate = Trim$(InputBox("Type the new template name.", "New Template", currentTemplate & "_Copy"))
    If newTemplate = "" Then Exit Sub
    If StrComp(newTemplate, currentTemplate, vbTextCompare) = 0 Then
        MsgBox "Use a different name for the new template.", vbExclamation
        Exit Sub
    End If
    If TemplateExists(setupWs, newTemplate) Then
        MsgBox "That template name already exists. Choose another name.", vbExclamation
        Exit Sub
    End If

    ApplyDropdownAppend setupWs
    lastSetupRow = LastTemplateRow(setupWs)
    nextRow = Application.Max(9, lastSetupRow + 1)

    For rowIndex = 9 To lastSetupRow
        If Trim$(CStr(setupWs.Cells(rowIndex, "A").Value)) = currentTemplate Then
            setupWs.Range("A" & rowIndex & ":M" & rowIndex).Copy Destination:=setupWs.Range("A" & nextRow)
            setupWs.Cells(nextRow, "A").Value = newTemplate
            nextRow = nextRow + 1
            copiedCount = copiedCount + 1
        End If
    Next rowIndex

    If copiedCount = 0 Then
        setupWs.Range("A" & nextRow & ":M" & nextRow).Value = Array( _
            newTemplate, _
            "Pivot_1", _
            "Pivot_Output", _
            "Save to input file", _
            "", _
            "", _
            "", _
            "", _
            "", _
            "", _
            "", _
            "", _
            "" _
        )
        copiedCount = 1
    End If

    setupWs.Range("B5").Value = newTemplate
    AddTemplateDropdown setupWs
    AddFieldValidations setupWs
    AddSaveBehaviorValidation setupWs
    ApplySaveBehaviorDisplay setupWs
    CaptureSetupCellCache setupWs
    setupWs.Activate
    setupWs.Range("B5").Select

    MsgBox "New template created: " & newTemplate & vbCrLf & "Edit its rows, then click Save Template.", vbInformation
End Sub

Public Sub ChooseSourceWorkbook()
    Dim setupWs As Worksheet
    Dim sourcePath As Variant
    Dim sourceWb As Workbook
    Dim firstSheet As String
    Dim headers As Variant
    Dim sheetNames As String

    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then
        MsgBox "Setup sheet not found. Run SetupPivotBuilderSheet first.", vbExclamation
        Exit Sub
    End If

    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    sourcePath = Application.GetOpenFilename( _
        FileFilter:="Excel or CSV files (*.xlsx;*.xlsm;*.xls;*.csv),*.xlsx;*.xlsm;*.xls;*.csv", _
        Title:="Choose source workbook or CSV file" _
    )

    If sourcePath = False Then Exit Sub

    Application.ScreenUpdating = False
    Set sourceWb = Workbooks.Open(CStr(sourcePath), ReadOnly:=True)
    firstSheet = FirstVisibleDataSheetName(sourceWb)
    sheetNames = VisibleSheetList(sourceWb)
    If firstSheet <> "" Then headers = NormalizedHeaders(sourceWb.Worksheets(firstSheet).UsedRange)
    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    If firstSheet = "" Then
        MsgBox "No visible sheet found in the selected workbook.", vbExclamation
        Exit Sub
    End If

    setupWs.Range("B3").Value = CStr(sourcePath)
    setupWs.Range("B4").Value = firstSheet
    AddListValidation setupWs.Range("B4"), sheetNames
    If IsArray(headers) Then PopulateFieldSuggestions setupWs, headers
    mLastAutoFieldKey = FieldRefreshKey(setupWs)
    ScheduleAutoFieldRefresh
    MsgBox "Source workbook selected:" & vbCrLf & CStr(sourcePath) & vbCrLf & vbCrLf & "Source sheet: " & firstSheet, vbInformation
End Sub

Public Sub RefreshFieldSuggestions()
    Dim setupWs As Worksheet

    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then
        MsgBox "Setup sheet not found. Run SetupPivotBuilderSheet first.", vbExclamation
        Exit Sub
    End If

    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    TryRefreshFieldSuggestions setupWs, True
End Sub

Private Function TryRefreshFieldSuggestions(ByVal setupWs As Worksheet, ByVal showMessage As Boolean) As Boolean
    Dim sourcePath As String
    Dim sourceSheet As String
    Dim sourceWb As Workbook
    Dim headers As Variant
    Dim errText As String

    On Error GoTo RefreshError

    sourcePath = NormalizeInputPath(Trim$(CStr(setupWs.Range("B3").Value)))
    sourceSheet = Trim$(CStr(setupWs.Range("B4").Value))

    If sourcePath = "" Or Dir(sourcePath) = "" Then
        If showMessage Then MsgBox "Choose a valid source workbook first.", vbExclamation
        Exit Function
    End If

    Application.ScreenUpdating = False
    Set sourceWb = Workbooks.Open(sourcePath, ReadOnly:=True)
    If sourceSheet = "" Then sourceSheet = FirstVisibleDataSheetName(sourceWb)
    If Not WorksheetExists(sourceWb, sourceSheet) Then
        sourceWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        If showMessage Then MsgBox "Source sheet not found: " & sourceSheet, vbExclamation
        Exit Function
    End If
    headers = NormalizedHeaders(sourceWb.Worksheets(sourceSheet).UsedRange)
    sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    PopulateFieldSuggestions setupWs, headers
    TryRefreshFieldSuggestions = True
    If showMessage Then MsgBox "Field suggestions refreshed from " & sourceSheet & ".", vbInformation
    Exit Function

RefreshError:
    errText = Err.Description
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    On Error GoTo 0
    If showMessage Then MsgBox "Field suggestions could not be refreshed: " & errText, vbExclamation
End Function

Private Sub StartAutoFieldRefresh(ByVal setupWs As Worksheet)
    mLastAutoFieldKey = FieldRefreshKey(setupWs)
    CaptureSetupCellCache setupWs
    ScheduleAutoFieldRefresh
End Sub

Private Sub ScheduleAutoFieldRefresh()
    On Error Resume Next
    If mNextAutoFieldRefresh <> 0 Then
        Application.OnTime EarliestTime:=mNextAutoFieldRefresh, Procedure:="AutoRefreshFieldSuggestionsTick", Schedule:=False
    End If
    On Error GoTo 0

    mNextAutoFieldRefresh = Now + TimeSerial(0, 0, 2)
    On Error Resume Next
    Application.OnTime EarliestTime:=mNextAutoFieldRefresh, Procedure:="AutoRefreshFieldSuggestionsTick", Schedule:=True
    On Error GoTo 0
End Sub

Public Sub AutoRefreshFieldSuggestionsTick()
    Dim setupWs As Worksheet
    Dim currentKey As String

    On Error GoTo Finish

    If WorksheetExists(ThisWorkbook, SETUP_SHEET) Then
        Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
        ApplyDropdownAppend setupWs
        ApplySaveBehaviorDisplay setupWs
        currentKey = FieldRefreshKey(setupWs)
        If currentKey <> mLastAutoFieldKey Then
            If TryRefreshFieldSuggestions(setupWs, False) Then
                mLastAutoFieldKey = currentKey
                CaptureSetupCellCache setupWs
            Else
                mLastAutoFieldKey = currentKey
            End If
        End If
    End If

Finish:
    ScheduleAutoFieldRefresh
End Sub

Private Function FieldRefreshKey(ByVal setupWs As Worksheet) As String
    FieldRefreshKey = Trim$(CStr(setupWs.Range("B3").Value)) & "|" & Trim$(CStr(setupWs.Range("B4").Value))
End Function

Private Sub CaptureSetupCellCache(ByVal setupWs As Worksheet)
    Dim cell As Range
    Set mSetupCellCache = CreateObject("Scripting.Dictionary")
    For Each cell In Application.Union(setupWs.Range("E9:E200"), setupWs.Range("G9:K200"))
        mSetupCellCache(cell.Address(False, False)) = CStr(cell.Value)
    Next cell
End Sub

Private Sub ApplyDropdownAppend(ByVal setupWs As Worksheet)
    Dim cell As Range
    Dim key As String
    Dim oldText As String
    Dim newText As String
    Dim finalText As String

    If mApplyingDropdownAppend Then Exit Sub
    If mSetupCellCache Is Nothing Then
        CaptureSetupCellCache setupWs
        Exit Sub
    End If

    mApplyingDropdownAppend = True
    For Each cell In Application.Union(setupWs.Range("E9:E200"), setupWs.Range("G9:K200"))
        key = cell.Address(False, False)
        oldText = ""
        If mSetupCellCache.Exists(key) Then oldText = CStr(mSetupCellCache(key))
        newText = CStr(cell.Value)

        If newText <> oldText Then
            finalText = AppendedDropdownText(oldText, newText, cell.Column)
            If finalText <> newText Then cell.Value = finalText
            mSetupCellCache(key) = CStr(cell.Value)
        End If
    Next cell
    mApplyingDropdownAppend = False
End Sub

Private Function AppendedDropdownText(ByVal oldText As String, ByVal newText As String, ByVal columnNumber As Long) As String
    Dim cleanOld As String
    Dim cleanNew As String
    Dim separator As String

    cleanOld = Trim$(oldText)
    cleanNew = Trim$(newText)
    AppendedDropdownText = newText

    If cleanNew = "" Then Exit Function
    If InStr(1, cleanNew, ",", vbTextCompare) > 0 Then Exit Function
    If InStr(1, cleanNew, ";", vbTextCompare) > 0 Then Exit Function
    If InStr(1, cleanNew, "=", vbTextCompare) > 0 Then Exit Function
    If Not IsSuggestedField(cleanNew) Then Exit Function

    If columnNumber = 7 Then
        If cleanOld = "" Then
            AppendedDropdownText = "Group 1:" & cleanNew & "="
        ElseIf ConditionHasField(cleanOld, cleanNew) Then
            AppendedDropdownText = cleanOld
        Else
            AppendedDropdownText = cleanOld & "; " & cleanNew & "="
        End If
    ElseIf columnNumber = 11 Then
        If cleanOld = "" Then
            AppendedDropdownText = cleanNew & "="
        ElseIf ConditionHasField(cleanOld, cleanNew) Then
            AppendedDropdownText = cleanOld
        Else
            AppendedDropdownText = cleanOld & "; " & cleanNew & "="
        End If
    Else
        If cleanOld = "" Then Exit Function
        If CommaListHasItem(cleanOld, cleanNew) Then
            AppendedDropdownText = cleanOld
        Else
            AppendedDropdownText = cleanOld & "," & cleanNew
        End If
    End If
End Function

Private Function IsSuggestedField(ByVal fieldName As String) As Boolean
    Dim setupWs As Worksheet
    Dim found As Range

    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then Exit Function
    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    Set found = setupWs.Range("N4:N500").Find(What:=fieldName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    IsSuggestedField = Not found Is Nothing
End Function

Private Function CommaListHasItem(ByVal listText As String, ByVal itemText As String) As Boolean
    Dim items As Variant
    Dim item As Variant

    items = SplitList(listText)
    For Each item In items
        If StrComp(Trim$(CStr(item)), Trim$(itemText), vbTextCompare) = 0 Then
            CommaListHasItem = True
            Exit Function
        End If
    Next item
End Function

Private Function ConditionHasField(ByVal conditionText As String, ByVal fieldName As String) As Boolean
    Dim conditions As Variant
    Dim condition As Variant
    Dim parts As Variant

    conditions = Split(conditionText, ";")
    For Each condition In conditions
        parts = Split(CStr(condition), "=", 2)
        If StrComp(Trim$(CStr(parts(0))), Trim$(fieldName), vbTextCompare) = 0 Then
            ConditionHasField = True
            Exit Function
        End If
    Next condition
End Function

Private Sub ApplySaveBehaviorDisplay(ByVal setupWs As Worksheet)
    Dim rowIndex As Long
    Dim behavior As String

    For rowIndex = 9 To 200
        behavior = NormalizedSaveBehavior(CStr(setupWs.Cells(rowIndex, "D").Value))
        If behavior = "EXPORT CSV" Then
            setupWs.Cells(rowIndex, "L").Interior.Color = RGB(255, 255, 255)
            setupWs.Cells(rowIndex, "L").Font.Color = RGB(0, 0, 0)
        Else
            setupWs.Cells(rowIndex, "L").ClearContents
            setupWs.Cells(rowIndex, "L").Interior.Color = RGB(229, 232, 237)
            setupWs.Cells(rowIndex, "L").Font.Color = RGB(229, 232, 237)
        End If
    Next rowIndex
End Sub

Private Sub AddSaveBehaviorValidation(ByVal setupWs As Worksheet)
    With setupWs.Range("D9:D200")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:="Save to input file,Export CSV"
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
    End With
End Sub

Private Function NormalizedSaveBehavior(ByVal textValue As String) As String
    Dim cleanValue As String
    cleanValue = UCase$(Trim$(textValue))
    If cleanValue = "" Then
        NormalizedSaveBehavior = "SAVE TO INPUT FILE"
    ElseIf cleanValue = "EXPORT CSV" Or cleanValue = "CSV" Then
        NormalizedSaveBehavior = "EXPORT CSV"
    Else
        NormalizedSaveBehavior = "SAVE TO INPUT FILE"
    End If
End Function

Private Function WorkbookFolder(ByVal wb As Workbook, ByVal fallbackPath As String) As String
    If Trim$(wb.Path) <> "" Then
        WorkbookFolder = wb.Path
    ElseIf InStrRev(fallbackPath, Application.PathSeparator) > 0 Then
        WorkbookFolder = Left$(fallbackPath, InStrRev(fallbackPath, Application.PathSeparator) - 1)
    End If
End Function

Private Function NormalizeInputPath(ByVal filePath As String) As String
    Dim result As String

    result = Trim$(filePath)
    If Left$(result, 1) = """" And Right$(result, 1) = """" Then
        result = Mid$(result, 2, Len(result) - 2)
    End If

    If Left$(result, 1) = "\" And Left$(result, 2) <> "\\" Then
        NormalizeInputPath = Environ$("SystemDrive") & result
    Else
        NormalizeInputPath = result
    End If
End Function

Private Function IsCsvPath(ByVal filePath As String) As Boolean
    IsCsvPath = LCase$(Right$(Trim$(filePath), 4)) = ".csv"
End Function

Private Function CsvConvertedWorkbookPath(ByVal csvPath As String) As String
    Dim folderPath As String
    Dim fileName As String
    Dim dotPos As Long
    Dim baseName As String

    folderPath = Left$(csvPath, InStrRev(csvPath, Application.PathSeparator))
    fileName = Mid$(csvPath, InStrRev(csvPath, Application.PathSeparator) + 1)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 1 Then
        baseName = Left$(fileName, dotPos - 1)
    Else
        baseName = fileName
    End If
    CsvConvertedWorkbookPath = UniqueFilePath(folderPath & baseName & "_PivotOutput.xlsx")
End Function

Private Function UniqueFilePath(ByVal desiredPath As String) As String
    Dim dotPos As Long
    Dim basePath As String
    Dim extensionText As String
    Dim index As Long
    Dim candidate As String

    If Dir(desiredPath) = "" Then
        UniqueFilePath = desiredPath
        Exit Function
    End If

    dotPos = InStrRev(desiredPath, ".")
    If dotPos > 0 Then
        basePath = Left$(desiredPath, dotPos - 1)
        extensionText = Mid$(desiredPath, dotPos)
    Else
        basePath = desiredPath
        extensionText = ""
    End If

    For index = 2 To 999
        candidate = basePath & "_" & CStr(index) & extensionText
        If Dir(candidate) = "" Then
            UniqueFilePath = candidate
            Exit Function
        End If
    Next index

    Err.Raise vbObjectError + 350, , "Could not create a unique output file path."
End Function

Private Sub OpenSelectedEditor()
    Dim setupWs As Worksheet
    Dim target As Range

    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then
        MsgBox "Setup sheet not found. Run SetupPivotBuilderSheet first.", vbExclamation
        Exit Sub
    End If

    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    If ActiveSheet.Name <> SETUP_SHEET Then
        MsgBox "Go to " & SETUP_SHEET & " and select a setup cell first.", vbExclamation
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Cells.CountLarge <> 1 Then
        MsgBox "Select one cell first.", vbExclamation
        Exit Sub
    End If

    Set target = Selection.Cells(1, 1)
    If target.Row < 9 Then
        MsgBox "Select a row in the main setup table first.", vbExclamation
        Exit Sub
    End If

    Select Case target.Column
        Case 5, 8, 9, 10
            OpenFieldChecklist target
        Case 11
            OpenConditionBuilder target
        Case Else
            MsgBox "Use the dropdown directly in Rows, Group Rules, Display Values, Values To Count/Sum, Filters, or Conditions.", vbInformation
    End Select
End Sub

Private Sub OpenFieldChecklist(ByVal target As Range)
    Dim pickerWs As Worksheet
    Dim setupWs As Worksheet
    Dim fields As Variant
    Dim currentItems As Variant
    Dim index As Long
    Dim rowIndex As Long
    Dim cb As CheckBox
    Dim shape As Shape

    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    If Application.WorksheetFunction.CountA(setupWs.Range("N4:N500")) = 0 Then
        TryRefreshFieldSuggestions setupWs, False
    End If

    fields = FieldSuggestionArray(setupWs)
    If IsEmpty(fields) Then
        MsgBox "No field suggestions found. Choose a source workbook first.", vbExclamation
        Exit Sub
    End If

    currentItems = SplitList(CStr(target.Value))
    Set pickerWs = GetOrCreateSheet(ThisWorkbook, FIELD_PICKER_SHEET)
    pickerWs.Visible = xlSheetVisible
    pickerWs.Cells.Clear
    DeleteShapesByPrefix pickerWs, "PivotBuilder"

    pickerWs.Range("A1").Value = "Pick Fields"
    pickerWs.Range("A1:D1").Merge
    pickerWs.Range("A1").Font.Bold = True
    pickerWs.Range("A1").Font.Size = 16
    pickerWs.Range("A1").Font.Color = RGB(255, 255, 255)
    pickerWs.Range("A1").Interior.Color = RGB(141, 2, 31)
    pickerWs.Range("A2").Value = "Target"
    pickerWs.Range("B2").Value = target.Address(False, False)
    pickerWs.Range("C2").Value = target.Worksheet.Name
    pickerWs.Range("A3").Value = "Use"
    pickerWs.Range("B3").Value = "Field"
    pickerWs.Range("A3:B3").Font.Bold = True
    pickerWs.Range("A3:B3").Font.Color = RGB(255, 255, 255)
    pickerWs.Range("A3:B3").Interior.Color = RGB(0, 0, 0)

    For index = LBound(fields) To UBound(fields)
        rowIndex = 4 + index
        pickerWs.Cells(rowIndex, "B").Value = CStr(fields(index))
        pickerWs.Cells(rowIndex, "C").Value = IsInList(currentItems, CStr(fields(index)))
        Set cb = pickerWs.CheckBoxes.Add(pickerWs.Cells(rowIndex, "A").Left + 4, pickerWs.Cells(rowIndex, "A").Top + 2, 14, 14)
        cb.Name = "PivotBuilderFieldCheck" & CStr(rowIndex)
        cb.Caption = ""
        cb.LinkedCell = pickerWs.Cells(rowIndex, "C").Address
        cb.Value = IIf(CBool(pickerWs.Cells(rowIndex, "C").Value), xlOn, xlOff)
    Next index

    pickerWs.Columns("A").ColumnWidth = 8
    pickerWs.Columns("B").ColumnWidth = 32
    pickerWs.Columns("C").Hidden = True
    pickerWs.Range("A1:D500").Interior.Color = RGB(242, 244, 247)
    pickerWs.Range("A4:B" & CStr(3 + UBound(fields) + 1)).Interior.Color = RGB(255, 255, 255)
    pickerWs.Range("A3:B" & CStr(3 + UBound(fields) + 1)).Borders.Color = RGB(200, 205, 212)

    Set shape = pickerWs.Shapes.AddShape(msoShapeRoundedRectangle, pickerWs.Range("D3").Left, pickerWs.Range("D3").Top, 90, 24)
    With shape
        .Name = "PivotBuilderApplyFieldPicker"
        .TextFrame.Characters.Text = "Apply"
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(141, 2, 31)
        .Line.ForeColor.RGB = RGB(141, 2, 31)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "ApplyFieldChecklist"
    End With

    Set shape = pickerWs.Shapes.AddShape(msoShapeRoundedRectangle, pickerWs.Range("D5").Left, pickerWs.Range("D5").Top, 90, 24)
    With shape
        .Name = "PivotBuilderBackFieldPicker"
        .TextFrame.Characters.Text = "Back"
        .Fill.ForeColor.RGB = RGB(242, 242, 242)
        .Line.ForeColor.RGB = RGB(166, 166, 166)
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .OnAction = "BackToSetup"
    End With

    pickerWs.Activate
    pickerWs.Range("A4").Select
End Sub

Private Sub ApplyFieldChecklist()
    Dim pickerWs As Worksheet
    Dim setupWs As Worksheet
    Dim targetAddress As String
    Dim rowIndex As Long
    Dim result As String
    Dim separator As String

    If Not WorksheetExists(ThisWorkbook, FIELD_PICKER_SHEET) Then Exit Sub
    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then Exit Sub

    Set pickerWs = ThisWorkbook.Worksheets(FIELD_PICKER_SHEET)
    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    targetAddress = Trim$(CStr(pickerWs.Range("B2").Value))
    If targetAddress = "" Then Exit Sub

    For rowIndex = 4 To pickerWs.Cells(pickerWs.Rows.Count, "B").End(xlUp).Row
        If CBool(pickerWs.Cells(rowIndex, "C").Value) Then
            result = result & separator & CStr(pickerWs.Cells(rowIndex, "B").Value)
            separator = ","
        End If
    Next rowIndex

    setupWs.Range(targetAddress).Value = result
    setupWs.Activate
    setupWs.Range(targetAddress).Select
    pickerWs.Visible = xlSheetHidden
End Sub

Private Sub OpenConditionBuilder(ByVal target As Range)
    Dim conditionWs As Worksheet
    Dim setupWs As Worksheet
    Dim fields As Variant
    Dim conditions As Variant
    Dim index As Long
    Dim rowIndex As Long
    Dim parts As Variant
    Dim shape As Shape

    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    If Application.WorksheetFunction.CountA(setupWs.Range("N4:N500")) = 0 Then
        TryRefreshFieldSuggestions setupWs, False
    End If

    fields = FieldSuggestionArray(setupWs)
    If IsEmpty(fields) Then
        MsgBox "No field suggestions found. Choose a source workbook first.", vbExclamation
        Exit Sub
    End If

    Set conditionWs = GetOrCreateSheet(ThisWorkbook, CONDITION_BUILDER_SHEET)
    conditionWs.Visible = xlSheetVisible
    conditionWs.Cells.Clear
    DeleteShapesByPrefix conditionWs, "PivotBuilder"

    conditionWs.Range("A1").Value = "Build Conditions"
    conditionWs.Range("A1:E1").Merge
    conditionWs.Range("A1").Font.Bold = True
    conditionWs.Range("A1").Font.Size = 16
    conditionWs.Range("A1").Font.Color = RGB(255, 255, 255)
    conditionWs.Range("A1").Interior.Color = RGB(141, 2, 31)
    conditionWs.Range("A2").Value = "Target"
    conditionWs.Range("B2").Value = target.Address(False, False)
    conditionWs.Range("A3:D3").Value = Array("Use", "Field", "Is", "Value")
    conditionWs.Range("A3:D3").Font.Bold = True
    conditionWs.Range("A3:D3").Font.Color = RGB(255, 255, 255)
    conditionWs.Range("A3:D3").Interior.Color = RGB(0, 0, 0)

    For index = LBound(fields) To UBound(fields)
        conditionWs.Cells(4 + index, "H").Value = CStr(fields(index))
    Next index

    conditions = Split(CStr(target.Value), ";")
    For index = LBound(conditions) To UBound(conditions)
        If Trim$(CStr(conditions(index))) <> "" Then
            rowIndex = 4 + index
            parts = Split(CStr(conditions(index)), "=", 2)
            conditionWs.Cells(rowIndex, "A").Value = "Yes"
            conditionWs.Cells(rowIndex, "B").Value = Trim$(CStr(parts(0)))
            conditionWs.Cells(rowIndex, "C").Value = "="
            If UBound(parts) >= 1 Then conditionWs.Cells(rowIndex, "D").Value = Trim$(CStr(parts(1)))
        End If
    Next index

    For rowIndex = 4 To 20
        If conditionWs.Cells(rowIndex, "A").Value = "" Then conditionWs.Cells(rowIndex, "A").Value = "No"
        conditionWs.Cells(rowIndex, "C").Value = "="
    Next rowIndex

    AddSuggestionValidation conditionWs.Range("B4:B20"), "=$H$4:$H$500"
    AddSuggestionValidation conditionWs.Range("A4:A20"), "Yes,No"
    conditionWs.Columns("A").ColumnWidth = 8
    conditionWs.Columns("B").ColumnWidth = 28
    conditionWs.Columns("C").ColumnWidth = 8
    conditionWs.Columns("D").ColumnWidth = 28
    conditionWs.Columns("H").Hidden = True
    conditionWs.Range("A1:E500").Interior.Color = RGB(242, 244, 247)
    conditionWs.Range("A4:D20").Interior.Color = RGB(255, 255, 255)
    conditionWs.Range("A3:D20").Borders.Color = RGB(200, 205, 212)

    Set shape = conditionWs.Shapes.AddShape(msoShapeRoundedRectangle, conditionWs.Range("F3").Left, conditionWs.Range("F3").Top, 90, 24)
    With shape
        .Name = "PivotBuilderApplyConditions"
        .TextFrame.Characters.Text = "Apply"
        .TextFrame.Characters.Font.Bold = True
        .Fill.ForeColor.RGB = RGB(141, 2, 31)
        .Line.ForeColor.RGB = RGB(141, 2, 31)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "ApplyConditionBuilder"
    End With

    Set shape = conditionWs.Shapes.AddShape(msoShapeRoundedRectangle, conditionWs.Range("F5").Left, conditionWs.Range("F5").Top, 90, 24)
    With shape
        .Name = "PivotBuilderBackConditions"
        .TextFrame.Characters.Text = "Back"
        .Fill.ForeColor.RGB = RGB(242, 242, 242)
        .Line.ForeColor.RGB = RGB(166, 166, 166)
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .OnAction = "BackToSetup"
    End With

    conditionWs.Activate
    conditionWs.Range("B4").Select
End Sub

Private Sub ApplyConditionBuilder()
    Dim conditionWs As Worksheet
    Dim setupWs As Worksheet
    Dim targetAddress As String
    Dim rowIndex As Long
    Dim result As String
    Dim separator As String
    Dim fieldName As String
    Dim fieldValue As String

    If Not WorksheetExists(ThisWorkbook, CONDITION_BUILDER_SHEET) Then Exit Sub
    If Not WorksheetExists(ThisWorkbook, SETUP_SHEET) Then Exit Sub

    Set conditionWs = ThisWorkbook.Worksheets(CONDITION_BUILDER_SHEET)
    Set setupWs = ThisWorkbook.Worksheets(SETUP_SHEET)
    targetAddress = Trim$(CStr(conditionWs.Range("B2").Value))
    If targetAddress = "" Then Exit Sub

    For rowIndex = 4 To 20
        If UCase$(Trim$(CStr(conditionWs.Cells(rowIndex, "A").Value))) = "YES" Then
            fieldName = Trim$(CStr(conditionWs.Cells(rowIndex, "B").Value))
            fieldValue = Trim$(CStr(conditionWs.Cells(rowIndex, "D").Value))
            If fieldName <> "" And fieldValue <> "" Then
                result = result & separator & fieldName & "=" & fieldValue
                separator = "; "
            End If
        End If
    Next rowIndex

    setupWs.Range(targetAddress).Value = result
    setupWs.Activate
    setupWs.Range(targetAddress).Select
    conditionWs.Visible = xlSheetHidden
End Sub

Private Sub BackToSetup()
    Dim activeName As String
    activeName = ActiveSheet.Name

    If WorksheetExists(ThisWorkbook, SETUP_SHEET) Then
        ThisWorkbook.Worksheets(SETUP_SHEET).Activate
    End If
    If activeName = FIELD_PICKER_SHEET And WorksheetExists(ThisWorkbook, FIELD_PICKER_SHEET) Then
        ThisWorkbook.Worksheets(FIELD_PICKER_SHEET).Visible = xlSheetHidden
    End If
    If activeName = CONDITION_BUILDER_SHEET And WorksheetExists(ThisWorkbook, CONDITION_BUILDER_SHEET) Then
        ThisWorkbook.Worksheets(CONDITION_BUILDER_SHEET).Visible = xlSheetHidden
    End If
End Sub

Public Sub BuildPivotTablesFromSetup()
    Dim controllerWb As Workbook
    Dim targetWb As Workbook
    Dim setupWs As Worksheet
    Dim dataWs As Worksheet
    Dim sourceWs As Worksheet
    Dim outputWs As Worksheet
    Dim selectedTemplate As String
    Dim dataSheetName As String
    Dim sourceWorkbookPath As String
    Dim outputSheetName As String
    Dim lastSetupRow As Long
    Dim rowIndex As Long
    Dim outputRow As Long
    Dim headers As Variant
    Dim sourceRange As Range
    Dim cache As PivotCache
    Dim pivot As PivotTable
    Dim builtCount As Long
    Dim csvCount As Long
    Dim csvFileName As String
    Dim saveBehavior As String
    Dim csvFolder As String
    Dim convertedFromCsv As Boolean
    Dim convertedWorkbookPath As String

    On Error GoTo BuildError

    Set controllerWb = ThisWorkbook

    If Not WorksheetExists(controllerWb, SETUP_SHEET) Then
        MsgBox "Setup sheet not found. Run SetupPivotBuilderSheet first.", vbExclamation
        Exit Sub
    End If

    Set setupWs = controllerWb.Worksheets(SETUP_SHEET)
    ApplyDropdownAppend setupWs
    selectedTemplate = Trim(CStr(setupWs.Range("B5").Value))
    dataSheetName = Trim(CStr(setupWs.Range("B4").Value))
    sourceWorkbookPath = NormalizeInputPath(Trim(CStr(setupWs.Range("B3").Value)))
    If sourceWorkbookPath <> Trim(CStr(setupWs.Range("B3").Value)) Then setupWs.Range("B3").Value = sourceWorkbookPath
    ApplySaveBehaviorDisplay setupWs

    If selectedTemplate = "" Then
        MsgBox "Enter a selected template in " & SETUP_SHEET & "!B5.", vbExclamation
        Exit Sub
    End If

    If dataSheetName = "" Then
        MsgBox "Enter the data sheet name in " & SETUP_SHEET & "!B4.", vbExclamation
        Exit Sub
    End If

    If sourceWorkbookPath = "" Then
        MsgBox "Choose a source workbook first. Click the Choose button beside B3.", vbExclamation
        Exit Sub
    End If

    If Dir(sourceWorkbookPath) = "" Then
        MsgBox "Source workbook not found:" & vbCrLf & sourceWorkbookPath, vbExclamation
        Exit Sub
    End If

    convertedFromCsv = IsCsvPath(sourceWorkbookPath)
    Set targetWb = Workbooks.Open(sourceWorkbookPath)
    If convertedFromCsv Then
        convertedWorkbookPath = CsvConvertedWorkbookPath(sourceWorkbookPath)
        Application.DisplayAlerts = False
        targetWb.SaveAs Filename:=convertedWorkbookPath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        sourceWorkbookPath = convertedWorkbookPath
        dataSheetName = targetWb.Worksheets(1).Name
        setupWs.Range("B3").Value = convertedWorkbookPath
        setupWs.Range("B4").Value = dataSheetName
    End If
    csvFolder = WorkbookFolder(targetWb, sourceWorkbookPath)

    If Not WorksheetExists(targetWb, dataSheetName) Then
        MsgBox "Data sheet not found: " & dataSheetName, vbExclamation
        targetWb.Close SaveChanges:=False
        Exit Sub
    End If

    Set dataWs = targetWb.Worksheets(dataSheetName)
    lastSetupRow = setupWs.Cells(setupWs.Rows.Count, "A").End(xlUp).Row
    headers = HeadersWithRowGroups(NormalizedHeaders(dataWs.UsedRange), setupWs, selectedTemplate, lastSetupRow)
    PopulateFieldSuggestions setupWs, headers
    Set sourceWs = CreateNormalizedSourceSheet(targetWb, dataWs, headers, setupWs, selectedTemplate, lastSetupRow)
    Set sourceRange = sourceWs.UsedRange
    Set cache = targetWb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceRange)

    outputSheetName = FirstOutputSheetName(setupWs, selectedTemplate, lastSetupRow)
    If outputSheetName = "" Then outputSheetName = "Pivot_Output"

    Set outputWs = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.Count))
    outputWs.Name = UniqueSheetName(targetWb, SafeSheetName(outputSheetName))

    outputWs.Range("A1").Value = "Pivot template: " & selectedTemplate
    outputWs.Range("A2").Value = "Created: " & Format(Now, "yyyy-mm-dd hh:mm")
    outputWs.Range("A1:A2").Font.Bold = True

    outputRow = 5
    builtCount = 0
    csvCount = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For rowIndex = 9 To lastSetupRow
        If IsBuildSetupRow(setupWs, rowIndex, selectedTemplate) Then

            Set pivot = CreateOnePivot( _
                cache, _
                outputWs, _
                outputRow, _
                CStr(setupWs.Cells(rowIndex, "B").Value), _
                CStr(setupWs.Cells(rowIndex, "E").Value), _
                CStr(setupWs.Cells(rowIndex, "F").Value), _
                CStr(setupWs.Cells(rowIndex, "G").Value), _
                CStr(setupWs.Cells(rowIndex, "H").Value), _
                CStr(setupWs.Cells(rowIndex, "I").Value), _
                CStr(setupWs.Cells(rowIndex, "J").Value), _
                CStr(setupWs.Cells(rowIndex, "K").Value), _
                headers _
            )

            builtCount = builtCount + 1
            saveBehavior = NormalizedSaveBehavior(CStr(setupWs.Cells(rowIndex, "D").Value))
            If saveBehavior = "EXPORT CSV" Then
                csvFileName = Trim$(CStr(setupWs.Cells(rowIndex, "L").Value))
                If csvFileName = "" Then csvFileName = Trim$(CStr(setupWs.Cells(rowIndex, "B").Value))
                If csvFileName = "" Then csvFileName = "Pivot_" & CStr(builtCount)
                If csvFolder = "" Then Err.Raise vbObjectError + 300, , "Cannot export CSV because the source workbook folder could not be found."
                ExportRangeToCsv pivot.TableRange2, BuildCsvPath(csvFolder, csvFileName)
                csvCount = csvCount + 1
            End If
            outputRow = pivot.TableRange2.Row + pivot.TableRange2.Rows.Count + 3
        End If
    Next rowIndex

    outputWs.Columns.AutoFit
    outputWs.Activate

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If builtCount = 0 Then
        MsgBox "No PivotTables found for template: " & selectedTemplate & vbCrLf & "Fill Pivot Name for each row you want to build.", vbExclamation
        targetWb.Close SaveChanges:=False
        Exit Sub
    End If

    targetWb.Save
    If convertedFromCsv Then
        MsgBox "Created " & builtCount & " PivotTable(s), exported " & csvCount & " CSV file(s), and saved an Excel copy:" & vbCrLf & convertedWorkbookPath, vbInformation
    Else
        MsgBox "Created " & builtCount & " PivotTable(s), exported " & csvCount & " CSV file(s), and saved the source workbook.", vbInformation
    End If

    targetWb.Close SaveChanges:=False
    Exit Sub

BuildError:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not targetWb Is Nothing Then targetWb.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "Could not build PivotTables:" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Function CreateOnePivot( _
    ByVal cache As PivotCache, _
    ByVal outputWs As Worksheet, _
    ByVal outputRow As Long, _
    ByVal pivotName As String, _
    ByVal rowFieldsText As String, _
    ByVal rowNamesText As String, _
    ByVal rowGroupRulesText As String, _
    ByVal displayFieldsText As String, _
    ByVal sumFieldsText As String, _
    ByVal filterFieldsText As String, _
    ByVal conditionsText As String, _
    ByVal headers As Variant _
) As PivotTable

    Dim pivot As PivotTable
    Dim rowFields As Variant
    Dim rowNames As Variant
    Dim rowGroups As Variant
    Dim displayFields As Variant
    Dim sumFields As Variant
    Dim filterFields As Variant
    Dim item As Variant
    Dim position As Long
    Dim finalName As String
    Dim displayName As String
    Dim sourceField As String
    Dim sourceInput As String
    Dim rowNameText As String
    Dim groupRulesText As String
    Dim combinedGroupText As String
    Dim rowCaption As String
    Dim rowSourceFields() As String
    Dim rowCaptions() As String
    Dim rowFieldCount As Long
    Dim filterCount As Long
    Dim reservedRows As Long
    Dim pivotStartRow As Long

    displayName = PivotDisplayName(pivotName, conditionsText, filterFieldsText, rowFieldsText, sumFieldsText)
    finalName = SafePivotName(displayName)
    If finalName = "" Then finalName = "Pivot_" & CStr(outputRow)

    rowFields = SplitList(rowFieldsText)
    rowNames = SplitList(rowNamesText)
    rowGroups = SplitGroupRulesList(rowGroupRulesText)
    displayFields = SplitList(displayFieldsText)
    sumFields = SplitList(sumFieldsText)
    filterFields = SplitList(filterFieldsText)
    filterCount = UBoundSafe(filterFields) + 1
    reservedRows = filterCount + 2
    If reservedRows < 2 Then reservedRows = 2
    pivotStartRow = outputRow + 2 + reservedRows

    outputWs.Cells(outputRow, 1).Value = displayName
    outputWs.Cells(outputRow, 1).Font.Bold = True
    With outputWs.Cells(outputRow + 1, 1)
        .Value = RowFieldDisplayText(rowFields, rowNames, rowGroups)
        .Font.Italic = True
        .Font.Color = RGB(90, 96, 104)
        .WrapText = True
        .ShrinkToFit = True
    End With

    Set pivot = cache.CreatePivotTable( _
        TableDestination:=outputWs.Cells(pivotStartRow, 1), _
        TableName:=finalName & "_" & CStr(outputRow) _
    )

    For position = 0 To MaxLong(UBoundSafe(rowFields), MaxLong(UBoundSafe(rowNames), UBoundSafe(rowGroups)))
        sourceInput = ""
        If position <= UBoundSafe(rowFields) Then sourceInput = CStr(rowFields(position))
        sourceField = sourceInput
        rowNameText = ""
        If position <= UBoundSafe(rowNames) Then rowNameText = CStr(rowNames(position))
        groupRulesText = ""
        If position <= UBoundSafe(rowGroups) Then groupRulesText = CStr(rowGroups(position))

        If Trim$(groupRulesText) <> "" Then
            combinedGroupText = BuildRowGroupText(rowNameText, sourceInput, groupRulesText)
            sourceField = RowGroupHeader(sourceInput, combinedGroupText)
            rowCaption = RowGroupCaption(combinedGroupText, sourceInput)
            If FieldExists(headers, rowCaption) Then rowCaption = rowCaption & " Group"
        ElseIf HasRowGroupRule(sourceField) Then
            rowNameText = sourceField
            sourceField = RowGroupHeader("", rowNameText)
            rowCaption = RowGroupCaption(rowNameText, "Condition Group")
            If FieldExists(headers, rowCaption) Then rowCaption = rowCaption & " Group"
        ElseIf HasRowGroupRule(rowNameText) Then
            sourceField = RowGroupHeader(sourceInput, rowNameText)
            rowCaption = RowGroupCaption(rowNameText, sourceInput)
        Else
            rowCaption = RowGroupCaption(rowNameText, sourceField)
        End If

        If Trim$(sourceField) <> "" And FieldExists(headers, sourceField) Then
            rowFieldCount = rowFieldCount + 1
            ReDim Preserve rowSourceFields(1 To rowFieldCount)
            ReDim Preserve rowCaptions(1 To rowFieldCount)
            rowSourceFields(rowFieldCount) = sourceField
            rowCaptions(rowFieldCount) = rowCaption
        Else
            Err.Raise vbObjectError + 100, , "Field not found: " & sourceField
        End If
    Next position

    For Each item In displayFields
        If FieldExists(headers, CStr(item)) Then
            rowFieldCount = rowFieldCount + 1
            ReDim Preserve rowSourceFields(1 To rowFieldCount)
            ReDim Preserve rowCaptions(1 To rowFieldCount)
            rowSourceFields(rowFieldCount) = CStr(item)
            rowCaptions(rowFieldCount) = ""
        Else
            Err.Raise vbObjectError + 100, , "Field not found: " & CStr(item)
        End If
    Next item

    If rowFieldCount > 0 Then
        ApplyPivotRowFields pivot, rowSourceFields, rowCaptions, rowFieldCount
    End If

    For Each item In sumFields
        If FieldExists(headers, CStr(item)) Then
            If PivotFieldLooksNumeric(pivot.PivotFields(CStr(item))) Then
                pivot.AddDataField pivot.PivotFields(CStr(item)), "Sum of " & CStr(item), xlSum
            Else
                pivot.AddDataField pivot.PivotFields(CStr(item)), "Count of " & CStr(item), xlCount
            End If
        Else
            Err.Raise vbObjectError + 101, , "Field not found: " & CStr(item)
        End If
    Next item

    position = 1
    For Each item In filterFields
        If FieldExists(headers, CStr(item)) Then
            With pivot.PivotFields(CStr(item))
                .Orientation = xlPageField
                .Position = position
            End With
            position = position + 1
        Else
            Err.Raise vbObjectError + 103, , "Filter field not found: " & CStr(item)
        End If
    Next item

    ApplyConditions pivot, conditionsText, headers

    On Error Resume Next
    pivot.RowAxisLayout xlTabularRow
    pivot.TableStyle2 = "PivotStyleMedium9"
    pivot.PivotCache.Refresh
    pivot.RefreshTable
    On Error GoTo 0

    Set CreateOnePivot = pivot
End Function

Private Function PivotFieldLooksNumeric(ByVal pf As PivotField) As Boolean
    Dim item As PivotItem
    Dim checkedCount As Long

    On Error GoTo NotNumeric
    For Each item In pf.PivotItems
        If Trim$(CStr(item.Name)) <> "" And Trim$(CStr(item.Name)) <> "(blank)" Then
            checkedCount = checkedCount + 1
            If Not IsNumeric(item.Name) Then GoTo NotNumeric
            If checkedCount >= 20 Then Exit For
        End If
    Next item

    PivotFieldLooksNumeric = checkedCount > 0
    Exit Function

NotNumeric:
    PivotFieldLooksNumeric = False
End Function

Private Sub ApplyPivotRowFields(ByVal pivot As PivotTable, ByRef rowSourceFields() As String, ByRef rowCaptions() As String, ByVal rowFieldCount As Long)
    Dim fieldsVariant As Variant
    Dim index As Long

    ReDim fieldsVariant(0 To rowFieldCount - 1)
    For index = 1 To rowFieldCount
        fieldsVariant(index - 1) = rowSourceFields(index)
    Next index

    On Error GoTo AddFieldsFallback
    pivot.AddFields RowFields:=fieldsVariant
    On Error GoTo 0

    For index = 1 To rowFieldCount
        With pivot.PivotFields(rowSourceFields(index))
            .Position = index
            If Trim$(rowCaptions(index)) <> "" Then
                On Error Resume Next
                .Caption = rowCaptions(index)
                On Error GoTo 0
            End If
        End With
    Next index
    Exit Sub

AddFieldsFallback:
    Err.Clear
    On Error GoTo 0
    For index = 1 To rowFieldCount
        With pivot.PivotFields(rowSourceFields(index))
            .Orientation = xlRowField
            .Position = index
            If Trim$(rowCaptions(index)) <> "" Then
                On Error Resume Next
                .Caption = rowCaptions(index)
                On Error GoTo 0
            End If
        End With
    Next index
End Sub

Private Function RowFieldDisplayText(ByVal rowFields As Variant, ByVal rowNames As Variant, ByVal rowGroups As Variant) As String
    Dim index As Long
    Dim result As String
    Dim separator As String
    Dim fieldName As String
    Dim displayName As String
    Dim groupRulesText As String
    Dim maxIndex As Long

    maxIndex = MaxLong(UBoundSafe(rowFields), MaxLong(UBoundSafe(rowNames), UBoundSafe(rowGroups)))
    If maxIndex < 0 Then
        RowFieldDisplayText = "Rows: none"
        Exit Function
    End If

    For index = 0 To maxIndex
        fieldName = Trim$(ListItemOrBlank(rowFields, index))
        displayName = Trim$(ListItemOrBlank(rowNames, index))
        groupRulesText = Trim$(ListItemOrBlank(rowGroups, index))

        If fieldName <> "" Or displayName <> "" Or groupRulesText <> "" Then
            If groupRulesText <> "" Then
                If fieldName <> "" Then
                    result = result & separator & fieldName & " = " & RowGroupCaption(BuildRowGroupText(displayName, fieldName, groupRulesText), fieldName)
                Else
                    result = result & separator & RowGroupCaption(BuildRowGroupText(displayName, fieldName, groupRulesText), "Condition Group")
                End If
            ElseIf HasRowGroupRule(fieldName) Then
                result = result & separator & RowGroupCaption(fieldName, "Condition Group")
            ElseIf HasRowGroupRule(displayName) Then
                result = result & separator & fieldName
                result = result & " = " & RowGroupCaption(displayName, fieldName)
            ElseIf displayName <> "" And fieldName <> "" Then
                result = result & separator & fieldName
                result = result & " = " & displayName
            Else
                result = result & separator & IIf(fieldName <> "", fieldName, displayName)
            End If
            separator = "; "
        End If
    Next index

    If result = "" Then result = "none"
    RowFieldDisplayText = "Rows: " & result
End Function

Private Function HeadersWithRowGroups(ByVal baseHeaders As Variant, ByVal setupWs As Worksheet, ByVal selectedTemplate As String, ByVal lastSetupRow As Long) As Variant
    Dim headers() As String
    Dim rowFields As Variant
    Dim rowNames As Variant
    Dim rowGroups As Variant
    Dim setupRow As Long
    Dim index As Long
    Dim count As Long
    Dim helperHeader As String
    Dim groupText As String
    Dim maxIndex As Long

    ReDim headers(LBound(baseHeaders) To UBound(baseHeaders))
    For index = LBound(baseHeaders) To UBound(baseHeaders)
        headers(index) = CStr(baseHeaders(index))
    Next index
    count = UBound(headers)

    For setupRow = 9 To lastSetupRow
        If IsBuildSetupRow(setupWs, setupRow, selectedTemplate) Then
            rowFields = SplitList(CStr(setupWs.Cells(setupRow, "E").Value))
            rowNames = SplitList(CStr(setupWs.Cells(setupRow, "F").Value))
            rowGroups = SplitGroupRulesList(CStr(setupWs.Cells(setupRow, "G").Value))
            maxIndex = MaxLong(UBoundSafe(rowFields), MaxLong(UBoundSafe(rowNames), UBoundSafe(rowGroups)))
            If maxIndex >= 0 Then
                For index = 0 To maxIndex
                    If index <= UBoundSafe(rowGroups) And Trim$(CStr(rowGroups(index))) <> "" Then
                        groupText = BuildRowGroupText(ListItemOrBlank(rowNames, index), ListItemOrBlank(rowFields, index), CStr(rowGroups(index)))
                        helperHeader = RowGroupHeader(ListItemOrBlank(rowFields, index), groupText)
                        If helperHeader <> "" And Not FieldExists(headers, helperHeader) Then
                            count = count + 1
                            ReDim Preserve headers(LBound(baseHeaders) To count)
                            headers(count) = helperHeader
                        End If
                    ElseIf index <= UBoundSafe(rowFields) And HasRowGroupRule(CStr(rowFields(index))) Then
                        helperHeader = RowGroupHeader("", CStr(rowFields(index)))
                        If helperHeader <> "" And Not FieldExists(headers, helperHeader) Then
                            count = count + 1
                            ReDim Preserve headers(LBound(baseHeaders) To count)
                            headers(count) = helperHeader
                        End If
                    ElseIf index <= UBoundSafe(rowNames) Then
                        If HasRowGroupRule(CStr(rowNames(index))) Then
                            helperHeader = RowGroupHeader(ListItemOrBlank(rowFields, index), CStr(rowNames(index)))
                            If helperHeader <> "" And Not FieldExists(headers, helperHeader) Then
                                count = count + 1
                                ReDim Preserve headers(LBound(baseHeaders) To count)
                                headers(count) = helperHeader
                            End If
                        End If
                    End If
                Next index
            End If
        End If
    Next setupRow

    HeadersWithRowGroups = headers
End Function

Private Function HasRowGroupRule(ByVal rowNameText As String) As Boolean
    HasRowGroupRule = InStr(1, rowNameText, "{", vbTextCompare) > 0 And InStr(1, rowNameText, "}", vbTextCompare) > InStr(1, rowNameText, "{", vbTextCompare)
End Function

Private Function RowGroupCaption(ByVal rowNameText As String, ByVal fieldName As String) As String
    Dim openPos As Long
    openPos = InStr(1, rowNameText, "{", vbTextCompare)
    If openPos > 1 Then
        RowGroupCaption = Trim$(Left$(rowNameText, openPos - 1))
    ElseIf HasRowGroupRule(rowNameText) Then
        If Trim$(fieldName) <> "" Then
            RowGroupCaption = Trim$(fieldName) & " Group"
        Else
            RowGroupCaption = "Condition Group"
        End If
    Else
        RowGroupCaption = Trim$(rowNameText)
    End If
End Function

Private Function RowGroupHeader(ByVal fieldName As String, ByVal rowNameText As String) As String
    Dim captionText As String

    captionText = RowGroupCaption(rowNameText, fieldName)
    If captionText = "" Then
        If Trim$(fieldName) <> "" Then
            captionText = Trim$(fieldName) & " Group"
        Else
            captionText = "Condition Group"
        End If
    End If

    If HasRowGroupRule(rowNameText) Then
        RowGroupHeader = UniquePivotGroupHeader(captionText)
    Else
        RowGroupHeader = captionText
    End If
End Function

Private Function UniquePivotGroupHeader(ByVal captionText As String) As String
    Dim result As String

    result = Trim$(captionText)
    If result = "" Then result = "Condition Group"
    result = Replace(result, "{", " ")
    result = Replace(result, "}", " ")
    result = Replace(result, "=", " ")
    result = Replace(result, ":", " ")
    result = Replace(result, ";", " ")
    result = Replace(result, "|", " ")
    result = Replace(result, "#", " No ")
    Do While InStr(1, result, "  ", vbTextCompare) > 0
        result = Replace(result, "  ", " ")
    Loop
    result = Trim$(result)
    If result = "" Then result = "Condition Group"

    UniquePivotGroupHeader = "PB Group - " & result
End Function

Private Function RowGroupRules(ByVal rowNameText As String) As String
    Dim openPos As Long
    Dim closePos As Long
    openPos = InStr(1, rowNameText, "{", vbTextCompare)
    closePos = InStrRev(rowNameText, "}")
    If openPos > 0 And closePos > openPos Then
        RowGroupRules = Mid$(rowNameText, openPos + 1, closePos - openPos - 1)
    End If
End Function

Private Function BuildRowGroupText(ByVal groupName As String, ByVal fieldName As String, ByVal rulesText As String) As String
    Dim captionText As String

    captionText = Trim$(groupName)
    If captionText = "" Then
        If Trim$(fieldName) <> "" Then
            captionText = Trim$(fieldName) & " Group"
        Else
            captionText = "Condition Group"
        End If
    End If

    BuildRowGroupText = captionText & "{" & Trim$(rulesText) & "}"
End Function

Private Function ListItemOrBlank(ByVal values As Variant, ByVal index As Long) As String
    If index >= 0 And index <= UBoundSafe(values) Then
        ListItemOrBlank = CStr(values(index))
    Else
        ListItemOrBlank = ""
    End If
End Function

Private Function MaxLong(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLong = leftValue
    Else
        MaxLong = rightValue
    End If
End Function

Private Function BuildCsvPath(ByVal folderPath As String, ByVal fileName As String) As String
    Dim cleanName As String
    cleanName = SafeFileName(fileName)
    If LCase$(Right$(cleanName, 4)) <> ".csv" Then cleanName = cleanName & ".csv"

    If Right$(folderPath, 1) = Application.PathSeparator Then
        BuildCsvPath = folderPath & cleanName
    Else
        BuildCsvPath = folderPath & Application.PathSeparator & cleanName
    End If
End Function

Private Function SafeFileName(ByVal fileName As String) As String
    Dim result As String
    result = Trim$(fileName)
    If result = "" Then result = "pivot.csv"
    result = Replace(result, "\", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, "?", "_")
    result = Replace(result, """", "_")
    result = Replace(result, "<", "_")
    result = Replace(result, ">", "_")
    result = Replace(result, "|", "_")
    SafeFileName = result
End Function

Private Function PivotDisplayName( _
    ByVal pivotName As String, _
    ByVal conditionsText As String, _
    ByVal filterFieldsText As String, _
    ByVal rowFieldsText As String, _
    ByVal sumFieldsText As String _
) As String
    Dim conditionName As String
    Dim fieldName As String

    If Trim$(pivotName) <> "" Then
        PivotDisplayName = Trim$(pivotName)
        Exit Function
    End If

    conditionName = ConditionsDisplayName(conditionsText)
    If conditionName <> "" Then
        PivotDisplayName = conditionName
        Exit Function
    End If

    If Trim$(filterFieldsText) <> "" Then
        PivotDisplayName = "Filtered by " & Replace(Trim$(filterFieldsText), ",", ", ")
        Exit Function
    End If

    If Trim$(rowFieldsText) <> "" Or Trim$(sumFieldsText) <> "" Then
        PivotDisplayName = "Rows " & Replace(Trim$(rowFieldsText), ",", ", ") & " Values " & Replace(Trim$(sumFieldsText), ",", ", ")
        PivotDisplayName = Trim$(PivotDisplayName)
        Exit Function
    End If

    PivotDisplayName = "Pivot"
End Function

Private Function ConditionsDisplayName(ByVal conditionsText As String) As String
    Dim conditions As Variant
    Dim condition As Variant
    Dim parts As Variant
    Dim fieldName As String
    Dim fieldValue As String
    Dim result As String
    Dim separator As String

    conditions = Split(conditionsText, ";")
    For Each condition In conditions
        If InStr(1, CStr(condition), "=", vbTextCompare) > 0 Then
            parts = Split(CStr(condition), "=", 2)
            fieldName = Trim$(CStr(parts(0)))
            fieldValue = Trim$(CStr(parts(1)))
            If fieldName <> "" And fieldValue <> "" Then
                result = result & separator & fieldName & " " & fieldValue
                separator = ", "
            End If
        End If
    Next condition

    ConditionsDisplayName = result
End Function

Private Sub ExportRangeToCsv(ByVal sourceRange As Range, ByVal outputPath As String)
    Dim fileNumber As Integer
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim lineText As String
    Dim valueText As String

    fileNumber = FreeFile
    Open outputPath For Output As #fileNumber

    For rowIndex = 1 To sourceRange.Rows.Count
        lineText = ""
        For colIndex = 1 To sourceRange.Columns.Count
            valueText = CsvEscape(CStr(sourceRange.Cells(rowIndex, colIndex).Text))
            If colIndex > 1 Then lineText = lineText & ","
            lineText = lineText & valueText
        Next colIndex
        Print #fileNumber, lineText
    Next rowIndex

    Close #fileNumber
End Sub

Private Function CsvEscape(ByVal valueText As String) As String
    If InStr(1, valueText, """", vbBinaryCompare) > 0 Then
        valueText = Replace(valueText, """", """""")
    End If
    If InStr(1, valueText, ",", vbBinaryCompare) > 0 _
        Or InStr(1, valueText, vbCr, vbBinaryCompare) > 0 _
        Or InStr(1, valueText, vbLf, vbBinaryCompare) > 0 _
        Or InStr(1, valueText, """", vbBinaryCompare) > 0 Then
        CsvEscape = """" & valueText & """"
    Else
        CsvEscape = valueText
    End If
End Function

Private Sub ApplyConditions(ByVal pivot As PivotTable, ByVal conditionsText As String, ByVal headers As Variant)
    Dim conditions As Variant
    Dim condition As Variant
    Dim parts As Variant
    Dim fieldName As String
    Dim fieldValue As String
    Dim pf As PivotField

    conditions = Split(conditionsText, ";")
    For Each condition In conditions
        If InStr(1, CStr(condition), "=", vbTextCompare) > 0 Then
            parts = Split(CStr(condition), "=", 2)
            fieldName = Trim$(CStr(parts(0)))
            fieldValue = Trim$(CStr(parts(1)))

            If fieldName <> "" And fieldValue <> "" Then
                If Not FieldExists(headers, fieldName) Then
                    Err.Raise vbObjectError + 102, , "Condition field not found: " & fieldName
                End If

                Set pf = pivot.PivotFields(fieldName)
                pf.Orientation = xlPageField
                On Error Resume Next
                pf.CurrentPage = fieldValue
                If Err.Number <> 0 Then
                    Err.Clear
                    pf.ClearAllFilters
                    pf.PivotFilters.Add Type:=xlCaptionEquals, Value1:=fieldValue
                End If
                On Error GoTo 0
            End If
        End If
    Next condition
End Sub

Private Function CreateNormalizedSourceSheet(ByVal wb As Workbook, ByVal dataWs As Worksheet, ByVal headers As Variant, ByVal setupWs As Worksheet, ByVal selectedTemplate As String, ByVal lastSetupRow As Long) As Worksheet
    Dim dataRange As Range
    Dim helperWs As Worksheet
    Dim rowsCount As Long
    Dim colsCount As Long
    Dim totalColsCount As Long
    Dim colIndex As Long

    Set dataRange = dataWs.UsedRange
    rowsCount = dataRange.Rows.Count
    colsCount = dataRange.Columns.Count
    totalColsCount = UBound(headers)

    Set helperWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    helperWs.Name = UniqueSheetName(wb, HELPER_PREFIX)

    For colIndex = 1 To totalColsCount
        helperWs.Cells(1, colIndex).Value = headers(colIndex)
    Next colIndex

    If rowsCount > 1 Then
        dataRange.Offset(1, 0).Resize(rowsCount - 1, colsCount).Copy Destination:=helperWs.Cells(2, 1)
        FillRowGroupColumns helperWs, dataRange, headers, setupWs, selectedTemplate, lastSetupRow
    End If

    dataWs.Activate
    helperWs.Visible = xlSheetVeryHidden
    Set CreateNormalizedSourceSheet = helperWs
End Function

Private Sub FillRowGroupColumns(ByVal helperWs As Worksheet, ByVal dataRange As Range, ByVal headers As Variant, ByVal setupWs As Worksheet, ByVal selectedTemplate As String, ByVal lastSetupRow As Long)
    Dim rowFields As Variant
    Dim rowNames As Variant
    Dim rowGroups As Variant
    Dim setupRow As Long
    Dim index As Long
    Dim sourceCol As Long
    Dim helperCol As Long
    Dim rowIndex As Long
    Dim rulesText As String
    Dim helperHeader As String
    Dim groupText As String
    Dim sourceField As String
    Dim groupRange As Range
    Dim maxIndex As Long

    For setupRow = 9 To lastSetupRow
        If IsBuildSetupRow(setupWs, setupRow, selectedTemplate) Then
            rowFields = SplitList(CStr(setupWs.Cells(setupRow, "E").Value))
            rowNames = SplitList(CStr(setupWs.Cells(setupRow, "F").Value))
            rowGroups = SplitGroupRulesList(CStr(setupWs.Cells(setupRow, "G").Value))
            maxIndex = MaxLong(UBoundSafe(rowFields), MaxLong(UBoundSafe(rowNames), UBoundSafe(rowGroups)))

            If maxIndex >= 0 Then
                For index = 0 To maxIndex
                    If index <= UBoundSafe(rowGroups) And Trim$(CStr(rowGroups(index))) <> "" Then
                        sourceField = ListItemOrBlank(rowFields, index)
                        sourceCol = HeaderIndex(headers, sourceField)
                        groupText = BuildRowGroupText(ListItemOrBlank(rowNames, index), sourceField, CStr(rowGroups(index)))
                        helperHeader = RowGroupHeader(sourceField, groupText)
                        helperCol = HeaderIndex(headers, helperHeader)
                        rulesText = RowGroupRules(groupText)

                        If helperCol > 0 Then
                            helperWs.Columns(helperCol).NumberFormat = "@"
                            For rowIndex = 2 To dataRange.Rows.Count
                                If sourceCol > 0 Then
                                    helperWs.Cells(rowIndex, helperCol).Value = ApplyRowGroupRules(dataRange.Cells(rowIndex, sourceCol).Value, rulesText, dataRange, headers, rowIndex, RowGroupCaption(groupText, sourceField))
                                Else
                                    helperWs.Cells(rowIndex, helperCol).Value = ApplyRowGroupRules("", rulesText, dataRange, headers, rowIndex, RowGroupCaption(groupText, sourceField))
                                End If
                            Next rowIndex
                            Set groupRange = helperWs.Cells(2, helperCol).Resize(dataRange.Rows.Count - 1, 1)
                            WarnIfGroupMatchedNothing groupRange, RowGroupCaption(groupText, sourceField)
                        End If
                    ElseIf index <= UBoundSafe(rowFields) And HasRowGroupRule(CStr(rowFields(index))) Then
                        sourceCol = 0
                        helperHeader = RowGroupHeader("", CStr(rowFields(index)))
                        helperCol = HeaderIndex(headers, helperHeader)
                        rulesText = RowGroupRules(CStr(rowFields(index)))

                        If helperCol > 0 Then
                            helperWs.Columns(helperCol).NumberFormat = "@"
                            For rowIndex = 2 To dataRange.Rows.Count
                                helperWs.Cells(rowIndex, helperCol).Value = ApplyRowGroupRules("", rulesText, dataRange, headers, rowIndex, RowGroupCaption(CStr(rowFields(index)), "Condition Group"))
                            Next rowIndex
                            Set groupRange = helperWs.Cells(2, helperCol).Resize(dataRange.Rows.Count - 1, 1)
                            WarnIfGroupMatchedNothing groupRange, RowGroupCaption(CStr(rowFields(index)), "Condition Group")
                        End If
                    ElseIf index <= UBoundSafe(rowNames) Then
                        If HasRowGroupRule(CStr(rowNames(index))) Then
                            sourceField = ListItemOrBlank(rowFields, index)
                            sourceCol = HeaderIndex(headers, sourceField)
                            helperHeader = RowGroupHeader(sourceField, CStr(rowNames(index)))
                            helperCol = HeaderIndex(headers, helperHeader)
                            rulesText = RowGroupRules(CStr(rowNames(index)))

                            If sourceCol > 0 And helperCol > 0 Then
                                helperWs.Columns(helperCol).NumberFormat = "@"
                                For rowIndex = 2 To dataRange.Rows.Count
                                    helperWs.Cells(rowIndex, helperCol).Value = ApplyRowGroupRules(dataRange.Cells(rowIndex, sourceCol).Value, rulesText, dataRange, headers, rowIndex, RowGroupCaption(CStr(rowNames(index)), sourceField))
                                Next rowIndex
                                Set groupRange = helperWs.Cells(2, helperCol).Resize(dataRange.Rows.Count - 1, 1)
                                WarnIfGroupMatchedNothing groupRange, RowGroupCaption(CStr(rowNames(index)), sourceField)
                            End If
                        End If
                    End If
                Next index
            End If
        End If
    Next setupRow
End Sub

Private Sub WarnIfGroupMatchedNothing(ByVal groupRange As Range, ByVal groupName As String)
    Dim nonOtherCount As Long

    On Error Resume Next
    nonOtherCount = Application.WorksheetFunction.CountIf(groupRange, "<>Other")
    On Error GoTo 0

    If nonOtherCount = 0 Then
        MsgBox "The row group '" & groupName & "' did not match any source rows." & vbCrLf & _
            "All rows were placed in Other. Check the field names and values inside the group rule.", vbExclamation
    End If
End Sub

Private Function ApplyRowGroupRules(ByVal rawValue As Variant, ByVal rulesText As String, ByVal dataRange As Range, ByVal headers As Variant, ByVal dataRow As Long, ByVal defaultLabel As String) As String
    Dim rules As Variant
    Dim rule As Variant
    Dim labelText As String
    Dim rangeText As String
    Dim bounds As Variant
    Dim valueMinutes As Long
    Dim startMinutes As Long
    Dim endMinutes As Long
    Dim equalPos As Long
    Dim colonPos As Long

    valueMinutes = TimeValueMinutes(rawValue)

    rules = Split(rulesText, "|")
    For Each rule In rules
        equalPos = InStr(1, CStr(rule), "=", vbTextCompare)
        colonPos = InStr(1, CStr(rule), ":", vbTextCompare)

        If colonPos > 0 And (equalPos = 0 Or colonPos < equalPos) Then
            labelText = Trim$(Left$(CStr(rule), colonPos - 1))
            If labelText <> "" And RowGroupConditionsMatch(dataRange, headers, dataRow, Mid$(CStr(rule), colonPos + 1)) Then
                ApplyRowGroupRules = labelText
                Exit Function
            End If
        ElseIf equalPos > 0 Then
            labelText = Trim$(Left$(CStr(rule), equalPos - 1))
            rangeText = Trim$(Mid$(CStr(rule), equalPos + 1))
            If valueMinutes >= 0 Then
                bounds = Split(rangeText, "-")
                If UBound(bounds) = 1 Then
                    startMinutes = TimeTextMinutes(CStr(bounds(0)))
                    endMinutes = TimeTextMinutes(CStr(bounds(1)))
                    If labelText <> "" And startMinutes >= 0 And endMinutes >= 0 Then
                        If TimeMinutesInRange(valueMinutes, startMinutes, endMinutes) Then
                            ApplyRowGroupRules = labelText
                            Exit Function
                        End If
                    End If
                End If
            End If
            If RowGroupConditionsMatch(dataRange, headers, dataRow, CStr(rule)) Then
                ApplyRowGroupRules = IIf(Trim$(defaultLabel) <> "", Trim$(defaultLabel), labelText)
                Exit Function
            End If
        End If
    Next rule

    ApplyRowGroupRules = "Other"
End Function

Private Function RowGroupConditionsMatch(ByVal dataRange As Range, ByVal headers As Variant, ByVal dataRow As Long, ByVal conditionsText As String) As Boolean
    Dim conditions As Variant
    Dim condition As Variant
    Dim fieldName As String
    Dim expectedText As String
    Dim sourceCol As Long
    Dim testedAny As Boolean
    Dim operatorText As String

    conditions = Split(conditionsText, ";")
    For Each condition In conditions
        operatorText = ConditionOperator(CStr(condition))
        If operatorText <> "" Then
            fieldName = Trim$(Left$(CStr(condition), InStr(1, CStr(condition), operatorText, vbTextCompare) - 1))
            expectedText = Trim$(Mid$(CStr(condition), InStr(1, CStr(condition), operatorText, vbTextCompare) + Len(operatorText)))
            If fieldName <> "" And expectedText <> "" Then
                testedAny = True
                sourceCol = HeaderIndex(headers, fieldName)
                If sourceCol = 0 Or sourceCol > dataRange.Columns.Count Then
                    RowGroupConditionsMatch = False
                    Exit Function
                End If
                If Not RowGroupValueMatches(dataRange.Cells(dataRow, sourceCol), operatorText, expectedText) Then
                    RowGroupConditionsMatch = False
                    Exit Function
                End If
            End If
        End If
    Next condition

    RowGroupConditionsMatch = testedAny
End Function

Private Function ConditionOperator(ByVal conditionText As String) As String
    If InStr(1, conditionText, ">=", vbTextCompare) > 0 Then
        ConditionOperator = ">="
    ElseIf InStr(1, conditionText, "<=", vbTextCompare) > 0 Then
        ConditionOperator = "<="
    ElseIf InStr(1, conditionText, "<>", vbTextCompare) > 0 Then
        ConditionOperator = "<>"
    ElseIf InStr(1, conditionText, ">", vbTextCompare) > 0 Then
        ConditionOperator = ">"
    ElseIf InStr(1, conditionText, "<", vbTextCompare) > 0 Then
        ConditionOperator = "<"
    ElseIf InStr(1, conditionText, "=", vbTextCompare) > 0 Then
        ConditionOperator = "="
    End If
End Function

Private Function RowGroupValueMatches(ByVal sourceCell As Range, ByVal operatorText As String, ByVal expectedText As String) As Boolean
    Dim bounds As Variant
    Dim valueMinutes As Long
    Dim startMinutes As Long
    Dim endMinutes As Long
    Dim sourceNumber As Double
    Dim expectedNumber As Double
    Dim lowerNumber As Double
    Dim upperNumber As Double

    If IsNumeric(sourceCell.Value) And IsNumeric(expectedText) Then
        sourceNumber = CDbl(sourceCell.Value)
        expectedNumber = CDbl(expectedText)
        Select Case operatorText
            Case "="
                RowGroupValueMatches = sourceNumber = expectedNumber
            Case "<>"
                RowGroupValueMatches = sourceNumber <> expectedNumber
            Case ">"
                RowGroupValueMatches = sourceNumber > expectedNumber
            Case "<"
                RowGroupValueMatches = sourceNumber < expectedNumber
            Case ">="
                RowGroupValueMatches = sourceNumber >= expectedNumber
            Case "<="
                RowGroupValueMatches = sourceNumber <= expectedNumber
        End Select
        Exit Function
    End If

    If operatorText = "=" And InStr(1, expectedText, "-", vbTextCompare) > 0 Then
        bounds = Split(expectedText, "-")
        If UBound(bounds) = 1 Then
            If IsNumeric(sourceCell.Value) And IsNumeric(Trim$(CStr(bounds(0)))) And IsNumeric(Trim$(CStr(bounds(1)))) Then
                sourceNumber = CDbl(sourceCell.Value)
                lowerNumber = CDbl(Trim$(CStr(bounds(0))))
                upperNumber = CDbl(Trim$(CStr(bounds(1))))
                RowGroupValueMatches = sourceNumber >= lowerNumber And sourceNumber <= upperNumber
                Exit Function
            End If
            valueMinutes = TimeValueMinutes(sourceCell.Value)
            startMinutes = TimeTextMinutes(CStr(bounds(0)))
            endMinutes = TimeTextMinutes(CStr(bounds(1)))
            If valueMinutes >= 0 And startMinutes >= 0 And endMinutes >= 0 Then
                RowGroupValueMatches = TimeMinutesInRange(valueMinutes, startMinutes, endMinutes)
                Exit Function
            End If
        End If
    End If

    If operatorText = "=" Then
        RowGroupValueMatches = StrComp(Trim$(CStr(sourceCell.Value)), expectedText, vbTextCompare) = 0 _
            Or StrComp(Trim$(CStr(sourceCell.Text)), expectedText, vbTextCompare) = 0
    ElseIf operatorText = "<>" Then
        RowGroupValueMatches = StrComp(Trim$(CStr(sourceCell.Value)), expectedText, vbTextCompare) <> 0 _
            And StrComp(Trim$(CStr(sourceCell.Text)), expectedText, vbTextCompare) <> 0
    Else
        RowGroupValueMatches = False
    End If
End Function

Private Function TimeValueMinutes(ByVal rawValue As Variant) As Long
    Dim textValue As String
    Dim timePart As String

    On Error GoTo NotTime
    If IsDate(rawValue) Then
        TimeValueMinutes = Hour(CDate(rawValue)) * 60 + Minute(CDate(rawValue))
        Exit Function
    End If

    textValue = Trim$(CStr(rawValue))
    If InStr(1, textValue, " ", vbTextCompare) > 0 Then
        timePart = Mid$(textValue, InStrRev(textValue, " ") + 1)
    Else
        timePart = textValue
    End If

    TimeValueMinutes = TimeTextMinutes(timePart)
    Exit Function

NotTime:
    TimeValueMinutes = -1
End Function

Private Function TimeTextMinutes(ByVal timeText As String) As Long
    Dim parts As Variant
    Dim hourValue As Long
    Dim minuteValue As Long

    On Error GoTo NotTime
    parts = Split(Trim$(timeText), ":")
    If UBound(parts) < 1 Then GoTo NotTime
    hourValue = CLng(parts(0))
    minuteValue = CLng(parts(1))
    If hourValue < 0 Or hourValue > 23 Or minuteValue < 0 Or minuteValue > 59 Then GoTo NotTime
    TimeTextMinutes = hourValue * 60 + minuteValue
    Exit Function

NotTime:
    TimeTextMinutes = -1
End Function

Private Function TimeMinutesInRange(ByVal valueMinutes As Long, ByVal startMinutes As Long, ByVal endMinutes As Long) As Boolean
    If startMinutes <= endMinutes Then
        TimeMinutesInRange = valueMinutes >= startMinutes And valueMinutes <= endMinutes
    Else
        TimeMinutesInRange = valueMinutes >= startMinutes Or valueMinutes <= endMinutes
    End If
End Function

Private Function NormalizedHeaders(ByVal dataRange As Range) As Variant
    Dim headers() As String
    Dim counts As Object
    Dim colIndex As Long
    Dim rawHeader As String
    Dim finalHeader As String

    Set counts = CreateObject("Scripting.Dictionary")
    ReDim headers(1 To dataRange.Columns.Count)

    For colIndex = 1 To dataRange.Columns.Count
        rawHeader = Trim$(CStr(dataRange.Cells(1, colIndex).Value))
        If rawHeader = "" Then rawHeader = "Column " & ColumnLetter(colIndex)

        If counts.Exists(rawHeader) Then
            counts(rawHeader) = CLng(counts(rawHeader)) + 1
            finalHeader = rawHeader & " (" & CStr(counts(rawHeader)) & ")"
        Else
            counts.Add rawHeader, 1
            finalHeader = rawHeader
        End If

        headers(colIndex) = finalHeader
    Next colIndex

    NormalizedHeaders = headers
End Function

Private Function SplitList(ByVal textValue As String) As Variant
    Dim rawItems As Variant
    Dim cleaned() As String
    Dim item As Variant
    Dim count As Long
    Dim normalizedText As String

    normalizedText = Replace(CStr(textValue), vbCrLf, ",")
    normalizedText = Replace(normalizedText, vbCr, ",")
    normalizedText = Replace(normalizedText, vbLf, ",")

    rawItems = Split(normalizedText, ",")
    ReDim cleaned(0 To 0)
    count = -1

    For Each item In rawItems
        If Trim$(CStr(item)) <> "" Then
            count = count + 1
            ReDim Preserve cleaned(0 To count)
            cleaned(count) = Trim$(CStr(item))
        End If
    Next item

    If count = -1 Then
        SplitList = Array()
    Else
        SplitList = cleaned
    End If
End Function

Private Function SplitGroupRulesList(ByVal textValue As String) As Variant
    Dim normalizedText As String
    Dim rawItems As Variant
    Dim cleaned() As String
    Dim item As Variant
    Dim count As Long

    normalizedText = Replace(CStr(textValue), vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)

    If InStr(1, normalizedText, vbLf, vbBinaryCompare) > 0 Then
        rawItems = Split(normalizedText, vbLf)
    Else
        rawItems = Split(normalizedText, ",")
    End If

    ReDim cleaned(0 To 0)
    count = -1

    For Each item In rawItems
        If Trim$(CStr(item)) <> "" Then
            count = count + 1
            ReDim Preserve cleaned(0 To count)
            cleaned(count) = Trim$(CStr(item))
        End If
    Next item

    If count = -1 Then
        SplitGroupRulesList = Array()
    Else
        SplitGroupRulesList = cleaned
    End If
End Function

Private Function UBoundSafe(ByVal values As Variant) As Long
    On Error GoTo EmptyArray
    UBoundSafe = UBound(values)
    Exit Function
EmptyArray:
    UBoundSafe = -1
End Function

Private Function FieldExists(ByVal headers As Variant, ByVal fieldName As String) As Boolean
    Dim index As Long
    For index = LBound(headers) To UBound(headers)
        If CStr(headers(index)) = fieldName Then
            FieldExists = True
            Exit Function
        End If
    Next index
    FieldExists = False
End Function

Private Function HeaderIndex(ByVal headers As Variant, ByVal fieldName As String) As Long
    Dim index As Long
    For index = LBound(headers) To UBound(headers)
        If CStr(headers(index)) = fieldName Then
            HeaderIndex = index
            Exit Function
        End If
    Next index
End Function

Private Function FieldSuggestionArray(ByVal setupWs As Worksheet) As Variant
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim fields() As String
    Dim count As Long

    lastRow = setupWs.Cells(setupWs.Rows.Count, "N").End(xlUp).Row
    If lastRow < 4 Then Exit Function

    count = -1
    For rowIndex = 4 To lastRow
        If Trim$(CStr(setupWs.Cells(rowIndex, "N").Value)) <> "" Then
            count = count + 1
            ReDim Preserve fields(0 To count)
            fields(count) = Trim$(CStr(setupWs.Cells(rowIndex, "N").Value))
        End If
    Next rowIndex

    If count >= 0 Then FieldSuggestionArray = fields
End Function

Private Function IsInList(ByVal values As Variant, ByVal textValue As String) As Boolean
    Dim item As Variant
    On Error GoTo Done
    For Each item In values
        If StrComp(Trim$(CStr(item)), Trim$(textValue), vbTextCompare) = 0 Then
            IsInList = True
            Exit Function
        End If
    Next item
Done:
End Function

Private Sub DeleteShapesByPrefix(ByVal ws As Worksheet, ByVal prefixText As String)
    Dim index As Long
    For index = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(index).Name, Len(prefixText)) = prefixText Then
            ws.Shapes(index).Delete
        End If
    Next index
End Sub

Private Function FirstOutputSheetName(ByVal setupWs As Worksheet, ByVal templateName As String, ByVal lastSetupRow As Long) As String
    Dim rowIndex As Long
    For rowIndex = 9 To lastSetupRow
        If IsBuildSetupRow(setupWs, rowIndex, templateName) Then
            FirstOutputSheetName = Trim$(CStr(setupWs.Cells(rowIndex, "C").Value))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function IsBuildSetupRow(ByVal setupWs As Worksheet, ByVal rowIndex As Long, ByVal templateName As String) As Boolean
    IsBuildSetupRow = StrComp(Trim$(CStr(setupWs.Cells(rowIndex, "A").Value)), templateName, vbTextCompare) = 0 _
        And Trim$(CStr(setupWs.Cells(rowIndex, "B").Value)) <> ""
End Function

Private Function TemplateExists(ByVal setupWs As Worksheet, ByVal templateName As String) As Boolean
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = LastTemplateRow(setupWs)
    For rowIndex = 9 To lastRow
        If StrComp(Trim$(CStr(setupWs.Cells(rowIndex, "A").Value)), templateName, vbTextCompare) = 0 Then
            TemplateExists = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function LastTemplateRow(ByVal setupWs As Worksheet) As Long
    Dim lastRow As Long
    lastRow = Application.Max( _
        setupWs.Cells(setupWs.Rows.Count, "A").End(xlUp).Row, _
        setupWs.Cells(setupWs.Rows.Count, "B").End(xlUp).Row, _
        setupWs.Cells(setupWs.Rows.Count, "C").End(xlUp).Row _
    )
    If lastRow < 9 Then lastRow = 9
    LastTemplateRow = lastRow
End Function

Private Function FirstSaveAsFlag(ByVal setupWs As Worksheet, ByVal templateName As String, ByVal lastSetupRow As Long) As Boolean
    Dim rowIndex As Long
    For rowIndex = 9 To lastSetupRow
        If IsBuildSetupRow(setupWs, rowIndex, templateName) Then
            FirstSaveAsFlag = NormalizedSaveBehavior(CStr(setupWs.Cells(rowIndex, "D").Value)) <> "EXPORT CSV"
            Exit Function
        End If
    Next rowIndex
    FirstSaveAsFlag = True
End Function

Private Function FirstVisibleDataSheetName(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible _
            And ws.Name <> SETUP_SHEET _
            And ws.Name <> FIELD_PICKER_SHEET _
            And ws.Name <> CONDITION_BUILDER_SHEET _
            And Left$(ws.Name, Len(HELPER_PREFIX)) <> HELPER_PREFIX Then
            FirstVisibleDataSheetName = ws.Name
            Exit Function
        End If
    Next ws
    FirstVisibleDataSheetName = ""
End Function

Private Function VisibleSheetList(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim result As String
    Dim separator As String

    separator = ""
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible _
            And ws.Name <> SETUP_SHEET _
            And ws.Name <> FIELD_PICKER_SHEET _
            And ws.Name <> CONDITION_BUILDER_SHEET _
            And Left$(ws.Name, Len(HELPER_PREFIX)) <> HELPER_PREFIX Then
            result = result & separator & ws.Name
            separator = ","
        End If
    Next ws

    VisibleSheetList = result
End Function

Private Sub AddListValidation(ByVal targetRange As Range, ByVal listText As String)
    targetRange.Validation.Delete
    If listText <> "" And Len(listText) <= 255 Then
        targetRange.Validation.Add Type:=xlValidateList, Formula1:=listText
    End If
End Sub

Private Sub AddDataSheetDropdown(ByVal setupWs As Worksheet, ByVal targetAddress As String)
    Dim ws As Worksheet
    Dim names As String
    Dim separator As String

    separator = ""
    For Each ws In setupWs.Parent.Worksheets
        If ws.Visible = xlSheetVisible _
            And ws.Name <> SETUP_SHEET _
            And ws.Name <> FIELD_PICKER_SHEET _
            And ws.Name <> CONDITION_BUILDER_SHEET _
            And Left$(ws.Name, Len(HELPER_PREFIX)) <> HELPER_PREFIX Then
            names = names & separator & ws.Name
            separator = ","
        End If
    Next ws

    AddListValidation setupWs.Range(targetAddress), names
End Sub

Private Sub AddTemplateDropdown(ByVal setupWs As Worksheet)
    setupWs.Range("B5").Validation.Delete
    setupWs.Range("B5").Validation.Add Type:=xlValidateList, Formula1:="=$A$9:$A$200"
    setupWs.Range("A9:A200").Validation.Delete
    setupWs.Range("A9:A200").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:="=$A$9:$A$200"
    setupWs.Range("A9:A200").Validation.ShowError = False
End Sub

Private Sub AddFieldValidations(ByVal setupWs As Worksheet)
    On Error Resume Next
    AddSuggestionValidation setupWs.Range("E9:E200"), "=$N$4:$N$500"
    AddSuggestionValidation setupWs.Range("G9:K200"), "=$N$4:$N$500"
    AddSuggestionValidation setupWs.Range("L9:L200"), "=$B$9:$B$200"
    setupWs.Range("F9:F200").Validation.Delete
    On Error GoTo 0
End Sub

Private Sub AddSuggestionValidation(ByVal targetRange As Range, ByVal formulaText As String)
    targetRange.Validation.Delete
    targetRange.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=formulaText
    targetRange.Validation.IgnoreBlank = True
    targetRange.Validation.InCellDropdown = True
    targetRange.Validation.ShowError = False
End Sub

Private Sub PopulateFieldSuggestions(ByVal setupWs As Worksheet, ByVal headers As Variant)
    Dim index As Long
    setupWs.Range("N4:N500").ClearContents
    For index = LBound(headers) To UBound(headers)
        setupWs.Cells(3 + index, "N").Value = CStr(headers(index))
    Next index
    AddFieldValidations setupWs
End Sub

Private Sub ApplyBrandLayout(ByVal ws As Worksheet)
    With ws
        .Cells.Font.Name = "Aptos"
        .Cells.Font.Size = 10
        .Cells.Interior.Color = RGB(242, 244, 247)

        .Range("A1:M1").Interior.Color = RGB(141, 2, 31)
        .Range("A1:M1").Font.Color = RGB(255, 255, 255)
        .Range("A1:M1").Font.Bold = True
        .Range("A1:M1").Font.Size = 16

        .Range("A3:D5").Interior.Color = RGB(255, 255, 255)
        .Range("A3:A6").Font.Bold = True

        .Range("A8:M8").Interior.Color = RGB(0, 0, 0)
        .Range("A8:M8").Font.Color = RGB(255, 255, 255)
        .Range("A8:M8").Font.Bold = True
        .Range("A9:M200").Interior.Color = RGB(255, 255, 255)
        .Range("A8:M200").Borders.Color = RGB(200, 205, 212)

        .Range("N1:Q1").Interior.Color = RGB(141, 2, 31)
        .Range("N1:Q1").Font.Color = RGB(255, 255, 255)
        .Range("N1:Q1").Font.Bold = True
        .Range("N3:O3").Interior.Color = RGB(0, 0, 0)
        .Range("N3:O3").Font.Color = RGB(255, 255, 255)
        .Range("N3:O3").Font.Bold = True
        .Range("N4:N500").Interior.Color = RGB(255, 255, 255)

        .Columns("A").ColumnWidth = 18
        .Columns("B").ColumnWidth = 22
        .Columns("C").ColumnWidth = 18
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 24
        .Columns("F").ColumnWidth = 22
        .Columns("G").ColumnWidth = 42
        .Columns("H").ColumnWidth = 16
        .Columns("I").ColumnWidth = 16
        .Columns("J").ColumnWidth = 16
        .Columns("K").ColumnWidth = 26
        .Columns("L").ColumnWidth = 18
        .Columns("M").ColumnWidth = 38
        .Columns("N:Q").ColumnWidth = 22
        .Range("A8:M200").WrapText = False
        .Range("E9:G200").WrapText = True
        .Range("K9:K200").WrapText = True
        .Rows("1:1").RowHeight = 28
        .Rows("3:6").RowHeight = 24
        .Rows("8:8").RowHeight = 22
        .Rows("9:200").RowHeight = 54
    End With
End Sub

Private Sub AddSetupComments(ByVal ws As Worksheet)
    ClearComment ws.Range("B3")
    ClearComment ws.Range("B4")
    ClearComment ws.Range("B5")
    ClearComment ws.Range("B8")
    ClearComment ws.Range("C8")
    ClearComment ws.Range("E8")
    ClearComment ws.Range("F8")
    ClearComment ws.Range("G8")
    ClearComment ws.Range("H8")
    ClearComment ws.Range("I8")
    ClearComment ws.Range("J8")
    ClearComment ws.Range("K8")
    ClearComment ws.Range("L8")

    AddCommentText ws.Range("B3"), "Click the small Choose button beside this box. This is the workbook where PivotTables will be created and saved."
    AddCommentText ws.Range("B4"), "Choose the source sheet inside the selected input workbook. Field suggestions refresh after this changes."
    AddCommentText ws.Range("B5"), "Choose which template name to run. Use New Template to copy the current template, then Save Template after editing."
    AddCommentText ws.Range("B8"), "Name this setup row/PivotTable, such as Open East Calls. Blank Pivot Name rows are skipped."
    AddCommentText ws.Range("E8"), "Rows group the data. Choose one or more source fields from the dropdown."
    AddCommentText ws.Range("F8"), "Optional friendly name for the row group, such as Time Period or id Group."
    AddCommentText ws.Range("G8"), "Optional grouping rules. For multiple row groups, put one rule set per line in the same order as Rows. Example line: Small:id<=20|Middle:id>=21;id<=50|Large:id>=51."
    AddCommentText ws.Range("H8"), "Display Values are shown as labels instead of summed. Dropdown choices are added to the current list."
    AddCommentText ws.Range("I8"), "Values To Count/Sum are calculated. Numeric fields are summed; text fields are counted."
    AddCommentText ws.Range("J8"), "Filters become PivotTable filter dropdowns at the top of the pivot."
    AddCommentText ws.Range("K8"), "Optional filter values. Choose a field from the dropdown; it adds Field=. Then type the value after the equals sign."
    AddCommentText ws.Range("L8"), "Used only when Save Behavior is Export CSV. If Save Behavior is Save to input file, this cell is cleared and grayed out."
End Sub

Private Sub ClearComment(ByVal target As Range)
    On Error Resume Next
    target.ClearComments
    On Error GoTo 0
End Sub

Private Sub AddCommentText(ByVal target As Range, ByVal textValue As String)
    On Error Resume Next
    target.AddComment textValue
    On Error GoTo 0
End Sub

Private Sub AddSetupButtons(ByVal setupWs As Worksheet)
    Dim shape As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim buildLeft As Double
    Dim buildTop As Double
    Const buttonWidth As Double = 92
    Const buttonHeight As Double = 20

    For Each shape In setupWs.Shapes
        If Left$(shape.Name, 13) = "PivotBuilder" Then
            shape.Delete
        End If
    Next shape

    leftPos = setupWs.Range("E3").Left
    topPos = setupWs.Range("E3").Top

    Set shape = setupWs.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, buttonWidth, buttonHeight)
    With shape
        .Name = "PivotBuilderChooseSourceButton"
        .TextFrame.Characters.Text = "Choose"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 8
        .Fill.ForeColor.RGB = RGB(141, 2, 31)
        .Line.ForeColor.RGB = RGB(141, 2, 31)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "ChooseSourceWorkbook"
    End With

    Set shape = setupWs.Shapes.AddShape(msoShapeRoundedRectangle, setupWs.Range("E5").Left, setupWs.Range("E5").Top, buttonWidth, buttonHeight)
    With shape
        .Name = "PivotBuilderSaveTemplateButton"
        .TextFrame.Characters.Text = "Save Template"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 8
        .Fill.ForeColor.RGB = RGB(141, 2, 31)
        .Line.ForeColor.RGB = RGB(141, 2, 31)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "SaveCurrentTemplate"
    End With

    Set shape = setupWs.Shapes.AddShape(msoShapeRoundedRectangle, setupWs.Range("E5").Left + 100, setupWs.Range("E5").Top, buttonWidth, buttonHeight)
    With shape
        .Name = "PivotBuilderNewTemplateButton"
        .TextFrame.Characters.Text = "New Template"
        .TextFrame.Characters.Font.Size = 8
        .Fill.ForeColor.RGB = RGB(242, 242, 242)
        .Line.ForeColor.RGB = RGB(166, 166, 166)
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .OnAction = "NewTemplateFromCurrent"
    End With

    buildLeft = setupWs.Range("E6").Left
    buildTop = setupWs.Range("E6").Top

    Set shape = setupWs.Shapes.AddShape(msoShapeRoundedRectangle, buildLeft, buildTop, buttonWidth, buttonHeight)
    With shape
        .Name = "PivotBuilderBuildButton"
        .TextFrame.Characters.Text = "Build Pivots"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Size = 9
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame.Characters.Font.Color = RGB(255, 255, 255)
        .OnAction = "BuildPivotTablesFromSetup"
    End With

    Set shape = setupWs.Shapes.AddShape(msoShapeRoundedRectangle, buildLeft + 100, buildTop, buttonWidth, buttonHeight)
    With shape
        .Name = "PivotBuilderResetButton"
        .TextFrame.Characters.Text = "Refresh Layout"
        .TextFrame.Characters.Font.Size = 9
        .Fill.ForeColor.RGB = RGB(217, 221, 226)
        .Line.ForeColor.RGB = RGB(166, 172, 181)
        .TextFrame.Characters.Font.Color = RGB(0, 0, 0)
        .OnAction = "SetupPivotBuilderSheet"
    End With

End Sub

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If WorksheetExists(wb, sheetName) Then
        Set GetOrCreateSheet = wb.Worksheets(sheetName)
    Else
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function UniqueSheetName(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim root As String
    Dim candidate As String
    Dim index As Long

    root = SafeSheetName(baseName)
    candidate = root

    If Not WorksheetExists(wb, candidate) Then
        UniqueSheetName = candidate
        Exit Function
    End If

    For index = 2 To 999
        candidate = Left$(root, 31 - Len("_" & CStr(index))) & "_" & CStr(index)
        If Not WorksheetExists(wb, candidate) Then
            UniqueSheetName = candidate
            Exit Function
        End If
    Next index

    Err.Raise vbObjectError + 200, , "Could not create a unique sheet name."
End Function

Private Function SafeSheetName(ByVal sheetName As String) As String
    Dim result As String
    result = Trim$(sheetName)
    If result = "" Then result = "Pivot_Output"

    result = Replace(result, "[", "_")
    result = Replace(result, "]", "_")
    result = Replace(result, "*", "_")
    result = Replace(result, ":", "_")
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    result = Replace(result, "?", "_")

    SafeSheetName = Left$(result, 31)
End Function

Private Function SafePivotName(ByVal pivotName As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String

    result = ""
    For i = 1 To Len(pivotName)
        ch = Mid$(pivotName, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            result = result & ch
        Else
            result = result & "_"
        End If
    Next i

    If result = "" Then result = "PivotTable"
    If Mid$(result, 1, 1) Like "[0-9]" Then result = "P_" & result
    SafePivotName = Left$(result, 240)
End Function

Private Function ColumnLetter(ByVal columnNumber As Long) As String
    Dim result As String
    Dim n As Long
    Dim remainder As Long

    n = columnNumber
    Do While n > 0
        remainder = (n - 1) Mod 26
        result = Chr$(65 + remainder) & result
        n = (n - remainder - 1) \ 26
    Loop

    ColumnLetter = result
End Function

Private Function DefaultOutputFileName(ByVal wb As Workbook) As String
    Dim basePath As String
    Dim baseName As String
    Dim dotPos As Long

    basePath = wb.Path
    baseName = wb.Name
    dotPos = InStrRev(baseName, ".")
    If dotPos > 0 Then baseName = Left$(baseName, dotPos - 1)

    If basePath = "" Then
        DefaultOutputFileName = baseName & "_with_pivots.xlsx"
    Else
        DefaultOutputFileName = basePath & Application.PathSeparator & baseName & "_with_pivots.xlsx"
    End If
End Function

Private Sub SaveWorkbookAsPath(ByVal wb As Workbook, ByVal outputPath As String)
    Dim ext As String
    ext = LCase$(Mid$(outputPath, InStrRev(outputPath, ".") + 1))

    Select Case ext
        Case "xlsm"
            wb.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Case "xls"
            wb.SaveAs Filename:=outputPath, FileFormat:=xlExcel8
        Case Else
            wb.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    End Select
End Sub
