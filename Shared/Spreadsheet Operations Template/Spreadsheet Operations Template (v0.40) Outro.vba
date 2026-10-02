' ************ '
' Alteration   '
' ************ '

Sub ApplyFormulaToCellOnWorksheet(ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal cellColumnLetter As String, ByVal cellRowNumber As Long, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFormulaToCellOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal cellColumnLetter As String, ByVal cellRowNumber As Long, ByVal worksheetName As String", methodName, "Alteration")

        Call ConfigureMethodSetting(methodName, "Keep Formula Active", 0, 0, 1)
    End If

    Dim validation As String
    Call ValidateNumeric(Array(cellRowNumber), "Long", Array(CLng(1)), Array(CLng(1048576)), "cellRowNumber", validation)
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaToApply & """" & ", " & IIf(useLocalFormula, "TRUE", "FALSE") & ", " & """" & cellColumnLetter & """" & ", " & cellRowNumber & ", " & """" & worksheetName & """", validation)

    Dim settings As Object
    Dim worksheet As Worksheet
    Dim cellColumnNumber As Long
    Dim cellRange As Range

    Set settings = methodRegistry(methodName)("Settings")
    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    cellColumnNumber = ConvertColumnLetterToColumnNumber(cellColumnLetter)
    Set cellRange = worksheet.Cells(cellRowNumber, cellColumnNumber)

    If CellStyleExists("Formula") = False Then
        Call LogConclusion("Failed", logConclusionData, "Cell style Formula not found.")
    End If

    cellRange.Style = "Formula"

    On Error GoTo ApplyFormulaToCellOnWorksheetError
    If useLocalFormula Then
        cellRange.FormulaLocal = formulaToApply
    Else
        cellRange.Formula = formulaToApply
    End If
    On Error GoTo 0

    If Application.Calculation <> xlCalculationAutomatic Then
        cellRange.Calculate
    End If

    If settings("Keep Formula Active")("Value") = False Then
        cellRange.Value = cellRange.Value
    End If

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
ApplyFormulaToCellOnWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub ApplyFormulaToColumnOnWorksheet(ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFormulaToColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal columnName As String, ByVal worksheetName As String", methodName, "Alteration")

        Call ConfigureMethodSetting(methodName, "Keep Formula Active", 0, 0, 1)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaToApply & """" & ", " & IIf(useLocalFormula, "TRUE", "FALSE") & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim settings As Object
    Dim worksheet As Worksheet
    Dim lastUsedRowNumber As Long
    Dim columnLetter As String
    Dim columnDataRange As Range

    If CellStyleExists("Formula") = False Then
        Call LogConclusion("Failed", logConclusionData, "Cell style Formula not found.")
    End If

    Set settings = methodRegistry(methodName)("Settings")
    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)

    If lastUsedRowNumber = 1 Then
        Call LogConclusion("Failed", logConclusionData, "No data available in worksheet.")
    End If

    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    Set columnDataRange = worksheet.Range(columnLetter & "2:" & columnLetter & LastUsedRowNumberOnWorksheet(worksheetName))
    columnDataRange.Style = "Formula"

    On Error GoTo ApplyFormulaToColumnOnWorksheetError
    If useLocalFormula Then
        columnDataRange.FormulaLocal = formulaToApply
    Else
        columnDataRange.Formula = formulaToApply
    End If
    On Error GoTo 0

    If Application.Calculation <> xlCalculationAutomatic Then
        columnDataRange.Calculate
    End If

    If settings("Keep Formula Active")("Value") = False Then
        columnDataRange.Value = columnDataRange.Value
    End If

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
ApplyFormulaToColumnOnWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub FindAndReplaceInColumnOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal columnName As String, ByVal worksheetName As String, Optional ByVal exactMatch As Boolean) ' Repeat Support: columnName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "FindAndReplaceInColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim columnNames As Variant
    Dim columnNamesHasMultipleValues As Boolean

    If Len(columnName) <> 0 And InStr(columnName, "|") Then
        columnNames = ParseMethodArgumentsIntoValues(columnName)
        If LBound(columnNames) < UBound(columnNames) Then columnNamesHasMultipleValues = True
    End If

    If columnNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal columnNames As String, ByVal worksheetName As String, ByVal exactMatch As Boolean", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """" & ", " & IIf(exactMatch, "TRUE", "FALSE"))

        Dim columnNameIndex As Long

        For columnNameIndex = LBound(columnNames) To UBound(columnNames)
            columnName = columnNames(columnNameIndex)

            Call FindAndReplaceInColumnOnWorksheet(findValue, replaceValue, columnName, worksheetName, exactMatch)
        Next columnNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal columnName As String, ByVal worksheetName As String, Optional ByVal exactMatch As Boolean", methodName, "Alteration")
    End If

    If Len(columnName) >= 3 And InStr(columnName, "|") And Left$(columnName, 1) = """" And Right$(columnName, 1) = """" Then
        columnName = Mid$(columnName, 2, Len(columnName) - 2)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """" & ", " & IIf(exactMatch, "TRUE", "FALSE"), validation)

    Dim worksheet As Worksheet
    Dim columnLetter As String
    Dim columnDataRange As Range

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    Set columnDataRange = worksheet.Range(columnLetter & "2:" & columnLetter & LastUsedRowNumberOnWorksheet(worksheetName))

    Dim lookAtMode As Long
    If exactMatch = True Then
        lookAtMode = xlWhole
    Else
        lookAtMode = xlPart
    End If

    Call columnDataRange.Replace(What:=findValue, Replacement:=replaceValue, LookAt:=lookAtMode, SearchOrder:=xlByRows, MatchCase:=False, MatchByte:=False, SearchFormat:=False, ReplaceFormat:=False)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub MoveColumnBesideColumnOnWorksheet(ByVal sourceColumnName As String, ByVal sideToInsertOn As String, ByVal anchorColumnName As String, ByVal worksheetName As String) ' Repeat Support: sourceColumnName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "MoveColumnBesideColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim sourceColumnNames As Variant
    Dim sourceColumnNamesHasMultipleValues As Boolean

    If Len(sourceColumnName) <> 0 And InStr(sourceColumnName, "|") Then
        sourceColumnNames = ParseMethodArgumentsIntoValues(sourceColumnName)
        If LBound(sourceColumnNames) < UBound(sourceColumnNames) Then sourceColumnNamesHasMultipleValues = True
    End If

    If sourceColumnNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal sourceColumnNames As String, ByVal sideToInsertOn As String, ByVal anchorColumnName As String, ByVal worksheetName As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & sourceColumnName & """" & ", " & """" & sideToInsertOn & """" & ", " & """" & anchorColumnName & """" & ", " & """" & worksheetName & """")

        Dim sourceColumnNameIndex As Long

        For sourceColumnNameIndex = LBound(sourceColumnNames) To UBound(sourceColumnNames)
            sourceColumnName = sourceColumnNames(sourceColumnNameIndex)

            Call MoveColumnBesideColumnOnWorksheet(sourceColumnName, sideToInsertOn, anchorColumnName, worksheetName)
        Next sourceColumnNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal sourceColumnName As String, ByVal sideToInsertOn As String, ByVal anchorColumnName As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    If Len(sourceColumnName) >= 3 And InStr(sourceColumnName, "|") And Left$(sourceColumnName, 1) = """" And Right$(sourceColumnName, 1) = """" Then
        sourceColumnName = Mid$(sourceColumnName, 2, Len(sourceColumnName) - 2)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(sourceColumnName, "sourceColumnName", worksheetName, "worksheetName", validation)
    Call ValidateWhitelist(sideToInsertOn, "sideToInsertOn", Array("Left", "Right"), validation)
    Call ValidateColumnOnWorksheet(anchorColumnName, "anchorColumnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & sourceColumnName & """" & ", " & """" & sideToInsertOn & """" & ", " & """" & anchorColumnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim sourceColumnLetter As String
    Dim anchorColumnLetter As String
    Dim sourceColumnNumber As Long
    Dim anchorColumnNumber As Long
    Dim lastColumnLetter As String
    Dim numberOfColumnsInUse As Long
    Dim worksheetVisibility As Long

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    sourceColumnLetter = FindColumnLetterOnWorksheet(sourceColumnName, worksheetName)
    anchorColumnLetter = FindColumnLetterOnWorksheet(anchorColumnName, worksheetName)
    sourceColumnNumber = ConvertColumnLetterToColumnNumber(sourceColumnLetter)
    anchorColumnNumber = ConvertColumnLetterToColumnNumber(anchorColumnLetter)
    lastColumnLetter = LastUsedColumnLetterOnWorksheet(worksheetName)
    numberOfColumnsInUse = ConvertColumnLetterToColumnNumber(lastColumnLetter)

    If numberOfColumnsInUse = 1 Then
        Call LogConclusion("Failed", logConclusionData, "Only one column available, unable to move it relative to anything else.")
    End If

    If sourceColumnLetter = anchorColumnLetter Then
        Call LogConclusion("Failed", logConclusionData, "Source column and anchor column are the same column.")
    End If

    If (sideToInsertOn = "Left" And sourceColumnNumber = anchorColumnNumber - 1) Or (sideToInsertOn = "Right" And sourceColumnNumber = anchorColumnNumber + 1) Then
        Call LogConclusion("Skipped", logConclusionData)

        Exit Sub
    End If

    Select Case worksheet.Visible
        Case xlSheetHidden, xlSheetVeryHidden
            worksheetVisibility = worksheet.Visible
        Case Else
            worksheetVisibility = -1
    End Select

    If worksheetVisibility <> -1 Then
        worksheet.Visible = xlSheetVisible
    End If

    worksheet.Select

    On Error GoTo MoveColumnBesideColumnOnWorksheetError
    Call worksheet.Columns(sourceColumnLetter).Select
    Call Selection.Cut
    On Error GoTo 0

    If sideToInsertOn = "Right" Then
        anchorColumnNumber = anchorColumnNumber + 1
        anchorColumnLetter = ConvertColumnNumberToColumnLetter(anchorColumnNumber)
    End If

    On Error GoTo MoveColumnBesideColumnOnWorksheetError
    Call worksheet.Columns(anchorColumnLetter).Select
    Call Selection.Insert(Shift:=xlToRight)
    Call worksheet.Range("A1").Select
    On Error GoTo 0

    sourceColumnLetter = FindColumnLetterOnWorksheet(sourceColumnName, worksheetName)
    If sourceColumnLetter = lastColumnLetter Or sourceColumnLetter = "A" Then
        If worksheet.AutoFilterMode = True Then
            worksheet.AutoFilterMode = False

            worksheet.Range("A1:" & lastColumnLetter & "1").AutoFilter
        End If
    End If

    If worksheetVisibility <> -1 Then
        worksheet.Visible = worksheetVisibility
    End If

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
MoveColumnBesideColumnOnWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub MoveWorksheetToEnd(ByVal worksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "MoveWorksheetToEnd"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & worksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call MoveWorksheetToEnd(worksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Alteration")
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim worksheetVisibility As Long

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    Select Case worksheet.Visible
        Case xlSheetHidden, xlSheetVeryHidden
            worksheetVisibility = worksheet.Visible
        Case Else
            worksheetVisibility = -1
    End Select

    If worksheetVisibility <> -1 Then
        worksheet.Visible = xlSheetVisible
    End If

    worksheet.Move After:=Sheets(Worksheets.Count)

    If worksheetVisibility <> -1 Then
        worksheet.Visible = worksheetVisibility
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RenameWorksheet(ByVal currentWorksheetName As String, ByVal newWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RenameWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal newWorksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateWorksheet(currentWorksheetName, "currentWorksheetName", validation)
    Call ValidateWorksheet(newWorksheetName, "newWorksheetName", validation, True)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & newWorksheetName & """", validation)

    Dim currentWorksheet As Worksheet

    Set currentWorksheet = mainWorkbook.Worksheets(currentWorksheetName)

    On Error GoTo RenameWorksheetError
    currentWorksheet.Name = newWorksheetName
    On Error GoTo 0

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
RenameWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub SetHeaderOnColumnOnWorksheet(ByVal columnHeader As String, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetHeaderOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnHeader As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnHeader, "columnHeader", worksheetName, "worksheetName", validation, True)
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnHeader & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim columnLetter As String

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)

    worksheet.Range(columnLetter & "1").Value = columnHeader

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetMethodSetting(ByVal settingMethod As String, ByVal settingName As String, ByVal settingValue As Long)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetMethodSetting"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal settingMethod As String, ByVal settingName As String, ByVal settingValue As Long", methodName, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & settingMethod & """" & ", " & """" & settingName & """" & ", " & settingValue)

    Call ConfigureMethodSetting(settingMethod, settingName, settingValue)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetVisibilityOfWorksheet(ByVal visibilityState As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetVisibilityOfWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal visibilityState As String, ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & visibilityState & """" & ", " & """" & worksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call SetVisibilityOfWorksheet(visibilityState, worksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal visibilityState As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateWhitelist(visibilityState, "visibilityState", Array("Hidden", "Very Hidden", "Visible"), validation)
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & visibilityState & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim alreadyMatchesRequestedVisibility As Boolean
    Dim currentSheet As Object
    Dim visibleSheetCount As Long

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    Select Case visibilityState
        Case "Hidden"
            alreadyMatchesRequestedVisibility = (worksheet.Visible = xlSheetHidden)
        Case "Very Hidden"
            alreadyMatchesRequestedVisibility = (worksheet.Visible = xlSheetVeryHidden)
        Case "Visible"
            alreadyMatchesRequestedVisibility = (worksheet.Visible = xlSheetVisible)
    End Select

    If alreadyMatchesRequestedVisibility = True Then
        Call LogConclusion("Skipped", logConclusionData)
        Exit Sub
    End If

    If visibilityState = "Hidden" Or visibilityState = "Very Hidden" Then
        For Each currentSheet In mainWorkbook.Sheets
            If currentSheet.Visible = xlSheetVisible Then
                visibleSheetCount = visibleSheetCount + 1
            End If
        Next currentSheet

        If worksheet.Visible = xlSheetVisible And visibleSheetCount = 1 Then
            Call LogConclusion("Failed", logConclusionData, "A workbook must contain at least one visible sheet.")
        End If
    End If

    Select Case visibilityState
        Case "Hidden"
            worksheet.Visible = xlSheetHidden
        Case "Very Hidden"
            worksheet.Visible = xlSheetVeryHidden
        Case "Visible"
            worksheet.Visible = xlSheetVisible
    End Select

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetWidthOnColumnOnWorksheet(ByVal columnWidth As Double, ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetWidthOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal columnWidth As Double, ByVal columnName As String, ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, columnWidth & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call SetWidthOnColumnOnWorksheet(columnWidth, columnName, worksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnWidth As Double, ByVal columnName As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateNumeric(Array(columnWidth), "Double", Array(CDbl(0)), Array(CDbl(255)), "columnWidth", validation)
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, columnWidth & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim columnLetter As String

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)

    On Error GoTo SetWidthOnColumnOnWorksheetError
    worksheet.Columns(columnLetter).ColumnWidth = columnWidth
    On Error GoTo 0

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
SetWidthOnColumnOnWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub SortColumnByOrderOnWorksheet(ByVal columnName As String, ByVal sortOrder As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName + worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SortColumnByOrderOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim columnNames As Variant
    Dim worksheetNames As Variant
    Dim columnNamesHasMultipleValues As Boolean
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(columnName) <> 0 And InStr(columnName, "|") Then
        columnNames = ParseMethodArgumentsIntoValues(columnName)
        If LBound(columnNames) < UBound(columnNames) Then columnNamesHasMultipleValues = True
    End If

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If columnNamesHasMultipleValues = True Or worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal columnNames As String, ByVal sortOrder As String, ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & sortOrder & """" & ", " & """" & worksheetName & """")

        Dim columnNameIndex As Long
        Dim worksheetNameIndex As Long

        If columnNamesHasMultipleValues = True And worksheetNamesHasMultipleValues = True Then
            For columnNameIndex = LBound(columnNames) To UBound(columnNames)
                For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
                    columnName = columnNames(columnNameIndex)
                    worksheetName = worksheetNames(worksheetNameIndex)

                    Call SortColumnByOrderOnWorksheet(columnName, sortOrder, worksheetName)
                Next worksheetNameIndex
            Next columnNameIndex
        ElseIf columnNamesHasMultipleValues = True And worksheetNamesHasMultipleValues = False Then
            For columnNameIndex = LBound(columnNames) To UBound(columnNames)
                columnName = columnNames(columnNameIndex)

                Call SortColumnByOrderOnWorksheet(columnName, sortOrder, worksheetName)
            Next columnNameIndex
        ElseIf columnNamesHasMultipleValues = False And worksheetNamesHasMultipleValues = True Then
            For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
                worksheetName = worksheetNames(worksheetNameIndex)

                Call SortColumnByOrderOnWorksheet(columnName, sortOrder, worksheetName)
            Next worksheetNameIndex
        End If

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal sortOrder As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    If Len(columnName) >= 3 And InStr(columnName, "|") And Left$(columnName, 1) = """" And Right$(columnName, 1) = """" Then
        columnName = Mid$(columnName, 2, Len(columnName) - 2)
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    Call ValidateWhitelist(sortOrder, "sortOrder", Array("Ascending", "Descending"), validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & sortOrder & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim columnLetter As String
    Dim keyRange As Range
    Dim sortRange As Range
    Dim sortXlOrder As XlSortOrder

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    Set keyRange = worksheet.Range(columnLetter & "2")
    Set sortRange = worksheet.Range("2:" & LastUsedRowNumberOnWorksheet(worksheetName))

    sortXlOrder = xlAscending
    If sortOrder = "Descending" Then
        sortXlOrder = xlDescending
    End If

    On Error GoTo SortColumnByOrderOnWorksheetError
    Call sortRange.Sort(Key1:=keyRange, Order1:=sortXlOrder, Header:=xlNo)
    On Error GoTo 0

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
SortColumnByOrderOnWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub SplitDelimitedTextOnWorksheet(ByVal columnDelimiter As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SplitDelimitedTextOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnDelimiter As String, ByVal worksheetName As String, Context columnsAfterSplit As Integer", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateRequiredText(columnDelimiter, "columnDelimiter", validation, 1, 1)
    Call ValidateWorksheet(worksheetName, "worksheetName", validation, False)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnDelimiter & """" & ", " & """" & worksheetName & """" & ", ", validation)

    Dim worksheet As Worksheet
    Dim lastUsedRowNumber As Long
    Dim sourceRange As Range
    Dim sourceValues As Variant
    Dim lineIndex As Long
    Dim lineText As String
    Dim textQualifier As String
    Dim lineLength As Long
    Dim delimiterCount As Long
    Dim maximumDelimiterCount As Long
    Dim sourceContainsQualifier As Boolean
    Dim lineParts As Variant
    Dim parsedValues As Variant
    Dim columnCount As Long
    Dim columnIndex As Long
    Dim destinationRange As Range

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If worksheet.Range("A1").Value = "" Then
        logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value = logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value & 0
        Call LogConclusion("Failed", logConclusionData, "No text to split found.")
    End If

    lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)

    Set sourceRange = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(lastUsedRowNumber, 1))
    sourceValues = sourceRange.Value

    textQualifier = """"

    If IsArray(sourceValues) = False Then
        ReDim sourceValues(1 To 1, 1 To 1)
        sourceValues(1, 1) = sourceRange.Value
    End If

    maximumDelimiterCount = 0
    sourceContainsQualifier = False
    For lineIndex = 1 To UBound(sourceValues, 1)
        lineText = CStr(sourceValues(lineIndex, 1) & "")

        If InStr(1, lineText, textQualifier, vbBinaryCompare) > 0 Then
            sourceContainsQualifier = True
            Exit For
        End If

        delimiterCount = Len(lineText) - Len(Replace(lineText, columnDelimiter, vbNullString))

        If delimiterCount > maximumDelimiterCount Then maximumDelimiterCount = delimiterCount
    Next lineIndex

    If sourceContainsQualifier = False Then
        columnCount = maximumDelimiterCount + 1

        ReDim parsedValues(1 To UBound(sourceValues, 1), 1 To columnCount)
        For lineIndex = 1 To UBound(sourceValues, 1)
            lineText = CStr(sourceValues(lineIndex, 1) & "")
            lineParts = Split(lineText, columnDelimiter)

            For columnIndex = LBound(lineParts) To UBound(lineParts)
                parsedValues(lineIndex, columnIndex + 1) = lineParts(columnIndex)
            Next columnIndex
        Next lineIndex
    Else
        Dim characterIndex As Long
        Dim scanIndex As Long
        Dim nextCharacter As String
        Dim qualifierCloseIndex As Long
        Dim innerText As String
        Dim currentField As String
        Dim usedQuotedField As Boolean
        Dim lineFields As Collection
        Dim parsedRows As Collection
        Dim fieldCount As Long
        Dim maximumFieldCount As Long
        Dim fieldIndex As Long

        Set parsedRows = New Collection
        maximumFieldCount = 1

        For lineIndex = 1 To UBound(sourceValues, 1)
            lineText = CStr(sourceValues(lineIndex, 1) & "")
            lineLength = Len(lineText)
            Set lineFields = New Collection

            If lineLength = 0 Then
                Call lineFields.Add("")
            Else
                characterIndex = 1
                Do While characterIndex <= lineLength
                    usedQuotedField = False
                    qualifierCloseIndex = 0

                    If Mid$(lineText, characterIndex, 1) = textQualifier Then
                        scanIndex = characterIndex + 1
                        Do While scanIndex <= lineLength
                            If Mid$(lineText, scanIndex, 1) = textQualifier Then
                                If scanIndex < lineLength Then
                                    nextCharacter = Mid$(lineText, scanIndex + 1, 1)
                                    If nextCharacter = textQualifier Then
                                        scanIndex = scanIndex + 2
                                    ElseIf nextCharacter = columnDelimiter Then
                                        qualifierCloseIndex = scanIndex
                                        Exit Do
                                    Else
                                        scanIndex = scanIndex + 1
                                    End If
                                Else
                                    qualifierCloseIndex = scanIndex
                                    Exit Do
                                End If
                            Else
                                scanIndex = scanIndex + 1
                            End If
                        Loop
                    End If

                    If qualifierCloseIndex > 0 Then
                        innerText = Mid$(lineText, characterIndex + 1, qualifierCloseIndex - characterIndex - 1)
                        If InStr(1, innerText, columnDelimiter, vbBinaryCompare) > 0 Then
                            currentField = Replace(innerText, textQualifier & textQualifier, textQualifier)
                        Else
                            currentField = Mid$(lineText, characterIndex, qualifierCloseIndex - characterIndex + 1)
                        End If
                        Call lineFields.Add(currentField)
                        characterIndex = qualifierCloseIndex + 1
                        usedQuotedField = True
                        If characterIndex <= lineLength Then
                            If Mid$(lineText, characterIndex, 1) = columnDelimiter Then
                                characterIndex = characterIndex + 1
                                If characterIndex > lineLength Then
                                    Call lineFields.Add("")
                                End If
                            End If
                        End If
                    End If

                    If usedQuotedField = False Then
                        currentField = ""
                        Do While characterIndex <= lineLength
                            If Mid$(lineText, characterIndex, 1) = columnDelimiter Then
                                Call lineFields.Add(currentField)
                                characterIndex = characterIndex + 1
                                If characterIndex > lineLength Then
                                    Call lineFields.Add("")
                                End If
                                Exit Do
                            End If
                            currentField = currentField & Mid$(lineText, characterIndex, 1)
                            characterIndex = characterIndex + 1
                            If characterIndex > lineLength Then
                                Call lineFields.Add(currentField)
                            End If
                        Loop
                    End If
                Loop
            End If

            fieldCount = lineFields.Count
            If fieldCount > maximumFieldCount Then maximumFieldCount = fieldCount
            Call parsedRows.Add(lineFields)
        Next lineIndex

        ReDim parsedValues(1 To parsedRows.Count, 1 To maximumFieldCount)
        For lineIndex = 1 To parsedRows.Count
            Set lineFields = parsedRows(lineIndex)
            For fieldIndex = 1 To lineFields.Count
                parsedValues(lineIndex, fieldIndex) = lineFields(fieldIndex)
            Next fieldIndex
        Next lineIndex
    End If

    On Error GoTo SplitDelimitedTextOnWorksheetError
    sourceRange.Clear
    Set destinationRange = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(UBound(parsedValues, 1), UBound(parsedValues, 2)))
    destinationRange.NumberFormat = "@"
    destinationRange.Value = parsedValues
    On Error GoTo 0

    Dim lastUsedColumnLetter
    Dim lastUsedColumnNumber

    lastUsedColumnLetter = LastUsedColumnLetterOnWorksheet(worksheetName)
    lastUsedColumnNumber = ConvertColumnLetterToColumnNumber(lastUsedColumnLetter)

    logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value = logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value & lastUsedColumnNumber

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
SplitDelimitedTextOnWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

' Functions: Alteration '

' ************ '
' Conjuration  '
' ************ '

Sub AddColumnOnWorksheet(ByVal columnName As String, ByVal columnWidth As Double, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName + worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "AddColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim columnNames As Variant
    Dim worksheetNames As Variant
    Dim columnNamesHasMultipleValues As Boolean
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(columnName) <> 0 And InStr(columnName, "|") Then
        columnNames = ParseMethodArgumentsIntoValues(columnName)
        If LBound(columnNames) < UBound(columnNames) Then columnNamesHasMultipleValues = True
    End If

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If columnNamesHasMultipleValues = True Or worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal columnNames As String, ByVal columnWidth As Double, ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & columnWidth & ", " & """" & worksheetName & """")

        Dim columnNameIndex As Long
        Dim worksheetNameIndex As Long

        If columnNamesHasMultipleValues = True And worksheetNamesHasMultipleValues = True Then
            For columnNameIndex = LBound(columnNames) To UBound(columnNames)
                For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
                    columnName = columnNames(columnNameIndex)
                    worksheetName = worksheetNames(worksheetNameIndex)

                    Call AddColumnOnWorksheet(columnName, columnWidth, worksheetName)
                Next worksheetNameIndex
            Next columnNameIndex
        ElseIf columnNamesHasMultipleValues = True And worksheetNamesHasMultipleValues = False Then
            For columnNameIndex = LBound(columnNames) To UBound(columnNames)
                columnName = columnNames(columnNameIndex)

                Call AddColumnOnWorksheet(columnName, columnWidth, worksheetName)
            Next columnNameIndex
        ElseIf columnNamesHasMultipleValues = False And worksheetNamesHasMultipleValues = True Then
            For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
                worksheetName = worksheetNames(worksheetNameIndex)

                Call AddColumnOnWorksheet(columnName, columnWidth, worksheetName)
            Next worksheetNameIndex
        End If

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal columnWidth As Double, ByVal worksheetName As String", methodName, "Conjuration")
    End If

    If Len(columnName) >= 3 And InStr(columnName, "|") And Left$(columnName, 1) = """" And Right$(columnName, 1) = """" Then
        columnName = Mid$(columnName, 2, Len(columnName) - 2)
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation, True)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & columnWidth & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim columnLetter As String
    Dim columnNumber As Long

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    columnLetter = LastUsedColumnLetterOnWorksheet(worksheetName)

    If worksheet.Range(columnLetter & "1").Value <> "" Then
        columnNumber = ConvertColumnLetterToColumnNumber(columnLetter)
        columnNumber = columnNumber + 1

        If columnNumber = 16385 Then
            Call LogConclusion("Failed", logConclusionData, "The worksheet already has the maximum of 16,384 columns. Another column can't be added.")
        End If

        columnLetter = ConvertColumnNumberToColumnLetter(columnNumber)
    End If

    If CellStyleExists("Header") = False Then
        Call LogConclusion("Failed", logConclusionData, "Cell style Header not found.")
    End If

    With worksheet.Range(columnLetter & "1")
        .ColumnWidth = columnWidth
        .Style = "Header"
        .Value = columnName
    End With

    If worksheet.AutoFilterMode = True Then
        worksheet.AutoFilterMode = False
    End If

    worksheet.Range("A1:" & columnLetter & "1").AutoFilter

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub CopyWorksheetAs(ByVal currentWorksheetName As String, ByVal newWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CopyWorksheetAs"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal newWorksheetName As String", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateWorksheet(currentWorksheetName, "currentWorksheetName", validation)
    Call ValidateWorksheet(newWorksheetName, "newWorksheetName", validation, True)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & newWorksheetName & """", validation)

    Dim currentWorksheet As Worksheet
    Dim lastWorksheetIndex As Long
    Dim newWorksheet As Worksheet

    Set currentWorksheet = mainWorkbook.Worksheets(currentWorksheetName)
    lastWorksheetIndex = mainWorkbook.Worksheets.Count

    On Error GoTo CopyWorksheetAsError
    Call currentWorksheet.Copy(After:=mainWorkbook.Worksheets(lastWorksheetIndex))
    Set newWorksheet = mainWorkbook.Worksheets(lastWorksheetIndex + 1)
    newWorksheet.Name = newWorksheetName
    On Error GoTo 0

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
CopyWorksheetAsError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub CreateWorksheet(ByVal worksheetName As String, Optional ByVal insertAfterWorksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CreateWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal worksheetNames As String, Optional ByVal insertAfterWorksheetName As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & worksheetName & """" & ", " & """" & insertAfterWorksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call CreateWorksheet(worksheetName, insertAfterWorksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String, Optional ByVal insertAfterWorksheetName As String", methodName, "Conjuration")
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation, True)
    If insertAfterWorksheetName <> "" Then Call ValidateWorksheet(insertAfterWorksheetName, "insertAfterWorksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """" & ", " & """" & insertAfterWorksheetName & """", validation)

    Dim worksheet As Worksheet

    On Error GoTo CreateWorksheetError
    If insertAfterWorksheetName = "" Then
        Set worksheet = mainWorkbook.Worksheets.Add(Before:=mainWorkbook.Worksheets(1))
    Else
        Set worksheet = mainWorkbook.Worksheets.Add(After:=mainWorkbook.Worksheets(insertAfterWorksheetName))
    End If

    worksheet.Name = worksheetName
    On Error GoTo 0

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
CreateWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub ImportTextFileOntoNewWorksheet(ByVal filePath As String, ByVal worksheetName As String, ByVal characterEncoding As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ImportTextFileOntoNewWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetName As String, ByVal characterEncoding As String, Context rowsImported As Long", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateFilePath(filePath, "filePath", validation)
    Call ValidateWorksheet(worksheetName, "worksheetName", validation, True)
    Call ValidateRequiredText(characterEncoding, "characterEncoding", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & worksheetName & """" & ", " & """" & characterEncoding & """" & ", ", validation)

    Dim worksheet As Worksheet
    Dim fileStream As Object
    Dim fileText As String
    Dim normalizedText As String
    Dim lineParts As Variant
    Dim lineCount As Long
    Dim lineValues As Variant
    Dim lineIndex As Long
    Dim destinationRange As Range

    On Error GoTo ImportTextFileOntoNewWorksheetError
    Set worksheet = mainWorkbook.Worksheets.Add(Before:=mainWorkbook.Worksheets(1))
    worksheet.Name = worksheetName

    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 2
    fileStream.Charset = characterEncoding
    fileStream.Open
    Call fileStream.LoadFromFile(filePath)
    fileText = fileStream.ReadText
    fileStream.Close
    Set fileStream = Nothing

    If Len(fileText) > 0 Then
        If AscW(Left$(fileText, 1)) = &HFEFF Then
            fileText = Mid$(fileText, 2)
        End If
    End If

    If Len(fileText) = 0 Then
        logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value = logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value & 0
        Call LogConclusion("Failed", logConclusionData, "Decoded file text is empty.")
    End If

    normalizedText = fileText
    normalizedText = Replace(normalizedText, vbCrLf, vbLf)
    normalizedText = Replace(normalizedText, vbCr, vbLf)
    lineParts = Split(normalizedText, vbLf)
    lineCount = UBound(lineParts) - LBound(lineParts) + 1

    ReDim lineValues(1 To lineCount, 1 To 1)
    For lineIndex = 1 To lineCount
        lineValues(lineIndex, 1) = lineParts(lineIndex - 1)
    Next lineIndex

    Set destinationRange = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(lineCount, 1))
    destinationRange.NumberFormat = "@"
    destinationRange.Value = lineValues
    On Error GoTo 0

    logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value = logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value & LastUsedRowNumberOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
ImportTextFileOntoNewWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub InsertValueOnNextEmptyRowInColumnOnWorksheet(ByVal value As String, ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "InsertValueOnNextEmptyRowInColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal value As String, ByVal columnName As String, ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & value & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call InsertValueOnNextEmptyRowInColumnOnWorksheet(value, columnName, worksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal value As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Conjuration")
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & value & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim columnLetter As String
    Dim lastEmptyRow As String

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    lastEmptyRow = LastUsedRowNumberOnWorksheet(worksheetName) + 1

    If lastEmptyRow = 1048577 Then
        Call LogConclusion("Failed", logConclusionData, "The column on the worksheet already has the maximum of 1,048,576 rows. Another row can't be inserted.")
    End If

    worksheet.Range(columnLetter & lastEmptyRow).Value = value

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub OpenWorkbook(ByVal filePath As String, ByRef workbook As Workbook, ByVal workbookVariableName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "OpenWorkbook"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, Withheld ByRef workbook As Workbook, ByVal workbookVariableName As String", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateFilePath(filePath, "filePath", validation)
    Call ValidateRequiredText(workbookVariableName, "workbookVariableName", validation, 1, 255)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & ", " & """" & workbookVariableName & """", validation)

    If workbook Is Nothing = False Then
        Call LogConclusion("Failed", logConclusionData, "workbook already assigned, unable to proceed.")
    End If

    Dim fileSizeInBytes As Double
    fileSizeInBytes = GetFileSizeInBytes(filePath)

    If fileSizeInBytes < 1024 Then
        Call LogConclusion("Failed", logConclusionData, "File size is under 1 KB which is too small to be a workbook.")
    End If

    On Error GoTo OpenWorkbookError
    Set workbook = Workbooks.Open(Filename:=filePath, UpdateLinks:=False, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    On Error GoTo 0

    mainWorkbook.Activate

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
OpenWorkbookError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub PopulateColumnFromWorksheetOntoWorksheet(ByVal columnNameToPopulate As String, ByVal fromKeyColumnName As String, ByVal fromWorksheetName As String, ByVal ontoKeyColumnName As String, ByVal ontoWorksheetName As String) ' Repeat Support: columnNameToPopulate. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "PopulateColumnFromWorksheetOntoWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim columnNamesToPopulate As Variant
    Dim columnNamesToPopulateHasMultipleValues As Boolean

    If Len(columnNameToPopulate) <> 0 And InStr(columnNameToPopulate, "|") Then
        columnNamesToPopulate = ParseMethodArgumentsIntoValues(columnNameToPopulate)
        If LBound(columnNamesToPopulate) < UBound(columnNamesToPopulate) Then columnNamesToPopulateHasMultipleValues = True
    End If

    If columnNamesToPopulateHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal columnNamesToPopulate As String, ByVal fromKeyColumnName As String, ByVal fromWorksheetName As String, ByVal ontoKeyColumnName As String, ByVal ontoWorksheetName As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & columnNameToPopulate & """" & ", " & """" & fromKeyColumnName & """" & ", " & """" & fromWorksheetName & """" & ", " & """" & ontoKeyColumnName & """" & ", " & """" & ontoWorksheetName & """")

        Dim columnNameToPopulateIndex As Long

        For columnNameToPopulateIndex = LBound(columnNamesToPopulate) To UBound(columnNamesToPopulate)
            columnNameToPopulate = columnNamesToPopulate(columnNameToPopulateIndex)

            Call PopulateColumnFromWorksheetOntoWorksheet(columnNameToPopulate, fromKeyColumnName, fromWorksheetName, ontoKeyColumnName, ontoWorksheetName)
        Next columnNameToPopulateIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNameToPopulate As String, ByVal fromKeyColumnName As String, ByVal fromWorksheetName As String, ByVal ontoKeyColumnName As String, ByVal ontoWorksheetName As String", methodName, "Conjuration")
    End If

    If Len(columnNameToPopulate) >= 3 And InStr(columnNameToPopulate, "|") And Left$(columnNameToPopulate, 1) = """" And Right$(columnNameToPopulate, 1) = """" Then
        columnNameToPopulate = Mid$(columnNameToPopulate, 2, Len(columnNameToPopulate) - 2)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnNameToPopulate, "columnNameToPopulate", fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateColumnOnWorksheet(columnNameToPopulate, "columnNameToPopulate", ontoWorksheetName, "ontoWorksheetName", validation, True)
    Call ValidateColumnOnWorksheet(fromKeyColumnName, "fromKeyColumnName", fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateColumnOnWorksheet(ontoKeyColumnName, "ontoKeyColumnName", ontoWorksheetName, "ontoWorksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNameToPopulate & """" & ", " & """" & fromKeyColumnName & """" & ", " & """" & fromWorksheetName & """" & ", " & """" & ontoKeyColumnName & """" & ", " & """" & ontoWorksheetName & """", validation)

    Dim catalogWorksheet As Worksheet
    Dim workingWorksheet As Worksheet
    Dim columnToPopulateWidth As Double
    Dim catalogKeyColumnLetter As String
    Dim catalogValueColumnLetter As String
    Dim workingKeyColumnLetter As String
    Dim catalogLastRowNumber As Long
    Dim workingLastRowNumber As Long
    Dim workingValueColumnLetter As String
    Dim workingKeyValues As Variant
    Dim populatedValues() As Variant
    Dim catalogValueByKey As Object
    Dim workingKey As Variant

    Set catalogWorksheet = mainWorkbook.Worksheets(fromWorksheetName)
    Set workingWorksheet = mainWorkbook.Worksheets(ontoWorksheetName)
    columnToPopulateWidth = catalogWorksheet.Range(FindColumnLetterOnWorksheet(columnNameToPopulate, fromWorksheetName) & "2").ColumnWidth
    catalogKeyColumnLetter = FindColumnLetterOnWorksheet(fromKeyColumnName, fromWorksheetName)
    catalogValueColumnLetter = FindColumnLetterOnWorksheet(columnNameToPopulate, fromWorksheetName)
    workingKeyColumnLetter = FindColumnLetterOnWorksheet(ontoKeyColumnName, ontoWorksheetName)
    catalogLastRowNumber = LastUsedRowNumberOnWorksheet(fromWorksheetName)
    workingLastRowNumber = LastUsedRowNumberOnWorksheet(ontoWorksheetName)

    If catalogLastRowNumber < 2 Then
        Call LogConclusion("Failed", logConclusionData, "Worksheet """ & fromWorksheetName & """ has no data rows.")
    End If

    If workingLastRowNumber < 2 Then
        Call LogConclusion("Failed", logConclusionData, "Worksheet """ & ontoWorksheetName & """ has no data rows.")
    End If

    Call AddColumnOnWorksheet(columnNameToPopulate, columnToPopulateWidth, ontoWorksheetName)
    workingValueColumnLetter = FindColumnLetterOnWorksheet(columnNameToPopulate, ontoWorksheetName)

    workingKeyValues = workingWorksheet.Range(workingKeyColumnLetter & "2:" & workingKeyColumnLetter & workingLastRowNumber).Value
    If IsArray(workingKeyValues) = False Then
        Dim workingKeyValuesSingleCell(1 To 1, 1 To 1) As Variant
        workingKeyValuesSingleCell(1, 1) = workingKeyValues
        workingKeyValues = workingKeyValuesSingleCell
    End If

    ReDim populatedValues(1 To UBound(workingKeyValues, 1), 1 To 1)

    Set catalogValueByKey = CreateObject("Scripting.Dictionary")
    catalogValueByKey.CompareMode = vbBinaryCompare

    If catalogLastRowNumber >= 2 Then
        Dim catalogKeyValues As Variant
        Dim catalogFieldValues As Variant
        catalogKeyValues = catalogWorksheet.Range(catalogKeyColumnLetter & "2:" & catalogKeyColumnLetter & catalogLastRowNumber).Value
        catalogFieldValues = catalogWorksheet.Range(catalogValueColumnLetter & "2:" & catalogValueColumnLetter & catalogLastRowNumber).Value

        If IsArray(catalogKeyValues) = False Then
            Dim catalogKeyValuesSingleCell(1 To 1, 1 To 1) As Variant
            catalogKeyValuesSingleCell(1, 1) = catalogKeyValues
            catalogKeyValues = catalogKeyValuesSingleCell
        End If

        If IsArray(catalogFieldValues) = False Then
            Dim catalogFieldValuesSingleCell(1 To 1, 1 To 1) As Variant
            catalogFieldValuesSingleCell(1, 1) = catalogFieldValues
            catalogFieldValues = catalogFieldValuesSingleCell
        End If

        Dim catalogRowIndex As Long
        Dim catalogKey As Variant
        For catalogRowIndex = 1 To UBound(catalogKeyValues, 1)
            catalogKey = catalogKeyValues(catalogRowIndex, 1)
            If IsError(catalogKey) = False Then
                If IsEmpty(catalogKey) = False Then
                    If Not (VarType(catalogKey) = vbString And catalogKey = "") Then
                        catalogValueByKey(catalogKey) = catalogFieldValues(catalogRowIndex, 1)
                    End If
                End If
            End If
        Next catalogRowIndex
    End If

    Dim index As Long
    For index = 1 To UBound(workingKeyValues, 1)
        populatedValues(index, 1) = ""
        workingKey = workingKeyValues(index, 1)
        If IsError(workingKey) = False Then
            If IsEmpty(workingKey) = False Then
                If Not (VarType(workingKey) = vbString And workingKey = "") Then
                    If catalogValueByKey.Exists(workingKey) = True Then
                        populatedValues(index, 1) = catalogValueByKey(workingKey)
                    End If
                End If
            End If
        End If
    Next index

    workingWorksheet.Range(workingValueColumnLetter & "2:" & workingValueColumnLetter & workingLastRowNumber).Value = populatedValues

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Functions: Conjuration '

Sub GetFilesFromDirectory(ByVal directoryPath As String, ByVal filterValue As String, ByRef files As Variant, ByVal filesVariableName As String)
    Const methodName As String = "GetFilesFromDirectory"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal directoryPath As String, ByVal filterValue As String, Withheld ByRef files As Variant, ByVal filesVariableName As String", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateDirectory(directoryPath, "directoryPath", validation)
    Call ValidateRequiredText(filesVariableName, "filesVariableName", validation, 1, 255)

    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & directoryPath & """" & ", " & """" & filterValue & """" & ", " & """" & filesVariableName & """", validation)
    End If

    Const defaultFilePattern As String = "*.*"
    Const initialArrayCapacity As Long = 16

    Dim currentFileName As String
    Dim fileListArray() As String
    Dim fileCount As Long
    Dim fileListCapacity As Long

    currentFileName = Dir$(directoryPath & defaultFilePattern)
    If Len(filterValue) <> 0 Then
        currentFileName = Dir$(directoryPath & filterValue)
    End If

    Do While Len(currentFileName) > 0
        If fileCount = fileListCapacity Then
            If fileListCapacity = 0 Then
                fileListCapacity = initialArrayCapacity
            Else
                fileListCapacity = fileListCapacity * 2
            End If

            ReDim Preserve fileListArray(0 To fileListCapacity - 1)
        End If

        fileListArray(fileCount) = directoryPath & currentFileName
        fileCount = fileCount + 1

        currentFileName = Dir$()
    Loop

    If fileCount = 0 Then
        files = Split(vbNullString)
    Else
        ReDim Preserve fileListArray(0 To fileCount - 1)
        files = fileListArray
    End If
End Sub

' ************ '
' Destruction  '
' ************ '

Sub CloseWorkbook(ByRef workbook As Workbook, ByVal workbookVariableName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CloseWorkbook"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("Withheld ByRef workbook As Workbook, ByVal workbookVariableName As String", methodName, "Destruction")
    End If

    Dim validation As String
    Call ValidateRequiredText(workbookVariableName, "workbookVariableName", validation, 1, 255)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & workbookVariableName & """", validation)

    If workbook Is Nothing = True Then
        Call LogConclusion("Failed", logConclusionData, "workbook not previously assigned, unable to proceed.")
    End If

    If TypeName(workbook) <> "Workbook" Then
        Call LogConclusion("Failed", logConclusionData, "workbook is no longer a valid open workbook, unable to proceed.")
    End If

    On Error GoTo CloseWorkbookError
    If StrComp(workbook.Name, "PERSONAL.XLSB", vbTextCompare) = 0 Then
        Call LogConclusion("Failed", logConclusionData, "Unable to close Personal Macro Workbook.")
    End If

    Call workbook.Close(SaveChanges:=False)
    On Error GoTo 0

    Set workbook = Nothing

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
CloseWorkbookError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub DeleteCellStyle(ByVal cellStyleName As String) ' Repeat Support: cellStyleName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteCellStyle"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim cellStyleNames As Variant
    Dim cellStyleNamesHasMultipleValues As Boolean

    If Len(cellStyleName) <> 0 And InStr(cellStyleName, "|") Then
        cellStyleNames = ParseMethodArgumentsIntoValues(cellStyleName)
        If LBound(cellStyleNames) < UBound(cellStyleNames) Then cellStyleNamesHasMultipleValues = True
    End If

    If cellStyleNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal cellStyleNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & cellStyleName & """")

        Dim cellStyleNameIndex As Long

        For cellStyleNameIndex = LBound(cellStyleNames) To UBound(cellStyleNames)
            cellStyleName = cellStyleNames(cellStyleNameIndex)

            Call DeleteCellStyle(cellStyleName)
        Next cellStyleNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyleName As String, Context regionalizedCellStyleName As String", methodName, "Destruction")
    End If

    If Len(cellStyleName) >= 3 And InStr(cellStyleName, "|") And Left$(cellStyleName, 1) = """" And Right$(cellStyleName, 1) = """" Then
        cellStyleName = Mid$(cellStyleName, 2, Len(cellStyleName) - 2)
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cellStyleName & """" & ", ")

    Dim regionalizedCellStyleName As String
    If cellStyles.Exists(cellStyleName) = True Then
        regionalizedCellStyleName = cellStyles(cellStyleName)
    End If

    logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value = logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value & """" & regionalizedCellStyleName & """"

    Dim deleteCellStyleName As String
    If CellStyleExists(cellStyleName) = True Then
        deleteCellStyleName = cellStyleName
    ElseIf regionalizedCellStyleName <> "" Then
        If CellStyleExists(regionalizedCellStyleName) = True Then
            deleteCellStyleName = regionalizedCellStyleName
        End If
    End If

    If deleteCellStyleName = "" Or Right$(deleteCellStyleName, 1) = " " Then
        Call LogConclusion("Skipped", logConclusionData)

        Exit Sub
    End If

    If deleteCellStyleName = "Normal" Or deleteCellStyleName = cellStyles("Normal") Then
        Call LogConclusion("Failed", logConclusionData, "The cell style """ & "Normal" & """ is reserved and can't be deleted.")
    End If

    mainWorkbook.Styles(deleteCellStyleName).Delete
    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub DeleteColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName + worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim columnNames As Variant
    Dim worksheetNames As Variant
    Dim columnNamesHasMultipleValues As Boolean
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(columnName) <> 0 And InStr(columnName, "|") Then
        columnNames = ParseMethodArgumentsIntoValues(columnName)
        If LBound(columnNames) < UBound(columnNames) Then columnNamesHasMultipleValues = True
    End If

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If columnNamesHasMultipleValues = True Or worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal columnNames As String, ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

        Dim columnNameIndex As Long
        Dim worksheetNameIndex As Long

        If columnNamesHasMultipleValues = True And worksheetNamesHasMultipleValues = True Then
            For columnNameIndex = LBound(columnNames) To UBound(columnNames)
                For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
                    columnName = columnNames(columnNameIndex)
                    worksheetName = worksheetNames(worksheetNameIndex)

                    Call DeleteColumnOnWorksheet(columnName, worksheetName)
                Next worksheetNameIndex
            Next columnNameIndex
        ElseIf columnNamesHasMultipleValues = True And worksheetNamesHasMultipleValues = False Then
            For columnNameIndex = LBound(columnNames) To UBound(columnNames)
                columnName = columnNames(columnNameIndex)

                Call DeleteColumnOnWorksheet(columnName, worksheetName)
            Next columnNameIndex
        ElseIf columnNamesHasMultipleValues = False And worksheetNamesHasMultipleValues = True Then
            For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
                worksheetName = worksheetNames(worksheetNameIndex)

                Call DeleteColumnOnWorksheet(columnName, worksheetName)
            Next worksheetNameIndex
        End If

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Destruction")
    End If

    If Len(columnName) >= 3 And InStr(columnName, "|") And Left$(columnName, 1) = """" And Right$(columnName, 1) = """" Then
        columnName = Mid$(columnName, 2, Len(columnName) - 2)
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim columnLetter As String

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)

    On Error GoTo DeleteColumnError
    Call worksheet.Columns(columnLetter).Delete(Shift:=xlToLeft)
    On Error GoTo 0

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
DeleteColumnError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub DeleteWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & worksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call DeleteWorksheet(worksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Destruction")
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Delete

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Functions: Destruction '

' ************ '
' Elementals   '
' ************ '

Sub SelectWorksheet(ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SelectWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If mainWorkbook.ActiveSheet.Name = worksheetName Then
        Call LogConclusion("Skipped", logConclusionData)
    Else
        worksheet.Select
        Call LogConclusion("Completed", logConclusionData)
    End If
End Sub

' Functions: Elementals '

Function ConvertColumnLetterToColumnNumber(ByVal columnLetterToConvert As String) As Long
    Dim tickCount As Currency
    Const methodName As String = "ConvertColumnLetterToColumnNumber"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnLetterToConvert As String", methodName, "Elementals")
    End If

    Const MAXIMUM_EXCEL_COLUMN_NUMBER As Long = 16384
    Const ASCII_CODE_BEFORE_UPPERCASE_A As Long = 64

    Dim normalizedColumnLetter As String
    Dim currentCharacter As String
    Dim characterValue As Long
    Dim columnNumber As Long

    ConvertColumnLetterToColumnNumber = 0
    normalizedColumnLetter = UCase$(columnLetterToConvert)

    If Len(normalizedColumnLetter) = 0 Then
        tickCount = GetTickCount64()
        Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnLetterToConvert & """", "Argument is empty, nothing can be done.")
    End If

    Dim characterIndex As Long
    For characterIndex = 1 To Len(normalizedColumnLetter)
        currentCharacter = Mid$(normalizedColumnLetter, characterIndex, 1)
        characterValue = AscW(currentCharacter) - ASCII_CODE_BEFORE_UPPERCASE_A

        If characterValue < 1 Or characterValue > 26 Then
            tickCount = GetTickCount64()
            Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnLetterToConvert & """", "Column letter " & Chr$(34) & columnLetterToConvert & Chr$(34) & " contains an invalid character.")
        End If

        columnNumber = columnNumber * 26 + characterValue

        If columnNumber > MAXIMUM_EXCEL_COLUMN_NUMBER Then
            tickCount = GetTickCount64()
            Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnLetterToConvert & """", "Column letter " & Chr$(34) & columnLetterToConvert & Chr$(34) & " exceeds Excel's maximum supported column.")
        End If
    Next characterIndex

    ConvertColumnLetterToColumnNumber = columnNumber
End Function

Function ConvertColumnNumberToColumnLetter(ByVal columnNumberToConvert As Long) As String
    Dim tickCount As Currency
    Const methodName As String = "ConvertColumnNumberToColumnLetter"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNumberToConvert As Long", methodName, "Elementals")
    End If

    Const MAXIMUM_EXCEL_COLUMN_NUMBER As Long = 16384
    Const ASCII_CODE_FOR_UPPERCASE_A As Long = 65

    Dim remainingColumnNumber As Long
    Dim remainderValue As Long
    Dim columnLetter As String

    ConvertColumnNumberToColumnLetter = ""

    If columnNumberToConvert < 1 Then
        tickCount = GetTickCount64()
        Call LogBeginning(methodName, tickCount, logConclusionData, CStr(columnNumberToConvert), "Column number " & CStr(columnNumberToConvert) & " is below Excel's minimum supported column.")
    ElseIf columnNumberToConvert > MAXIMUM_EXCEL_COLUMN_NUMBER Then
        tickCount = GetTickCount64()
        Call LogBeginning(methodName, tickCount, logConclusionData, CStr(columnNumberToConvert), "Column number " & CStr(columnNumberToConvert) & " exceeds Excel's maximum supported column limit of 16,384 columns.")
    End If

    remainingColumnNumber = columnNumberToConvert
    Do While remainingColumnNumber > 0
        remainderValue = (remainingColumnNumber - 1) Mod 26
        columnLetter = ChrW$(ASCII_CODE_FOR_UPPERCASE_A + remainderValue) & columnLetter
        remainingColumnNumber = (remainingColumnNumber - 1) \ 26
    Loop

    ConvertColumnNumberToColumnLetter = columnLetter
End Function

Function FindColumnLetterOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    Const methodName As String = "FindColumnLetterOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & worksheetName & """", validation)
    End If

    Dim worksheet As Worksheet
    Dim headerSearchRange As Range
    Dim foundHeaderCell As Range

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    Set headerSearchRange = worksheet.Rows(1)
    Set foundHeaderCell = headerSearchRange.Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)

    FindColumnLetterOnWorksheet = Split(foundHeaderCell.Address, "$")(1)
End Function

Function GetFileSizeInBytes(ByVal filePath As String) As Double
    Const methodName As String = "GetFileSizeInBytes"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateFilePath(filePath, "filePath", validation)
    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & filePath & """", validation)
    End If

    Const GENERIC_READ As Long = &H80000000
    Const FILE_SHARE_READ As Long = 1
    Const FILE_SHARE_WRITE As Long = 2
    Const OPEN_EXISTING As Long = 3
    Const FILE_ATTRIBUTE_NORMAL As Long = &H80
    Const INVALID_HANDLE_VALUE As LongPtr = -1

    Dim fileHandle As LongPtr
    Dim rawFileSize As Currency
    Dim sizeWasRetrieved As Long

    fileHandle = CreateFileW(StrPtr(filePath), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If fileHandle = INVALID_HANDLE_VALUE Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & filePath & """", "Unable to open the file.")
    End If

    sizeWasRetrieved = GetFileSizeEx(fileHandle, rawFileSize)
    Call CloseHandle(fileHandle)

    If sizeWasRetrieved = 0 Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & filePath & """", "Unable to read the size of the file.")
    End If

    GetFileSizeInBytes = rawFileSize * 10000
End Function

Function GetUtcDateTime() As Date
    Dim systemTime As SystemTimeStructure

    Call GetSystemTime(systemTime)
    
    GetUtcDateTime = DateSerial(systemTime.Year, systemTime.Month, systemTime.Day) + TimeSerial(systemTime.Hour, systemTime.Minute, systemTime.Second) + (systemTime.Milliseconds / 86400000#)
End Function

Function GetUtcTimestamp() As String
    Dim systemTime As SystemTimeStructure
    
    Call GetSystemTime(systemTime)
    
    GetUtcTimestamp = Format$(systemTime.Year, "0000") & "-" & Format$(systemTime.Month, "00") & "-" & Format$(systemTime.Day, "00") & " " & _
        Format$(systemTime.Hour, "00") & ":" & Format$(systemTime.Minute, "00") & ":" & Format$(systemTime.Second, "00") & "." & Format$(systemTime.Milliseconds, "000")
End Function

Function LastUsedColumnLetterOnWorksheet(ByVal worksheetName As String) As String
    Const methodName As String = "LastUsedColumnLetterOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & worksheetName & """", validation)
    End If

    Dim worksheet As Worksheet
    Dim lastUsedColumnNumber As Long
    Dim workingColumnNumber As Long
    Dim resultingColumnLetter As String

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If IsEmpty(worksheet.Cells(1, worksheet.Columns.Count).Value) = False Then
        lastUsedColumnNumber = worksheet.Columns.Count
    Else
        lastUsedColumnNumber = worksheet.Cells(1, worksheet.Columns.Count).End(xlToLeft).Column
    End If

    workingColumnNumber = lastUsedColumnNumber
    Do While workingColumnNumber > 0
        workingColumnNumber = workingColumnNumber - 1
        resultingColumnLetter = Chr$(65 + (workingColumnNumber Mod 26)) & resultingColumnLetter
        workingColumnNumber = workingColumnNumber \ 26
    Loop

    LastUsedColumnLetterOnWorksheet = resultingColumnLetter
End Function

Function LastUsedRowNumberOnWorksheet(ByVal worksheetName As String) As Long
    Const methodName As String = "LastUsedRowNumberOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & worksheetName & """", validation)
    End If

    Dim worksheet As Worksheet
    Dim lastUsedRowNumber As Long

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If IsEmpty(worksheet.Cells(worksheet.Rows.Count, 1).Value) = False Then
        lastUsedRowNumber = worksheet.Cells(worksheet.Rows.Count, 1).Row
    Else
        lastUsedRowNumber = worksheet.Cells(worksheet.Rows.Count, 1).End(xlUp).Row
    End If

    LastUsedRowNumberOnWorksheet = lastUsedRowNumber
End Function

' Functions: Elementals, Formulas '

Function AdjacentRange(ByVal columnName As String, ByVal worksheetName As String) As String
    Const methodName As String = "AdjacentRange"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)
    End If

    Dim columnLetter As String

    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)

    AdjacentRange = columnLetter & "1:" & columnLetter & "3"
End Function

Function ContainsCriteria(ByVal columnName As String, ByVal worksheetName As String) As String
    Const methodName As String = "ContainsCriteria"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)
    End If

    Dim columnLetter As String

    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)

    ContainsCriteria = """*;"" & " & columnLetter & "2 & "";*"""
End Function

Function DataRange(ByVal columnName As String, ByVal worksheetName As String) As String
    Const methodName As String = "DataRange"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)
    End If

    Dim columnLetter As String
    Dim lastUsedRowNumber As Long

    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)

    DataRange = "'" & Replace(worksheetName, "'", "''") & "'!" & columnLetter & "$2:" & columnLetter & "$" & lastUsedRowNumber
End Function

Function StartCell(ByVal columnName As String, ByVal worksheetName As String) As String
    Const methodName As String = "StartCell"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    If validation <> "" Then
        Call LogBeginning(methodName, GetTickCount64(), logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)
    End If

    Dim columnLetter As String

    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)

    StartCell = columnLetter & "2"
End Function

' Helpers: Elementals '

Private Sub ApplyBoldBeforeFirstColon(ByVal targetCells As Range)
    Dim currentCell As Range
    Dim firstColonPosition As Long

    For Each currentCell In targetCells.Cells
        If currentCell.MergeCells = False Or currentCell.Address = currentCell.MergeArea.Cells(1, 1).Address Then
            firstColonPosition = InStr(currentCell.Text, ":")
            currentCell.Font.Bold = False
            currentCell.Characters(Start:=1, Length:=firstColonPosition - 1).Font.Bold = True
        End If
    Next currentCell
End Sub

Private Sub ApplyConsolasAfterFirstColonAndSpace(ByVal targetCells As Range)
    Dim currentCell As Range
    Dim firstColonPosition As Long

    For Each currentCell In targetCells.Cells
        If currentCell.MergeCells = False Or currentCell.Address = currentCell.MergeArea.Cells(1, 1).Address Then
            firstColonPosition = InStr(currentCell.Text, ":")
            currentCell.Characters(Start:=firstColonPosition + 2).Font.Name = "Consolas"
        End If
    Next currentCell
End Sub

' ************ '
' Formatting   '
' ************ '

Sub ApplyCellStyleToColumnOnWorksheet(ByVal cellStyle As String, ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName + worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyCellStyleToColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim columnNames As Variant
    Dim worksheetNames As Variant
    Dim columnNamesHasMultipleValues As Boolean
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(columnName) <> 0 And InStr(columnName, "|") Then
        columnNames = ParseMethodArgumentsIntoValues(columnName)
        If LBound(columnNames) < UBound(columnNames) Then columnNamesHasMultipleValues = True
    End If

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If columnNamesHasMultipleValues = True Or worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal cellStyle As String, ByVal columnNames As String, ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & cellStyle & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

        Dim columnNameIndex As Long
        Dim worksheetNameIndex As Long

        If columnNamesHasMultipleValues = True And worksheetNamesHasMultipleValues = True Then
            For columnNameIndex = LBound(columnNames) To UBound(columnNames)
                For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
                    columnName = columnNames(columnNameIndex)
                    worksheetName = worksheetNames(worksheetNameIndex)

                    Call ApplyCellStyleToColumnOnWorksheet(cellStyle, columnName, worksheetName)
                Next worksheetNameIndex
            Next columnNameIndex
        ElseIf columnNamesHasMultipleValues = True And worksheetNamesHasMultipleValues = False Then
            For columnNameIndex = LBound(columnNames) To UBound(columnNames)
                columnName = columnNames(columnNameIndex)

                Call ApplyCellStyleToColumnOnWorksheet(cellStyle, columnName, worksheetName)
            Next columnNameIndex
        ElseIf columnNamesHasMultipleValues = False And worksheetNamesHasMultipleValues = True Then
            For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
                worksheetName = worksheetNames(worksheetNameIndex)

                Call ApplyCellStyleToColumnOnWorksheet(cellStyle, columnName, worksheetName)
            Next worksheetNameIndex
        End If

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyle As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Formatting")
    End If

    If Len(columnName) >= 3 And InStr(columnName, "|") And Left$(columnName, 1) = """" And Right$(columnName, 1) = """" Then
        columnName = Mid$(columnName, 2, Len(columnName) - 2)
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cellStyle & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    validation = "Parameter """ & "cellStyle" & """ failed validation. Cell style can't be found."
    If cellStyles.Exists(cellStyle) Then
        If CellStyleExists(cellStyle) = False And CellStyleExists(cellStyles(cellStyle)) = False Then
            Call LogConclusion("Failed", logConclusionData, validation)
        End If

        If CellStyleExists(cellStyle) = False And CellStyleExists(cellStyles(cellStyle)) = True Then
            cellStyle = cellStyles(cellStyle)
        End If
    Else
        If CellStyleExists(cellStyle) = False Then
            Call LogConclusion("Failed", logConclusionData, validation)
        End If
    End If

    Dim columnLetter As String
    Dim lastUsedRowNumber As Long

    columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)

    With worksheet.Range(columnLetter & "2:" & columnLetter & lastUsedRowNumber)
        .Style = cellStyle
        .Value = .Value
    End With

    If lastUsedRowNumber <> worksheet.Rows.Count Then
        With worksheet.Range(columnLetter & (lastUsedRowNumber + 1) & ":" & columnLetter & worksheet.Rows.Count)
            .Style = cellStyle
        End With
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ColorWorksheet(ByVal colorName As String, ByRef colorDictionary As Object, ByVal colorDictionaryVariableName As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal colorName As String, Withheld ByRef colorDictionary As Object, ByVal colorDictionaryVariableName As String, ByVal worksheetName As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & colorName & """" & ", " & """" & colorDictionaryVariableName & """" & ", " & """" & worksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call ColorWorksheet(colorName, colorDictionary, colorDictionaryVariableName, worksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, Withheld ByRef colorDictionary As Object, ByVal colorDictionaryVariableName As String, ByVal worksheetName As String, Context hex As String", methodName, "Formatting")
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateRequiredText(colorName, "colorName", validation, 1)
    Call ValidateRequiredText(colorDictionaryVariableName, "colorDictionaryVariableName", validation, 1, 255)
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & colorName & """" & ", " & """" & colorDictionaryVariableName & """" & ", " & """" & worksheetName & """" & ", ", validation)

    Dim worksheet As Worksheet
    Dim hex As String
    Dim color As Long
    Dim originalValueLength As Long

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If TypeName(colorDictionary) <> "Dictionary" Then
        logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value = logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value & """#"""
        Call LogConclusion("Failed", logConclusionData, "Passed in colorDictionary is not a dictionary as expected.")
    End If

    If colorDictionary.Exists(colorName) = False Then
        logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value = logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2).Value & """#"""
        Call LogConclusion("Failed", logConclusionData, "Color not found in colorDictionary.")
    End If

    On Error GoTo ColorWorksheetError
    hex = colorDictionary(colorName)("Hex")
    color = colorDictionary(colorName)("Long")

    With logWorksheet.Cells(logConclusionData.operationSequenceNumber, 2)
        originalValueLength = Len(.Value)
        .Value = .Value & """" & hex & """"
        .Characters(Start:=originalValueLength + 2, Length:=Len("""" & hex & """") - 2).Font.Color = color
    End With

    worksheet.Tab.Color = color
    On Error GoTo 0

    Call LogConclusion("Completed", logConclusionData)
    Exit Sub
ColorWorksheetError:
    Call LogConclusion("Failed", logConclusionData, "Error " & Err.Number & ": " & Err.Description)
End Sub

Sub CreateColorDictionaryFromColumnsOnWorksheet(ByRef colorDictionary As Object, ByVal colorDictionaryVariableName As String, ByVal colorColumnName As String, ByVal hexColumnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CreateColorDictionaryFromColumnsOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("Withheld ByRef colorDictionary As Object, ByVal colorDictionaryVariableName As String, ByVal colorColumnName As String, ByVal hexColumnName As String, ByVal worksheetName As String", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateRequiredText(colorDictionaryVariableName, "colorDictionaryVariableName", validation, 1, 255)
    Call ValidateColumnOnWorksheet(colorColumnName, "colorColumnName", worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(hexColumnName, "hexColumnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & colorDictionaryVariableName & """" & ", " & """" & colorColumnName & """" & ", " & """" & hexColumnName & """" & ", " & """" & worksheetName & """", validation)

    If colorDictionary Is Nothing = False Then
        Call LogConclusion("Failed", logConclusionData, "colorDictionary already assigned, unable to proceed.")
    End If

    Dim worksheet As Worksheet
    Dim lastUsedRowNumber As Long
    Dim colorColumnLetter As String
    Dim hexColumnLetter As String
    Dim colorDataRange As Range
    Dim hexDataRange As Range
    Dim colorValues As Variant
    Dim hexValues As Variant

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)
    colorColumnLetter = FindColumnLetterOnWorksheet(colorColumnName, worksheetName)
    hexColumnLetter = FindColumnLetterOnWorksheet(hexColumnName, worksheetName)

    If lastUsedRowNumber = 1 Then
        Call LogConclusion("Failed", logConclusionData, "No data available in worksheet.")
    End If

    Set colorDataRange = worksheet.Range(colorColumnLetter & "2:" & colorColumnLetter & lastUsedRowNumber)
    Set hexDataRange = worksheet.Range(hexColumnLetter & "2:" & hexColumnLetter & lastUsedRowNumber)

    colorValues = colorDataRange.Value
    hexValues = hexDataRange.Value

    If IsArray(colorValues) = False Then
        Dim colorValuesWrapped(1 To 1, 1 To 1) As Variant
        colorValuesWrapped(1, 1) = colorValues
        colorValues = colorValuesWrapped
    End If

    If IsArray(hexValues) = False Then
        Dim hexValuesWrapped(1 To 1, 1 To 1) As Variant
        hexValuesWrapped(1, 1) = hexValues
        hexValues = hexValuesWrapped
    End If

    Set colorDictionary = CreateObject("Scripting.Dictionary")
    colorDictionary.CompareMode = vbTextCompare

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(colorValues, 1)
        Dim colorLabel As String
        Dim hexadecimalColor As String

        colorLabel = CStr(colorValues(rowIndex, 1))
        hexadecimalColor = CStr(hexValues(rowIndex, 1))

        If Len(colorLabel) = 0 And Len(hexadecimalColor) = 0 Then
            ' Sheet last row is below these columns.
        ElseIf Len(colorLabel) = 0 Then
            Call LogConclusion("Failed", logConclusionData, "Row " & CStr(rowIndex + 1) & " Color is missing data.")
        ElseIf Len(hexadecimalColor) = 0 Then
            Call LogConclusion("Failed", logConclusionData, "Row " & CStr(rowIndex + 1) & " Hex is missing data.")
        ElseIf Left$(colorLabel, 1) = " " Then
            Call LogConclusion("Failed", logConclusionData, "Row " & CStr(rowIndex + 1) & " Color starts with a space.")
        ElseIf Right$(colorLabel, 1) = " " Then
            Call LogConclusion("Failed", logConclusionData, "Row " & CStr(rowIndex + 1) & " Color ends with a space.")
        Else
            Dim redComponent As Long
            Dim greenComponent As Long
            Dim blueComponent As Long
            Dim excelColorValue As Long
            Dim colorDefinition As Object

            Call ValidateHexColor(hexadecimalColor, "hexadecimalColor", validation)

            If Len(validation) > 0 Then
                validation = Replace(validation, "Parameter """, "Variable """)
                Call LogConclusion("Failed", logConclusionData, validation)
            End If

            If colorDictionary.Exists(colorLabel) = True Then
                Call LogConclusion("Failed", logConclusionData, "Duplicate color label """ & colorLabel & """.")
            End If

            redComponent = CLng("&H" & Mid$(hexadecimalColor, 2, 2))
            greenComponent = CLng("&H" & Mid$(hexadecimalColor, 4, 2))
            blueComponent = CLng("&H" & Mid$(hexadecimalColor, 6, 2))
            excelColorValue = RGB(redComponent, greenComponent, blueComponent)

            Set colorDefinition = CreateObject("Scripting.Dictionary")
            colorDefinition.CompareMode = vbTextCompare
            colorDefinition("Hex") = UCase$(hexadecimalColor)
            colorDefinition("Long") = excelColorValue
            colorDefinition("Red") = redComponent
            colorDefinition("Green") = greenComponent
            colorDefinition("Blue") = blueComponent

            Set colorDictionary(colorLabel) = colorDefinition
        End If
    Next rowIndex

    If colorDictionary.Count = 0 Then
        Call LogConclusion("Failed", logConclusionData, "No color rows were loaded.")
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub FormatCoreWorksheet(ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "FormatCoreWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateWhitelist(worksheetName, "worksheetName", Array("About", "Log", "Run Status"), validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If worksheetName = "About" Then
        With worksheet
            .Cells.RowHeight = 16.5
            .Rows("1:4").RowHeight = 33.0
            .Columns("A:B").ColumnWidth = 67.29
            .Columns("C").ColumnWidth = 31.29

            .Range("A1:C1").WrapText = True
            .Range("B2:C2").WrapText = True
            .Range("A3:C4").WrapText = True

            .Range("B3:B4").VerticalAlignment = xlTop

            .Range("A1:C4").Font.Name = "Tahoma"
            .Range("A1:C4").Font.Size = 8
            .Range("A1:B2").Font.Size = 12

            .Range("ReportName").Font.Bold = True
            .Range("TemplateVersion").Font.Bold = True

            Call ApplyBoldBeforeFirstColon(Union(.Range("FoundationCheckpoints"), .Range("AugmentationCheckpoints"), .Range("DependenciesList"), .Range("CreationDate"), .Range("EditionName"), .Range("DurationMilliseconds"), .Range("LogSummary")))
        End With
    End If

    If worksheetName = "Log" Then
        With worksheet
            If .Cells.RowHeight <> 16.5 Or _
            .Rows("1:1").RowHeight <> 49.5 Or _
            .AutoFilterMode <> True Then
                Call NormalizeLayoutOnWorksheet("Log")
                Call SetFrozenPanesOnWorksheet(1, 0, "Log")
            End If

            If .Columns("A").ColumnWidth <> 50.14 Or .Columns("B").ColumnWidth <> 75.86 Or .Columns("C").ColumnWidth <> 6.14 Or .Columns("D").ColumnWidth <> 8.71 Or .Columns("E").ColumnWidth <> 8.71 Or .Columns("F").ColumnWidth <> 10.43 Then
                .Columns("A").ColumnWidth = 50.14
                .Columns("B").ColumnWidth = 75.86
                .Columns("C").ColumnWidth = 6.14
                .Columns("D").ColumnWidth = 8.71
                .Columns("E").ColumnWidth = 8.71
                .Columns("F").ColumnWidth = 10.43
            End If
        End With

        If worksheet.Range("C2").Style <> "Integer" Then
            Call ApplyCellStyleToColumnOnWorksheet("Integer", "Depth", "Log")
        End If

        If worksheet.Range("D2").Style <> "Integer" Then
            Call ApplyCellStyleToColumnOnWorksheet("Integer", "Tick Count", "Log")
        End If

        If worksheet.Range("E2").Style <> "Integer" Then
            Call ApplyCellStyleToColumnOnWorksheet("Integer", "Duration", "Log")
        End If
    End If

    If worksheetName = "Run Status" Then
        With worksheet
            .Cells.Font.Name = "Segoe UI"
            .Cells.Font.Size = 10
            .Cells.RowHeight = 16.5
            .Cells.HorizontalAlignment = xlCenter
            .Range("A1:A28").HorizontalAlignment = xlLeft
            .Columns("A").ColumnWidth = 74.71
            .Columns("B:E").ColumnWidth = 22.43

            .Range("A4:A8").VerticalAlignment = xlTop
            .Range("A1:A12").WrapText = True
            
            Call ApplyBoldBeforeFirstColon(.Range("A1:A28"))
            Call ApplyConsolasAfterFirstColonAndSpace(Union(.Range("A19"), .Range("A22:A24"), .Range("A26:A27")))

            .Range("A8").Characters(Start:=InStrRev(.Range("A8").Value, " ") + 1).Font.Name = "Consolas"
            Union(.Range("C11"), .Range("C12"), .Range("C17"), .Range("C20"), .Range("C23"), .Range("E4"), .Range("E13"), .Range("E17")).Style = "Integer"
            With Union(.Range("B1"), .Range("B10"), .Range("B14"), .Range("B24"), .Range("D1"), .Range("D19"), .Range("D22"))
                    .Font.Bold = True
                    .Font.Name = "Segoe UI"
                    .Font.Size = 12
            End With
        End With
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub NormalizeLayoutOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "NormalizeLayoutOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, """" & worksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call NormalizeLayoutOnWorksheet(worksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Formatting")

        Call ConfigureMethodSetting(methodName, "Apply Auto Filter", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Header Style", 1, 0, 1)
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim settings As Object
    Dim worksheet As Worksheet
    Dim lastUsedColumnLetter As String

    Set settings = methodRegistry(methodName)("Settings")
    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    lastUsedColumnLetter = LastUsedColumnLetterOnWorksheet(worksheetName)

    With worksheet
        .Cells.RowHeight = 16.5
        .Rows("1:1").RowHeight = 49.5
    End With

    If settings("Apply Auto Filter")("Value") = True Then
        worksheet.AutoFilterMode = False

        If worksheet.Range(lastUsedColumnLetter & "1").Value <> "" Then
            worksheet.Range("A1:" & lastUsedColumnLetter & "1").AutoFilter
        End If
    End If

    If settings("Header Style")("Value") = True Then
        If CellStyleExists("Header") = False Then
            Call LogConclusion("Failed", logConclusionData, "Cell style Header not found.")
        End If

        If worksheet.Range(lastUsedColumnLetter & "1").Value <> "" Then
            With worksheet.Range("A1:" & lastUsedColumnLetter & "1")
                .Style = "Header"
            End With
        End If
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetFrozenPanesOnWorksheet(ByVal frozenRowCount As Long, ByVal frozenColumnCount As Long, ByVal worksheetName As String) ' Repeat Support: worksheetName. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetFrozenPanesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim worksheetNames As Variant
    Dim worksheetNamesHasMultipleValues As Boolean

    If Len(worksheetName) <> 0 And InStr(worksheetName, "|") Then
        worksheetNames = ParseMethodArgumentsIntoValues(worksheetName)
        If LBound(worksheetNames) < UBound(worksheetNames) Then worksheetNamesHasMultipleValues = True
    End If

    If worksheetNamesHasMultipleValues = True Then
        If methodRegistry.Exists("Repeat" & methodName) = True Then If methodRegistry("Repeat" & methodName).Exists("Registered") = True Then isRegistered = True
        If isRegistered = False Then
            Call RegisterMethod("ByVal frozenRowCount As Long, ByVal frozenColumnCount As Long, ByVal worksheetNames As String", "Repeat" & methodName, "Repetition")
        End If

        Call LogBeginning("Repeat" & methodName, tickCount, logConclusionData, frozenRowCount & ", " & frozenColumnCount & ", " & """" & worksheetName & """")

        Dim worksheetNameIndex As Long

        For worksheetNameIndex = LBound(worksheetNames) To UBound(worksheetNames)
            worksheetName = worksheetNames(worksheetNameIndex)

            Call SetFrozenPanesOnWorksheet(frozenRowCount, frozenColumnCount, worksheetName)
        Next worksheetNameIndex

        Call LogConclusion("Completed", logConclusionData)
        Exit Sub
    End If

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal frozenRowCount As Long, ByVal frozenColumnCount As Long, ByVal worksheetName As String", methodName, "Formatting")
    End If

    If Len(worksheetName) >= 3 And InStr(worksheetName, "|") And Left$(worksheetName, 1) = """" And Right$(worksheetName, 1) = """" Then
        worksheetName = Mid$(worksheetName, 2, Len(worksheetName) - 2)
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, frozenRowCount & ", " & frozenColumnCount & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    mainWorkbook.Activate
    worksheet.Activate

    Dim targetWindow As Window
    Set targetWindow = ActiveWindow

    If targetWindow.WindowState = xlMinimized Then
        targetWindow.WindowState = xlNormal
    End If

    targetWindow.ScrollRow = 1
    targetWindow.ScrollColumn = 1

    If targetWindow.View <> xlNormalView Then
        targetWindow.View = xlNormalView
    End If

    If targetWindow.FreezePanes = True Then
        targetWindow.FreezePanes = False

        If targetWindow.Split = True Then
            targetWindow.Split = False
        End If
    End If

    If frozenRowCount > 0 Or frozenColumnCount > 0 Then
        targetWindow.SplitRow = frozenRowCount
        targetWindow.SplitColumn = frozenColumnCount
        targetWindow.FreezePanes = True

        If targetWindow.FreezePanes = False Then
            Call LogConclusion("Failed", logConclusionData, "Freeze Panes was not set.")
        End If
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Functions: Formatting '

' ************ '
' Logging      '
' ************ '

Sub CreateAboutWorksheet()
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CreateAboutWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("", methodName, "Logging")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, "")

    Dim namedRanges As Object
    Set namedRanges = CreateObject("Scripting.Dictionary")

    namedRanges("ReportDetails")       = NamedRangeExists("ReportDetails")
    namedRanges("TemplateDetails")     = NamedRangeExists("TemplateDetails")
    namedRanges("ProgressionStatus")   = NamedRangeExists("ProgressionStatus")
    namedRanges("AugmentationModules") = NamedRangeExists("AugmentationModules")
    namedRanges("RetrievedDate")       = NamedRangeExists("RetrievedDate")
    namedRanges("ScriptDuration")      = NamedRangeExists("ScriptDuration")

    namedRanges("ReportName")              = NamedRangeExists("ReportName")
    namedRanges("TemplateVersion")         = NamedRangeExists("TemplateVersion")
    namedRanges("FoundationCheckpoints")   = NamedRangeExists("FoundationCheckpoints")
    namedRanges("AugmentationCheckpoints") = NamedRangeExists("AugmentationCheckpoints")
    namedRanges("ReportVision")            = NamedRangeExists("ReportVision")
    namedRanges("DependenciesList")        = NamedRangeExists("DependenciesList")
    namedRanges("CreationDate")            = NamedRangeExists("CreationDate")
    namedRanges("EditionName")             = NamedRangeExists("EditionName")
    namedRanges("DurationMilliseconds")    = NamedRangeExists("DurationMilliseconds")
    namedRanges("LogSummary")              = NamedRangeExists("LogSummary")

    If WorksheetExists("About") = False Then
        If namedRanges("ReportDetails") = True Then mainWorkbook.Names("ReportDetails").Delete
        If namedRanges("TemplateDetails") = True Then mainWorkbook.Names("TemplateDetails").Delete
        If namedRanges("ProgressionStatus") = True Then mainWorkbook.Names("ProgressionStatus").Delete
        If namedRanges("AugmentationModules") = True Then mainWorkbook.Names("AugmentationModules").Delete
        If namedRanges("RetrievedDate") = True Then mainWorkbook.Names("RetrievedDate").Delete
        If namedRanges("ScriptDuration") = True Then mainWorkbook.Names("ScriptDuration").Delete

        If namedRanges("ReportName") = True Then mainWorkbook.Names("ReportName").Delete
        If namedRanges("TemplateVersion") = True Then mainWorkbook.Names("TemplateVersion").Delete
        If namedRanges("FoundationCheckpoints") = True Then mainWorkbook.Names("FoundationCheckpoints").Delete
        If namedRanges("AugmentationCheckpoints") = True Then mainWorkbook.Names("AugmentationCheckpoints").Delete
        If namedRanges("ReportVision") = True Then mainWorkbook.Names("ReportVision").Delete
        If namedRanges("DependenciesList") = True Then mainWorkbook.Names("DependenciesList").Delete
        If namedRanges("CreationDate") = True Then mainWorkbook.Names("CreationDate").Delete
        If namedRanges("EditionName") = True Then mainWorkbook.Names("EditionName").Delete
        If namedRanges("DurationMilliseconds") = True Then mainWorkbook.Names("DurationMilliseconds").Delete
        If namedRanges("LogSummary") = True Then mainWorkbook.Names("LogSummary").Delete

        Set aboutWorksheet = mainWorkbook.Worksheets.Add(Before:=mainWorkbook.Worksheets(mainWorkbook.Worksheets.Count))
        aboutWorksheet.Name = "About"
    Else
        Set aboutWorksheet = mainWorkbook.Worksheets("About")

        If namedRanges("ReportName") = True And namedRanges("TemplateVersion") = True And namedRanges("FoundationCheckpoints") = True And namedRanges("AugmentationCheckpoints") = True And _
        namedRanges("ReportVision") = True And NamedRangeExists("DependenciesList") = True And namedRanges("CreationDate") = True And namedRanges("EditionName") = True And _
        namedRanges("DurationMilliseconds") = True And namedRanges("LogSummary") = True And namedRanges("ReportDetails") = False And namedRanges("TemplateDetails") = False And _
        namedRanges("ProgressionStatus") = False And namedRanges("AugmentationModules") = False And namedRanges("RetrievedDate") = False And namedRanges("ScriptDuration") = False Then

            If aboutWorksheet.Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & templateVersion & ")" Then
                If aboutWorksheet.Visible = xlSheetHidden Then
                    aboutWorksheet.Visible = xlSheetVisible
                End If

                Call LogConclusion("Skipped", logConclusionData)
                Exit Sub
            End If
        End If
    End If
    
    If namedRanges("ReportDetails") = True Then mainWorkbook.Names("ReportDetails").Delete
    If namedRanges("TemplateDetails") = True Then mainWorkbook.Names("TemplateDetails").Delete
    If namedRanges("ProgressionStatus") = True Then mainWorkbook.Names("ProgressionStatus").Delete
    If namedRanges("AugmentationModules") = True Then mainWorkbook.Names("AugmentationModules").Delete
    If namedRanges("RetrievedDate") = True Then mainWorkbook.Names("RetrievedDate").Delete
    If namedRanges("ScriptDuration") = True Then mainWorkbook.Names("ScriptDuration").Delete

    If namedRanges("ReportName") = False Then mainWorkbook.Names.Add Name:="ReportName", RefersTo:="=About!$A$1"
    If namedRanges("TemplateVersion") = False Then mainWorkbook.Names.Add Name:="TemplateVersion", RefersTo:="=About!$A$2"
    If namedRanges("FoundationCheckpoints") = False Then mainWorkbook.Names.Add Name:="FoundationCheckpoints", RefersTo:="=About!$A$3"
    If namedRanges("AugmentationCheckpoints") = False Then mainWorkbook.Names.Add Name:="AugmentationCheckpoints", RefersTo:="=About!$A$4"
    If namedRanges("ReportVision") = False Then mainWorkbook.Names.Add Name:="ReportVision", RefersTo:="=About!$B$1"
    If namedRanges("DependenciesList") = False Then mainWorkbook.Names.Add Name:="DependenciesList", RefersTo:="=About!$B$3"
    If namedRanges("CreationDate") = False Then mainWorkbook.Names.Add Name:="CreationDate", RefersTo:="=About!$C$1"
    If namedRanges("EditionName") = False Then mainWorkbook.Names.Add Name:="EditionName", RefersTo:="=About!$C$2"
    If namedRanges("DurationMilliseconds") = False Then mainWorkbook.Names.Add Name:="DurationMilliseconds", RefersTo:="=About!$C$3"
    If namedRanges("LogSummary") = False Then mainWorkbook.Names.Add Name:="LogSummary", RefersTo:="=About!$C$4"

    With aboutWorksheet
        If .Range("TemplateVersion").Value = "Spreadsheet Operations Template (v0.39, 16.02.2024)" Then
            Dim legacyRetrievedDate As String

            .Range("FoundationCheckpoints").Value = Replace(.Range("FoundationCheckpoints").Value, "Progression Status: ", "Foundation Checkpoints: ")
            .Range("AugmentationCheckpoints").Value = Replace(.Range("AugmentationCheckpoints").Value, "Augmentation Modules: ", "Augmentation Checkpoints: ")

            .Range("CreationDate").Value = Replace(.Range("CreationDate").Value, "Retrieved", "Creation")
            legacyRetrievedDate = GetAboutNamedRange("Creation Date")
            Call SetAboutNamedRange(Right(legacyRetrievedDate, 4) & "-" & Mid(legacyRetrievedDate, 4, 2) & "-" & Left(legacyRetrievedDate, 2), "Creation Date")

            .Range("DurationMilliseconds").Value = "Duration (Milliseconds): 0."
            .Range("LogSummary").Value = "Log Summary: 0 Runs. 0 Checkpoints. 0 Rows."
        End If

        If .Visible = xlSheetHidden Then
            .Visible = xlSheetVisible
        End If

        If report.Exists("Name") Then
            If .Range("ReportName").Value = "" Then .Range("ReportName").Value = report("Name")
        Else
            If .Range("ReportName").Value = "" Then .Range("ReportName").Value = "N/A"
        End If

        .Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & templateVersion & ")"

        If .Range("FoundationCheckpoints").Value = "" Then .Range("FoundationCheckpoints").Value = "Foundation Checkpoints: N/A."
        If .Range("AugmentationCheckpoints").Value = "" Then .Range("AugmentationCheckpoints").Value = "Augmentation Checkpoints: N/A."

        If report.Exists("Vision") Then
            If .Range("ReportVision").Value = "" Then .Range("ReportVision").Value = report("Vision")
        Else
            If .Range("ReportVision").Value = "" Then .Range("ReportVision").Value = "N/A."
        End If

        If report.Exists("Dependencies") Then
            If .Range("DependenciesList").Value = "" Then .Range("DependenciesList").Value = "Dependencies List: " & report("Dependencies") & "."
        Else
            If .Range("DependenciesList").Value = "" Then .Range("DependenciesList").Value = "Dependencies List: N/A."
        End If

        If report.Exists("Creation Date") = True Then
            If .Range("CreationDate").Value = "" Then .Range("CreationDate").Value = "Creation Date: " & report("Creation Date") & "."
        End If

        If report.Exists("Edition") Then
            If .Range("EditionName").Value = "" Then .Range("EditionName").Value = "Edition Name: " & report("Edition") & "."
        Else
            If .Range("EditionName").Value = "" Then .Range("EditionName").Value = "Edition Name: N/A."
        End If

        If .Range("DurationMilliseconds").Value = "" Then .Range("DurationMilliseconds").Value = "Duration (Milliseconds): 0."
        If .Range("LogSummary").Value = "" Then .Range("LogSummary").Value = "Log Summary: 0 Runs. 0 Checkpoints. 0 Rows."

        If .Range("B2").MergeArea.Cells.Count = 1 Then
            .Range("B1:B2").Merge
        End If

        If .Range("B4").MergeArea.Cells.Count = 1 Then
            .Range("B3:B4").Merge
        End If
    End With

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub CreateRunStatusWorksheet()
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CreateRunStatusWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("", methodName, "Logging")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, "")

    If WorksheetExists("Run Status") = True Then
        mainWorkbook.Worksheets("Run Status").Delete
    End if

    Set runStatusWorksheet = mainWorkbook.Worksheets.Add(Before:=mainWorkbook.Worksheets("About"))
    runStatusWorksheet.Name = "Run Status"

    With runStatusWorksheet
        .Range("A7").Value = "Checkpoints: N/A"
        .Range("A8").Value = "Date Runtime: N/A"

        .Range("A13").Resize(16).Value = Application.Transpose(Array( _
            "Run Identifier: " & environment("Run Identifier"), "Computer Name: " & environment("Computer Name"), "Username: " & environment("Username"), "Display Resolution: " & environment("Display Resolution"), _
            "DPI Scale: " & environment("DPI Scale"), "Operating System: " & environment("Operating System"), "User Interface Language Code Identifier: " & environment("User Interface Language Code Identifier"), _
            "Time Zone Key Name: " & environment("Time Zone Key Name"), "Time Zone UTC Offset: " & environment("Time Zone UTC Offset"), "QPC Frequency: " & environment("QPC Frequency"), _
            "Base QPC: " & report("Base QPC"), "Base Tick Count: " & baseTickCount, "Base UTC Timestamp: " & report("Base UTC Timestamp"), _
            "Telemetry QPC: " & telemetry("QPC Midpoint Timestamp"), "Telemetry Tick Count: " & telemetry("Tick Count"), "Telemetry UTC Timestamp: " & telemetry("UTC Timestamp Precise")))

        .Range("B1").Resize(28).Value = Application.Transpose(Array( _
            "Brackets and Braces", "Left Brace", "Left Bracket", "Lower Case Column Letter", "Lower Case Row Letter", "Right Brace", "Right Bracket", "Upper Case Column Letter", "Upper Case Row Letter", _
            "Country/Region Settings", "Country Code", "Country Setting", "General Format Name", _
            "Currency", "Currency Before", "Currency Code", "Currency Digits", "Currency Leading Zeros", "Currency Minus Sign", "Currency Negative", "Currency Space Before", "Currency Trailing Zeros", "Noncurrency Digits", _
            "Windows Locale Settings", "Display Language", "Regional Format", "Input Language", "Keyboard Layout"))

        .Range("C1").Resize(28).Value = Application.Transpose(Array( _
            "", international("Left Brace"), international("Left Bracket"), international("Lower Case Column Letter"), international("Lower Case Row Letter"), _
            international("Right Brace"), international("Right Bracket"), international("Upper Case Column Letter"), international("Upper Case Row Letter"), _
            "", international("Country Code"), international("Country Setting"), international("General Format Name"), _
            "", IIf(international("Currency Before"), "TRUE", "FALSE"), international("Currency Code"), international("Currency Digits"), IIf(international("Currency Leading Zeros"), "TRUE", "FALSE"), IIf(international("Currency Minus Sign"), "TRUE", "FALSE"), _
            international("Currency Negative"), IIf(international("Currency Space Before"), "TRUE", "FALSE"), IIf(international("Currency Trailing Zeros"), "TRUE", "FALSE"), international("Noncurrency Digits"), _
            "", environment("Display Language"), environment("Regional Format"), environment("Input Language"), environment("Keyboard Layout")))

        .Range("D1").Resize(28).Value = Application.Transpose(Array( _
            "Date and Time", "24 Hour Clock", "4 Digit Years", "Date Order", "Date Separator", "Day Code", "Day Leading Zero", "Hour Code", "MDY", "Minute Code", _
            "Month Code", "Month Leading Zero", "Month Name Chars", "Second Code", "Time Separator", "Time Leading Zero", "Weekday Name Chars", "Year Code", _
            "Measurement Systems", "Metric", "Non English Functions", _
            "Separators", "Alternate Array Separator", "Column Separator", "Decimal Separator", "List Separator", "Row Separator", "Thousands Separator"))

        .Range("E1").Resize(28).Value = Application.Transpose(Array( _
            "", IIf(international("24 Hour Clock"), "TRUE", "FALSE"), IIf(international("4 Digit Years"), "TRUE", "FALSE"), international("Date Order"), international("Date Separator"), international("Day Code"), _
            IIf(international("Day Leading Zero"), "TRUE", "FALSE"), international("Hour Code"), IIf(international("MDY"), "TRUE", "FALSE"), international("Minute Code"), international("Month Code"), IIf(international("Month Leading Zero"), "TRUE", "FALSE"), _
            international("Month Name Chars"), international("Second Code"), international("Time Separator"), IIf(international("Time Leading Zero"), "TRUE", "FALSE"), international("Weekday Name Chars"), international("Year Code"), _
            "", IIf(international("Metric"), "TRUE", "FALSE"), IIf(international("Non English Functions"), "TRUE", "FALSE"), _
            "", international("Alternate Array Separator"), international("Column Separator"), international("Decimal Separator"), international("List Separator"), international("Row Separator"), international("Thousands Separator")))

        If .Range("A1").MergeArea.Cells.Count = 1 Then .Range("A1:A3").Merge
        If .Range("A5").MergeArea.Cells.Count = 1 Then .Range("A5:A6").Merge
        If .Range("A9").MergeArea.Cells.Count = 1 Then .Range("A9:A12").Merge
        If .Range("B1").MergeArea.Cells.Count = 1 Then .Range("B1:C1").Merge
        If .Range("B10").MergeArea.Cells.Count = 1 Then .Range("B10:C10").Merge
        If .Range("B14").MergeArea.Cells.Count = 1 Then .Range("B14:C14").Merge
        If .Range("B24").MergeArea.Cells.Count = 1 Then .Range("B24:C24").Merge
        If .Range("D1").MergeArea.Cells.Count = 1 Then .Range("D1:E1").Merge
        If .Range("D19").MergeArea.Cells.Count = 1 Then .Range("D19:E19").Merge
        If .Range("D22").MergeArea.Cells.Count = 1 Then .Range("D22:E22").Merge
    End With

    Dim cellAddresses As Variant
    cellAddresses = Array("B1", "B10", "B14", "D1", "D19", "D22")

    Dim cellAddressIndex As Long
    For cellAddressIndex = LBound(cellAddresses) To UBound(cellAddresses)
        Call runStatusWorksheet.Hyperlinks.Add(Anchor:=runStatusWorksheet.Range(cellAddresses(cellAddressIndex)), Address:="https://learn.microsoft.com/en-us/office/vba/api/excel.application.international#remarks")
    Next cellAddressIndex

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub LogCheckpoint(ByVal checkpointType As String, ByVal checkpointName As String, ByVal checkpointStatus As String, ByVal reportName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "LogCheckpoint"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    Dim qpc As Currency: Call QueryPerformanceCounter(qpc)
    Dim utcTimestamp As String: utcTimestamp = GetUtcTimestamp()

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal checkpointType As String, ByVal checkpointName As String, ByVal checkpointStatus As String, ByVal reportName As String, Context qpc As Double, Context utcTimestamp As String", methodName, "Logging")
    End If

    If checkpointStatus = "Conclusion" Then
        Dim worksheetCount As Long
        worksheetCount = mainWorkbook.Worksheets.Count

        Dim lastWorksheetName As String
        Dim secondToLastWorksheetName As String

        lastWorksheetName = mainWorkbook.Worksheets(worksheetCount).Name
        If worksheetCount >= 2 Then
            secondToLastWorksheetName = mainWorkbook.Worksheets(worksheetCount - 1).Name
        End If

        If worksheetCount >= 2 Then
            If lastWorksheetName = "About" And secondToLastWorksheetName = "Log" Then
                Call MoveWorksheetToEnd("Log")
            ElseIf lastWorksheetName <> "Log" And secondToLastWorksheetName <> "About" Then
                Call MoveWorksheetToEnd("About")
                Call MoveWorksheetToEnd("Log")
            End If
        End If
    End If

    Dim validation As String
    Call ValidateWhitelist(checkpointType, "checkpointType", Array("Foundation", "Augmentation"), validation)
    Call ValidateWhitelist(checkpointStatus, "checkpointStatus", Array("Beginning", "Conclusion"), validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & checkpointType & """" & ", " & """" & checkpointName & """" & ", " & """" & checkpointStatus & """" & ", " & """" & reportName & """" & ", " & _
        CDbl(qpc * 10000) & ", " & """" & utcTimestamp & """", validation)

    Static checkpointBeginningTickCount As Double

    With runStatusWorksheet
        .Range("A1").Value = "Declaration: " & methodRegistry(logConclusionData.methodName)("Declaration")
        .Range("A4").Value = "Parameters: " & methodRegistry(logConclusionData.methodName)("Parameters")
        .Range("A5").Value = "Arguments: " & logConclusionData.arguments      
        .Range("A8").Value = "Date Runtime: " & utcTimestamp & " (UTC), " & CDbl(qpc * 10000)

        .Range("A4").Font.Size = 10
        .Range("A4").Characters(Start:=13).Font.Size = 8

        Call ApplyBoldBeforeFirstColon(.Range("A1:A8"))

        Dim lastSpacePosition As Long
        
        lastSpacePosition = InStrRev(.Range("A8").Value, " ")
        If lastSpacePosition > 0 Then
            .Range("A8").Characters(Start:=lastSpacePosition + 1).Font.Name = "Consolas"
        End If
    End With

    If checkpointStatus = "Beginning" Then
        checkpointBeginningTickCount = logConclusionData.tickCount

        With runStatusWorksheet
            If .Range("A7").Value = "Checkpoints: N/A" Then
                .Range("A7").Value = "Checkpoints: " & checkpointName
            Else
                .Range("A7").Value = Left$(.Range("A7").Value, Len(.Range("A7").Value) - 1) & ", " & checkpointName
            End If

            .Range("A7").Font.Bold = False
            .Range("A7").Characters(Start:=1, Length:=11).Font.Bold = True

            .Range("A9").Value = ""
            .Range("A9").Font.Bold = False
        End With

        mainWorkbook.Worksheets("Log").Visible = xlSheetVisible

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.EnableEvents = False
    End If

    Dim checkpointColorCode As Long

    If checkpointStatus = "Beginning" And checkpointType ="Foundation" Then checkpointColorCode = -16737281
    If checkpointStatus = "Conclusion" And checkpointType ="Foundation" Then checkpointColorCode = -10040320

    If checkpointStatus = "Beginning" And checkpointType ="Augmentation" Then checkpointColorCode = -13056
    If checkpointStatus = "Conclusion" And checkpointType ="Augmentation" Then checkpointColorCode = -1897831

    mainWorkbook.Worksheets("Log").Range("A" & logConclusionData.operationSequenceNumber & ":" & "B" & logConclusionData.operationSequenceNumber).Font.Color = checkpointColorCode

    Dim lastRowCheckpoint As Long
    lastRowCheckpoint = LastUsedRowNumberOnWorksheet("Log")
    If lastRowCheckpoint <= 30 Then lastRowCheckpoint = 31

    mainWorkbook.Worksheets("Log").Select
    Call Application.Goto(Reference:=ActiveSheet.Cells.SpecialCells(xlCellTypeVisible).Range("A" & (lastRowCheckpoint - 30)), Scroll:=True)
    mainWorkbook.Worksheets("Log").Range("A" & logConclusionData.operationSequenceNumber & ":B" & logConclusionData.operationSequenceNumber).Select

    If checkpointStatus = "Conclusion" Then
        If checkpointType = "Foundation" Then
            Call SetAboutNamedRange(checkpointName, "Foundation Checkpoints")
        ElseIf checkpointType = "Augmentation" Then
            Call SetAboutNamedRange(checkpointName, "Augmentation Checkpoints")
        End If

        Call SetAboutNamedRange("0 Runs. 1 Checkpoint. 0 Rows.", "Log Summary")

        With runStatusWorksheet
            .Range("A7").Value = .Range("A7").Value & "."

            .Range("A7").Font.Bold = False
            .Range("A7").Characters(Start:=1, Length:=11).Font.Bold = True
        End With
    End If

    Call LogConclusion("Completed", logConclusionData)

    If checkpointStatus = "Conclusion" Then
        Dim durationRange As Range
        Dim durationMilliseconds As Double
        Set durationRange = logWorksheet.Cells(logConclusionData.operationSequenceNumber, 5)
        durationMilliseconds = (logConclusionData.tickCount + durationRange.Value) - checkpointBeginningTickCount

        Call SetAboutNamedRange(CStr(durationMilliseconds), "Duration (Milliseconds)")

        With runStatusWorksheet
            .Range("A9").Value = "Success Output: The run started at " & report("Base UTC Timestamp") & " (UTC) and finished successfully at " & utcTimestamp & " (UTC). The entire run took " & (logConclusionData.tickCount + durationRange.Value) & " milliseconds."
            .Range("A9").Font.Bold = False
            .Range("A9").Characters(Start:=1, Length:=14).Font.Bold = True
        End With

        mainWorkbook.Worksheets("Log").Visible = xlSheetHidden

        checkpointBeginningTickCount = 0

        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
    End If

    With logWorksheet
        .Rows(logConclusionData.operationSequenceNumber).RowHeight = 33
        .Cells(logConclusionData.operationSequenceNumber, 2).WrapText = True
    End With
End Sub

Sub LogEngine()
If logEngineActive = True Then Exit Sub

    Const methodName As String = "LogEngine"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("Context templateVersion As String, Context runIdentifier As String, Context qpcFrequency As Double, Context baseQpc As Double, Context baseTickCount As Double, Context baseUtcTimestamp As String", methodName, "Logging")

        Call ConfigureMethodSetting(methodName, "Configure Core Cell Styles", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Add Custom Cell Styles", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Delete Empty Worksheets", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Format Log Worksheet", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Format About Worksheet", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Format Run Status Worksheet", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Delete Unnecessary Cell Styles", 1, 0, 1)
    End If

    Dim logWorksheetAlreadyExists As Boolean

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    If WorksheetExists("Log") = True Then
        With mainWorkbook.Worksheets("Log")
            Dim legacyLogHeaderValues As Variant
            legacyLogHeaderValues = Array("Task", "Arguments", "Category", "Status", "Active Worksheet", "Original Order", "Date Time Start", "Date Time End", "Stopwatch Start", "Stopwatch End")

            If Join(Application.Index(.Range("A1:J1").Value, 1, 0), "|") = Join(legacyLogHeaderValues, "|") Then
                .Delete
            Else
                Dim logHeaderValues As Variant
                logHeaderValues = Array("Method", "Arguments", "Depth", "Tick Count", "Duration", "Outcome")

                If Join(Application.Index(.Range("A1:F1").Value, 1, 0), "|") = Join(logHeaderValues, "|") Then
                    operationSequenceNumber = CDbl(LastUsedRowNumberOnWorksheet("Log"))
                    logWorksheetAlreadyExists = True
                Else
                    .Delete
                End If
            End If
        End With
    End If

    If logWorksheetAlreadyExists = True Then
        Set logWorksheet = mainWorkbook.Worksheets("Log")
    Else
        Set logWorksheet = mainWorkbook.Worksheets.Add(After:=mainWorkbook.Worksheets(mainWorkbook.Worksheets.Count))
        logWorksheet.Name = "Log"

        logWorksheet.Cells(1, 1).Resize(1, 6).Value = Array("Method", "Arguments", "Depth", "Tick Count", "Duration", "Outcome")
    End If
  
    Call LogBeginning(methodName, CCur(baseTickCount / 10000), logConclusionData, """" & templateVersion & """" & ", " & """" & environment("Run Identifier") & """" & ", " & _
        environment("QPC Frequency") & ", " & report("Base QPC") & ", " & baseTickCount & ", " & """" & report("Base UTC Timestamp") & """")

    Dim settings As Object

    Set settings = methodRegistry(methodName)("Settings")

    If report("Settings").Count <> 0 Then
        Dim settingArray As Variant
        For Each settingArray In report("Settings")           
            Call SetMethodSetting(settingArray(0), settingArray(1), settingArray(2))
        Next settingArray
    End If

    If settings("Configure Core Cell Styles")("Value") = True Then
        Dim normalCellStyle As String: normalCellStyle = cellStyles("Normal")
        If CellStyleExists(normalCellStyle) = False Then
            normalCellStyle = "Normal"
        End If

        With mainWorkbook.Styles(normalCellStyle)
            .Font.ColorIndex = xlAutomatic
            .Font.Name = "Arial"
            .Font.Size = 10
            .IncludeBorder = False
            .IncludePatterns = False
            .IncludeProtection = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        mainWorkbook.Styles(normalCellStyle).NumberFormat = "@"

        Dim percentCellStyle As String: percentCellStyle = "Percent"
        If CellStyleExists(percentCellStyle) = False Then
            percentCellStyle = cellStyles("Percent")
        End If

        If CellStyleExists(percentCellStyle) = False Then
            Call mainWorkbook.Styles.Add(Name:=percentCellStyle)
        End If

        With mainWorkbook.Styles(percentCellStyle)
            .IncludeAlignment = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ReadingOrder = xlContext
        End With
        mainWorkbook.Styles(percentCellStyle).NumberFormat = "0.00 %"

        Dim hyperlinkCellStyle As String: hyperlinkCellStyle = "Hyperlink"
        If CellStyleExists(hyperlinkCellStyle) = False And CellStyleExists(cellStyles("Hyperlink")) = False Then
            Call logWorksheet.Hyperlinks.Add(logWorksheet.Range("B1"), "", "'" & logWorksheet.Name & "'!A1")
        End If

        If CellStyleExists(hyperlinkCellStyle) = False Then
            hyperlinkCellStyle = cellStyles("Hyperlink")
        End If

        With mainWorkbook.Styles(hyperlinkCellStyle)
            .IncludeAlignment = True
            .IncludeNumber = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ReadingOrder = xlContext
        End With
        mainWorkbook.Styles(hyperlinkCellStyle).NumberFormat = "@"

        Call logWorksheet.Range("B1").ClearHyperlinks

        If CellStyleExists("Followed Hyperlink") = False And CellStyleExists(cellStyles("Followed Hyperlink")) = False Then
            Dim cellStyleEntry As Style

            Dim cellStyleIndex As Integer
            For cellStyleIndex = mainWorkbook.Styles.Count To 1 Step -1
                Set cellStyleEntry = mainWorkbook.Styles(cellStyleIndex)
                
                If cellStyleEntry.Font.Color = 8216726 And cellStyleEntry.Font.Underline = 2 Then
                    Call DeleteCellStyle(cellStyleEntry)
                    
                    Exit For
                End If
            Next cellStyleIndex

            With mainWorkbook.Styles.Add(Name:=cellStyles("Followed Hyperlink"))
                .Font.Color = 8216726
                .Font.Underline = 2
            End With
        End If
    End If

    If settings("Add Custom Cell Styles")("Value") = True Then
        Dim customCellStylesArray As Variant: customCellStylesArray = Array("Date", "Date Time", "Decimal", "Formula", "Header", "Integer")
        Dim customCellStyle As String

        Dim customCellStyleIndex As Integer
        For customCellStyleIndex = LBound(customCellStylesArray) To UBound(customCellStylesArray)
            customCellStyle = customCellStylesArray(customCellStyleIndex)

            If CellStyleExists(customCellStyle) = False Then
                Call mainWorkbook.Styles.Add(Name:=customCellStyle)

                If customCellStyle = "Date" Then
                    mainWorkbook.Styles("Date").NumberFormat = "yyyy-mm-dd"
                End If

                If customCellStyle = "Date Time" Then
                    mainWorkbook.Styles("Date Time").NumberFormat = "yyyy-mm-dd HH:mm:ss"
                End If

                If customCellStyle = "Decimal" Then
                    mainWorkbook.Styles("Decimal").NumberFormat = "0.00"
                End If

                If customCellStyle = "Formula" Then
                    mainWorkbook.Styles("Formula").NumberFormat = "General"
                End If

                If customCellStyle = "Header" Then
                    With mainWorkbook.Styles("Header")
                        .Font.Bold = True
                        .WrapText = True
                    End With
                End If

                If customCellStyle = "Integer" Then
                    With mainWorkbook.Styles("Integer").Font
                        .Name = "Consolas"
                    End With
                    mainWorkbook.Styles("Integer").NumberFormat = "0"
                End If
            End If
        Next customCellStyleIndex
    End If

    If settings("Delete Empty Worksheets")("Value") = True Then
        Dim worksheetEntry As Worksheet

        For Each worksheetEntry In mainWorkbook.Worksheets
            If WorksheetIsEmpty(worksheetEntry.Name) Then
                mainWorkbook.Worksheets(worksheetEntry.Name).Delete
            End If
        Next worksheetEntry
    End If

    If WorksheetExists("About") = False Then
        Call CreateAboutWorksheet()
    Else
        Dim namedRanges As Object
        Set namedRanges = CreateObject("Scripting.Dictionary")

        namedRanges("ReportName")              = NamedRangeExists("ReportName")
        namedRanges("TemplateVersion")         = NamedRangeExists("TemplateVersion")
        namedRanges("FoundationCheckpoints")   = NamedRangeExists("FoundationCheckpoints")
        namedRanges("AugmentationCheckpoints") = NamedRangeExists("AugmentationCheckpoints")
        namedRanges("ReportVision")            = NamedRangeExists("ReportVision")
        namedRanges("DependenciesList")        = NamedRangeExists("DependenciesList")
        namedRanges("CreationDate")            = NamedRangeExists("CreationDate")
        namedRanges("EditionName")             = NamedRangeExists("EditionName")
        namedRanges("DurationMilliseconds")    = NamedRangeExists("DurationMilliseconds")
        namedRanges("LogSummary")              = NamedRangeExists("LogSummary")

        Dim currentNamedRangeName As Variant
        For Each currentNamedRangeName In namedRanges.Keys
            If namedRanges(currentNamedRangeName) = False Then
                Call CreateAboutWorksheet()
                Exit For
            End If
        Next currentNamedRangeName

        If mainWorkbook.Worksheets("About").Range("TemplateVersion").Value <> "Spreadsheet Operations Template (" & templateVersion & ")" Then
            Call CreateAboutWorksheet()
        End If
    End If

    Set aboutWorksheet = mainWorkbook.Worksheets("About")
    Call SetAboutNamedRange("1 Run. 0 Checkpoints. 0 Rows.", "Log Summary")

    Call CreateRunStatusWorksheet()
    Set runStatusWorksheet = mainWorkbook.Worksheets("Run Status")
    With runStatusWorksheet
        .Range("A1").Value = "Declaration: " & methodRegistry(logConclusionData.methodName)("Declaration")
        .Range("A4").Value = "Parameters: " & methodRegistry(logConclusionData.methodName)("Parameters")
        .Range("A5").Value = "Arguments: " & logConclusionData.arguments
        .Range("A8").Value = "Date Runtime: " & report("Base UTC Timestamp") & " (UTC), " & report("Base QPC")
    End With

    If settings("Format Log Worksheet")("Value") = True Then
        If logWorksheetAlreadyExists = False Then
            Call FormatCoreWorksheet("Log")
        End If
    End If

    If settings("Format About Worksheet")("Value") = True Then
        Call FormatCoreWorksheet("About")
    End If

    If settings("Format Run Status Worksheet")("Value") = True Then
        Call FormatCoreWorksheet("Run Status")
    End If

    If settings("Delete Unnecessary Cell Styles")("Value") = True Then
        Dim cellStyleKey As Variant
        For Each cellStyleKey In cellStyles.Keys
            Select Case cellStyleKey
                Case "Followed Hyperlink", "Hyperlink", "Normal", "Percent"
                    ' Core styles to keep.
                Case Else
                    If CellStyleExists(cellStyleKey) = True Or CellStyleExists(cellStyles(cellStyleKey)) = True Then
                        Call DeleteCellStyle(cellStyleKey)
                    End If
            End Select
        Next cellStyleKey
    End If

    logEngineActive = True

    Call LogConclusion("Completed", logConclusionData)

    Call SetAboutNamedRange(CStr(logWorksheet.Cells(logConclusionData.operationSequenceNumber, 5).Value), "Duration (Milliseconds)")

    With logWorksheet
        .Rows(logConclusionData.operationSequenceNumber).RowHeight = 33
        .Cells(logConclusionData.operationSequenceNumber, 2).WrapText = True
    End With
End Sub

' Functions: Logging '

Function ParseMethodArgumentsIntoValues(ByVal argument As String) As Variant
    Const methodName As String = "ParseMethodArgumentsIntoValues"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal argument As String", methodName, "Logging")
    End If

    Dim textQualifier As String
    Dim argumentLength As Long
    Dim argumentValues() As String
    Dim collectedValues As Collection
    Dim unquotedParts As Variant
    Dim partIndex As Long
    Dim valueCount As Long
    Dim valueIndex As Long
    Dim characterIndex As Long
    Dim scanIndex As Long
    Dim nextCharacter As String
    Dim qualifierCloseIndex As Long
    Dim currentValue As String
    Dim usedQuotedValue As Boolean

    textQualifier = """"
    argumentLength = Len(argument)

    If argumentLength = 0 Then
        ReDim argumentValues(1 To 1)
        argumentValues(1) = ""
    ElseIf InStr(1, argument, textQualifier, vbBinaryCompare) = 0 Then
        unquotedParts = Split(argument, delimiter)
        valueCount = UBound(unquotedParts) - LBound(unquotedParts) + 1
        ReDim argumentValues(1 To valueCount)
        For partIndex = LBound(unquotedParts) To UBound(unquotedParts)
            argumentValues(partIndex - LBound(unquotedParts) + 1) = unquotedParts(partIndex)
        Next partIndex
    Else
        Set collectedValues = New Collection
        characterIndex = 1
        Do While characterIndex <= argumentLength
            usedQuotedValue = False
            qualifierCloseIndex = 0
            If Mid$(argument, characterIndex, 1) = textQualifier Then
                scanIndex = characterIndex + 1
                Do While scanIndex <= argumentLength
                    If Mid$(argument, scanIndex, 1) = textQualifier Then
                        If scanIndex < argumentLength Then
                            nextCharacter = Mid$(argument, scanIndex + 1, 1)
                            If nextCharacter = textQualifier Then
                                scanIndex = scanIndex + 2
                            ElseIf nextCharacter = delimiter Then
                                qualifierCloseIndex = scanIndex
                                Exit Do
                            Else
                                scanIndex = scanIndex + 1
                            End If
                        Else
                            qualifierCloseIndex = scanIndex
                            Exit Do
                        End If
                    Else
                        scanIndex = scanIndex + 1
                    End If
                Loop
            End If

            If qualifierCloseIndex > 0 Then
                currentValue = Mid$(argument, characterIndex, qualifierCloseIndex - characterIndex + 1)

                Call collectedValues.Add(currentValue)
                characterIndex = qualifierCloseIndex + 1
                usedQuotedValue = True

                If characterIndex <= argumentLength Then
                    If Mid$(argument, characterIndex, 1) = delimiter Then
                        characterIndex = characterIndex + 1
                        If characterIndex > argumentLength Then
                            Call collectedValues.Add("")
                        End If
                    End If
                End If
            End If

            If usedQuotedValue = False Then
                currentValue = ""
                Do While characterIndex <= argumentLength
                    If Mid$(argument, characterIndex, 1) = delimiter Then
                        Call collectedValues.Add(currentValue)
                        characterIndex = characterIndex + 1
                        If characterIndex > argumentLength Then
                            Call collectedValues.Add("")
                        End If
                        Exit Do
                    End If
                    currentValue = currentValue & Mid$(argument, characterIndex, 1)
                    characterIndex = characterIndex + 1
                    If characterIndex > argumentLength Then
                        Call collectedValues.Add(currentValue)
                    End If
                Loop
            End If
        Loop

        valueCount = collectedValues.Count

        ReDim argumentValues(1 To valueCount)
        For valueIndex = 1 To valueCount
            argumentValues(valueIndex) = collectedValues(valueIndex)
        Next valueIndex
    End If

    ParseMethodArgumentsIntoValues = argumentValues
End Function

' Core: Logging '

Private Sub LogBeginning(ByVal methodName As String, ByVal tickCount As Currency, ByRef logConclusionData As LogEntry, ByVal arguments As String, Optional ByVal errorMessage As String)
    operationSequenceNumber = operationSequenceNumber + 1

    logConclusionData.operationSequenceNumber = operationSequenceNumber
    logConclusionData.methodName = methodName
    logConclusionData.arguments = arguments
    logConclusionData.tickCount = CDbl(tickCount * 10000) - baseTickCount

    logWorksheet.Cells(logConclusionData.operationSequenceNumber, 1).Resize(1, 6).Value = Array(logConclusionData.methodName, logConclusionData.arguments, depth, logConclusionData.tickCount, "", "Beginning")

    depth = depth + 1

    If errorMessage <> "" Then
        Call LogConclusion("Failed", logConclusionData, errorMessage)
    End If
End Sub

Private Sub LogConclusion(ByVal conclusionStatus As String, ByRef logConclusionData As LogEntry, Optional ByVal errorMessage As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Dim duration As Double

    If errorMessage <> "" Then       
        Dim qpc As Currency: Call QueryPerformanceCounter(qpc)
        Dim utcTimestamp As String: utcTimestamp = GetUtcTimestamp()

        duration = (CDbl(tickCount * 10000) - baseTickCount) - logConclusionData.tickCount
        logWorksheet.Cells(logConclusionData.operationSequenceNumber, 5).Resize(1, 2).Value = Array(duration, conclusionStatus)

        Call SetAboutNamedRange("0 Runs. 0 Checkpoint. 0 Rows.", "Log Summary")

        Dim lastRowCheckpoint As Long
        Dim horizontalRange As Range
        Dim firstRowBeginning As Long
        Dim verticalRange As Range
        Dim finalSelection As Range

        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True

        lastRowCheckpoint = logConclusionData.operationSequenceNumber
        If lastRowCheckpoint <= 30 Then lastRowCheckpoint = 31             

        logWorksheet.Select
        Call Application.Goto(Reference:=ActiveSheet.Cells.SpecialCells(xlCellTypeVisible).Range("A" & (lastRowCheckpoint - 30)), Scroll:=True)

        Set horizontalRange = logWorksheet.Range("A" & logConclusionData.operationSequenceNumber & ":F" & logConclusionData.operationSequenceNumber)

        Dim currentRowNumber As Long
        For currentRowNumber = 2 To logConclusionData.operationSequenceNumber
            If StrComp(CStr(logWorksheet.Cells(currentRowNumber, "F").Value), "Beginning", vbBinaryCompare) = 0 Then
                firstRowBeginning = logWorksheet.Cells(currentRowNumber, "F").Row
                Exit For
            End If
        Next currentRowNumber

        If firstRowBeginning <> 0 Then
            Set verticalRange = logWorksheet.Range("F" & firstRowBeginning & ":F" & logConclusionData.operationSequenceNumber)
            
            Set finalSelection = Union(horizontalRange, verticalRange)
        Else
            Set finalSelection = horizontalRange
        End If

        finalSelection.Select
        finalSelection.Font.Color = RGB(244, 28, 80)

        With runStatusWorksheet
            .Tab.Color = RGB(244, 28, 80)

            .Range("A1").Value = "Declaration: " & methodRegistry(logConclusionData.methodName)("Declaration")

            If methodRegistry(logConclusionData.methodName).Exists("Contract") Then
                .Range("A4").Value = "Parameters: " & methodRegistry(logConclusionData.methodName)("Parameters")
                .Range("A5").Value = "Arguments: " & logConclusionData.arguments
            Else
                .Range("A4").Value = "Parameters: N/A"
                .Range("A5").Value = "Arguments: N/A"
            End If

            .Range("A8").Value = "Date Runtime: " & utcTimestamp & " (UTC), " & CDbl(qpc * 10000)
            .Range("A9").Value = "Error Output: " & errorMessage

            .Range("A4").Font.Size = 10
            .Range("A4").Characters(Start:=13).Font.Size = 8

            Call ApplyBoldBeforeFirstColon(.Range("A1:A12"))

            Dim lastSpacePosition As Long
            
            lastSpacePosition = InStrRev(.Range("A8").Value, " ")
            If lastSpacePosition > 0 Then
                .Range("A8").Characters(Start:=lastSpacePosition + 1).Font.Name = "Consolas"
            End If
        End With

        Call mainWorkbook.Worksheets("Run Status").Move(After:=mainWorkbook.Worksheets("Log"))
        logWorksheet.Select

        errorMessage = methodRegistry(logConclusionData.methodName)("Declaration") & ". " & errorMessage
        Err.Raise 1000, Description:=errorMessage
    End If

    duration = (CDbl(tickCount * 10000) - baseTickCount) - logConclusionData.tickCount
    logWorksheet.Cells(logConclusionData.operationSequenceNumber, 5).Resize(1, 2).Value = Array(duration, conclusionStatus)

    depth = depth - 1
End Sub

' ************ '
' Sequencing   '
' ************ '

Sub AppendGroupBasedOnFormulaOnColumnOnWorksheet(ByVal groupName As String, ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "AppendGroupBasedOnFormulaOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal groupName As String, ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal columnName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateRequiredText(groupName, "groupName", validation, 1)
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & groupName & """" & ", " & """" & formulaToApply & """" & ", " & IIf(useLocalFormula, "TRUE", "FALSE") & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim groupColumnLetter As String
    Dim lastUsedRowNumber As Long
    Dim formulaColumnLetter As String
    Dim groupColumnRange As Range
    Dim formulaColumnRange As Range
    Dim groupListValues As Variant
    Dim formulaResultValues As Variant
    Dim temporarySingleCellValues(1 To 1, 1 To 1) As Variant
    Dim rowIndex As Long
    Dim formulaResultValue As Variant
    Dim isRowSelected As Boolean
    Dim currentGroupList As String

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    groupColumnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)

    If lastUsedRowNumber = 1 Then
        Call LogConclusion("Failed", logConclusionData, "No data available in worksheet.")
    End If

    Call AddColumnOnWorksheet(formulaColumn, 7.00, worksheetName)
    Call ApplyFormulaToColumnOnWorksheet(formulaToApply, useLocalFormula, formulaColumn, worksheetName)
    formulaColumnLetter = FindColumnLetterOnWorksheet(formulaColumn, worksheetName)

    Set groupColumnRange = worksheet.Range(groupColumnLetter & "2:" & groupColumnLetter & lastUsedRowNumber)
    Set formulaColumnRange = worksheet.Range(formulaColumnLetter & "2:" & formulaColumnLetter & lastUsedRowNumber)

    groupListValues = groupColumnRange.Value
    formulaResultValues = formulaColumnRange.Value

    If IsArray(groupListValues) = False Then
        temporarySingleCellValues(1, 1) = groupListValues
        groupListValues = temporarySingleCellValues
    End If
    If IsArray(formulaResultValues) = False Then
        temporarySingleCellValues(1, 1) = formulaResultValues
        formulaResultValues = temporarySingleCellValues
    End If

    For rowIndex = 1 To UBound(formulaResultValues, 1)
        formulaResultValue = formulaResultValues(rowIndex, 1)
        isRowSelected = False

        If IsError(formulaResultValue) = False Then
            If VarType(formulaResultValue) = vbBoolean Then
                isRowSelected = formulaResultValue
            ElseIf IsNumeric(formulaResultValue) = True Then
                isRowSelected = (CDbl(formulaResultValue) <> 0)
            End If
        End If

        If isRowSelected = True Then
            If IsError(groupListValues(rowIndex, 1)) = False Then
                If IsNull(groupListValues(rowIndex, 1)) = True Then
                    currentGroupList = ""
                Else
                    currentGroupList = CStr(groupListValues(rowIndex, 1))
                End If

                If Len(currentGroupList) = 0 Then
                    groupListValues(rowIndex, 1) = ";" & groupName & ";"
                ElseIf Right$(currentGroupList, 1) = ";" Then
                    groupListValues(rowIndex, 1) = currentGroupList & groupName & ";"
                Else
                    groupListValues(rowIndex, 1) = currentGroupList & ";" & groupName & ";"
                End If
            End If
        End If
    Next rowIndex

    groupColumnRange.Value = groupListValues
    Call DeleteColumnOnWorksheet(formulaColumn, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub AssignGroupBasedOnFormulaOnColumnOnWorksheet(ByVal groupName As String, ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "AssignGroupBasedOnFormulaOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal groupName As String, ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal columnName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateRequiredText(groupName, "groupName", validation, 1)
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & groupName & """" & ", " & """" & formulaToApply & """" & ", " & IIf(useLocalFormula, "TRUE", "FALSE") & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim groupColumnLetter As String
    Dim lastUsedRowNumber As Long
    Dim formulaColumnLetter As String
    Dim groupColumnRange As Range
    Dim formulaColumnRange As Range
    Dim groupColumnValues As Variant
    Dim formulaResultValues As Variant
    Dim temporarySingleCellValues(1 To 1, 1 To 1) As Variant
    Dim rowIndex As Long
    Dim formulaResultValue As Variant
    Dim isRowSelected As Boolean

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    groupColumnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)

    If lastUsedRowNumber = 1 Then
        Call LogConclusion("Failed", logConclusionData, "No data available in worksheet.")
    End If

    Call AddColumnOnWorksheet(formulaColumn, 7.00, worksheetName)
    Call ApplyFormulaToColumnOnWorksheet(formulaToApply, useLocalFormula, formulaColumn, worksheetName)
    formulaColumnLetter = FindColumnLetterOnWorksheet(formulaColumn, worksheetName)

    Set groupColumnRange = worksheet.Range(groupColumnLetter & "2:" & groupColumnLetter & lastUsedRowNumber)
    Set formulaColumnRange = worksheet.Range(formulaColumnLetter & "2:" & formulaColumnLetter & lastUsedRowNumber)

    groupColumnValues = groupColumnRange.Value
    formulaResultValues = formulaColumnRange.Value

    If IsArray(groupColumnValues) = False Then
        temporarySingleCellValues(1, 1) = groupColumnValues
        groupColumnValues = temporarySingleCellValues
    End If
    If IsArray(formulaResultValues) = False Then
        temporarySingleCellValues(1, 1) = formulaResultValues
        formulaResultValues = temporarySingleCellValues
    End If

    For rowIndex = 1 To UBound(formulaResultValues, 1)
        formulaResultValue = formulaResultValues(rowIndex, 1)
        isRowSelected = False

        If IsError(formulaResultValue) = False Then
            If VarType(formulaResultValue) = vbBoolean Then
                isRowSelected = formulaResultValue
            ElseIf IsNumeric(formulaResultValue) = True Then
                isRowSelected = (CDbl(formulaResultValue) <> 0)
            End If
        End If

        If isRowSelected = True Then
            If IsError(groupColumnValues(rowIndex, 1)) = False Then
                If IsNull(groupColumnValues(rowIndex, 1)) = True Then
                    groupColumnValues(rowIndex, 1) = groupName
                ElseIf Len(CStr(groupColumnValues(rowIndex, 1))) = 0 Then
                    groupColumnValues(rowIndex, 1) = groupName
                End If
            End If
        End If
    Next rowIndex

    groupColumnRange.Value = groupColumnValues
    Call DeleteColumnOnWorksheet(formulaColumn, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub DeleteRowsBasedOnFormulaOnWorksheet(ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteRowsBasedOnFormulaOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaToApply As String, ByVal useLocalFormula As Boolean, ByVal worksheetName As String", methodName, "Sequencing")

        Call ConfigureMethodSetting(methodName, "Number of Rows per Delete Batch", 8192, 1, 81920)
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaToApply & """" & ", " & IIf(useLocalFormula, "TRUE", "FALSE") & ", " & """" & worksheetName & """", validation)

    Call ValidateColumnOnWorksheet(helperColumn, "helperColumn", worksheetName, "worksheetName", validation, True)
    Call ValidateColumnOnWorksheet(sortingColumn, "sortingColumn", worksheetName, "worksheetName", validation, True)
    validation = Replace(validation, "Parameter """, "Variable """)

    If validation <> "" Then
        Call LogConclusion("Failed", logConclusionData, validation)
    End If

    Dim settings As Object
    Dim worksheet As Worksheet
    Dim lastUsedRowNumber As Long

    Set settings = methodRegistry(methodName)("Settings")
    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)

    If lastUsedRowNumber = 1 Then
        Call LogConclusion("Skipped", logConclusionData)
    Else
        Dim formulaColumnLetter As String
        Dim helperColumnLetter As String
        Dim trueCount As Double

        Call AddColumnOnWorksheet(helperColumn & "|" & formulaColumn & "|" & sortingColumn, 7.00, worksheetName)
        formulaColumnLetter = FindColumnLetterOnWorksheet(formulaColumn, worksheetName)
        helperColumnLetter = FindColumnLetterOnWorksheet(helperColumn, worksheetName)
        Call ApplyFormulaToColumnOnWorksheet(formulaToApply, useLocalFormula, formulaColumn, worksheetName)
        Call ApplyFormulaToColumnOnWorksheet("=" & formulaColumnLetter & "2*1", False, helperColumn, worksheetName)
        trueCount = Application.WorksheetFunction.Sum(worksheet.Range(helperColumnLetter & "2:" & helperColumnLetter & lastUsedRowNumber))

        If trueCount >= 1 Then
            Dim firstOneRowNumber As Long
            Dim lastOneRowNumber As Long
            Dim batchTopRowNumber As Long

            Call ApplyFormulaToColumnOnWorksheet("=ROWS($A$2:A2)", False, sortingColumn, worksheetName)
            Call SortColumnByOrderOnWorksheet(helperColumn, "Ascending", worksheetName)

            firstOneRowNumber = lastUsedRowNumber - CLng(trueCount) + 1
            lastOneRowNumber = lastUsedRowNumber

            Do While lastOneRowNumber >= firstOneRowNumber
                batchTopRowNumber = lastOneRowNumber - settings("Number of Rows per Delete Batch")("Value") + 1
                If batchTopRowNumber < firstOneRowNumber Then
                    batchTopRowNumber = firstOneRowNumber
                End If

                worksheet.Rows(batchTopRowNumber & ":" & lastOneRowNumber).Delete
                lastOneRowNumber = batchTopRowNumber - 1
            Loop

            lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)
            If lastUsedRowNumber <> 1 Then
                Call SortColumnByOrderOnWorksheet(sortingColumn, "Ascending", worksheetName)
            End If
        End If

        Call DeleteColumnOnWorksheet(helperColumn & "|" & formulaColumn & "|" & sortingColumn, worksheetName)

        Call LogConclusion("Completed", logConclusionData)
    End If
End Sub

' Functions: Sequencing'

' ************ '
' Background   '
' ************ '

Sub ConfigureMethodSetting(ByVal methodName As String, ByVal settingName As String, ByVal settingValue As Long, Optional ByVal floor As Long, Optional ByVal ceiling As Long)
    Dim setMethodSettingOnly As Boolean
    Dim methodDictionary As Object
    Dim methodSettingsDictionary As Object
    Dim methodSubSettingDictionary As Object

    If floor = 0 And ceiling = 0 Then
        setMethodSettingOnly = True
    End If

    If methodRegistry.Exists(methodName) = False Then
        Set methodDictionary = CreateObject("Scripting.Dictionary")
        Set methodRegistry(methodName) = methodDictionary
    Else
        Set methodDictionary = methodRegistry(methodName)
    End If
	
    If methodDictionary.Exists("Settings") = False Then
        Set methodSettingsDictionary = CreateObject("Scripting.Dictionary")
        Set methodDictionary("Settings") = methodSettingsDictionary
    Else
        Set methodSettingsDictionary = methodDictionary("Settings")
    End If

    If methodSettingsDictionary.Exists(settingName) = False Then
        Set methodSubSettingDictionary = CreateObject("Scripting.Dictionary")
        Set methodSettingsDictionary(settingName) = methodSubSettingDictionary
    Else
        Set methodSubSettingDictionary = methodSettingsDictionary(settingName)
    End If

    If setMethodSettingOnly = False Then
        methodSubSettingDictionary("Default") = settingValue
        methodSubSettingDictionary("Floor")   = floor
        methodSubSettingDictionary("Ceiling") = ceiling
    End If

    If methodSubSettingDictionary.Exists("Value") = False Then
        methodSubSettingDictionary("Value") = settingValue
    Else
        If setMethodSettingOnly = True Then
            methodSubSettingDictionary("Value") = settingValue
        End If
    End If


    If methodSubSettingDictionary.Exists("Default") Then
        If VarType(methodSubSettingDictionary("Default")) = vbBoolean Then
            methodSubSettingDictionary("Default") = Abs(CLng(methodSubSettingDictionary("Default")))
        End If

        If VarType(methodSubSettingDictionary("Floor")) = vbBoolean Then
            methodSubSettingDictionary("Floor") = Abs(CLng(methodSubSettingDictionary("Floor")))
        End If

        If VarType(methodSubSettingDictionary("Ceiling")) = vbBoolean Then
            methodSubSettingDictionary("Ceiling") = Abs(CLng(methodSubSettingDictionary("Ceiling")))
        End If

        If VarType(methodSubSettingDictionary("Value")) = vbBoolean Then
            methodSubSettingDictionary("Value") = Abs(CLng(methodSubSettingDictionary("Value")))
        End If

        If methodSubSettingDictionary("Value") > methodSubSettingDictionary("Ceiling") Then
            methodSubSettingDictionary("Value") = methodSubSettingDictionary("Ceiling")
        ElseIf methodSubSettingDictionary("Value") < methodSubSettingDictionary("Floor") Then
            methodSubSettingDictionary("Value") = methodSubSettingDictionary("Floor")
        End If
    End If

    If methodSubSettingDictionary.Exists("Default") Then
        If methodSubSettingDictionary("Floor") = 0 And methodSubSettingDictionary("Ceiling") = 1 Then
            methodSubSettingDictionary("Default") = CBool(methodSubSettingDictionary("Default"))
            methodSubSettingDictionary("Floor") = CBool(methodSubSettingDictionary("Floor"))
            methodSubSettingDictionary("Ceiling") = CBool(methodSubSettingDictionary("Ceiling"))
            methodSubSettingDictionary("Value") = CBool(methodSubSettingDictionary("Value"))
        End If
    End If
End Sub

Sub Intermission(ByVal intermissionsArray As Variant, ByVal checkpointName As String)
If IsArray(intermissionsArray) = False Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' https://learn.microsoft.com/en-us/office/vba/api/excel.range.replace
    mainWorkbook.Worksheets("About").Range("A1").Replace What:="", Replacement:="", LookAt:=xlPart

    Dim intermissionState As String

    Dim index As Integer
    For index = LBound(intermissionsArray) To UBound(intermissionsArray)
        intermissionState = intermissionsArray(index)

        Select Case intermissionState
            Case "Break Script", "Break", "BS", "B"
                End
            Case "Delete About Names", "DAN"
                Dim aboutNamedRanges As String: aboutNamedRanges = "AugmentationCheckpoints|AugmentationModules|CreationDate|DependenciesList|DurationMilliseconds|EditionName|FoundationCheckpoints|LogSummary|" & _
                    "ProgressionStatus|ReportDetails|ReportName|ReportVision|RetrievedDate|ScriptDuration|TemplateDetails|TemplateVersion"
                Dim aboutNamedRangesArray() As String: aboutNamedRangesArray = Split(aboutNamedRanges, "|")
                Dim aboutNamedRange As String

                Dim aboutNameIndex As Integer
                For aboutNameIndex = LBound(aboutNamedRangesArray) To UBound(aboutNamedRangesArray)
                    aboutNamedRange = aboutNamedRangesArray(aboutNameIndex)

                    If NamedRangeExists(aboutNamedRange) Then
                        mainWorkbook.Names(aboutNamedRange).Delete
                    End If
                Next aboutNameIndex
            Case "Delete Worksheet Log", "DWL"
                If WorksheetExists("Log") Then
                    mainWorkbook.Worksheets("Log").Delete
                End If
            Case "Delete Worksheet Run Status", "DWRS"
                If WorksheetExists("Run Status") Then
                    mainWorkbook.Worksheets("Run Status").Delete
                End If
            Case "Duplicate Workbook", "Duplicate", "DW", "D"
                mainWorkbook.SaveCopyAs Left(mainWorkbook.FullName, Len(mainWorkbook.FullName) - 5) & " (" & checkpointName & ")" & ".xlsx"
            Case "End Workbook", "End", "EW", "E"
                mainWorkbook.Close
            Case "Open Workbook", "Open", "OW", "O"
                Dim closingWorkbook As Workbook
                Set closingWorkbook = ActiveWorkbook
                
                Workbooks.Open checkpointName
                Set mainWorkbook = ActiveWorkbook
                closingWorkbook.Close
            Case "Quit Excel", "Quit", "QE", "Q"
                Excel.Application.Quit
                Workbooks(2).Close SaveChanges:=False
                Workbooks(1).Close SaveChanges:=False
            Case "Reset View", "Reset", "RW", "R"
                Dim worksheetCount As Long: worksheetCount = mainWorkbook.Worksheets.Count
                Dim worksheetIsHidden As Boolean

                Dim indexResetView As Long
                For indexResetView = 1 To worksheetCount
                    If mainWorkbook.Worksheets(indexResetView).Name <> "Log" Then
                        worksheetIsHidden = False
                        If mainWorkbook.Worksheets(indexResetView).Visible = xlSheetHidden Then worksheetIsHidden = True

                        If worksheetIsHidden = True Then
                            mainWorkbook.Worksheets(indexResetView).Visible = xlSheetVisible
                        End If

                        mainWorkbook.Worksheets(indexResetView).Select
                        Application.Goto Reference:=ActiveSheet.Cells.SpecialCells(xlCellTypeVisible).Range("A1"), Scroll:=True

                        If worksheetIsHidden = True Then
                            mainWorkbook.Worksheets(indexResetView).Visible = xlSheetHidden
                        End If
                    End If
                Next indexResetView

                mainWorkbook.Activate
                mainWorkbook.Worksheets("About").Select
                mainWorkbook.Worksheets("About").Activate
            Case "Save Workbook", "Save", "SW", "S"
                mainWorkbook.Save
            Case "Testing Mode", "Testing", "TM", "T"
                If NamedRangeExists("Foundation Checkpoints") Then
                    Call SetAboutNamedRange("N/A", "Foundation Checkpoints", True)
                End If

                If NamedRangeExists("Augmentation Checkpoints") Then
                    Call SetAboutNamedRange("N/A", "Augmentation Checkpoints", True)
                End If
        End Select
    Next index
End Sub

Sub RegisterMethod(ByVal contract As String, ByVal methodName As String, ByVal categoryName As String)
    Dim methodDictionary As Object
    
    If methodRegistry.Exists(methodName) = False Then
        Set methodDictionary = CreateObject("Scripting.Dictionary")
        Set methodRegistry(methodName) = methodDictionary
    Else
        Set methodDictionary = methodRegistry(methodName)
    End If
    
    methodDictionary("Category")    = categoryName
    methodDictionary("Declaration") = methodName & "(" & contract & ") @ " & categoryName & " (" & templateVersion & ")"
    methodDictionary("Registered")  = True

    If contract <> "" Then
        methodDictionary("Contract") = contract
    Else
        Exit Sub
    End If

    Dim parameters() As String
    Dim parameter As String
    Dim positionOfAs As Integer
    Dim result As String
    
    parameters = Split(contract, ",")
    
    Dim index As Integer
    For index = LBound(parameters) To UBound(parameters)
        parameter = Trim(parameters(index))

        If Left(parameter, 8) = "Context " Then
            parameter = Trim(Mid(parameter, 9))
        End If

        If Left(parameter, 9) = "Withheld " Then
            parameter = Trim(Mid(parameter, 10))
        End If

        If Left(parameter, 9) = "Optional " Then
            parameter = Trim(Mid(parameter, 10))
        End If

        If Left(parameter, 6) = "ByVal " Then
            parameter = Trim(Mid(parameter, 7))
        End If

        If Left(parameter, 6) = "ByRef " Then
            parameter = Trim(Mid(parameter, 7))
        End If

        If Left(parameter, 11) = "ParamArray " Then
            parameter = Trim(Mid(parameter, 12))
        End If

        positionOfAs = InStr(1, parameter, " As ", vbTextCompare)
        parameter = Left(parameter, positionOfAs - 1)
    
        result = result & parameter & ", "

        If index = UBound(parameters) Then
            result = Left(result, Len(result) - 2)
        End If
    Next index

    methodDictionary("Parameters") = result
End Sub

Sub SaveWorkbook(ByVal workbookName As String, ByVal directoryPath As String)
    Dim fullFilePath As String
    Dim previousDisplayAlerts As Boolean: previousDisplayAlerts = Application.DisplayAlerts
    
    If Right$(directoryPath, 1) <> "\" Then
        directoryPath = directoryPath & "\"
    End If

    fullFilePath = directoryPath & workbookName & ".xlsx"
    
    If Dir(fullFilePath) <> "" Then
        Application.DisplayAlerts = False
    End If

    Call mainWorkbook.SaveAs(Filename:=fullFilePath, FileFormat:=xlOpenXMLWorkbook)
    
    Application.DisplayAlerts = previousDisplayAlerts
End Sub

Sub SetAboutNamedRange(ByVal namedRangeValue As String, ByVal aboutNamedRange As String, Optional ByVal overwrite As Boolean)
    Dim currentNamedRangeValue As String

    With mainWorkbook.Worksheets("About")
        Select Case aboutNamedRange
            Case "ReportName", "ReportName", "Report", "Name"
                If NamedRangeExists("ReportName") Then
                    .Range("ReportName").Value = namedRangeValue
                    .Range("ReportName").Font.Bold = True
                End If
            Case "TemplateVersion", "Template Version", "Template", "Version"
                If NamedRangeExists("TemplateVersion") Then
                    .Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & namedRangeValue & ")"
                    .Range("TemplateVersion").Font.Bold = True
                End If
            Case "FoundationCheckpoints", "Foundation Checkpoints", "Foundation"
                If NamedRangeExists("FoundationCheckpoints") Then
                    currentNamedRangeValue = GetAboutNamedRange("Foundation Checkpoints")

                    If overwrite = True Or currentNamedRangeValue = "N/A" Then
                        .Range("FoundationCheckpoints").Value = "Foundation Checkpoints: " & namedRangeValue & "."
                    Else
                        .Range("FoundationCheckpoints").Value = "Foundation Checkpoints: " & currentNamedRangeValue & ", " & namedRangeValue & "."
                    End If

                    .Range("FoundationCheckpoints").Font.Bold = False
                    .Range("FoundationCheckpoints").Characters(Start:=1, Length:=22).Font.Bold = True
                End If
            Case "AugmentationCheckpoints", "Augmentation Checkpoints", "Augmentation"
                If NamedRangeExists("AugmentationCheckpoints") Then
                    currentNamedRangeValue = GetAboutNamedRange("Augmentation Checkpoints")

                    If overwrite = True Or currentNamedRangeValue = "N/A" Then
                        .Range("AugmentationCheckpoints").Value = "Augmentation Checkpoints: " & namedRangeValue & "."
                    Else
                        .Range("AugmentationCheckpoints").Value = "Augmentation Checkpoints: " & currentNamedRangeValue & ", " & namedRangeValue & "."
                    End If

                    .Range("AugmentationCheckpoints").Font.Bold = False
                    .Range("AugmentationCheckpoints").Characters(Start:=1, Length:=24).Font.Bold = True
                End If
            Case "ReportVision", "Report Vision", "Vision"
                If NamedRangeExists("ReportVision") Then
                    .Range("ReportVision").Value = namedRangeValue
                    .Range("ReportVision").Font.Bold = False
                End If
            Case "DependenciesList", "Dependencies List", "Dependencies"
                If NamedRangeExists("DependenciesList") Then
                    currentNamedRangeValue = GetAboutNamedRange("Dependencies List")

                    If overwrite = True Or currentNamedRangeValue = "N/A" Then
                        .Range("DependenciesList").Value = "Dependencies List: " & namedRangeValue & "."
                    Else
                        .Range("DependenciesList").Value = "Dependencies List: " & currentNamedRangeValue & ", " & namedRangeValue & "."
                    End If

                    .Range("DependenciesList").Font.Bold = False
                    .Range("DependenciesList").Characters(Start:=1, Length:=17).Font.Bold = True
                End If
            Case "CreationDate", "Creation Date", "Creation", "CD"
                If NamedRangeExists("CreationDate") Then
                    .Range("CreationDate").Value = "Creation Date: " & namedRangeValue & "."
                    .Range("CreationDate").Font.Bold = False
                    .Range("CreationDate").Characters(Start:=1, Length:=13).Font.Bold = True
                End If
            Case "EditionName", "Edition Name", "Edition"
                If NamedRangeExists("EditionName") Then
                    .Range("EditionName").Value = "Edition Name: " & namedRangeValue & "."
                    .Range("EditionName").Font.Bold = False
                    .Range("EditionName").Characters(Start:=1, Length:=12).Font.Bold = True
                End If
            Case "Duration (Milliseconds)", "Duration Milliseconds", "Duration"
                If NamedRangeExists("DurationMilliseconds") Then
                    currentNamedRangeValue = GetAboutNamedRange("Duration (Milliseconds)")

                    If overwrite = True Then
                        .Range("DurationMilliseconds").Value = "Duration (Milliseconds): " & namedRangeValue & "."
                    Else
                        .Range("DurationMilliseconds").Value = "Duration (Milliseconds): " & (CDbl(currentNamedRangeValue) + CDbl(namedRangeValue)) & "."
                    End If

                    .Range("DurationMilliseconds").Font.Bold = False
                    .Range("DurationMilliseconds").Characters(Start:=1, Length:=23).Font.Bold = True
                End If
            Case "LogSummary", "Log Summary", "Summary"
                If NamedRangeExists("LogSummary") Then
                    currentNamedRangeValue = GetAboutNamedRange("Log Summary")

                    Dim logRows As Long
                    Dim logRowsSummary As String
                    
                    If WorksheetExists("Log") = False Then
                        logRows = 0
                        logRowsSummary = logRows & " Rows."
                    Else
                        Dim logWorksheet As Worksheet: Set logWorksheet = mainWorkbook.Worksheets("Log")
                        logRows = logWorksheet.Cells(logWorksheet.Rows.Count, 1).End(xlUp).Row - 1

                        If logRows = 1 Then
                            logRowsSummary = logRows & " Row."
                        Else
                            logRowsSummary = logRows & " Rows."
                        End If
                    End If

                    Dim currentRunsValue As Long
                    Dim currentCheckpointsValue As Long

                    Dim argumentRunsValue As Long
                    Dim argumentCheckpointsValue As Long

                    Dim currentParts() As String
                    Dim argumentParts() As String

                    If currentNamedRangeValue = "N/A" Then
                        currentNamedRangeValue = "0 Runs. 0 Checkpoints. 0 Rows."
                    End If

                    currentParts  = Split(currentNamedRangeValue, ". ")
                    argumentParts = Split(namedRangeValue, ". ")

                    currentRunsValue        = CLng(Val(currentParts(0)))
                    currentCheckpointsValue = CLng(Val(currentParts(1)))

                    argumentRunsValue        = CLng(Val(argumentParts(0)))
                    argumentCheckpointsValue = CLng(Val(argumentParts(1)))

                    Dim combinedRunsValue As Long: combinedRunsValue = currentRunsValue + argumentRunsValue
                    Dim runsSummary As String

                    If combinedRunsValue = 1 Then
                        runsSummary = combinedRunsValue & " Run. "
                    Else
                        runsSummary = combinedRunsValue & " Runs. "
                    End If

                    Dim combinedCheckpointsValue As Long: combinedCheckpointsValue = currentCheckpointsValue + argumentCheckpointsValue
                    Dim checkpointsSummary As String

                    If combinedCheckpointsValue = 1 Then
                        checkpointsSummary = combinedCheckpointsValue & " Checkpoint. "
                    Else
                        checkpointsSummary = combinedCheckpointsValue & " Checkpoints. "
                    End If

                    .Range("LogSummary").Value = "Log Summary: " & runsSummary & checkpointsSummary & logRowsSummary
                    .Range("LogSummary").Font.Bold = False
                    .Range("LogSummary").Characters(Start:=1, Length:=11).Font.Bold = True
                End If
        End Select
    End With
End Sub

' Functions: Background '

Function CheckpointIsNew(ByVal checkpointName As String) As Boolean
    If WorksheetExists("About") = False Then
        CheckpointIsNew = True

        Exit Function
    End If

    Dim foundationCheckpointsValue As String
    Dim augmentationCheckpointsValue As String

    foundationCheckpointsValue = mainWorkbook.Worksheets("About").Range("A3").Value
    foundationCheckpointsValue = Replace(foundationCheckpointsValue, "Progression Status: ", "")
    foundationCheckpointsValue = Replace(foundationCheckpointsValue, "Foundation Checkpoints: ", "")

    augmentationCheckpointsValue = mainWorkbook.Worksheets("About").Range("A4").Value
    augmentationCheckpointsValue = Replace(augmentationCheckpointsValue, "Augmentation Modules: ", "")
    augmentationCheckpointsValue = Replace(augmentationCheckpointsValue, "Augmentation Checkpoints: ", "")
    
    If InStr(foundationCheckpointsValue, checkpointName) Or InStr(augmentationCheckpointsValue, checkpointName) Then
        CheckpointIsNew = False
    Else
        CheckpointIsNew = True
    End If
End Function

Function GetAboutNamedRange(ByVal aboutNamedRange As String) As String
    Dim namedRangeValue As String

    Select Case aboutNamedRange
        Case "ReportDetails", "Report Details", "ReportName", "ReportName", "Report", "Name"
            If NamedRangeExists("ReportDetails") Then
                namedRangeValue = aboutWorksheet.Range("ReportDetails").Value
            ElseIf NamedRangeExists("ReportName") Then
                namedRangeValue = aboutWorksheet.Range("ReportName").Value
            End If
        Case "TemplateDetails", "Template Details", "TemplateVersion", "Template Version", "Template", "Version"
            If NamedRangeExists("TemplateDetails") Then
                namedRangeValue = aboutWorksheet.Range("TemplateDetails").Value
            ElseIf NamedRangeExists("TemplateVersion") Then
                namedRangeValue = aboutWorksheet.Range("TemplateVersion").Value
            End If

            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, "(") + 1)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "FoundationCheckpoints", "Foundation Checkpoints", "Foundation", "ProgressionStatus", "Progression Status", "Progression"
            If NamedRangeExists("FoundationCheckpoints") Then
                namedRangeValue = aboutWorksheet.Range("FoundationCheckpoints").Value
            ElseIf NamedRangeExists("ProgressionStatus") Then
                namedRangeValue = aboutWorksheet.Range("ProgressionStatus").Value
            End If

            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "AugmentationCheckpoints", "Augmentation Checkpoints", "Augmentation", "AugmentationModules", "Augmentation Modules"
            If NamedRangeExists("AugmentationCheckpoints") Then
                namedRangeValue = aboutWorksheet.Range("AugmentationCheckpoints").Value
            ElseIf NamedRangeExists("AugmentationModules") Then
                namedRangeValue = aboutWorksheet.Range("AugmentationModules").Value
            End If

            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "ReportVision", "Report Vision", "Vision"
            namedRangeValue = aboutWorksheet.Range("ReportVision").Value
        Case "DependenciesList", "Dependencies List", "Dependencies"
            namedRangeValue = aboutWorksheet.Range("DependenciesList").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "CreationDate", "Creation Date", "Creation", "CD", "RetrievedDate", "Retrieved Date", "Retrieved"
            If NamedRangeExists("CreationDate") Then
                namedRangeValue = aboutWorksheet.Range("CreationDate").Value
            ElseIf NamedRangeExists("RetrievedDate") Then
                namedRangeValue = aboutWorksheet.Range("RetrievedDate").Value
            End If

            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "EditionName", "Edition Name", "Edition"
            namedRangeValue = aboutWorksheet.Range("EditionName").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "ScriptDuration", "Script Duration", "Duration (Milliseconds)", "Duration Milliseconds", "Duration"
            namedRangeValue = aboutWorksheet.Range("DurationMilliseconds").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "LogSummary", "Log Summary", "Summary"
            namedRangeValue = aboutWorksheet.Range("LogSummary").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case Else
            namedRangeValue = ""
    End Select

    GetAboutNamedRange = namedRangeValue
End Function

Function GetQueryPerformanceCounter() As Double
    Dim queryPerformanceCounterValue As Currency
    Call QueryPerformanceCounter(queryPerformanceCounterValue)
    GetQueryPerformanceCounter = CDbl(queryPerformanceCounterValue) * 10000#
End Function

' ************ '
' Validation   '
' ************ '

Sub ValidateColumnOnWorksheet(ByVal columnName As String, ByVal columnParameterName As String, ByVal worksheetName As String, ByVal worksheetParameterName As String, ByRef validation As String, Optional ByVal validColumn As Boolean)
    Dim validationMessage As String
    Dim worksheet As Worksheet
    Dim lastColumn As Long
    Dim headerRowRange As Range
    Dim occurrenceCount As Long

    Call ValidateWorksheet(worksheetName, worksheetParameterName, validationMessage)

    If validationMessage <> "" Then
        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If

        Exit Sub
    End If

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If Len(columnName) = 0 Then
        validationMessage = "Column name can't be blank."
    ElseIf Len(Trim$(columnName)) = 0 Then
        validationMessage = "Column name cannot consist only of spaces."
    ElseIf Len(columnName) > 255 Then
        validationMessage = "Column name can't be longer than 255 characters."
    ElseIf InStr(columnName, "*") > 0 Then
        validationMessage = "Column name can't contain an asterisk."
    ElseIf InStr(columnName, "?") > 0 Then
        validationMessage = "Column name can't contain a question mark."
    ElseIf worksheet.Range("A1").Value = "" And validColumn = False Then
        validationMessage = "First column doesn't have a header."
    End If

    If validationMessage = "" Then
        lastColumn = worksheet.Cells(1, worksheet.Columns.Count).End(xlToLeft).Column
        Set headerRowRange = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, lastColumn))
        occurrenceCount = Application.WorksheetFunction.CountIf(headerRowRange, columnName)

        If occurrenceCount = 0 And validColumn = False Then
            validationMessage = "Column name doesn't exist on the worksheet."
        ElseIf occurrenceCount = 1 And validColumn = True Then
            validationMessage = "Column name already exists on the worksheet."
        ElseIf occurrenceCount >= 2 Then
            validationMessage = "Column name appears more than once on the worksheet."
        End If
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & columnParameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

Sub ValidateDirectory(ByVal directoryPath As String, ByVal parameterName As String, ByRef validation As String, Optional ByVal validDirectoryPath As Boolean)
    Dim validationMessage As String
    Const invalidDirectoryCharacters As String = "/*?""<>|"
    Dim currentInvalidCharacter As String
    Dim foundPosition As Long
    Dim earliestPosition As Long
    Dim matchedInvalidCharacter As String
    Dim isLegalDriveColon As Boolean
    Dim firstPathCharacter As String
    Dim currentComponentStart As Long
    Dim currentDirectoryPosition As Long
    Dim isComponentBoundary As Boolean
    Dim currentComponent As String
    Dim isLeadingSeparator As Boolean
    Dim isUncSecondSeparator As Boolean
    Dim isTrailingSeparator As Boolean
    Dim isDrivePrefixComponent As Boolean
    Dim directoryComponentUpperCase As String
    Dim reservedDeviceNames As Variant
    Dim reservedDeviceName As String
    Dim lastCharacter As String
    Dim currentDirectoryCharacterCode As Long
    Dim fileSystemObject As Object

    reservedDeviceNames = Array("CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "COM" & ChrW(&HB9), "COM" & ChrW(&HB2), "COM" & ChrW(&HB3), _
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9", "LPT" & ChrW(&HB9), "LPT" & ChrW(&HB2), "LPT" & ChrW(&HB3))

    If Len(directoryPath) = 0 Then
        validationMessage = "Directory Path can't be blank."
    ElseIf Right$(directoryPath, 1) <> "\" Then
        validationMessage = "Directory Path must end with a backslash."
    Else
        Dim invalidDirectoryIndex As Long
        For invalidDirectoryIndex = 1 To Len(invalidDirectoryCharacters)
            currentInvalidCharacter = Mid$(invalidDirectoryCharacters, invalidDirectoryIndex, 1)
            foundPosition = InStr(1, directoryPath, currentInvalidCharacter, vbBinaryCompare)

            If foundPosition > 0 Then
                If earliestPosition = 0 Or foundPosition < earliestPosition Then
                    earliestPosition = foundPosition
                    matchedInvalidCharacter = currentInvalidCharacter
                End If
            End If
        Next invalidDirectoryIndex

        foundPosition = InStr(1, directoryPath, ":", vbBinaryCompare)
        Do While foundPosition > 0
            isLegalDriveColon = False

            If foundPosition = 2 Then
                firstPathCharacter = UCase$(Left$(directoryPath, 1))
                If firstPathCharacter >= "A" And firstPathCharacter <= "Z" Then
                    If Mid$(directoryPath, 3, 1) = "\" Then
                        isLegalDriveColon = True
                    End If
                End If
            End If

            If isLegalDriveColon = False Then
                If earliestPosition = 0 Or foundPosition < earliestPosition Then
                    earliestPosition = foundPosition
                    matchedInvalidCharacter = ":"
                End If
            End If

            foundPosition = InStr(foundPosition + 1, directoryPath, ":", vbBinaryCompare)
        Loop

        If earliestPosition > 0 Then
            validationMessage = "Directory Path contains the invalid character """ & matchedInvalidCharacter & """ at position " & CStr(earliestPosition) & "."
        Else
            currentComponentStart = 1
            For currentDirectoryPosition = 1 To Len(directoryPath) + 1
                If validationMessage <> "" Then
                    Exit For
                End If

                isComponentBoundary = False
                If currentDirectoryPosition > Len(directoryPath) Then
                    isComponentBoundary = True
                    currentComponent = Mid$(directoryPath, currentComponentStart)
                ElseIf Mid$(directoryPath, currentDirectoryPosition, 1) = "\" Then
                    isComponentBoundary = True
                    currentComponent = Mid$(directoryPath, currentComponentStart, currentDirectoryPosition - currentComponentStart)
                End If

                If isComponentBoundary = True Then
                    If Len(currentComponent) = 0 Then
                        isLeadingSeparator = (currentDirectoryPosition = 1)
                        isUncSecondSeparator = (currentDirectoryPosition = 2 And Left$(directoryPath, 2) = "\\")
                        isTrailingSeparator = (currentDirectoryPosition > Len(directoryPath))

                        If isLeadingSeparator = False And isUncSecondSeparator = False And isTrailingSeparator = False Then
                            validationMessage = "Directory Path contains an empty component."
                        End If
                    Else
                        isDrivePrefixComponent = (currentComponentStart = 1 And Len(currentComponent) = 2 And Right$(currentComponent, 1) = ":")

                        If isDrivePrefixComponent = False Then
                            directoryComponentUpperCase = UCase$(currentComponent)

                            Dim reservedDeviceNameIndex As Long
                            For reservedDeviceNameIndex = LBound(reservedDeviceNames) To UBound(reservedDeviceNames)
                                reservedDeviceName = CStr(reservedDeviceNames(reservedDeviceNameIndex))

                                If directoryComponentUpperCase = reservedDeviceName Then
                                    validationMessage = "Directory Path uses the reserved device name """ & currentComponent & """."
                                    Exit For
                                ElseIf Left$(directoryComponentUpperCase, Len(reservedDeviceName) + 1) = reservedDeviceName & "." Then
                                    validationMessage = "Directory Path uses the reserved device name """ & Left$(currentComponent, Len(reservedDeviceName)) & """."
                                    Exit For
                                End If
                            Next reservedDeviceNameIndex
                        End If
                    End If

                    currentComponentStart = currentDirectoryPosition + 1
                End If
            Next currentDirectoryPosition
        End If
    End If

    If validationMessage = "" Then
        currentComponentStart = 1
        For currentDirectoryPosition = 1 To Len(directoryPath) + 1
            If validationMessage <> "" Then
                Exit For
            End If

            isComponentBoundary = False
            If currentDirectoryPosition > Len(directoryPath) Then
                isComponentBoundary = True
                currentComponent = Mid$(directoryPath, currentComponentStart)
            ElseIf Mid$(directoryPath, currentDirectoryPosition, 1) = "\" Then
                isComponentBoundary = True
                currentComponent = Mid$(directoryPath, currentComponentStart, currentDirectoryPosition - currentComponentStart)
            End If

            If isComponentBoundary = True Then
                If Len(currentComponent) > 0 And currentComponent <> "." And currentComponent <> ".." Then
                    lastCharacter = Right$(currentComponent, 1)

                    If lastCharacter = "." Then
                        validationMessage = "Directory Path can't end with a period."
                    ElseIf lastCharacter = " " Then
                        validationMessage = "Directory Path can't end with a space."
                    End If
                End If

                currentComponentStart = currentDirectoryPosition + 1
            End If
        Next currentDirectoryPosition
    End If

    If validationMessage = "" Then
        For currentDirectoryPosition = 1 To Len(directoryPath)
            currentDirectoryCharacterCode = AscW(Mid$(directoryPath, currentDirectoryPosition, 1))

            If currentDirectoryCharacterCode >= 0 And currentDirectoryCharacterCode <= 31 Then
                validationMessage = "Directory Path contains an invalid control character " & CStr(currentDirectoryCharacterCode) & " at position " & CStr(currentDirectoryPosition) & "."
                Exit For
            End If
        Next currentDirectoryPosition
    End If

    If validationMessage = "" And validDirectoryPath = False Then
        Set fileSystemObject = CreateObject("Scripting.FileSystemObject")

        If fileSystemObject.FolderExists(directoryPath) = False Then
            validationMessage = "Directory Path doesn't exist."
        End If

        Set fileSystemObject = Nothing
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

Sub ValidateFilename(ByVal filename As String, ByVal parameterName As String, ByRef validation As String, Optional ByVal extension As String)
    Dim validationMessage As String
    Const invalidFilenameCharacters As String = "\/:*?""<>|"
    Dim currentInvalidCharacter As String
    Dim foundPosition As Long
    Dim earliestPosition As Long
    Dim matchedInvalidCharacter As String
    Dim filenameUpperCase As String
    Dim reservedDeviceNames As Variant
    Dim reservedDeviceName As String
    Dim lastCharacter As String
    Dim currentFilenameCharacterCode As Long

    reservedDeviceNames = Array("CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "COM" & ChrW(&HB9), "COM" & ChrW(&HB2), "COM" & ChrW(&HB3), _
        "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9", "LPT" & ChrW(&HB9), "LPT" & ChrW(&HB2), "LPT" & ChrW(&HB3))

    If Len(filename) = 0 Then
        validationMessage = "Filename can't be blank."
    Else
        Dim invalidFilenameIndex As Long
        For invalidFilenameIndex = 1 To Len(invalidFilenameCharacters)
            currentInvalidCharacter = Mid$(invalidFilenameCharacters, invalidFilenameIndex, 1)
            foundPosition = InStr(1, filename, currentInvalidCharacter, vbBinaryCompare)

            If foundPosition > 0 Then
                If earliestPosition = 0 Or foundPosition < earliestPosition Then
                    earliestPosition = foundPosition
                    matchedInvalidCharacter = currentInvalidCharacter
                End If
            End If
        Next invalidFilenameIndex

        If earliestPosition > 0 Then
            validationMessage = "Filename contains the invalid character """ & matchedInvalidCharacter & """ at position " & CStr(earliestPosition) & "."
        Else
            filenameUpperCase = UCase$(filename)

            Dim reservedDeviceNameIndex As Long
            For reservedDeviceNameIndex = LBound(reservedDeviceNames) To UBound(reservedDeviceNames)
                reservedDeviceName = UCase$(CStr(reservedDeviceNames(reservedDeviceNameIndex)))

                If filenameUpperCase = reservedDeviceName Then
                    validationMessage = "Filename uses the reserved device name """ & filename & """."
                    Exit For
                ElseIf Left$(filenameUpperCase, Len(reservedDeviceName) + 1) = reservedDeviceName & "." Then
                    validationMessage = "Filename uses the reserved device name """ & Left$(filename, Len(reservedDeviceName)) & """."
                    Exit For
                End If
            Next reservedDeviceNameIndex
        End If
    End If

    If validationMessage = "" Then
        lastCharacter = Right$(filename, 1)

        If lastCharacter = "." Then
            validationMessage = "Filename can't end with a period."
        ElseIf lastCharacter = " " Then
            validationMessage = "Filename can't end with a space."
        End If
    End If

    If validationMessage = "" Then
        Dim currentFilenamePosition As Long
        For currentFilenamePosition = 1 To Len(filename)
            currentFilenameCharacterCode = AscW(Mid$(filename, currentFilenamePosition, 1))

            If currentFilenameCharacterCode >= 0 And currentFilenameCharacterCode <= 31 Then
                validationMessage = "Filename contains an invalid control character " & CStr(currentFilenameCharacterCode) & " at position " & CStr(currentFilenamePosition) & "."
                Exit For
            End If
        Next currentFilenamePosition
    End If

    If validationMessage = "" And Len(extension) > 0 Then
        If Left$(extension, 1) = "." Then
            validationMessage = "Filename extension can't start with a period."
        ElseIf StrComp(Right$(filename, Len(extension) + 1), "." & extension, vbTextCompare) <> 0 Then
            validationMessage = "Filename doesn't end with the required extension '." & extension & "'."
        End If
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

Sub ValidateFilePath(ByVal filePath As String, ByVal parameterName As String, ByRef validation As String, Optional ByVal extension As String)
    Dim validationMessage As String
    Dim lastSeparator As Long
    Dim directoryPath As String
    Dim filename As String
    Dim fileSystemObject As Object

    lastSeparator = InStrRev(filePath, "\")

    If Len(filePath) = 0 Then
        validationMessage = "File Path can't be blank."
    ElseIf lastSeparator = 0 Then
        validationMessage = "File Path has no separator character (\)."
    End If

     If validationmessage = "" Then
        directoryPath = Left$(filePath, lastSeparator - 1)

        If directoryPath = "" Then
            directoryPath = "\"
        Else
            directoryPath = directoryPath & "\"
        End If

        Call ValidateDirectory(directoryPath, parameterName, validationMessage)

        If validationMessage <> "" Then
            If validation = "" Then
                validation = validationMessage
            Else
                validation = validation & " " & validationMessage
            End If

            Exit Sub
        End If
    End If

    If validationmessage = "" Then
        filename = Mid(filePath, lastSeparator + 1)

        Call ValidateFilename(filename, parameterName, validationMessage, extension)

        If validationMessage <> "" Then
            If validation = "" Then
                validation = validationMessage
            Else
                validation = validation & " " & validationMessage
            End If

            Exit Sub
        End If
    End If

    If validationMessage = "" Then
        Set fileSystemObject = CreateObject("Scripting.FileSystemObject")

        If fileSystemObject.FileExists(filePath) = False Then
            validationMessage = "File Path doesn't refer to an existing file."
        End If

        Set fileSystemObject = Nothing
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

Sub ValidateHexColor(ByVal hexColor As String, ByVal parameterName As String, ByRef validation As String)
    Dim validationMessage As String
    Dim hexadecimalBody As String
    Dim characterIndex As Long
    Dim currentCharacter As String
    Const validHexadecimalCharacters As String = "0123456789ABCDEFabcdef"

    If Len(hexColor) = 0 Then
        validationMessage = "Hex Color can't be blank."
    ElseIf Left$(hexColor, 1) <> "#" Then
        validationMessage = "Hex Color must start with the hash character (#)."
    ElseIf Len(hexColor) <> 7 Then
        validationMessage = "Hex Color must be 7 characters long, including the leading hash character (#)."
    End If

    If validationMessage = "" Then
        hexadecimalBody = Mid$(hexColor, 2)

        For characterIndex = 1 To Len(hexadecimalBody)
            currentCharacter = Mid$(hexadecimalBody, characterIndex, 1)

            If InStr(1, validHexadecimalCharacters, currentCharacter, vbBinaryCompare) = 0 Then
                validationMessage = "Hex Color contains the invalid character """ & currentCharacter & """ at position " & CStr(characterIndex + 1) & "."
                Exit For
            End If
        Next characterIndex
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

Sub ValidateNumeric(ByVal numeric As Variant, ByVal numericDataType As String, ByVal floor As Variant, ByVal ceiling As Variant, ByVal parameterName As String, ByRef validation As String)
    Dim validationMessage As String
    Dim actualNumericDataType As String
    Dim actualFloorDataType As String
    Dim actualCeilingDataType As String
    Dim numericValue As Variant
    Dim floorValue As Variant
    Dim ceilingValue As Variant

    Call ValidateWhitelist(numericDataType, parameterName, Array("Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal"), validationMessage)

    If validationMessage <> "" Then
        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If

        Exit Sub
    End If

    If IsArray(numeric) = False Or IsArray(floor) = False Or IsArray(ceiling) = False Then
        If IsArray(numeric) = False Then
            validationMessage = "Numeric isn't passed in via an array."
        ElseIf IsArray(floor) = False Then
            validationMessage = "Floor isn't passed in via an array."
        ElseIf IsArray(ceiling) = False Then
            validationMessage = "Ceiling isn't passed in via an array."
        End If
    Else
        If UBound(numeric) > LBound(numeric) Then
            validationMessage = "Numeric array has more than one entry."
        ElseIf UBound(floor) > LBound(floor) Then
            validationMessage = "Floor array has more than one entry."
        ElseIf UBound(ceiling) > LBound(ceiling) Then
            validationMessage = "Ceiling array has more than one entry."
        End If
    End If

    If validationMessage <> "" Then
        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If

        Exit Sub
    End If

    numericValue = numeric(LBound(numeric))
    floorValue = floor(LBound(floor))
    ceilingValue = ceiling(LBound(ceiling))

    Select Case VarType(numericValue)
        Case vbByte
            actualNumericDataType = "Byte"
        Case vbInteger
            actualNumericDataType = "Integer"
        Case vbLong
            actualNumericDataType = "Long"
        Case vbSingle
            actualNumericDataType = "Single"
        Case vbDouble
            actualNumericDataType = "Double"
        Case vbCurrency
            actualNumericDataType = "Currency"
        Case vbDecimal
            actualNumericDataType = "Decimal"
    End Select

    Select Case VarType(floorValue)
        Case vbByte
            actualFloorDataType = "Byte"
        Case vbInteger
            actualFloorDataType = "Integer"
        Case vbLong
            actualFloorDataType = "Long"
        Case vbSingle
            actualFloorDataType = "Single"
        Case vbDouble
            actualFloorDataType = "Double"
        Case vbCurrency
            actualFloorDataType = "Currency"
        Case vbDecimal
            actualFloorDataType = "Decimal"
    End Select

    Select Case VarType(ceilingValue)
        Case vbByte
            actualCeilingDataType = "Byte"
        Case vbInteger
            actualCeilingDataType = "Integer"
        Case vbLong
            actualCeilingDataType = "Long"
        Case vbSingle
            actualCeilingDataType = "Single"
        Case vbDouble
            actualCeilingDataType = "Double"
        Case vbCurrency
            actualCeilingDataType = "Currency"
        Case vbDecimal
            actualCeilingDataType = "Decimal"
    End Select

    If actualNumericDataType <> numericDataType Then
        validationMessage = "Mismatch between numeric and expected data type."
    ElseIf actualFloorDataType <> numericDataType Then
        validationMessage = "Mismatch between floor and expected data type."
    ElseIf actualCeilingDataType <> numericDataType Then
        validationMessage = "Mismatch between ceiling and expected data type."
    End If

    If validationMessage = "" Then
        If numericValue < floorValue Then
            validationMessage = "Numeric must be equal to or greater than " & floorValue & "."
        ElseIf numericValue > ceilingValue Then
            validationMessage = "Numeric must be equal to or less than " & ceilingValue & "."
        End If
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

Sub ValidateRequiredText(ByVal requiredText As String, ByVal parameterName As String, ByRef validation As String, Optional ByVal floor As Long, Optional ByVal ceiling As Long)
    Dim validationMessage As String

    If Len(requiredText) = 0 Then
        validationMessage = "Required Text can't be blank."
    ElseIf Len(Trim$(requiredText)) = 0 Then
        validationMessage = "Required Text can't be whitespace only."
    ElseIf floor > 0 And Len(requiredText) < floor Then
        validationMessage = "Required Text is shorter than " & CStr(floor) & " character" & IIf(floor = 1, ".", "s.")
    ElseIf ceiling > 0 And Len(requiredText) > ceiling Then
        validationMessage = "Required Text is longer than " & CStr(ceiling) & " character" & IIf(ceiling = 1, ".", "s.")
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

Sub ValidateWhitelist(ByVal argument As String, ByVal parameterName As String, ByVal whitelist As Variant, ByRef validation As String)
    Dim validationMessage As String
    Dim argumentIsWhitelisted As Boolean
    Dim currentWhitelistedWord As String
    Dim whitelistValues As String

    Dim index As Long
    For index = LBound(whitelist) To UBound(whitelist)
        currentWhitelistedWord = CStr(whitelist(index))

        If StrComp(currentWhitelistedWord, argument, vbBinaryCompare) = 0 Then
            argumentIsWhitelisted = True
            Exit For
        End If
    Next index

    If argumentIsWhitelisted = False Then
        whitelistValues = Join(whitelist, ", ")

        validationMessage = "Value is not in the whitelist (" & whitelistValues & ")."
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

Sub ValidateWorksheet(ByVal worksheetName As String, ByVal parameterName As String, ByRef validation As String, Optional ByVal validWorksheet As Boolean)
    Dim validationMessage As String
    Dim worksheetNameExists As Boolean
    Const invalidWorksheetCharacters As String = "/\?*:[]"
    Dim currentInvalidCharacter As String

    worksheetNameExists = WorksheetExists(worksheetName)

    If worksheetNameExists = True And validWorksheet = True Then
        validationMessage = "Worksheet already exists."
    ElseIf Len(worksheetName) >= 27 Then
        validationMessage = "Worksheet name is too long, unable to process further."
    ElseIf Len(worksheetName) = 0 Then
        validationMessage = "Worksheet name can't be blank."
    ElseIf Left$(worksheetName, 1) = "'" Then
        validationMessage = "Worksheet name can't start with the apostrophe character (')."
    ElseIf Right$(worksheetName, 1) = "'" Then
        validationMessage = "Worksheet name can't end with the apostrophe character (')."
    ElseIf StrComp(worksheetName, "History", vbTextCompare) = 0 Then
        validationMessage = "Worksheet name is reserved."
    End If

    If validationMessage = "" Then   
        Dim index As Long
        For index = 1 To Len(invalidWorksheetCharacters)
            currentInvalidCharacter = Mid$(invalidWorksheetCharacters, index, 1)
            
            If InStr(1, worksheetName, currentInvalidCharacter, vbBinaryCompare) > 0 Then
                validationMessage = "Worksheet name contains the invalid character """ & currentInvalidCharacter & """ at position " & CStr(index) & "."
                Exit For
            End If
        Next index
    End If

    If validationMessage = "" And worksheetNameExists = False And validWorksheet = False Then
        validationMessage = "Worksheet name not found."
    End If

    If validationMessage <> "" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

' Functions: Validation '

Function CellStyleExists(ByVal cellStyleName As String) As Boolean
    Dim cellStyleEntry As Style
        
    For Each cellStyleEntry In mainWorkbook.Styles
        If StrComp(cellStyleEntry.Name, cellStyleName, vbTextCompare) = 0 Then
            CellStyleExists = True
            Exit Function
        End If
    Next cellStyleEntry
    
    CellStyleExists = False
End Function

Function NamedRangeExists(ByVal namedRange As String) As Boolean
    Dim workbookNamedRange As Name

    For Each workbookNamedRange In mainWorkbook.Names
        If StrComp(workbookNamedRange.Name, namedRange, vbTextCompare) = 0 Then
            NamedRangeExists = True
            Exit Function
        End If
    Next workbookNamedRange

    NamedRangeExists = False
End Function

Function WorksheetExists(ByVal worksheetName As String) As Boolean
    Dim worksheetEntry As Worksheet

    For Each worksheetEntry In mainWorkbook.Worksheets
        If StrComp(worksheetEntry.Name, worksheetName, vbTextCompare) = 0 Then
            WorksheetExists = True
            Exit Function
        End If
    Next worksheetEntry

    WorksheetExists = False
End Function

Function WorksheetIsEmpty(ByVal worksheetName As String) As Variant
    Dim worksheetFound As Boolean
    Dim worksheet As Worksheet

    worksheetFound = WorksheetExists(worksheetName)

    If worksheetFound = False Then
        WorksheetIsEmpty = Null
        Exit Function
    End If

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    
    WorksheetIsEmpty = Application.WorksheetFunction.CountA(worksheet.Cells) = 0 And worksheet.Shapes.Count = 0
End Function

' ************ '
' ************ '

Sub Run()
    Dim logEngineQpc As Currency
    Dim logEngineTickCount As Currency
    Dim logEngineUtcTimestamp As String

    Call QueryPerformanceCounter(logEngineQpc)
    logEngineTickCount = GetTickCount64()
    logEngineUtcTimestamp = GetUtcTimestamp()

    Set mainWorkbook = ActiveWorkbook

    Set cellStyles = CreateObject("Scripting.Dictionary")
    Set environment = CreateObject("Scripting.Dictionary")
    Set international = CreateObject("Scripting.Dictionary")
    Set methodRegistry = CreateObject("Scripting.Dictionary")
    Set paths = CreateObject("Scripting.Dictionary")
    Set report = CreateObject("Scripting.Dictionary")
    Set telemetry = CreateObject("Scripting.Dictionary")

    baseTickCount = CDbl(logEngineTickCount * 10000)
    operationSequenceNumber = 1&
    logEngineActive = False
    delimiter = "|"
    helperColumn = "Helper Column"
    formulaColumn = "Formula Column"
    sortingColumn = "Sorting Column"
    templateVersion = "v0.40, 2026-08-06"

    report("Base QPC") = CDbl(logEngineQpc * 10000)
    report("Base UTC Timestamp") = logEngineUtcTimestamp
    report("Creation Date") = Format$(Date, "yyyy-mm-dd")
    report("Original Workbook") = mainWorkbook.FullName

    Set report("Settings") = New Collection

    Call Startup()
    Call Master()
End Sub

' Script is finished when "About" is placed last. '
' When saving the document, choose yes at prompt. '