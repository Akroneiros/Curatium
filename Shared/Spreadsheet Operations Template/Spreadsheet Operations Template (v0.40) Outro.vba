' ************ '
' Alteration   '
' ************ '

Sub AddColumnOnWorksheet(ByVal columnName As String, ByVal columnWidth As Double, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatAddColumnOnWorksheet(columnName, columnWidth, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "AddColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal columnWidth As Double, ByVal worksheetName As String", methodName, "Alteration")
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
        columnLetter = ConvertColumnNumberToColumnLetter(columnNumber)
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
End If

End Sub

Sub ApplyFormulaToCellWithCellStyleOnWorksheet(ByVal formulaValue As String, ByVal cellValue As String, ByVal cellStyle As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFormulaToCellWithCellStyleOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal cellValue As String, ByVal cellStyle As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaValue & """" & ", " & """" & cellValue & """" & ", " & """" & cellStyle & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If CellStyleExists(cellStyle) = False Then
        Call LogConclusion("Failed", logConclusionData, "Cell style """ & cellStyle & """ does not exist.")
    End If

    If worksheet.Range(cellValue).Style <> "Formula" Then
        With worksheet.Range(cellValue)
            .Style = "Formula"
            .Value = .Value
        End With
    End If
        
    Dim index As Long
    Dim cellValueLetter As String
    Dim cellValueNumber As String
    Dim cellValueNumberConversion As Long

    For index = 1 To Len(cellValue)
        If Not(IsNumeric(Mid(cellValue, index, 1))) And cellValueLetter = "" Then 
            cellValueLetter = Mid(cellValue, index, 1)
        ElseIf Not(IsNumeric(Mid(cellValue, index, 1))) Then
            cellValueLetter = cellValueLetter & Mid(cellValue, index, 1)
        End If

        If IsNumeric(Mid(cellValue, index, 1)) And cellValueNumber = "" Then 
            cellValueNumber = Mid(cellValue, index, 1)
        ElseIf IsNumeric(Mid(cellValue, index, 1)) Then
            cellValueNumber = cellValueNumber & Mid(cellValue, index, 1)
        End If
    Next index

    cellValueNumberConversion = CInt(cellValueNumber)
    cellValueNumberConversion = cellValueNumberConversion + 1
    cellValueNumber = CStr(cellValueNumberConversion)

    Call SelectWorksheet(worksheetName)

    On Error Resume Next
    worksheet.Range(cellValue).Formula2 = formulaValue
    If Err.Number <> 0 Then
        Err.Clear
        worksheet.Range(cellValue).FormulaLocal = formulaValue
    End If
    
    If worksheet.Range(cellValueLetter & cellValueNumber) <> "" Then
        worksheet.Range(Sheets(worksheetName).Range(cellValue), worksheet.Range(cellValue).End(xlDown)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    Else
        worksheet.Range(cellValue).Copy
        worksheet.Range(cellValue).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If

    With worksheet.Range(cellValue)
        .Style = cellStyle
        .Value = .Value
    End With

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyFormulaToColumnOnWorksheet(ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFormulaToColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If columnName = helperColumn Then Call EnsureHelperColumnOnWorksheet(worksheetName)

    If worksheet.Range(FirstColumnInRange(columnName) & FirstRowInRange(columnName)).Style <> "Formula" Then Call ApplyCellStyleToColumnOnWorksheet("Formula", columnName, worksheetName)
    worksheet.Range(FirstColumnInRange(columnName) & FirstRowInRange(columnName)).FormulaLocal = formulaValue

    Call SelectWorksheet(worksheetName)

    If worksheet.Range("A2").Value = "" Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & worksheetName & """" & " doesn't appear to have any valid data, can't apply formula.")

    If Range(columnName).Rows.Count <> 1 Then
        worksheet.Range(columnName).FillDown
    End If
    
    ' To avoid bugs with number formatting, i.e. leading zeroes. '
    worksheet.Range(columnName).Copy
    worksheet.Range(columnName).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyHyperlinkToCellOnWorksheet(ByVal hyperlinkValue As String, ByVal cellValue As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyHyperlinkToCellOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal hyperlinkValue As String, ByVal cellValue As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & hyperlinkValue & """" & ", " & """" & cellValue & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim currentCellValue As String
    
    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    currentCellValue = worksheet.Range(cellValue).Value

    If InputContainsValue(hyperlinkValue, "http") Or InputContainsValue(hyperlinkValue, "https") Then
        worksheet.Hyperlinks.Add Anchor:=Sheets(worksheetName).Range(cellValue), Address:="", SubAddress:=hyperlinkValue, TextToDisplay:=currentCellValue
    Else
        worksheet.Hyperlinks.Add Anchor:=Sheets(worksheetName).Range(cellValue), Address:="", SubAddress:="'" & hyperlinkValue & "'!A1", TextToDisplay:=currentCellValue
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub FindAndReplaceInColumnOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal columnName As String, ByVal worksheetName As String, Optional ByVal exactMatch As Boolean)  ' Repeat Support: columnName. '

If InputContainsValue(columnName, "|") Then
    Call RepeatFindAndReplaceInColumnOnWorksheet(findValue, replaceValue, columnName, worksheetName, exactMatch)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "FindAndReplaceInColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal columnName As String, ByVal worksheetName As String, Optional ByVal exactMatch As Boolean", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """" & ", " & exactMatch, validation)

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
End If

End Sub

Sub HideWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatHideWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "HideWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Visible = xlSheetHidden

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertDataValuesOnWorksheet(ByVal dataValues As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatInsertDataValuesOnWorksheet(dataValues, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "InsertDataValuesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal dataValues As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & dataValues & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim delimitedStringArray() As String
    Dim currentLastDataRow As Long

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    delimitedStringArray = Split(dataValues, "|")
    currentLastDataRow = LastUsedRowNumberOnWorksheet(worksheetName) + 1

    Dim index As Long
    For index = 0 To UBound(delimitedStringArray)
        worksheet.Range(ConvertColumnNumberToColumnLetter(index + 1) & currentLastDataRow).Value = delimitedStringArray(index)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertHeaderValuesOnWorksheet(ByVal headerValues As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatInsertHeaderValuesOnWorksheet(headerValues, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "InsertHeaderValuesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal headerValues As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & headerValues & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    Dim delimitedStringArray() As String: delimitedStringArray = Split(headerValues, "|")

    Dim index As Long
    For index = 0 To UBound(delimitedStringArray)
        worksheet.Range(ConvertColumnNumberToColumnLetter(index + 1) & "1").Value = delimitedStringArray(index)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub MoveWorksheetToEnd(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatMoveWorksheetToEnd(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "MoveWorksheetToEnd"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim worksheetIsHidden As Boolean

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    worksheetIsHidden = False

    If worksheet.Visible = xlSheetHidden Then worksheetIsHidden = True

    If worksheetIsHidden = True Then
        worksheet.Visible = xlSheetVisible
    End If

    worksheet.Move After:=Sheets(Worksheets.Count)

    If worksheetIsHidden = True Then
        worksheet.Visible = xlSheetHidden
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RenameWorksheetToValue(ByVal currentWorksheetName As String, ByVal newWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RenameWorksheetToValue"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal newWorksheetName As String", methodName, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & newWorksheetName & """")
    
    If WorksheetExists(currentWorksheetName) = False Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & currentWorksheetName & """" & " doesn't exist.")
    If WorksheetExists(newWorksheetName) Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & newWorksheetName & """" & " already exists.")
    If Len(newWorksheetName) > 31 Then Call LogConclusion("Failed", logConclusionData, "Worksheet name " & """" & newWorksheetName & """" & " is too long, max is 31 characters.")

    mainWorkbook.Worksheets(currentWorksheetName).Name = newWorksheetName

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ResetColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ResetColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If worksheet.Range("A2").Value <> "" Then worksheet.Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName)).FormulaArray = ";;"

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

Sub SetNewHeaderOnColumnOnWorksheet(ByVal newHeader As String, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetNewHeaderOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal newHeader As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & newHeader & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Range(columnName & "1").Value = newHeader

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetNewWidthOnColumnOnWorksheet(ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatSetNewWidthOnColumnOnWorksheet(newWidth, columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetNewWidthOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetName As String", methodName, "Alteration")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, newWidth & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Range(columnName & "1").ColumnWidth = newWidth

    Call LogConclusion("Completed", logConclusionData)
End If
End Sub

' Functions: Alteration '

Function AlterDelimitedStringBeforeAndAfter(delimitedString As String, beforeAlteration As String, afterAlteration As String) As String
    Dim delimitedStringArray() As String: delimitedStringArray = Split(delimitedString, "|")
    Dim index As Long

    For index = 0 To UBound(delimitedStringArray)
        AlterDelimitedStringBeforeAndAfter = AlterDelimitedStringBeforeAndAfter & beforeAlteration & delimitedStringArray(index) & afterAlteration & "|"
    Next index

    AlterDelimitedStringBeforeAndAfter = Left(AlterDelimitedStringBeforeAndAfter, Len(AlterDelimitedStringBeforeAndAfter) - 1)
End Function

Function DecreaseLetterOnce(ByVal letterToDecrease As String) As String
    Dim decreasedNumber As Long

    decreasedNumber = ConvertColumnLetterToColumnNumber(letterToDecrease)
    decreasedNumber = decreasedNumber - 1
    DecreaseLetterOnce = ConvertColumnNumberToColumnLetter(decreasedNumber)
End Function

Function IfStringIsNotEmptyReturnValue(stringValue As String, returnValue As String) As String
    If stringValue = "" Then IfStringIsNotEmptyReturnValue = ""
    If stringValue <> "" Then IfStringIsNotEmptyReturnValue = returnValue
End Function

Function IfWorksheetExistsReturnValue(ByVal worksheetName As String, ByVal returnValue As String) As String
    IfWorksheetExistsReturnValue = ""
    If WorksheetExists(worksheetName) Then IfWorksheetExistsReturnValue = returnValue
End Function

Function IncreaseLetterOnce(ByVal letterToIncrease As String) As String
    Dim increasedNumber As Long

    increasedNumber = ConvertColumnLetterToColumnNumber(letterToIncrease)
    increasedNumber = increasedNumber + 1   
    IncreaseLetterOnce = ConvertColumnNumberToColumnLetter(increasedNumber)
End Function

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
    For index = 0 To UBound(intermissionsArray)
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
                For aboutNameIndex = 0 To UBound(aboutNamedRangesArray)
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
    methodDictionary("Declaration") = methodName & "(" & contract & ") @ " & categoryName & " (" & report("Template Version") & ")"
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
    For index = 0 To UBound(parameters)
        parameter = Trim(parameters(index))

        If Left(parameter, 8) = "Context " Then
            parameter = Trim(Mid(parameter, 9))
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

Sub SaveWorkbook(ByVal workbookName As String, ByVal folderPath As String)
    Dim fullFilePath As String
    Dim previousDisplayAlerts As Boolean: previousDisplayAlerts = Application.DisplayAlerts
    
    If Right$(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    fullFilePath = folderPath & workbookName & ".xlsx"
    
    If Dir(fullFilePath) <> "" Then
        Application.DisplayAlerts = False
    End If

    mainWorkbook.SaveAs Filename:=fullFilePath, FileFormat:=xlOpenXMLWorkbook
    
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
    
    If InputContainsValue(foundationCheckpointsValue, checkpointName) Or InputContainsValue(augmentationCheckpointsValue, checkpointName) Then
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

Public Function GetQueryPerformanceCounter() As Double
    Dim queryPerformanceCounterValue As Currency
    Call QueryPerformanceCounter(queryPerformanceCounterValue)
    GetQueryPerformanceCounter = CDbl(queryPerformanceCounterValue) * 10000#
End Function

' ************ '
' Conjuration  '
' ************ '

Sub CloseGuestWorkbook()
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CloseGuestWorkbook"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("", methodName, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, "")

    On Error Resume Next
    
    importWorkbook.Close SaveChanges:=False

    If Err.Number <> 0 Then
        Err.Clear
        ' Call LogTaskWarning("Guest Workbook is already closed or has not been previously opened.", logOrderLocal)
    End If

    On Error GoTo 0

    Set importWorkbook = Nothing
    mainWorkbook.Activate

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub CreateWorksheet(ByVal worksheetName As String, Optional ByVal insertAfterWorksheetName As String) ' Repeat Support: worksheetName '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatCreateWorksheet(worksheetName, insertAfterWorksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CreateWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String, Optional ByVal insertAfterWorksheetName As String", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation, True)
    If insertAfterWorksheetName <> "" Then Call ValidateWorksheet(insertAfterWorksheetName, "insertAfterWorksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """" & ", " & """" & insertAfterWorksheetName & """", validation)

    Dim worksheet As Worksheet

    If insertAfterWorksheetName = "" Then
        Set worksheet = mainWorkbook.Worksheets.Add(Before:=mainWorkbook.Worksheets(1))
    Else
        Set worksheet = mainWorkbook.Worksheets.Add(After:=mainWorkbook.Worksheets(insertAfterWorksheetName))
    End If

    worksheet.Name = worksheetName

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub CreateFileListArrayFromDirectory(ByRef fileListArray As Variant, ByVal folderPath As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CreateFileListArrayFromDirectory"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByRef fileListArray As Variant, ByVal folderPath As String", methodName, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, "[" & "fileListArray" & "]" & " ," & """" & folderPath & """")

    If (VBA.IsArray(fileListArray) = False) Then Call LogConclusion("Failed", logConclusionData, "Input value for fileListArray is not an array.")

    Dim filename As String: filename = Dir(folderPath & "\" & "*.*")
    Dim fileCount As Long

    Do Until fileName = ""
        fileCount = fileCount + 1
        ReDim Preserve fileListArray(1 to fileCount)
        fileListArray(fileCount) = fileName
        fileName = Dir
    Loop

    If Len(Join(fileListArray)) = 0 Then
        Call LogConclusion("Failed", logConclusionData, "The directory """ & folderPath & """" & " contains no files.")
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub 

Sub DuplicateColumnFromWorksheetToWorksheet(ByVal columnName As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String) ' Repeat Support: columnName '

If InputContainsValue(columnName, "|") Then
    Call RepeatDuplicateColumnFromWorksheetToWorksheet(columnName, fromWorksheetName, toWorksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DuplicateColumnFromWorksheetToWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateWorksheet(toWorksheetName, "toWorksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, "columnName", fromWorksheetName, "fromWorksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & fromWorksheetName & """" & ", " & """" & toWorksheetName & """", validation)

    ' Remember to sort the two primary columns prior to using this method, or there will be errors unless they are sorted from before. '

    Dim cellStyleToCopy As String
    Dim widthToCopy As Double
    Dim originalColumnName As String
    originalColumnName = columnName

    Dim sourceColumnName As String
    Dim matchColumnName As String

    sourceColumnName = mainWorkbook.Worksheets(toWorksheetName).Range("A1").Value
    matchColumnName = mainWorkbook.Worksheets(fromWorksheetName).Range("A1").Value

    Dim deleteDuplicateColumn As Boolean
    deleteDuplicateColumn = False

    If sourceColumnName = "" Then
        Call AddColumnOnWorksheet("Sorting Column", 7, fromWorksheetName)
        Call ApplyFormulaToColumnOnWorksheet("=ROW()-1", "Sorting Column", fromWorksheetName)
        Call SortColumnByOrderOnWorksheet(matchColumnName, "Ascending", fromWorksheetName)
        mainWorkbook.Worksheets(fromWorksheetName).Range("A1:A" & LastUsedRowNumberOnWorksheet(fromWorksheetName)).Copy
        mainWorkbook.Worksheets(toWorksheetName).Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        mainWorkbook.Worksheets(toWorksheetName).Range("A1").Value = "Temporary"
        sourceColumnName = "Temporary"
        deleteDuplicateColumn = True
    End If

    cellStyleToCopy = mainWorkbook.Worksheets(fromWorksheetName).Range(columnName & "2").Style
    widthToCopy = mainWorkbook.Worksheets(fromWorksheetName).Range(columnName & "2").ColumnWidth

    ' If ColumnOnWorksheetExists(columnName, toWorksheetName) = False Then Call AddColumnOnWorksheet(originalColumnName, widthToCopy, toWorksheetName)

    Dim lookupValue As String
    lookupValue = "XLOOKUP(" & ColumnStartOnWorksheet(sourceColumnName, toWorksheetName) & ";" & ColumnDataOnWorksheet(matchColumnName, fromWorksheetName) & ";" & ColumnDataOnWorksheet(originalColumnName, fromWorksheetName)

    Call ApplyFormulaToColumnOnWorksheet("=IF(" & lookupValue & ";"""";0;2)<>"""";" & lookupValue & ";"""";0;2);"""")", originalColumnName, toWorksheetName)
    ' [search_mode]: 2 => Perform a binary search that relies on lookup_array being sorted in ascending order. If not sorted, invalid results will be returned. '
    Call ApplyCellStyleToColumnOnWorksheet(cellStyleToCopy, originalColumnName, toWorksheetName)

    If deleteDuplicateColumn = True Then Call DeleteColumnOnWorksheet("Temporary", toWorksheetName)

    mainWorkbook.Worksheets(fromWorksheetName).Range(FindColumnLetterOnWorksheet(originalColumnName, fromWorksheetName) & "1").Copy
    mainWorkbook.Worksheets(toWorksheetName).Range(FindColumnLetterOnWorksheet(originalColumnName, toWorksheetName) & "1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    If deleteDuplicateColumn = True Then
        Call SortColumnByOrderOnWorksheet("Sorting Column", "Ascending", fromWorksheetName)
        Call DeleteColumnOnWorksheet("Sorting Column", fromWorksheetName)
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DuplicateHeadersFromWorksheetToName(ByVal currentWorksheetName As String, ByVal toWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DuplicateHeadersFromWorksheetToName"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal toWorksheetName As String", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateWorksheet(currentWorksheetName, "currentWorksheetName", validation)
    Call ValidateWorksheet(toWorksheetName, "toWorksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & toWorksheetName & """", validation)

    If mainWorkbook.Worksheets(toWorksheetName).Range("A1").Value <> "" Then Call LogConclusion("Failed", logConclusionData, "Worksheet """ & toWorksheetName & """ is not blank.")

    mainWorkbook.Worksheets(currentWorksheetName).Range("A1:" & LastUsedColumnLetterOnWorksheet(currentWorksheetName) & "1").Copy
    mainWorkbook.Worksheets(toWorksheetName).Range("A1").PasteSpecial xlPasteAllUsingSourceTheme
    mainWorkbook.Worksheets(toWorksheetName).Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    Call ApplyAutoFilterOnWorksheet(toWorksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub DuplicateWorksheetToName(ByVal currentWorksheetName As String, ByVal newWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DuplicateWorksheetToName"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal newWorksheetName As String", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateWorksheet(currentWorksheetName, "currentWorksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & newWorksheetName & """", validation)

    If currentWorksheetName = newWorksheetName Then newWorksheetName = currentWorksheetName & "!"

    mainWorkbook.Worksheets(currentWorksheetName).Copy After:=Sheets(Worksheets.Count)
    ActiveSheet.Name = newWorksheetName

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ExportWorksheetToBlankWorkbook(ByVal worksheetName As String, ByVal filePath As String) ' Plural Support. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ExportWorksheetToBlankWorkbook"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String, ByVal filePath As String", methodName, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """" & ", " & """" & filePath & """")

    Workbooks.Add
    Set importWorkbook = ActiveWorkbook
    importWorkbook.SaveAs filePath

    Dim filterValueArray() As String: filterValueArray = Split(worksheetName, "|")

    Dim index As Integer
    For index = 0 To UBound(filterValueArray)
        mainWorkbook.Activate
        Dim validation As String
        Call ValidateWorksheet(filterValueArray(index), "worksheetName", validation)
        If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

        mainWorkbook.Worksheets(filterValueArray(index)).Copy Before:=importWorkbook.Sheets(1)
        importWorkbook.Activate

        importWorkbook.Sheets(filterValueArray(index)).Cells.Borders.LineStyle = xlLineStyleNone
        importWorkbook.Sheets(filterValueArray(index)).Cells.Interior.Pattern = xlNone
        importWorkbook.Sheets(filterValueArray(index)).Cells.Interior.TintAndShade = 0
        importWorkbook.Sheets(filterValueArray(index)).Cells.Interior.PatternTintAndShade = 0

        importWorkbook.Sheets(filterValueArray(index)).Cells.RowHeight = 16.5
        importWorkbook.Sheets(filterValueArray(index)).Rows("1:1").RowHeight = 49.5
        importWorkbook.Sheets(filterValueArray(index)).Cells.Style = "Normal"

        importWorkbook.Sheets(filterValueArray(index)).Cells.Font.Name = "Arial"
        importWorkbook.Sheets(filterValueArray(index)).Cells.Font.Size = 10
        importWorkbook.Sheets(filterValueArray(index)).Cells.HorizontalAlignment = xlCenter
        importWorkbook.Sheets(filterValueArray(index)).Cells.VerticalAlignment = xlCenter

        mainWorkbook.Activate
    Next index

    importWorkbook.Styles("Normal 2").Delete
    importWorkbook.Sheets("Sheet1").Delete
    importWorkbook.SaveAs filePath

    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ImportExcelFileWorksheets(ByVal filePath As String, ByVal worksheetsValues As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ImportExcelFileWorksheets"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetsValues As String", methodName, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & worksheetsValues & """")

    Dim worksheetsArray() As String

    Call OpenGuestWorkbook(filePath)

    If worksheetsValues = "" Then
        Dim excelFileWorksheetsCollection As New Collection

        Dim indexWorksheet As Integer
        For indexWorksheet = 1 To importWorkbook.Worksheets.Count
            If importWorkbook.Worksheets(indexWorksheet).Name <> "About" And importWorkbook.Worksheets(indexWorksheet).Name <> "Log" Then worksheetsValues = worksheetsValues & importWorkbook.Worksheets(indexWorksheet).Name & "|"
        Next indexWorksheet

        worksheetsValues = Left(worksheetsValues, Len(worksheetsValues) - 1)
    End If

    worksheetsArray = Split(worksheetsValues, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetsArray)
        Dim excelFileWorksheetExists As Boolean

        importWorkbook.Activate
        excelFileWorksheetExists = WorksheetExists(worksheetsArray(index))
            
        If excelFileWorksheetExists = True Then
            mainWorkbook.Activate
            Call CreateWorksheet(worksheetsArray(index))
            importWorkbook.Sheets(worksheetsArray(index)).Cells.Copy mainWorkbook.Worksheets(worksheetsArray(index)).Cells
        Else
            Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & worksheetsArray(index) & """" & " not found in the workbook " & """" & importWorkbook.Name & """" & ".")
        End If
    Next index

    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ImportExcelFileAndRenameWorksheet(ByVal filePath As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ImportExcelFileAndRenameWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetName As String", methodName, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & worksheetName & """")

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    Call OpenGuestWorkbook(filePath)

    If importWorkbook.Worksheets.Count > 1 Then Call LogConclusion("Failed", logConclusionData, "This method only supports importing one worksheet per workbook, the workbook " & """" & importWorkbook.Name & """" & " contains " & importWorkbook.Worksheets.Count & " worksheets.")

    importWorkbook.Sheets(1).Name = worksheetName
    mainWorkbook.Activate
    If WorksheetExists(worksheetName) Then Call LogConclusion("Failed", logConclusionData, "Worksheet named " & """" & worksheetName & """" & " already exists.")
    Call CreateWorksheet(worksheetName)

    importWorkbook.Sheets(worksheetName).Cells.Copy worksheet.Cells
    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ImportTextFileToWorksheet(ByVal filePath As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ImportTextFileToWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetName As String", methodName, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & worksheetName & """")

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    Call OpenGuestWorkbook(filePath)
    mainWorkbook.Activate
    Call CreateWorksheet(worksheetName)
    importWorkbook.Sheets(1).Cells.Copy worksheet.Cells
    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub OpenGuestWorkbook(ByVal filePath As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "OpenGuestWorkbook"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String", methodName, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """")

    Call AssertFilePathNotEmptyForVariable(filePath, "importWorkbookFilePath")

    Set importWorkbook = Workbooks.Open(filePath)
    importWorkbook.Activate

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub TransferDataFromWorksheetToWorksheet(ByVal fromWorksheetName As String, toWorksheetName As String)

If InputContainsValue(fromWorksheetName, "|") Then
    Call RepeatTransferDataFromWorksheetToWorksheet(fromWorksheetName, toWorksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "TransferDataFromWorksheetToWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetName As String, toWorksheetName As String", methodName, "Conjuration")
    End If

    Dim validation As String
    Call ValidateWorksheet(fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateWorksheet(toWorksheetName, "toWorksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & fromWorksheetName & """" & " ," & """" & toWorksheetName & """", validation)

    Call EnsureHelperColumnDeletedOnWorksheet(fromWorksheetName)
    Call EnsureHelperColumnDeletedOnWorksheet(toWorksheetName)

    Dim fromWorksheetNamePrimaryColumnName As String: fromWorksheetNamePrimaryColumnName = GetPrimaryColumnHeaderNameOnWorksheet(fromWorksheetName)
    Dim toWorksheetNamePrimaryColumnName As String: toWorksheetNamePrimaryColumnName = GetPrimaryColumnHeaderNameOnWorksheet(toWorksheetName)

    Call SortColumnByOrderOnWorksheet(fromWorksheetNamePrimaryColumnName, "Ascending", fromWorksheetName)
    Call SortColumnByOrderOnWorksheet(toWorksheetNamePrimaryColumnName, "Ascending", toWorksheetName)
    Call DuplicateColumnFromWorksheetToWorksheet(CreateDelimitedHeaderStringExcludingOnWorksheet(fromWorksheetNamePrimaryColumnName, fromWorksheetName), fromWorksheetName, toWorksheetName)
    Call DeleteWorksheet(fromWorksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

' Functions: Conjuration '

Function SelectFileOfTypeWithTitle(fileOfType As String, dialogTitle As String) As String
    Dim selectFileDialog As FileDialog
    Set selectFileDialog = Application.FileDialog(msoFileDialogFilePicker)

    With selectFileDialog
        .Filters.Clear
        .InitialFileName = "C:\Users\" & Environ$("Username") & "\Desktop"
        If fileOfType = "CSV" Then .Filters.Add "Comma-Separated Values File", "*.csv"
        If fileOfType = "Excel" Then .Filters.Add "Excel File", "*.xlsx"
        If fileOfType = "ODS" Then .Filters.Add "Open Document Spreadsheet File", "*.ods"
        If fileOfType = "Text" Then .Filters.Add "Text File", "*.txt"
        .FilterIndex = 1
        .Title = dialogTitle
        .AllowMultiSelect = False
        If .Show = True Then SelectFileOfTypeWithTitle = selectFileDialog.SelectedItems(1)
    End With
End Function

Function SelectFolderWithTitle(dialogTitle As String) As String
    Dim selectFolderDialog As FileDialog
    Set selectFolderDialog = Application.FileDialog(msoFileDialogFolderPicker)

    With selectFolderDialog
        .Filters.Clear
        .InitialFileName = "C:\Users\" & Environ$("Username") & "\Desktop"
        .Title = dialogTitle
        .AllowMultiSelect = False
        If .Show = True Then SelectFolderWithTitle = selectFolderDialog.SelectedItems(1)
    End With
End Function

' ************ '
' Destruction  '
' ************ '

Sub DeleteCellStyle(ByVal cellStyleName As String) ' Repeat Support: cellStyleName. '

If InputContainsValue(cellStyleName, "|") Then
    Call RepeatDeleteCellStyle(cellStyleName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteCellStyle"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyleName As String, Context regionalizedCellStyleName As String", methodName, "Destruction")
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

    If deleteCellStyleName = "" Then
        Call LogConclusion("Skipped", logConclusionData)

        Exit Sub
    End If

    If deleteCellStyleName = "Normal" Or deleteCellStyleName = cellStyles("Normal") Then
        Call LogConclusion("Failed", logConclusionData, "The cell style """ & "Normal" & """ is reserved and can't be deleted.")
    End If

    mainWorkbook.Styles(deleteCellStyleName).Delete
    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DeleteColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteColumnOnWorksheet(columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Destruction")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Columns(columnName & ":" & columnName).EntireColumn.Delete Shift:=xlToLeft

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DeleteRowsOnWorksheet(ByVal numberOfRows As Long, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteRowsOnWorksheet(numberOfRows, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteRowsOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal numberOfRows As Long, ByVal worksheetName As String", methodName, "Destruction")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, numberOfRows & " ," & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Rows("1:" & numberOfRows).Delete Shift:=xlUp

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DeleteWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Destruction")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Delete

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RemoveBlankLinesOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatRemoveBlankLinesOnWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RemoveBlankLinesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If WorksheetExists(worksheetName) Then
        Dim lastRowNumberForWorksheet As String: lastRowNumberForWorksheet = LastUsedRowNumberOnWorksheet(worksheetName)

        worksheet.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        worksheet.Range("A1").Value = "Remove Blank Lines Header( Temporary)"
        worksheet.Range("A2" & ":" & "A" & lastRowNumberForWorksheet).Value = ";;"
        Call RemoveDataBasedOnFormulaOnWorksheet("=OR(B2="""";CHAR(32)=B2)", worksheetName) ' Code 32 (decimal) is a nonprinting spacing character. (https://web.archive.org/web/20221004120012/http://www.columbia.edu/kermit/ascii.html) '
        worksheet.Columns("A:A").Delete Shift:=xlToLeft
    Else
        ' Call LogTaskWarning("Worksheet " & """" & worksheetName & """" & " does not exist.", logOrderLocal)
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RemoveEmptyColumnsOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatRemoveEmptyColumnsOnWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RemoveEmptyColumnsOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Destruction")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim emptyColumnArray() As String

    emptyColumnArray = Split(CreateDelimitedHeaderStringExcludingOnWorksheet("", worksheetName), "|")

    Dim index As Integer
    For index = 0 To UBound(emptyColumnArray)
        ' If ColumnIsEmptyOnWorksheet(FindColumnLetterOnWorksheet(emptyColumnArray(index), worksheetName), worksheetName) Then Call DeleteColumnOnWorksheet(emptyColumnArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RemoveDataBasedOnFormulaOnWorksheet(ByVal formulaValue As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RemoveDataBasedOnFormulaOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal worksheetName As String", methodName, "Destruction")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaValue & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    Call ApplyFormulaToColumnOnWorksheet(formulaValue, helperColumn, worksheetName)
    Call ApplyFilterToColumnOnWorksheet(True, helperColumn, worksheetName)

    If NumberOfVisibleCells(worksheetName) <> 0 Then
        Call ClearFilterOnWorksheet(worksheetName)
        Call AddColumnOnWorksheet("Original Order (Temporary)", 13, worksheetName)
        Call ApplyFormulaToColumnOnWorksheet("=ROW(2:2)-1", "Original Order (Temporary)", worksheetName)
        Call ApplyCellStyleToColumnOnWorksheet("Integer", "Original Order (Temporary)", worksheetName)
        Call ApplyFilterToColumnOnWorksheet(True, helperColumn, worksheetName)
    End If

    If NumberOfVisibleCells(worksheetName) <> 0 Then
        If worksheet.FilterMode Then
            ' This approach is used due to the instability of excel when dealing with large datasets. Just deleting the data can crash the application. '

            Call SelectWorksheet(worksheetName)

            ActiveWindow.ActivateNext
            ActiveWindow.WindowState = xlMaximized

            Dim savedFreezePanesMode As String
            savedFreezePanesMode = ActiveWindow.SplitColumn & ", " & ActiveWindow.SplitRow

            worksheet.Rows("1:1").EntireRow.Hidden = True
            worksheet.Range(ColumnRangeTypeOnWorksheet("A", "Data", worksheetName)).SpecialCells(xlCellTypeConstants, 23).ClearContents ' Header hidden before or bug occurs when only one line gets deleted. '
            worksheet.Rows("1:1").EntireRow.Hidden = False
            Call ClearFilterOnWorksheet(worksheetName)
            Call SortColumnByOrderOnWorksheet("A", "Ascending", worksheetName)
            worksheet.Name = "!"
            Call CreateWorksheet(worksheetName)
            mainWorkbook.Worksheets("!").Range("A1:" & LastUsedColumnLetterOnWorksheet("!") & LastUsedRowNumberOnWorksheet("!")).Copy mainWorkbook.Worksheets(worksheetName).Range("A1")
            mainWorkbook.Worksheets("!").Range("A1:" & LastUsedColumnLetterOnWorksheet("!") & "1").Copy
            worksheet.Range("A1").PasteSpecial xlPasteColumnWidths
            Application.CutCopyMode = False
            Call ApplyAutoFilterOnWorksheet(worksheetName)
            Call SortColumnByOrderOnWorksheet("Original Order (Temporary)", "Ascending", worksheetName)
            Call DeleteColumnOnWorksheet("Original Order (Temporary)", worksheetName)

            If savedFreezePanesMode <> "0, 0" Then Call SetFrozenPanesOnWorksheet(ActiveWindow.SplitColumn, ActiveWindow.SplitRow, worksheetName)

            If mainWorkbook.Worksheets("!").Tab.Color <> False Then
                worksheet.Tab.Color = mainWorkbook.Worksheets("!").Tab.Color
            End If

            Call DeleteWorksheet("!")
        Else
            ' Call LogTaskWarning("Worksheet " & """" & worksheetName & """" & " does not have an active filter.", logOrderLocal)
        End If
    End If

    If NumberOfVisibleCells(worksheetName) = 0 Then Call ClearFilterOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Sub SynchronizeWorksheetOnColumnFromWorksheet(ByVal synchronizeWorksheetName As String, ByVal columnName As String, ByVal fromWorksheetName As String)
'     Dim tickCount As Currency: tickCount = GetTickCount64()
'     Const methodName As String = "SynchronizeWorksheetOnColumnFromWorksheet"
'     Dim isRegistered As Boolean
'     Dim logConclusionData As LogEntry

'     If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
'     If isRegistered = False Then
'         Call RegisterMethod("ByVal synchronizeWorksheetName As String, ByVal columnName As String, ByVal fromWorksheetName As String", methodName, "Destruction")
'     End If

'     Dim validation As String
'     Call ValidateColumnOnWorksheet(synchronizeWorksheetColumnName, "synchronizeWorksheetColumnName", synchronizeWorksheetName, "synchronizeWorksheetName", validation)
'     Call ValidateColumnOnWorksheet(fromWorksheetColumnName, "fromWorksheetColumnName", fromWorksheetName, "fromWorksheetName", validation)

'     Call LogBeginning(methodName, tickCount, logConclusionData, """" & synchronizeWorksheetName & """" & ", " & """" & columnName & """" & ", " & """" & fromWorksheetName & """", validation)

'     Dim synchronizeWorksheetColumnName As String
'     Dim fromWorksheetColumnName As String

'     synchronizeWorksheetColumnName = columnName
'     fromWorksheetColumnName = columnName

'     Dim synchronizeWorksheetColumnStart As String
'     Dim fromWorksheetColumnData As String

'     synchronizeWorksheetColumnStart = ColumnStartOnWorksheet(synchronizeWorksheetColumnName, synchronizeWorksheetName)
'     fromWorksheetColumnData = ColumnDataOnWorksheet(fromWorksheetColumnName, fromWorksheetName)

'     Call RemoveDataBasedOnFormulaOnWorksheet("=ISNA(XLOOKUP(" & synchronizeWorksheetColumnStart & ";" & fromWorksheetColumnData & ";" & fromWorksheetColumnData & "))", synchronizeWorksheetName)

'     Call LogConclusion("Completed", logConclusionData)
' End Sub

' Functions: Destruction '

' ************ '
' Elementals   '
' ************ '

Sub ApplyFilterToColumnOnWorksheet(ByVal filterValue As String, ByVal columnName As String, ByVal worksheetName As String) ' Plural Support. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFilterToColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filterValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filterValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

If InputContainsValue(filterValue, "|") Then
    Dim filterValueArray() As String: filterValueArray = Split(filterValue, "|")

    worksheet.Range(columnName & "1:" & columnName & "1").AutoFilter Field:=ConvertColumnLetterToColumnNumber(columnName), Criteria1:=filterValueArray, Operator:=xlFilterValues
Else
    worksheet.Range(columnName & "1:" & columnName & "1").AutoFilter Field:=ConvertColumnLetterToColumnNumber(columnName), Criteria1:=filterValue, Operator:=xlAnd
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ClearFilterOnWorksheet(ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ClearFilterOnWorksheet"
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

    worksheet.AutoFilter.Sort.SortFields.Clear
    If worksheet.AutoFilterMode Then worksheet.AutoFilter.ShowAllData

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub CutColumnAndPasteAtColumnOnWorksheet(ByVal cutColumnName As String, ByVal pasteColumnName As String, ByVal worksheetName As String) ' Repeat Support: cutColumnName. '

If InputContainsValue(cutColumnName, "|") Then
    Call RepeatCutColumnAndPasteAtColumnOnWorksheet(cutColumnName, pasteColumnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CutColumnAndPasteAtColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal cutColumnName As String, ByVal pasteColumnName As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(cutColumnName, "cutColumnName", worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(pasteColumnName, "pasteColumnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cutColumnName & """" & ", " & """" & pasteColumnName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.Columns(cutColumnName & ":" & cutColumnName).Cut
    worksheet.Columns(pasteColumnName & ":" & pasteColumnName).Insert Shift:=xlToRight

    Call ResetAutoFilterOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

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

Sub SortColumnByOrderOnWorksheet(ByVal columnName As String, ByVal sortOrder As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatSortColumnByOrderOnWorksheet(columnName, sortOrder, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SortColumnByOrderOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal sortOrder As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & sortOrder & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If sortOrder <> "Ascending" And sortOrder <> "Descending" Then
        sortOrder = "Ascending"
        ' Call LogTaskWarning("Sorting order " & """" & sortOrder & """" & " not Ascending or Descending, setting to Ascending.", logOrderLocal)
    End If
      
    Call SelectWorksheet(worksheetName) ' Prevents selection bug. '
    If Not worksheet.AutoFilterMode Then Call LogConclusion("Failed", logConclusionData, "Worksheet """ & worksheetName & """ does not have a filter applied.")

    If sortOrder ="Ascending" Then worksheet.AutoFilter.Sort.SortFields.Add Key:=Range(columnName & "1:" & columnName & "1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    If sortOrder ="Descending" Then worksheet.AutoFilter.Sort.SortFields.Add Key:=Range(columnName & "1:" & columnName & "1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    With worksheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    worksheet.AutoFilter.Sort.SortFields.Clear

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ManualCodeSection(ByVal sectionName As String, ByVal startOrFinish As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ManualCodeSection"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal sectionName As String, ByVal startOrFinish As String", methodName, "Elementals")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & sectionName & """" & ", " & """" & startOrFinish & """")

    If startOrFinish <> "Start" And startOrFinish <> "Finish" Then Call LogConclusion("Failed", logConclusionData, "Only valid values are ""Start"" and ""Finish"", as opposed to input value of: """ & startOrFinish & """.")

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub TextToColumnsModeOnWorksheet(ByVal modeValue As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "TextToColumnsModeOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal modeValue As String, ByVal worksheetName As String", methodName, "Elementals")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & modeValue & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If modeValue = "" Then modeValue = "Semicolon"

    If modeValue = "Comma" Then worksheet.Range("A1:A" & LastUsedRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Comma:=True
    If modeValue = "Semicolon" Then worksheet.Range("A1:A" & LastUsedRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Semicolon:=True
    If modeValue = "Tab" Then worksheet.Range("A1:A" & LastUsedRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Tab:=True

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Functions: Elementals '

Function ConvertDateTimeToSerial(ByVal columnLetter As String, ByVal startingRow As Long) As String
    ConvertDateTimeToSerial = "=DATEVALUE(" & columnLetter & startingRow & ") + RIGHT(LEFT(" & columnLetter & startingRow & ";13);2)/24 + LEFT(RIGHT(" & columnLetter & startingRow & ";5);2)/1440 + RIGHT(" & columnLetter & startingRow & ";2)/86400"
End Function

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
        Call LogBeginning(methodName, tickCount, logConclusionData, CStr(columnNumberToConvert), "Column number " & CStr(columnNumberToConvert) & " exceeds Excel's maximum supported column.")
    End If

    remainingColumnNumber = columnNumberToConvert
    Do While remainingColumnNumber > 0
        remainderValue = (remainingColumnNumber - 1) Mod 26
        columnLetter = ChrW$(ASCII_CODE_FOR_UPPERCASE_A + remainderValue) & columnLetter
        remainingColumnNumber = (remainingColumnNumber - 1) \ 26
    Loop

    ConvertColumnNumberToColumnLetter = columnLetter
End Function

Function CreateDelimitedDataStringBasedOnColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    Dim worksheetColumnRange As Range
    Dim indexRange As Range

    Set worksheetColumnRange = mainWorkbook.Worksheets(worksheetName).Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName))
    For Each indexRange In worksheetColumnRange.SpecialCells(xlCellTypeVisible)
        CreateDelimitedDataStringBasedOnColumnOnWorksheet = CreateDelimitedDataStringBasedOnColumnOnWorksheet & indexRange.Value & "|"
    Next indexRange

    CreateDelimitedDataStringBasedOnColumnOnWorksheet = Left(CreateDelimitedDataStringBasedOnColumnOnWorksheet, Len(CreateDelimitedDataStringBasedOnColumnOnWorksheet) - 1)
End Function

Function CreateDelimitedHeaderStringExcludingOnWorksheet(ByVal exclusionHeaders As String, ByVal worksheetName As String) As String
    Dim exclusionHeadersArray() As String
    Dim indexHeaders As Integer
    Dim includeHeader As Boolean

    If exclusionHeaders <> "" Then
        exclusionHeadersArray = Split(exclusionHeaders, "|")
    End If

   Dim worksheetHeaderRange As Range
   Dim indexRange As Range

    Set worksheetHeaderRange = mainWorkbook.Worksheets(worksheetName).Range("A1:" & LastUsedColumnLetterOnWorksheet(worksheetName) & "1")
    For Each indexRange In worksheetHeaderRange
        includeHeader = True

        If exclusionHeaders <> "" Then
            For indexHeaders = 0 To UBound(exclusionHeadersArray)
                If indexRange.Value = exclusionHeadersArray(indexHeaders) Then
                    includeHeader = False
                End If
             Next indexHeaders
        End If

        If includeHeader = True Then
            CreateDelimitedHeaderStringExcludingOnWorksheet = CreateDelimitedHeaderStringExcludingOnWorksheet & indexRange.Value & "|"
        End If
    Next indexRange

    CreateDelimitedHeaderStringExcludingOnWorksheet = Left(CreateDelimitedHeaderStringExcludingOnWorksheet, Len(CreateDelimitedHeaderStringExcludingOnWorksheet) - 1)
End Function

Function CreateDelimitedWorksheetStringExcluding(ByVal exclusionWorksheets As String) As String
    Dim exclusionWorksheetsArray() As String
    Dim indexWorksheets As Integer
    Dim includeWorksheet As Boolean

    If exclusionWorksheets <> "" Then
        exclusionWorksheetsArray = Split(exclusionWorksheets, "|")
    End If

    Dim worksheetCount As Long
    Dim index As Long

    worksheetCount = mainWorkbook.Worksheets.Count
    For index = 1 To worksheetCount
        includeWorksheet = True

        If exclusionWorksheets <> "" Then
            For indexWorksheets = 0 To UBound(exclusionWorksheetsArray)
                If mainWorkbook.Worksheets(index).Name = exclusionWorksheetsArray(indexWorksheets) Then
                    includeWorksheet = False
                End If
             Next indexWorksheets
        End If

        If includeWorksheet = True Then
            CreateDelimitedWorksheetStringExcluding = CreateDelimitedWorksheetStringExcluding & mainWorkbook.Worksheets(index).Name & "|"
        End If
    Next index

    CreateDelimitedWorksheetStringExcluding = Left(CreateDelimitedWorksheetStringExcluding, Len(CreateDelimitedWorksheetStringExcluding) - 1)
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

Function FirstColumnInRange(ByVal rangeValue As String) As String
    If Len(rangeValue) <= 3 Then
        FirstColumnInRange = rangeValue
        Exit Function
    Else
        rangeValue = Left(rangeValue, InStr(1, rangeValue, ":") - 1)

        Dim numericalCheck As String

        Dim index As Integer
        For index = 1 To Len(rangeValue)
            numericalCheck = Right(rangeValue, Len(rangeValue) - index)

            If IsNumeric(numericalCheck) Then
                rangeValue = Left(rangeValue, index)
                FirstColumnInRange = rangeValue
                Exit Function
            End If
        Next index
    End If
End Function

Function FirstRowInRange(ByVal rangeValue As String) As String
    If Len(rangeValue) <= 3 Then
        FirstRowInRange = "2"
    Else
        rangeValue = Left(rangeValue, InStr(1, rangeValue, ":") - 1)

        Dim numericalCheck As String

        Dim index As Integer
        For index = 1 To Len(rangeValue)
            numericalCheck = Right(rangeValue, Len(rangeValue) - index)

            If IsNumeric(numericalCheck) Then
                FirstRowInRange = numericalCheck
                Exit Function
            End If
        Next index
    End If
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

Function LastBlankRowNumberOnWorksheet(ByVal worksheetName As String) As Long
    LastBlankRowNumberOnWorksheet = mainWorkbook.Worksheets(worksheetName).Cells(Rows.Count, "A").End(xlUp).Row + 1
End Function

Function LastColumnNumberOnWorksheet(ByVal worksheetName As String) As Long
    LastColumnNumberOnWorksheet = mainWorkbook.Worksheets(worksheetName).Cells(1, Columns.Count).End(xlToLeft).Column
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

    lastUsedColumnNumber = worksheet.Cells(1, worksheet.Columns.Count).End(xlToLeft).Column
    workingColumnNumber  = lastUsedColumnNumber

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

    lastUsedRowNumber = worksheet.Cells(worksheet.Rows.Count, 1).End(xlUp).Row

    If lastUsedRowNumber = 1 And IsEmpty(worksheet.Cells(1, 1).Value) Then
        lastUsedRowNumber = 0
    End If

    LastUsedRowNumberOnWorksheet = lastUsedRowNumber
End Function

Function NumberOfVisibleCells(worksheetName As String) As Long
    NumberOfVisibleCells = mainWorkbook.Worksheets(worksheetName).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
End Function

' Functions: Elementals, Columns '

Function ColumnRangeTypeOnWorksheet(ByVal columnName As String, ByVal rangeType As String, ByVal worksheetName As String) As String
    If worksheetName = "" Then worksheetName = ActiveSheet.Name
    If Len(columnName) <= 3 Then
        ' If ColumnLetterValid(columnName) = False Then Call LogFunctionError("ColumnRangeTypeOnWorksheet(ByVal columnName As String, ByVal rangeType As String, ByVal worksheetName As String) As String", "Column letter " & """" & columnName & """" & " on worksheet " & """" & worksheetName & """" & " not valid.")
    Else
        columnName = FindColumnLetterOnWorksheet(columnName, worksheetName)
    End If

    If rangeType = "Adjacent" Then ColumnRangeTypeOnWorksheet = columnName & "1:" & columnName & "3"
    If rangeType = "Contains" Then ColumnRangeTypeOnWorksheet = """*;""" & " & " & "'" & worksheetName & "'!" & columnName & "2" & " & """ & ";*"""
    If rangeType = "Data" Then ColumnRangeTypeOnWorksheet = columnName & "$2:" & columnName & "$" & LastUsedRowNumberOnWorksheet(worksheetName)
    If rangeType = "Full" Then ColumnRangeTypeOnWorksheet = columnName & "$2:" & columnName & "$1048576"
    If rangeType = "Header" Then ColumnRangeTypeOnWorksheet = columnName & "$1:" & columnName & "$1"
    If rangeType = "Last" Then ColumnRangeTypeOnWorksheet = columnName & LastBlankRowNumberOnWorksheet(worksheetName) & ":" & columnName & LastBlankRowNumberOnWorksheet(worksheetName)
    If rangeType = "Start" Then ColumnRangeTypeOnWorksheet = columnName & "2"
End Function

Function ColumnAdjacentOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    ColumnAdjacentOnWorksheet = "'" & worksheetName & "'!" & ColumnRangeTypeOnWorksheet(columnName, "Adjacent", worksheetName)
End Function

Function ColumnContainsOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    ColumnContainsOnWorksheet = ColumnRangeTypeOnWorksheet(columnName, "Contains", worksheetName)
End Function

Function ColumnDataOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    ColumnDataOnWorksheet = "'" & worksheetName & "'!" & ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName)
End Function

Function ColumnFullOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    ColumnFullOnWorksheet = "'" & worksheetName & "'!" & ColumnRangeTypeOnWorksheet(columnName, "Full", worksheetName)
End Function

Function ColumnHeaderOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    ColumnHeaderOnWorksheet = "'" & worksheetName & "'!" & ColumnRangeTypeOnWorksheet(columnName, "Header", worksheetName)
End Function

Function ColumnLastOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    ColumnLastOnWorksheet = "'" & worksheetName & "'!" & ColumnRangeTypeOnWorksheet(columnName, "Last", worksheetName)
End Function

Function ColumnStartOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    ColumnStartOnWorksheet = "'" & worksheetName & "'!" & ColumnRangeTypeOnWorksheet(columnName, "Start", worksheetName)
End Function

Function GetPrimaryColumnHeaderNameOnWorksheet(ByVal worksheetName As String) As String
    GetPrimaryColumnHeaderNameOnWorksheet = mainWorkbook.Worksheets(worksheetName).Range("A1").Value
End Function

' ************ '
' Formatting   '
' ************ '

Sub ApplyAutoFilterOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatApplyAutoFilterOnWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyAutoFilterOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If Not worksheet.AutoFilterMode Then
        worksheet.Range("A1:" & LastUsedColumnLetterOnWorksheet(worksheetName) & "1").AutoFilter
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ApplyCellStyleToColumnOnWorksheet(ByVal cellStyle As String, ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatApplyCellStyleToColumnOnWorksheet(cellStyle, columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyCellStyleToColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyle As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Formatting")
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

    Dim columnLetter As String: columnLetter = FindColumnLetterOnWorksheet(columnName, worksheetName)
    Dim lastUsedRowNumber As Long: lastUsedRowNumber = LastUsedRowNumberOnWorksheet(worksheetName)

    With worksheet.Range(columnLetter & "2:" & columnLetter & lastUsedRowNumber)
        .Style = cellStyle
        .Value = .Value
    End With

    If lastUsedRowNumber <> 1048576 Then
        With worksheet.Range(columnLetter & (lastUsedRowNumber + 1) & ":" & columnLetter & worksheet.Rows.Count)
            .Style = cellStyle
        End With
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ConvertColumnToDateFormattingOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName '

If InputContainsValue(columnName, "|") Then
    Call RepeatConvertColumnToDateFormattingOnWorksheet(columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ConvertColumnToDateFormattingOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Call EnsureHelperColumnOnWorksheet(worksheetName)

    Call ApplyFormulaToColumnOnWorksheet("=IF(" & ColumnStartOnWorksheet(columnName, worksheetName) & "<>"""";INT(" & ColumnStartOnWorksheet(columnName, worksheetName) & ");"""")", helperColumn, worksheetName)
    Call ApplyFormulaToColumnOnWorksheet("=IF(" & ColumnStartOnWorksheet(helperColumn, worksheetName) & "<>"""";" & ColumnStartOnWorksheet(helperColumn, worksheetName) & ";"""")", columnName, worksheetName)
    Call ApplyCellStyleToColumnOnWorksheet("Date", columnName, worksheetName)

    Call EnsureHelperColumnDeletedOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ConvertColumnToDecimalFormattingOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName. '

If InputContainsValue(columnName, "|") Then
    Call RepeatConvertColumnToDecimalFormattingOnWorksheet(columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ConvertColumnToDecimalFormattingOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Call EnsureHelperColumnOnWorksheet(worksheetName)

    Dim columnData As String
    Dim helperData As String
    columnData = ColumnRangeTypeOnWorksheet(columnName, "Start", worksheetName)
    helperData = ColumnRangeTypeOnWorksheet(helperColumn, "Start", worksheetName)

    Call ApplyFormulaToColumnOnWorksheet("=IF(NOT(ISBLANK(" & columnData & "));IF(NOT(ISERROR(VALUE(" & columnData & ")));VALUE(" & columnData & ");"""");"""")", helperColumn, worksheetName)
    Call ApplyFormulaToColumnOnWorksheet("=IF(" & helperData & "<>"""";" & helperData & ";"""")", columnName, worksheetName)
    
    Call EnsureHelperColumnDeletedOnWorksheet(worksheetName)

    Call ApplyCellStyleToColumnOnWorksheet("Decimal", columnName, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ColorColumnOfTypeAndElementOnWorksheet(ByVal colorName As String, ByVal columnName As String, ByVal columnType As String, ByVal elementType As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorColumnOfTypeAndElementOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal columnName As String, ByVal columnType As String, ByVal elementType As String, ByVal worksheetName As String", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & colorName & """" & ", " & """" & columnName & """" & ", " & """" & columnType & """" & ", " & """" & elementType & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim colorRange As String

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If columnType = "Data" Then colorRange = columnName & "$2:" & columnName & "$" & LastUsedRowNumberOnWorksheet(worksheetName)
    If columnType = "Full" Then colorRange = columnName & "$2:" & columnName & "$1048576"
    If columnType = "Header" Then colorRange = columnName & "$1:" & columnName & "$1"

    If columnType <> "Data" And columnType <> "Full" And columnType <> "Header" Then Call LogConclusion("Failed", logConclusionData, "Column type " & """" & columnType & """" & " is not a valid predefined column type.")

    Dim colorValue As Long
    colorValue = 2

    colorValue = StyleColor(colorName)
    If colorValue = 2 Then Call LogConclusion("Failed", logConclusionData, "Color name " & """" & colorName & """" & " is not a valid predefined color name.")

    If elementType = "Background" Or elementType = "Border" Or elementType = "Font" Then
        ' OK. '
    Else
        Call LogConclusion("Failed", logConclusionData, "Element type " & """" & elementType & """" & " is not a valid predefined element type.")
    End If

    If elementType = "Background" Then worksheet.Range(colorRange).Interior.Color = colorValue
    If elementType = "Border" Then
        With worksheet.Range(colorRange).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With worksheet.Range(colorRange).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With worksheet.Range(colorRange).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With worksheet.Range(colorRange).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
    End If
    If elementType = "Font" Then worksheet.Range(colorRange).Font.Color = colorValue

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ColorRangeBackgroundAndFontAndBorderOnWorksheet(ByVal rangeToColor As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal borderColorName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorRangeBackgroundAndFontAndBorderOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal rangeToColor As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal borderColorName As String, ByVal worksheetName As String", methodName, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & rangeToColor & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & borderColorName & """" & ", " & """" & worksheetName & """")

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If backgroundColorName = "" Then backgroundColorName = "White"
    If fontColorName = "" Then fontColorName = "Black"

    Dim backgroundColorValue As Long: backgroundColorValue = 2
    Dim fontColorValue As Long: fontColorValue = 2
    Dim borderColorValue As Long: borderColorValue = 2

    backgroundColorValue = StyleColor(backgroundColorName)
    If backgroundColorValue = 2 Then Call LogConclusion("Failed", logConclusionData, "Color name " & """" & backgroundColorValue & """" & " is not a valid predefined color name.")
    fontColorValue = StyleColor(fontColorName)
    If fontColorValue = 2 Then Call LogConclusion("Failed", logConclusionData, "Color name " & """" & fontColorValue & """" & " is not a valid predefined color name.")
    borderColorValue = StyleColor(borderColorName)
    If borderColorValue = 2 Then Call LogConclusion("Failed", logConclusionData, "Color name " & """" & borderColorValue & """" & " is not a valid predefined color name.")

    worksheet.Range(rangeToColor).Interior.Color = backgroundColorValue
    worksheet.Range(rangeToColor).Font.Color = fontColorValue

    With worksheet.Range(rangeToColor).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With worksheet.Range(rangeToColor).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With worksheet.Range(rangeToColor).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With worksheet.Range(rangeToColor).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ColorWorksheet(ByVal colorName As String, ByVal worksheetName As String) ' Repeat Support: worksheetName '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatColorWorksheet(colorName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal worksheetName As String", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & colorName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet
    Dim colorValue As Long

    Set worksheet = mainWorkbook.Worksheets(worksheetName)
    colorValue = 2

    colorValue = StyleColor(colorName)
    If colorValue = 2 Then Call LogConclusion("Failed", logConclusionData, "Color name " & """" & colorName & """" & " is not a valid predefined color name.")

    worksheet.Tab.Color = colorValue

    Call LogConclusion("Completed", logConclusionData)
End If

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
            .Columns("A:B").ColumnWidth = 68.14
            .Columns("C").ColumnWidth = 29.57

            .Range("A1:C1").WrapText = True
            .Range("B2:C2").WrapText = True
            .Range("A3:C4").WrapText = True

            .Range("B3:B4").VerticalAlignment = xlTop

            .Range("A1:C4").Font.Name = "Tahoma"
            .Range("A1:C4").Font.Size = 8
            .Range("A1:B2").Font.Size = 12

            .Range("ReportName").Font.Bold = False
            .Range("TemplateVersion").Font.Bold = False
            .Range("FoundationCheckpoints").Font.Bold = False
            .Range("AugmentationCheckpoints").Font.Bold = False
            .Range("DependenciesList").Font.Bold = False
            .Range("CreationDate").Font.Bold = False
            .Range("EditionName").Font.Bold = False
            .Range("DurationMilliseconds").Font.Bold = False
            .Range("LogSummary").Font.Bold = False

            .Range("ReportName").Font.Bold = True
            .Range("TemplateVersion").Font.Bold = True

            If Len(.Range("FoundationCheckpoints").Value) >= 22 Then
                .Range("FoundationCheckpoints").Characters(Start:=1, Length:=22).Font.Bold = True
            End If

            If Len(.Range("AugmentationCheckpoints").Value) >= 24 Then
                .Range("AugmentationCheckpoints").Characters(Start:=1, Length:=24).Font.Bold = True
            End If

            If Len(.Range("DependenciesList").Value) >= 17 Then
                .Range("DependenciesList").Characters(Start:=1, Length:=17).Font.Bold = True
            End If

            If Len(.Range("CreationDate").Value) >= 13 Then
                .Range("CreationDate").Characters(Start:=1, Length:=13).Font.Bold = True
            End If

            If Len(.Range("EditionName").Value) >= 12 Then
                .Range("EditionName").Characters(Start:=1, Length:=12).Font.Bold = True
            End If

            If Len(.Range("DurationMilliseconds").Value) >= 23 Then
                .Range("DurationMilliseconds").Characters(Start:=1, Length:=23).Font.Bold = True
            End If

            If Len(.Range("LogSummary").Value) >= 11 Then
                .Range("LogSummary").Characters(Start:=1, Length:=11).Font.Bold = True
            End If
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

            If .Columns("A").ColumnWidth <> 45.0 Or .Columns("B").ColumnWidth <> 77.86 Or .Columns("C").ColumnWidth <> 9.57 Or .Columns("D").ColumnWidth <> 9.0 Or .Columns("E").ColumnWidth <> 9.0 Or .Columns("F").ColumnWidth <> 9.29 Then
                .Columns("A").ColumnWidth = 45.0
                .Columns("B").ColumnWidth = 77.86
                .Columns("C").ColumnWidth = 9.57
                .Columns("D").ColumnWidth = 9.0
                .Columns("E").ColumnWidth = 9.0
                .Columns("F").ColumnWidth = 9.29
            End If
        End With

        If worksheet.Range("E2").Style <> "Integer" Then
            Call ApplyCellStyleToColumnOnWorksheet("Integer", "Tick Count", "Log")
        End If

        If worksheet.Range("F2").Style <> "Integer" Then
            Call ApplyCellStyleToColumnOnWorksheet("Integer", "Duration", "Log")
        End If
    End If

    If worksheetName = "Run Status" Then
        Dim lastSpacePosition As Long
        Dim currentCell As Range
        Dim colonPosition As Long

        With worksheet
            .Cells.Font.Bold = False
            .Cells.Font.Name = "Segoe UI"
            .Cells.Font.Size = 10
            .Cells.RowHeight = 16.5
            .Cells.HorizontalAlignment = xlCenter
            .Range("A1:A28").HorizontalAlignment = xlLeft
            .Columns("A").ColumnWidth = 73.57
            .Columns("B:E").ColumnWidth = 22.71

            .Range("A4:A8").VerticalAlignment = xlTop
            .Range("A1:A12").WrapText = True
            
            lastSpacePosition = InStrRev(.Range("A8").Value, " ")
            .Range("A8").Characters(Start:=lastSpacePosition + 1).Font.Name = "Consolas"

            For Each currentCell In Union(.Range("A1"), .Range("A4:A5"), .Range("A7:A8"), .Range("A13:A28"))
                colonPosition = InStr(currentCell.Value, ":")
                currentCell.Characters(Start:=1, Length:=colonPosition - 1).Font.Bold = True
            Next currentCell

            For Each currentCell In Union(.Range("A19"), .Range("A21:A23"), .Range("A25:A27"))
                colonPosition = InStr(currentCell.Value, ":")
                currentCell.Characters(Start:=colonPosition + 2).Font.Name = "Consolas"
            Next currentCell

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

Sub InsertImageOnWorksheetWithLeftAndTopValues(ByVal imageFilePath As String, ByVal worksheetName As String, ByVal leftValue As Double, ByVal topValue As Double)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "InsertImageOnWorksheetWithLeftAndTopValues"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal imageFilePath As String, ByVal worksheetName As String, ByVal leftValue As Double, ByVal topValue As Double", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & imageFilePath & """" & ", " & """" & worksheetName & """" & ", " & leftValue & ", " & topValue, validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    Call SelectWorksheet(worksheetName)
    worksheet.Pictures.Insert(imageFilePath).Select
    Selection.ShapeRange.IncrementLeft leftValue
    Selection.ShapeRange.IncrementTop topValue

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub NormalizeLayoutOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatNormalizeLayoutOnWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "NormalizeLayoutOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Formatting")

        Call ConfigureMethodSetting(methodName, "Apply Auto Filter", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Header Style", 1, 0, 1)
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
        If worksheet.Range(lastUsedColumnLetter & "1").Value <> "" Then
            With worksheet.Range("A1:" & lastUsedColumnLetter & "1")
                .Style = "Header"
            End With
        End If
    End If

    Call LogConclusion("Completed", logConclusionData)
End If
    
End Sub

Sub ResetAutoFilterOnWorksheet(ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ResetAutoFilterOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, "Formatting")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    worksheet.AutoFilterMode = False
    worksheet.Range("A1:" & LastUsedColumnLetterOnWorksheet(worksheetName) & "1").AutoFilter

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetFrozenPanesOnWorksheet(ByVal frozenRowCount As Long, ByVal frozenColumnCount As Long, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatSetFrozenPanesOnWorksheet(frozenRowCount, frozenColumnCount, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetFrozenPanesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal frozenRowCount As Long, ByVal frozenColumnCount As Long, ByVal worksheetName As String", methodName, "Formatting")
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
End If

End Sub

' Functions: Formatting '

Function StyleColor(ByVal styleColorName As String) As Long
    StyleColor = 2 ' If no match found the number two needs to be handled as error in method. '

    ' Primary color palette '
    If styleColorName = "Core Purple" Then StyleColor = 14879385     ' #990AE3: R: 153, G: 10, B: 227. '
    If styleColorName = "Light Grey" Then StyleColor = 15921906      ' #F2F2F2: R: 242, G: 242, B: 242. '
    If styleColorName = "White" Then StyleColor = 16777215           ' #FFFFFF: R: 255, G: 255, B: 255. '
    If styleColorName = "Black" Then StyleColor = 0                  ' #000000: R: 0, G: 0, B: 0. '

    ' Accent color palette '
    If styleColorName = "Black Purple" Then StyleColor = 3080479     ' #1F012F: R: 31, G: 1, B: 47. '
    If styleColorName = "Deep Purple" Then StyleColor = 5505848      ' #380354: R: 56, G: 3, B: 84. '
    If styleColorName = "Dark Purple" Then StyleColor = 8389975      ' #570580: R: 87, G: 5, B: 128. '
    If styleColorName = "Dark Violet" Then StyleColor = 10158235     ' #9B009B: R: 155, G: 0, B: 155.
    If styleColorName = "Bright Violet" Then StyleColor = 16711884   ' #CC00FF: R: 204, G: 0, B: 255.
    If styleColorName = "Dark Pink" Then StyleColor = 12135890       ' #D22DB9: R: 210, G: 45, B: 185. '
    If styleColorName = "Pink" Then StyleColor = 13435135            ' #FF00CD: R: 255, G: 0, B: 205. '
    If styleColorName = "Blue" Then StyleColor = 16750848            ' #0099FF: R: 0, G: 153, B: 255. '
    If styleColorName = "Dark Teal" Then StyleColor = 10066176       ' #009999: R: 0, G: 153, B: 153. '
    If styleColorName = "Green" Then StyleColor = 6737920            ' #00D066: R: 0, G: 208, B: 102. '
    If styleColorName = "Light Green" Then StyleColor = 6619033      ' #99FF64: R: 153, G: 255, B: 100. '
    If styleColorName = "Red" Then StyleColor = 6035428              ' #E4175C: R: 228, G: 23, B: 92. '
    If styleColorName = "Light Red" Then StyleColor = 6566655        ' #FF3264: R: 255, G: 50, B: 100.
    If styleColorName = "Orange" Then StyleColor = 39935             ' #FF9B00: R: 255, G: 155, B: 0. '
    If styleColorName = "Dark Grey" Then StyleColor = 10526880       ' #A0A0A0: R: 160, G: 160, B: 160. '

    ' Pebble color palette '
    If styleColorName = "Warm Purple" Then StyleColor = 15353779     ' #B347EA: R: 179, G: 71, B: 234. '
    If styleColorName = "Light Violet " Then StyleColor = 16728281   ' #D940FF: R: 217, G: 64, B: 255. '
    If styleColorName = "Pale Purple" Then StyleColor = 16745411     ' #C383FF: R: 195, G: 131, B: 255. '
    If styleColorName = "Cold Purple" Then StyleColor = 16724889     ' #9933FF: R: 153, G: 51, B: 255. '
    If styleColorName = "Light Purple" Then StyleColor = 16734895    ' #AF5AFF: R: 175, G: 90, B: 255. '
    If styleColorName = "Bright Blue" Then StyleColor = 16757568     ' #40B3FF: R: 64, G: 179, B: 255. '
    If styleColorName = "Light Turquoise" Then StyleColor = 16776960 ' #00FFFF: R: 0, G: 255, B: 255. '
    If styleColorName = "Light Blue" Then StyleColor = 16764160      ' #00CDFF: R: 0, G: 205, B: 255. '
    If styleColorName = "Teal" Then StyleColor = 13487360            ' #00CDCD: R: 0, G: 205, B: 205. '
    If styleColorName = "Light Teal" Then StyleColor = 13500160      ' #00FFCD: R: 0, G: 255, B: 205. '
    If styleColorName = "Deep Teal" Then StyleColor = 11776832       ' #40B3B3: R: 64, G: 179, B: 179. '
    If styleColorName = "Lime" Then StyleColor = 3342285             ' #CDFF32: R: 205, G: 255, B: 50. '
    If styleColorName = "Light Lime" Then StyleColor = 10092493      ' #CDFF99: R: 205, G: 255, B: 153. '
    If styleColorName = "Pale Green" Then StyleColor = 14351561      ' #C9FCDA: R: 201, G: 252, B: 218. '
    If styleColorName = "Bright Pink" Then StyleColor = 14303487     ' #FF40DA: R: 255, G: 64, B: 218. '
    If styleColorName = "Pale Pink" Then StyleColor = 16763647       ' #FFCAFF: R: 255, G: 202, B: 255. '
    If styleColorName = "Bright Orange" Then StyleColor = 4240639    ' #FFB440: R: 255, G: 180, B: 64. '
    If styleColorName = "Pale Orange" Then StyleColor = 9165567      ' #FFDA8B: R: 255, G: 218, B: 139. '
    If styleColorName = "Cool Grey" Then StyleColor = 12040119       ' #B7B7B7: R: 183, G: 183, B: 183. '
    If styleColorName = "Cool Red" Then StyleColor = 9135615         ' #FF658B: R: 255, G: 101, B: 139. '
    If styleColorName = "Pale Red" Then StyleColor = 9145343         ' #FF8B8B: R: 255, G: 139, B: 139. '
End Function

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

            If aboutWorksheet.Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & report("Template Version") & ")" Then
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

    If namedRanges("ReportName") = False Then mainWorkbook.Names.Add Name:="ReportName", RefersToR1C1:="=About!R1C1"
    If namedRanges("TemplateVersion") = False Then mainWorkbook.Names.Add Name:="TemplateVersion", RefersToR1C1:="=About!R2C1"
    If namedRanges("FoundationCheckpoints") = False Then mainWorkbook.Names.Add Name:="FoundationCheckpoints", RefersToR1C1:="=About!R3C1"
    If namedRanges("AugmentationCheckpoints") = False Then mainWorkbook.Names.Add Name:="AugmentationCheckpoints", RefersToR1C1:="=About!R4C1"
    If namedRanges("ReportVision") = False Then mainWorkbook.Names.Add Name:="ReportVision", RefersToR1C1:="=About!R1C2"
    If namedRanges("DependenciesList") = False Then mainWorkbook.Names.Add Name:="DependenciesList", RefersToR1C1:="=About!R3C2"
    If namedRanges("CreationDate") = False Then mainWorkbook.Names.Add Name:="CreationDate", RefersToR1C1:="=About!R1C3"
    If namedRanges("EditionName") = False Then mainWorkbook.Names.Add Name:="EditionName", RefersToR1C1:="=About!R2C3"
    If namedRanges("DurationMilliseconds") = False Then mainWorkbook.Names.Add Name:="DurationMilliseconds", RefersToR1C1:="=About!R3C3"
    If namedRanges("LogSummary") = False Then mainWorkbook.Names.Add Name:="LogSummary", RefersToR1C1:="=About!R4C3"

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

        .Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & report("Template Version") & ")"

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
            "Time Zone: " & environment("Time Zone"), "QPC Frequency: " & environment("QPC Frequency"), "Base QPC: " & baseQpc, "Base Tick Count: " & baseTickCount, "Base UTC Timestamp: " & baseUtcTimestamp, _
            "Telemetry QPC Midpoint Timestamp: " & telemetry("QPC Midpoint Timestamp"), "Telemetry Tick Count: " & telemetry("Tick Count"), _
            "Telemetry UTC Timestamp Integer: " & telemetry("UTC Timestamp Integer"), "Telemetry UTC Timestamp Precise: " & telemetry("UTC Timestamp Precise")))

        .Range("B1").Resize(28).Value = Application.Transpose(Array( _
            "Brackets and Braces", "Left Brace", "Left Bracket", "Lower Case Column Letter", "Lower Case Row Letter", "Right Brace", "Right Bracket", "Upper Case Column Letter", "Upper Case Row Letter", _
            "Country/Region Settings", "Country Code", "Country Setting", "General Format Name", _
            "Currency", "Currency Before", "Currency Code", "Currency Digits", "Currency Leading Zeros", "Currency Minus Sign", "Currency Negative", "Currency Space Before", "Currency Trailing Zeros", "Noncurrency Digits", _
            "Windows Locale Settings", "Display Language", "Input Language", "Keyboard Layout", "Regional Format"))

        .Range("C1").Resize(28).Value = Application.Transpose(Array( _
            "", international("Left Brace"), international("Left Bracket"), international("Lower Case Column Letter"), international("Lower Case Row Letter"), _
            international("Right Brace"), international("Right Bracket"), international("Upper Case Column Letter"), international("Upper Case Row Letter"), _
            "", international("Country Code"), international("Country Setting"), international("General Format Name"), _
            "", CBool(international("Currency Before")), international("Currency Code"), international("Currency Digits"), CBool(international("Currency Leading Zeros")), CBool(international("Currency Minus Sign")), _
            international("Currency Negative"), CBool(international("Currency Space Before")), CBool(international("Currency Trailing Zeros")), international("Noncurrency Digits"), _
            "", environment("Display Language"), environment("Input Language"), environment("Keyboard Layout"), environment("Regional Format")))

        .Range("D1").Resize(28).Value = Application.Transpose(Array( _
            "Date and Time", "24 Hour Clock", "4 Digit Years", "Date Order", "Date Separator", "Day Code", "Day Leading Zero", "Hour Code", "MDY", "Minute Code", _
            "Month Code", "Month Leading Zero", "Month Name Chars", "Second Code", "Time Separator", "Time Leading Zero", "Weekday Name Chars", "Year Code", _
            "Measurement Systems", "Metric", "Non English Functions", _
            "Separators", "Alternate Array Separator", "Column Separator", "Decimal Separator", "List Separator", "Row Separator", "Thousands Separator"))

        .Range("E2").Resize(28).Value = Application.Transpose(Array( _
            "", CBool(international("24 Hour Clock")), CBool(international("4 Digit Years")), international("Date Order"), international("Date Separator"), international("Day Code"), _
            CBool(international("Day Leading Zero")), international("Hour Code"), CBool(international("MDY")), international("Minute Code"), international("Month Code"), CBool(international("Month Leading Zero")), _
            international("Month Name Chars"), international("Second Code"), international("Time Separator"), CBool(international("Time Leading Zero")), international("Weekday Name Chars"), international("Year Code"), _
            "", CBool(international("Metric")), CBool(international("Non English Functions")), _
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
        .Range("A1").Font.Bold = False
        .Range("A1").Characters(Start:=1, Length:=11).Font.Bold = True

        .Range("A4").Value = "Parameters: " & methodRegistry(logConclusionData.methodName)("Parameters")
        .Range("A4").Font.Bold = False
        .Range("A4").Font.Size = 10
        .Range("A4").Characters(Start:=1, Length:=10).Font.Bold = True
        .Range("A4").Characters(Start:=13).Font.Size = 8

        .Range("A5").Value = "Arguments: " & logConclusionData.arguments      
        .Range("A5").Font.Bold = False
        .Range("A5").Characters(Start:=1, Length:=9).Font.Bold = True

        .Range("A8").Value = "Date Runtime: " & utcTimestamp & " (UTC), " & CDbl(qpc * 10000)
        .Range("A8").Font.Bold = False
        .Range("A8").Font.Name = "Segoe UI"
        .Range("A8").Characters(Start:=1, Length:=12).Font.Bold = True

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
            .Range("A9").Value = "Success Output: The run started at " & baseUtcTimestamp & " (UTC) and finished successfully at " & utcTimestamp & " (UTC). The entire run took " & (logConclusionData.tickCount + durationRange.Value) & " milliseconds."
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
If report("Log Engine Active") = True Then Exit Sub

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
                logHeaderValues = Array("Method", "Arguments", "Category", "Tick Count", "Duration", "Outcome")

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

        logWorksheet.Cells(1, 1).Resize(1, 6).Value = Array("Method", "Arguments", "Category", "Tick Count", "Duration", "Outcome")
    End If
  
    Call LogBeginning(methodName, CCur(baseTickCount / 10000), logConclusionData, """" & report("Template Version") & """" & ", " & """" & environment("Run Identifier") & """" & ", " & _
        environment("QPC Frequency") & ", " & baseQpc & ", " & baseTickCount & ", " & """" & baseUtcTimestamp & """")

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

        Dim temporaryCell As Range: Set temporaryCell = logWorksheet.Range("B1")

        Call logWorksheet.Hyperlinks.Add(temporaryCell, "", "'" & logWorksheet.Name & "'!A1")

        Dim hyperlinkCellStyle As String: hyperlinkCellStyle = "Hyperlink"
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

        Call temporaryCell.Hyperlinks.Delete

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
        For customCellStyleIndex = 0 To UBound(customCellStylesArray)
            customCellStyle = customCellStylesArray(customCellStyleIndex)

            If CellStyleExists(customCellStyle) = False Then
                Call mainWorkbook.Styles.Add(Name:=customCellStyle)

                If customCellStyle = "Date" Then
                    mainWorkbook.Styles("Date").NumberFormat = "dd/mm/yyyy"
                End If

                If customCellStyle = "Date Time" Then
                    mainWorkbook.Styles("Date Time").NumberFormat = "dd/mm/yyyy HH:mm:ss"
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

    Call CreateAboutWorksheet()
    Set aboutWorksheet = mainWorkbook.Worksheets("About")
    Call SetAboutNamedRange("1 Run. 0 Checkpoints. 0 Rows.", "Log Summary")

    Call CreateRunStatusWorksheet()
    Set runStatusWorksheet = mainWorkbook.Worksheets("Run Status")
    With runStatusWorksheet
        .Range("A1").Value = "Declaration: " & methodRegistry(logConclusionData.methodName)("Declaration")
        .Range("A1").Characters(Start:=1, Length:=11).Font.Bold = True

        .Range("A4").Value = "Parameters: " & methodRegistry(logConclusionData.methodName)("Parameters")
        .Range("A4").Font.Bold = False
        .Range("A4").Font.Size = 10
        .Range("A4").Characters(Start:=1, Length:=10).Font.Bold = True
        .Range("A4").Characters(Start:=13).Font.Size = 8

        .Range("A5").Value = "Arguments: " & logConclusionData.arguments
        .Range("A5").Font.Bold = False
        .Range("A5").Characters(Start:=1, Length:=9).Font.Bold = True

        .Range("A8").Value = "Date Runtime: " & baseUtcTimestamp & " (UTC), " & baseQpc
        .Range("A8").Font.Bold = False
        .Range("A8").Font.Name = "Segoe UI"
        .Range("A8").Characters(Start:=1, Length:=12).Font.Bold = True
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

    report("Log Engine Active") = True

    Call LogConclusion("Completed", logConclusionData)

    Call SetAboutNamedRange(CStr(logWorksheet.Cells(logConclusionData.operationSequenceNumber, 5).Value), "Duration (Milliseconds)")

    With logWorksheet
        .Rows(logConclusionData.operationSequenceNumber).RowHeight = 33
        .Cells(logConclusionData.operationSequenceNumber, 2).WrapText = True
    End With
End Sub

Sub LogAssertFailure(ByVal errorMessage As String)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    Err.Raise vbObjectError + 513, "Assert Failure", "Code execution halted due to Assert Failure: " & errorMessage
End Sub

Sub LogFunctionError(ByVal functionName As String, ByVal errorMessage As String)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    Err.Raise vbObjectError + 513, functionName, "Code execution halted due to Function Error: " & errorMessage
End Sub

Sub LogTaskWarning(ByVal warningMessage As String, ByVal taskLogOrder As Long)
    mainWorkbook.Worksheets("Log").Range("D" & (taskLogOrder)).Value = "Warning"
    mainWorkbook.Worksheets("Log").Range("H" & (taskLogOrder)).Value = Format(Now, "dd/mm/yyyy HH:mm:ss")
    mainWorkbook.Worksheets("Log").Range("J" & (taskLogOrder)).Value = Round(Timer - stopwatchTimer, 2)

    If mainWorkbook.Worksheets("Log").Range("D" & (taskLogOrder)).Comment Is Nothing Then
        mainWorkbook.Worksheets("Log").Range("D" & (taskLogOrder)).AddComment (warningMessage)
    Else
        mainWorkbook.Worksheets("Log").Range("D" & (taskLogOrder)).Comment.Text Text:=Sheets("Log").Range("D" & (taskLogOrder)).Comment.Text & " " & warningMessage
    End If

    mainWorkbook.Worksheets("Log").Range("D" & taskLogOrder).Comment.Shape.TextFrame.AutoSize = True
End Sub

' Core: Logging '

Private Sub LogBeginning(ByVal methodName As String, ByVal tickCount As Currency, ByRef logConclusionData As LogEntry, ByVal arguments As String, Optional ByVal errorMessage As String)
    operationSequenceNumber = operationSequenceNumber + 1

    logConclusionData.operationSequenceNumber = operationSequenceNumber
    logConclusionData.methodName = methodName
    logConclusionData.arguments = arguments
    logConclusionData.categoryName = methodRegistry(methodName)("Category")
    logConclusionData.tickCount = CDbl(tickCount * 10000) - baseTickCount
    logConclusionData.outcome = "Beginning"

    logWorksheet.Cells(logConclusionData.operationSequenceNumber, 1).Resize(1, 6).Value = Array(logConclusionData.methodName, logConclusionData.arguments, logConclusionData.categoryName, logConclusionData.tickCount, "", logConclusionData.outcome)

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
            .Range("A1").Value = "Declaration: " & methodRegistry(logConclusionData.methodName)("Declaration")
            .Range("A1").Font.Bold = False
            .Range("A1").Characters(Start:=1, Length:=11).Font.Bold = True

            If methodRegistry(logConclusionData.methodName).Exists("Contract") Then
                .Range("A4").Value = "Parameters: " & methodRegistry(logConclusionData.methodName)("Parameters")
                .Range("A5").Value = "Arguments: " & logConclusionData.arguments
            Else
                .Range("A4").Value = "Parameters: N/A"
                .Range("A5").Value = "Arguments: N/A"
            End If

            .Range("A4").Font.Bold = False
            .Range("A4").Font.Size = 10
            .Range("A4").Characters(Start:=1, Length:=10).Font.Bold = True
            .Range("A4").Characters(Start:=13).Font.Size = 8

            .Range("A5").Font.Bold = False
            .Range("A5").Characters(Start:=1, Length:=9).Font.Bold = True

            .Range("A9").Value = "Error Output: " & errorMessage
            .Range("A9").Font.Bold = False
            .Range("A9").Characters(Start:=1, Length:=12).Font.Bold = True

            .Tab.Color = RGB(244, 28, 80)

            .Range("A8").Value = "Date Runtime: " & utcTimestamp & " (UTC), " & CDbl(qpc * 10000)
            .Range("A8").Font.Bold = False
            .Range("A8").Font.Name = "Segoe UI"
            .Range("A8").Characters(Start:=1, Length:=12).Font.Bold = True

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
End Sub

' ************ '
' Repetition   '
' ************ '

Sub RepeatAnonymizeNumbersOnColumnOnWorksheet(ByVal columnNames As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatAnonymizeNumbersOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetName As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """", validation)

    Dim columnNamesArray() As String
    Dim columnName As String

    columnNamesArray = Split(columnNames, "|")

    Dim index As Integer
    For index = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(index)

        Call AnonymizeNumbersOnColumnOnWorksheet(columnName, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatApplyAutoFilterOnWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatApplyAutoFilterOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call ApplyAutoFilterOnWorksheet(worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatApplyCellStyleToColumnOnWorksheet(ByVal cellStyle As String, ByVal columnNames As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatApplyCellStyleToColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyle As String, ByVal columnNames As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cellStyle & """" & ", " & """" & columnNames & """" & ", " & """" & worksheetNames & """", validation)

    Dim columnNamesArray() As String
    Dim worksheetNamesArray() As String
    Dim columnName As String
    Dim worksheetName As String

    Dim index As Integer
    If InputContainsValue(columnNames, "|") Then
        columnNamesArray = Split(columnNames, "|")
        For index = 0 To UBound(columnNamesArray)
            columnName = columnNamesArray(index)

            Call ApplyCellStyleToColumnOnWorksheet(cellStyle, columnName, worksheetNames)
        Next index
    End If

    If InputContainsValue(worksheetNames, "|") Then
        worksheetNamesArray = Split(worksheetNames, "|")
        For index = 0 To UBound(worksheetNamesArray)
            worksheetName = worksheetNamesArray(index)

            Call ApplyCellStyleToColumnOnWorksheet(cellStyle, columnNames, worksheetName)
        Next index
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatSetFrozenPanesOnWorksheet(ByVal frozenRowCount As Long, ByVal frozenColumnCount As Long, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatSetFrozenPanesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal frozenRowCount As Long, ByVal frozenColumnCount As Long, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, frozenRowCount & ", " & frozenColumnCount & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call SetFrozenPanesOnWorksheet(frozenRowCount, frozenColumnCount, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatCutColumnAndPasteAtColumnOnWorksheet(ByVal cutColumnNames As String, ByVal pasteColumnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatCutColumnAndPasteAtColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal cutColumnNames As String, ByVal pasteColumnName As String, ByVal worksheetName As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(cutColumnNames, "cutColumnNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cutColumnNames & """" & ", " & """" & pasteColumnName & """" & ", " & """" & worksheetName & """", validation)

    Dim cutColumnNamesArray() As String
    Dim cutColumnName As String

    Dim index As Integer
    If InputContainsValue(cutColumnNames, "|") Then
        cutColumnNamesArray = Split(cutColumnNames, "|")
        For index = 0 To UBound(cutColumnNamesArray)
            cutColumnName = cutColumnNamesArray(index)

            Call CutColumnAndPasteAtColumnOnWorksheet(cutColumnName, pasteColumnName, worksheetName)
        Next index
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorBorderOuterAndInnerOnWorksheet(ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatColorBorderOuterAndInnerOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & outerColorName & """" & ", " & """" & innerColorName & """" & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call ColorBorderOuterAndInnerOnWorksheet(outerColorName, innerColorName, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorDataBackgroundAndFontOnWorksheet(ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatColorDataBackgroundAndFontOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String
    Dim indexWorksheetName As Integer
    Dim indexColumnName As Integer
    Dim columnName As String
    Dim worksheetName As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")

    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            columnName = columnNamesArray(indexColumnName)
            worksheetName = worksheetNamesArray(indexWorksheetName)

            Call ColorDataBackgroundAndFontOnWorksheet(columnName, backgroundColorName, fontColorName, worksheetName)
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(indexColumnName)

        Call ColorDataBackgroundAndFontOnWorksheet(columnName, backgroundColorName, fontColorName, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(indexWorksheetName)

        Call ColorDataBackgroundAndFontOnWorksheet(columnNames, backgroundColorName, fontColorName, worksheetName)
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorHeaderBackgroundAndFontOnWorksheet(ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatColorHeaderBackgroundAndFontOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String
    Dim indexWorksheetName As Integer
    Dim indexColumnName As Integer
    Dim columnName As String
    Dim worksheetName As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")

    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            columnName = columnNamesArray(indexColumnName)
            worksheetName = worksheetNamesArray(indexWorksheetName)

            Call ColorHeaderBackgroundAndFontOnWorksheet(columnName, backgroundColorName, fontColorName, worksheetName)
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(indexColumnName)

        Call ColorHeaderBackgroundAndFontOnWorksheet(columnName, backgroundColorName, fontColorName, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(indexWorksheetName)

        Call ColorHeaderBackgroundAndFontOnWorksheet(columnNames, backgroundColorName, fontColorName, worksheetName)
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorWorksheet(ByVal colorName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatColorWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & colorName & """" & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call ColorWorksheet(colorName, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatConvertColumnToDateFormattingOnWorksheet(ByVal columnNames As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatConvertColumnToDateFormattingOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetName As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """", validation)

    Dim columnNamesArray() As String
    Dim columnName As String

    columnNamesArray = Split(columnNames, "|")

    Dim index As Integer
    For index = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(index)

        Call ConvertColumnToDateFormattingOnWorksheet(columnName, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatConvertColumnToDecimalFormattingOnWorksheet(ByVal columnNames As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatConvertColumnToDecimalFormattingOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetName As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """", validation)

    Dim columnNamesArray() As String
    Dim columnName As String

    columnNamesArray = Split(columnNames, "|")

    Dim index As Integer
    For index = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(index)

        Call ConvertColumnToDecimalFormattingOnWorksheet(columnName, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatCreateWorksheet(ByVal worksheetNames As String, ByVal insertAfterWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatCreateWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String, ByVal insertAfterWorksheetName As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """" & ", " & """" & insertAfterWorksheetName & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call CreateWorksheet(worksheetName, insertAfterWorksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteColumnOnWorksheet(ByVal columnNames As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDeleteColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String
    Dim indexWorksheetName As Integer
    Dim indexColumnName As Integer
    Dim columnName As String
    Dim worksheetName As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")

    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            columnName = columnNamesArray(indexColumnName)
            worksheetName = worksheetNamesArray(indexWorksheetName)

            Call DeleteColumnOnWorksheet(columnName, worksheetName)
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(indexColumnName)

        Call DeleteColumnOnWorksheet(columnName, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(indexWorksheetName)

        Call DeleteColumnOnWorksheet(columnNames, worksheetName)
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteRowsOnWorksheet(ByVal numberOfRows As Long, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDeleteRowsOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal numberOfRows As Long, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, numberOfRows & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call DeleteRowsOnWorksheet(numberOfRows, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDeleteWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call DeleteWorksheet(worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDuplicateColumnFromWorksheetToWorksheet(ByVal columnNames As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDuplicateColumnFromWorksheetToWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & fromWorksheetName & """" & ", " & """" & toWorksheetName & """", validation)

    Dim columnNamesArray() As String
    Dim columnName As String

    columnNamesArray = Split(columnNames, "|")

    Dim index As Integer
    For index = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(index)

        Call DuplicateColumnFromWorksheetToWorksheet(columnName, fromWorksheetName, toWorksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatFindAndReplaceInColumnOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal columnNames As String, ByVal worksheetName As String, ByVal exactMatch As Boolean)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatFindAndReplaceInColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal columnNames As String, ByVal worksheetName As String, ByVal exactMatch As Boolean", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & columnNames & """" & ", " & """" & worksheetName & """" & ", " & exactMatch, validation)

    Dim columnNamesArray() As String
    Dim columnName As String

    columnNamesArray = Split(columnNames, "|")

    Dim index As Integer
    For index = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(index)

        Call FindAndReplaceInColumnOnWorksheet(findValue, replaceValue, columnName, worksheetName, exactMatch)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatHideWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatHideWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call HideWorksheet(worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatInsertDataValuesOnWorksheet(ByVal dataValues As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatInsertDataValuesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal dataValues As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(dataValues, "dataValues", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & dataValues & """" & ", "  & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call InsertDataValuesOnWorksheet(dataValues, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatAddColumnOnWorksheet(ByVal columnNames As String, ByVal columnWidth As Double, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatAddColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal columnWidth As Double, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", "  & columnWidth & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String
    Dim indexWorksheetName As Integer
    Dim indexColumnName As Integer
    Dim columnName As String
    Dim worksheetName As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")

    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            columnName = columnNamesArray(indexColumnName)
            worksheetName = worksheetNamesArray(indexWorksheetName)

            Call AddColumnOnWorksheet(columnName, columnWidth, worksheetName)
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(indexColumnName)

        Call AddColumnOnWorksheet(columnName, columnWidth, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(indexWorksheetName)

        Call AddColumnOnWorksheet(columnNames, columnWidth, worksheetName)
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatInsertHeaderValuesOnWorksheet(ByVal headerValues As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatInsertHeaderValuesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal headerValues As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(headerValues, "headerValues", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & headerValues & """" & ", "  & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call InsertHeaderValuesOnWorksheet(headerValues, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatMoveWorksheetToEnd(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatMoveWorksheetToEnd"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call MoveWorksheetToEnd(worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatNormalizeLayoutOnWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatNormalizeLayoutOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call NormalizeLayoutOnWorksheet(worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatRemoveBlankLinesOnWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatRemoveBlankLinesOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """", validation)

    Dim worksheetsArray() As String
    Dim worksheetName As String

    worksheetsArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetsArray)
        worksheetName = worksheetsArray(index)

        Call RemoveBlankLinesOnWorksheet(worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteCellStyle(ByVal cellStyleNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDeleteCellStyle"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyleNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(cellStyleNames, "cellStyleNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cellStyleNames & """", validation)

    Dim cellStyleNamesArray() As String
    Dim cellStyle As String

    cellStyleNamesArray = Split(cellStyleNames, "|")

    Dim index As Integer
    For index = 0 To UBound(cellStyleNamesArray)
        cellStyle = cellStyleNamesArray(index)

        Call DeleteCellStyle(cellStyle)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatRemoveEmptyColumnsOnWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatRemoveEmptyColumnsOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call RemoveEmptyColumnsOnWorksheet(worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatSetNewWidthOnColumnOnWorksheet(ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatSetNewWidthOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim worksheetName As String

    worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(index)

        Call SetNewWidthOnColumnOnWorksheet(newWidth, columnName, worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatSortColumnByOrderOnWorksheet(ByVal columnNames As String, ByVal sortOrder As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatSortColumnByOrderOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal sortOrder As String, ByVal worksheetNames As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & sortOrder & """" & ", " & """" & worksheetNames & """", validation)

    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String
    Dim indexWorksheetName As Integer
    Dim indexColumnName As Integer
    Dim columnName As String
    Dim worksheetName As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")
    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            columnName = columnNamesArray(indexColumnName)
            worksheetName = worksheetNamesArray(indexWorksheetName)

            Call SortColumnByOrderOnWorksheet(columnName, sortOrder, worksheetName)
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        columnName = columnNamesArray(indexColumnName)

        Call SortColumnByOrderOnWorksheet(columnName, sortOrder, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        worksheetName = worksheetNamesArray(indexWorksheetName)

        Call SortColumnByOrderOnWorksheet(columnNames, sortOrder, worksheetName)
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatTransferDataFromWorksheetToWorksheet(ByVal fromWorksheetNames As String, toWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatTransferDataFromWorksheetToWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetNames As String, toWorksheetName As String", methodName, "Repetition")
    End If

    Dim validation As String
    Call ValidateRepeatArgument(fromWorksheetNames, "fromWorksheetNames", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & fromWorksheetNames & """" & ", " & """" & toWorksheetName & """", validation)

    Dim fromWorksheetNamesArray() As String
    Dim fromWorksheetName As String

    fromWorksheetNamesArray = Split(fromWorksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(fromWorksheetNamesArray)
        fromWorksheetName = fromWorksheetNamesArray(index)

        Call TransferDataFromWorksheetToWorksheet(fromWorksheetName, toWorksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

' No logging. '

Sub RepeatAssertMinimumRowsOfDataOnWorksheet(ByVal minimumRowsOfData As Long, ByVal worksheetNames As String)
    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")

    Dim index As Integer
    For index = 0 To UBound(worksheetNamesArray)
        Call AssertMinimumRowsOfDataOnWorksheet(minimumRowsOfData, worksheetNamesArray(index))
    Next index
End Sub

' Functions: Repetition '

' ************ '
' Sequencing   '
' ************ '

Sub AnonymizeNumbersOnColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName. '

If InputContainsValue(columnName, "|") Then
    Call RepeatAnonymizeNumbersOnColumnOnWorksheet(columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "AnonymizeNumbersOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Dim formulaColumn As String
    Dim columnReference As String
    formulaColumn = "Formula Column"
    columnReference = ColumnStartOnWorksheet(columnName, worksheetName)

    ' https://web.archive.org/web/20201028212900/https://heroes.thelazy.net/index.php/Creature
    Call AddColumnOnWorksheet(formulaColumn, 13.57, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!1000+", "=" & columnReference & ">=1000", formulaColumn, worksheetName)
    ' https://www.thepunctuationguide.com/en-dash.html: The en dash (�) is slightly wider than the hyphen (-) but narrower than the em dash (�). The en dash is used to represent a span or range of numbers, dates, or time. '
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!500" & WorksheetFunction.Unichar(8211) & "999", "=AND(" & columnReference & ">=500;" & columnReference & "<=999)", formulaColumn, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!250" & WorksheetFunction.Unichar(8211) & "499", "=AND(" & columnReference & ">=250;" & columnReference & "<=499)", formulaColumn, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!100" & WorksheetFunction.Unichar(8211) & "249", "=AND(" & columnReference & ">=100;" & columnReference & "<=249)", formulaColumn, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!50" & WorksheetFunction.Unichar(8211) & "99", "=AND(" & columnReference & ">=50;" & columnReference & "<=99)", formulaColumn, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!20" & WorksheetFunction.Unichar(8211) & "49", "=AND(" & columnReference & ">=20;" & columnReference & "<=49)", formulaColumn, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!10" & WorksheetFunction.Unichar(8211) & "19", "=AND(" & columnReference & ">=10;" & columnReference & "<=19)", formulaColumn, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!5" & WorksheetFunction.Unichar(8211) & "9", "=AND(" & columnReference & ">=5;" & columnReference & "<=9)", formulaColumn, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!1" & WorksheetFunction.Unichar(8211) & "4", "=AND(" & columnReference & ">=1;" & columnReference & "<=4)", formulaColumn, worksheetName)
    Call ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet("!0", "=" & columnReference & "=0", formulaColumn, worksheetName)

    Call ApplyFormulaToColumnOnWorksheet("=SUBSTITUTE(TEXT(" & ColumnStartOnWorksheet(formulaColumn, worksheetName) & ";""@"");""!"";"""")", columnName, worksheetName)
    Call ApplyCellStyleToColumnOnWorksheet("Normal", columnName, worksheetName)
    Call DeleteColumnOnWorksheet(formulaColumn, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ApplyConsecutiveGroupBasedOnFormulaOnColumnOnWorksheet(ByVal consecutiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyConsecutiveGroupBasedOnFormulaOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal consecutiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & consecutiveGroupName & """" & ", " & """" & formulaValue & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    Call ApplyFormulaToColumnOnWorksheet(formulaValue, helperColumn, worksheetName)
    Call ApplyFilterToColumnOnWorksheet(True, helperColumn, worksheetName)

    Call FindAndReplaceInColumnOnWorksheet(";;", ";" & consecutiveGroupName & ";;", columnName, worksheetName)
    Call ClearFilterOnWorksheet(worksheetName)

    Call SelectWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet(ByVal definitiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal definitiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & definitiveGroupName & """" & ", " & """" & formulaValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """", validation)

    If formulaValue = "" Then formulaValue = "=True"

    Call ApplyFormulaToColumnOnWorksheet(formulaValue, helperColumn, worksheetName)
    Call ApplyFilterToColumnOnWorksheet(True, helperColumn, worksheetName)
    Call FindAndReplaceInColumnOnWorksheet(";;", definitiveGroupName, columnName, worksheetName)
    Call ClearFilterOnWorksheet(worksheetName)

    Call SelectWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ColorBorderOuterAndInnerOnWorksheet(ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatColorBorderOuterAndInnerOnWorksheet(outerColorName, innerColorName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorBorderOuterAndInnerOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & outerColorName & """" & ", " & """" & innerColorName & """" & ", " & """" & worksheetName & """", validation)

    Dim worksheet As Worksheet

    Set worksheet = mainWorkbook.Worksheets(worksheetName)

    If outerColorName = "" Then outerColorName = "Light Grey"
    If innerColorName = "" Then innerColorName = "Light Grey"

    Dim lastColumnLetterPlusOne As String
    lastColumnLetterPlusOne = LastUsedColumnLetterOnWorksheet(worksheetName)
    lastColumnLetterPlusOne = IncreaseLetterOnce(lastColumnLetterPlusOne)
    lastColumnLetterPlusOne = IncreaseLetterOnce(lastColumnLetterPlusOne)

    ' Outer Border. '
    With worksheet.Cells.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = StyleColor(outerColorName)
        .Weight = xlThin
    End With
    With worksheet.Cells.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = StyleColor(outerColorName)
        .Weight = xlThin
    End With

    Dim lastRowNumberPlusOne As Long
    lastRowNumberPlusOne = LastUsedRowNumberOnWorksheet(worksheetName) + 1

    worksheet.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    worksheet.Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Inner Border. '
    With worksheet.Range("A1:" & lastColumnLetterPlusOne & lastRowNumberPlusOne + 1).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = StyleColor(innerColorName)
        .Weight = xlThin
    End With
    With worksheet.Range("A1:" & lastColumnLetterPlusOne & lastRowNumberPlusOne + 1).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = StyleColor(innerColorName)
        .Weight = xlThin
    End With

    ' Clean up of bug or weird behaviour which doesn't apply the border everywhere. '
    lastColumnLetterPlusOne = DecreaseLetterOnce(lastColumnLetterPlusOne)
    worksheet.Columns("A:A").Delete Shift:=xlToLeft
    worksheet.Rows("1:1").Delete Shift:=xlUp
    worksheet.Columns(lastColumnLetterPlusOne & ":" & lastColumnLetterPlusOne).Delete Shift:=xlToLeft
    worksheet.Columns(lastColumnLetterPlusOne & ":" & lastColumnLetterPlusOne).Delete Shift:=xlToLeft
    worksheet.Rows(lastRowNumberPlusOne & ":" & lastRowNumberPlusOne).Delete Shift:=xlUp
    worksheet.Rows(lastRowNumberPlusOne & ":" & lastRowNumberPlusOne).Delete Shift:=xlUp

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ColorDataBackgroundAndFontOnWorksheet(ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatColorHeaderBackgroundAndFontOnWorksheet(columnName, backgroundColorName, fontColorName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorDataBackgroundAndFontOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetName & """", validation)

    If backgroundColorName = "" Then backgroundColorName = "White"
    If fontColorName = "" Then fontColorName = "Black"

    If NumberOfVisibleCells(worksheetName) <> 0 Then
        Call ColorColumnOfTypeAndElementOnWorksheet(backgroundColorName, columnName, "Data", "Background", worksheetName)
        Call ColorColumnOfTypeAndElementOnWorksheet(fontColorName, columnName, "Data", "Font", worksheetName)
    Else
        ' Call LogTaskWarning("Worksheet " & """" & worksheetName & """" & " currently has no visible data rows.", logOrderLocal)
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ColorHeaderBackgroundAndFontOnWorksheet(ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatColorHeaderBackgroundAndFontOnWorksheet(columnName, backgroundColorName, fontColorName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorHeaderBackgroundAndFontOnWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetName & """", validation)

    If backgroundColorName = "" Then backgroundColorName = "White"
    If fontColorName = "" Then fontColorName = "Black"

    Call ColorColumnOfTypeAndElementOnWorksheet(backgroundColorName, columnName, "Header", "Background", worksheetName)
    Call ColorColumnOfTypeAndElementOnWorksheet(fontColorName, columnName, "Header", "Font", worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ImportExcelFileWithDependencyAndRenameWorksheet(ByVal filePath As String, ByVal dependency As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ImportExcelFileWithDependencyAndRenameWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal dependency As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & dependency & """" & " ," & """" & worksheetName & """")

    Call ImportExcelFileAndRenameWorksheet(filePath, worksheetName)
    Call SetAboutNamedRange(dependency, "Dependencies List")

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub MoveDataFromWorksheetToWorksheet(ByVal fromWorksheetName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "MoveDataFromWorksheetToWorksheet"
    Dim isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If methodRegistry.Exists(methodName) = True Then If methodRegistry(methodName).Exists("Registered") = True Then isRegistered = True
    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetName As String, ByVal worksheetName As String", methodName, "Sequencing")
    End If

    Dim validation As String
    Call ValidateWorksheet(fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & fromWorksheetName & """" & ", " & """" & worksheetName & """", validation)

    Call EnsureHelperColumnDeletedOnWorksheet(fromWorksheetName)
    Call EnsureHelperColumnDeletedOnWorksheet(worksheetName)
   
    Dim numberOfColumnsFromWorksheetName As Long: numberOfColumnsFromWorksheetName = ConvertColumnLetterToColumnNumber(LastUsedColumnLetterOnWorksheet(fromWorksheetName))
    Dim numberOfColumnsWorksheetName As Long: numberOfColumnsWorksheetName = ConvertColumnLetterToColumnNumber(LastUsedColumnLetterOnWorksheet(worksheetName))

    If numberOfColumnsFromWorksheetName <> numberOfColumnsWorksheetName Then Call LogConclusion("Failed", logConclusionData, "Number of columns in worksheet """ & fromWorksheetName & """ doesn't match up with the worksheet """ & worksheetName & """.")

    Dim pasteRow As Long: pasteRow = LastUsedRowNumberOnWorksheet(worksheetName) + 1

    If NumberOfVisibleCells(fromWorksheetName) <> 0 Then
        mainWorkbook.Worksheets(fromWorksheetName).Range("A2:" & LastUsedColumnLetterOnWorksheet(fromWorksheetName) & LastUsedRowNumberOnWorksheet(fromWorksheetName)).Copy
        mainWorkbook.Worksheets(worksheetName).Range("A" & pasteRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If

    Call DeleteWorksheet(fromWorksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Functions: Sequencing'

Function Lookup(ByVal sourceColumnName As String, ByVal sourceWorksheetName As String, ByVal matchColumnName As String, ByVal targetColumnName As String, ByVal targetWorksheetName As String) As String
    Dim lookupValue As String
    lookupValue = ColumnStartOnWorksheet(sourceColumnName, sourceWorksheetName) & ";" & ColumnDataOnWorksheet(matchColumnName, targetWorksheetName) & ";" & ColumnDataOnWorksheet(targetColumnName, targetWorksheetName)

    Lookup = "=IF(AND(NOT(ISBLANK(XLOOKUP(" & lookupValue & ")));NOT(ISNA(XLOOKUP(" & lookupValue & "))));XLOOKUP(" & lookupValue & ");"""")"
End Function

Function LookupIsNotAvailable(ByVal sourceColumnName As String, ByVal sourceWorksheetName As String, ByVal matchColumnName As String, ByVal targetColumnName As String, ByVal targetWorksheetName As String) As String
    Dim lookupValue As String
    lookupValue = ColumnStartOnWorksheet(sourceColumnName, sourceWorksheetName) & ";" & ColumnDataOnWorksheet(matchColumnName, targetWorksheetName) & ";" & ColumnDataOnWorksheet(targetColumnName, targetWorksheetName)

    LookupIsNotAvailable = "=ISNA(XLOOKUP(" & lookupValue & "))"
End Function

' ************ '
' Validation   '
' ************ '

Sub AssertFilePathNotEmptyForVariable(ByVal filePath As String, ByVal variableDescription)
    If filePath = "" Then
        Call LogAssertFailure("Filepath is blank for the variable " & """" & variableDescription & """" & ".")
    End If
End Sub

Sub AssertMatchingInputs(ByVal firstValue As String, ByVal secondValue As String)
    If firstValue <> secondValue Then Call LogAssertFailure("First value " & """" & firstValue & """" & " does not match second value " & """" & secondValue & """" & ".")
End Sub

Sub AssertMinimumRowsOfDataOnWorksheet(ByVal minimumRowsOfData As Long, ByVal worksheetName As String)

If InputContainsValue(worksheetName, "|") Then
    Call RepeatAssertMinimumRowsOfDataOnWorksheet(minimumRowsOfData, worksheetName)
Else
    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)

    Dim rowsOfDataFound As Long
    rowsOfDataFound = LastUsedRowNumberOnWorksheet(worksheetName) - 1

    If rowsOfDataFound < minimumRowsOfData Then
        Call LogAssertFailure("Worksheet " & """" & worksheetName & """" & " has " & rowsOfDataFound & " row(s) of data, less than the " & minimumRowsOfData & " row(s) of data expected.")
    End If
End If

End Sub

Sub EnsureHelperColumnOnWorksheet(ByVal worksheetName As String)
    ' If ColumnOnWorksheetExists(helperColumn, worksheetName) Then
    '     Call ResetColumnOnWorksheet(helperColumn, worksheetName)
    ' Else
    '     Call AddColumnOnWorksheet(helperColumn, 7.00, worksheetName)
    ' End If
End Sub

Sub EnsureHelperColumnDeletedOnWorksheet(ByVal worksheetName As String)
    ' If ColumnOnWorksheetExists(helperColumn, worksheetName) Then
    '     Call DeleteColumnOnWorksheet(helperColumn, worksheetName)
    ' End If
End Sub

Sub EnsureHelperColumnsDeletedOnMainWorkbook()
    Dim worksheetNumber As Long
    Dim index As Long

    worksheetNumber = mainWorkbook.Worksheets.count

    For index = 1 To worksheetNumber
        Call EnsureHelperColumnDeletedOnWorksheet(mainWorkbook.Worksheets(index).Name)
    Next
End Sub

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
    ElseIf InStr(columnName, "*") > 0 Then
        validationMessage = "Column name can't contain an asterisk."
    ElseIf InStr(columnName, "?") > 0 Then
        validationMessage = "Column name can't contain a question mark."
    ElseIf mainWorkbook.Worksheets(worksheetName).Range("A1").Value = "" And validColumn = False Then
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

Sub ValidateRepeatArgument(ByVal argument As String, ByVal parameterName As String, ByRef validation As String)
    Dim validationMessage As String

    If InStr(argument, "||") <> 0 Then
        validationMessage = "Argument contains multiple instances of ""||"" next to each other."
    ElseIf Right(argument, 1) = "|" Then
        validationMessage = "Argument contains instance of ""|"" at the very end."
    ElseIf Left(argument, 1) = "|" Then
        validationMessage = "Argument contains instance of ""|"" at the very start."
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
    Dim worksheetNameExists As Boolean: worksheetNameExists = WorksheetExists(worksheetName)

    If worksheetNameExists = True And validWorksheet = True Then
        validationMessage = "Worksheet name already exists."
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
        Dim forbiddenCharacters As String: forbiddenCharacters = "/\?*:[]"
        Dim characterIndex As Integer
        Dim currentForbiddenCharacter As String
    
        For characterIndex = 1 To Len(forbiddenCharacters)
            currentForbiddenCharacter = Mid$(forbiddenCharacters, characterIndex, 1)
            
            If InStr(1, worksheetName, currentForbiddenCharacter, vbBinaryCompare) > 0 Then
                validationMessage = "Worksheet name has a forbidden character: " & currentForbiddenCharacter & "."
                Exit For
            End If
        Next characterIndex

        If validationMessage = "" And worksheetNameExists = False And validWorksheet = False Then
            validationMessage = "Worksheet name not found."
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

Function InputContainsValue(ByVal inputText As String, ByVal searchValue As String, Optional ByVal caseSensitive As Boolean) As Boolean
    Dim compareMode As VbCompareMethod

    If Len(searchValue) = 0 Then
        InputContainsValue = False
        Exit Function
    End If

    If caseSensitive Then
        compareMode = vbBinaryCompare
    Else
        compareMode = vbTextCompare
    End If

    InputContainsValue = InStr(1, inputText, searchValue, compareMode) > 0
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
    Set report = CreateObject("Scripting.Dictionary")
    Set telemetry = CreateObject("Scripting.Dictionary")

    baseQpc = CDbl(logEngineQpc * 10000)
    baseTickCount = CDbl(logEngineTickCount * 10000)
    baseUtcTimestamp = logEngineUtcTimestamp
    helperColumn = "Helper Column"
    operationSequenceNumber = 1&

    report("Log Engine Active") = False
    report("Original Workbook") = mainWorkbook.FullName
    report("Creation Date") = Format$(Date, "yyyy-mm-dd")
    report("Template Version") = "v0.40, 2026-08-06"

    Set report("Settings") = New Collection

    Call Startup()
    Call Master()
End Sub

' Script is finished when "About" is placed last. '
' When saving the document, choose yes at prompt. '