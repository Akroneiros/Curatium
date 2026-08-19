' ************ '
' Alteration   '
' ************ '

Sub ApplyActiveFormulaToRangeOnWorksheet(ByVal activeFormulaValue As String, ByVal rangeValue As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyActiveFormulaToRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal activeFormulaValue As String, ByVal rangeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & activeFormulaValue & """" & ", " & """" & rangeValue & """" & ", " & """" & worksheetName & """")
    
    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If rangeValue = report("Helper Column") Then Call EnsureHelperColumnOnWorksheet(worksheetName)
    Call ValidateRangeOnWorksheet(rangeValue, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If mainWorkbook.Worksheets(worksheetName).Range(FirstColumnInRange(rangeValue) & FirstRowInRange(rangeValue)).Style <> "Formula" Then Call ApplyCellStyleToColumnOnWorksheet("Formula", rangeValue, worksheetName)
    mainWorkbook.Worksheets(worksheetName).Range(FirstColumnInRange(rangeValue) & FirstRowInRange(RangeValue)).FormulaLocal = activeFormulaValue

    Call SelectWorksheet(worksheetName)

    If mainWorkbook.Worksheets(worksheetName).Range("A2").Value = "" Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & worksheetName & """" & " doesn't appear to have any valid data, can't apply formula.")

    If Range(rangeValue).Rows.Count <> 1 Then
        mainWorkbook.Worksheets(worksheetName).Range(rangeValue).FillDown
    End If
    
    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyFormulaToCellWithCellStyleOnWorksheet(ByVal formulaValue As String, ByVal cellValue As String, ByVal cellStyle As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFormulaToCellWithCellStyleOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal cellValue As String, ByVal cellStyle As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaValue & """" & ", " & """" & cellValue & """" & ", " & """" & cellStyle & """" & ", " & """" & worksheetName & """")
    
    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If CellStyleExists(cellStyle) = False Then
        Call LogConclusion("Failed", logConclusionData, "Cell style """ & cellStyle & """ does not exist.")
    End If

    If mainWorkbook.Worksheets(worksheetName).Range(cellValue).Style <> "Formula" Then
        With mainWorkbook.Worksheets(worksheetName).Range(cellValue)
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
    mainWorkbook.Worksheets(worksheetName).Range(cellValue).Formula2 = formulaValue
    If Err.Number <> 0 Then
        Err.Clear
        mainWorkbook.Worksheets(worksheetName).Range(cellValue).FormulaLocal = formulaValue
    End If
    
    If mainWorkbook.Worksheets(worksheetName).Range(cellValueLetter & cellValueNumber) <> "" Then
        mainWorkbook.Worksheets(worksheetName).Range(Sheets(worksheetName).Range(cellValue), mainWorkbook.Worksheets(worksheetName).Range(cellValue).End(xlDown)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    Else
        mainWorkbook.Worksheets(worksheetName).Range(cellValue).Copy
        mainWorkbook.Worksheets(worksheetName).Range(cellValue).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If

    With mainWorkbook.Worksheets(worksheetName).Range(cellValue)
        .Style = cellStyle
        .Value = .Value
    End With

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyFormulaToColumnOnWorksheet(ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFormulaToColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")
    
    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If columnName = report("Helper Column") Then Call EnsureHelperColumnOnWorksheet(worksheetName)
    Call ValidateRangeOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If mainWorkbook.Worksheets(worksheetName).Range(FirstColumnInRange(columnName) & FirstRowInRange(columnName)).Style <> "Formula" Then Call ApplyCellStyleToColumnOnWorksheet("Formula", columnName, worksheetName)
    mainWorkbook.Worksheets(worksheetName).Range(FirstColumnInRange(columnName) & FirstRowInRange(columnName)).FormulaLocal = formulaValue

    Call SelectWorksheet(worksheetName)

    If mainWorkbook.Worksheets(worksheetName).Range("A2").Value = "" Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & worksheetName & """" & " doesn't appear to have any valid data, can't apply formula.")

    If Range(columnName).Rows.Count <> 1 Then
        mainWorkbook.Worksheets(worksheetName).Range(columnName).FillDown
    End If
    
    ' To avoid bugs with number formatting, i.e. leading zeroes. '
    mainWorkbook.Worksheets(worksheetName).Range(columnName).Copy
    mainWorkbook.Worksheets(worksheetName).Range(columnName).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyHyperlinkToCellOnWorksheet(ByVal hyperlinkValue As String, ByVal cellValue As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyHyperlinkToCellOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal hyperlinkValue As String, ByVal cellValue As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & hyperlinkValue & """" & ", " & """" & cellValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim currentCellValue As String: currentCellValue = mainWorkbook.Worksheets(worksheetName).Range(cellValue).Value

    If InputContainsValue(hyperlinkValue, "http") Or InputContainsValue(hyperlinkValue, "https") Then
        mainWorkbook.Worksheets(worksheetName).Hyperlinks.Add Anchor:=Sheets(worksheetName).Range(cellValue), Address:="", SubAddress:=hyperlinkValue, TextToDisplay:=currentCellValue
    Else
        mainWorkbook.Worksheets(worksheetName).Hyperlinks.Add Anchor:=Sheets(worksheetName).Range(cellValue), Address:="", SubAddress:="'" & hyperlinkValue & "'!A1", TextToDisplay:=currentCellValue
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub FindAndReplaceOnRangeOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValue As String, ByVal worksheetName As String)  ' Repeat Support: rangeValue. '

If InputContainsValue(rangeValue, "|") Then
    Call RepeatFindAndReplaceOnRangeOnWorksheet(findValue, replaceValue, rangeValue, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "FindAndReplaceOnRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & rangeValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateRangeOnWorksheet(rangeValue, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Range(rangeValue).Replace What:=findValue, Replacement:=replaceValue, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub HideWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatHideWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "HideWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Visible = xlSheetHidden

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertDataValuesOnWorksheet(ByVal dataValues As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatInsertDataValuesOnWorksheet(dataValues, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "InsertDataValuesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal dataValues As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & dataValues & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim delimitedStringArray() As String: delimitedStringArray = Split(dataValues, "|")
    Dim currentLastDataRow As Long: currentLastDataRow = LastUsedRowNumberOnWorksheet(worksheetName) + 1
    Dim index As Long

    For index = 0 To UBound(delimitedStringArray)
        mainWorkbook.Worksheets(worksheetName).Range(ConvertNumberToLetter(index + 1) & currentLastDataRow).Value = delimitedStringArray(index)
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal headerValues As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & headerValues & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    Dim delimitedStringArray() As String: delimitedStringArray = Split(headerValues, "|")
    Dim index As Long

    For index = 0 To UBound(delimitedStringArray)
        mainWorkbook.Worksheets(worksheetName).Range(ConvertNumberToLetter(index + 1) & "1").Value = delimitedStringArray(index)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertNewColumnAndSetWidthOnWorksheet(ByVal columnName As String, ByVal setWidth As Double, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatInsertNewColumnAndSetWidthOnWorksheet(columnName, setWidth, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "InsertNewColumnAndSetWidthOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal setWidth As Double, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & setWidth & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateUniqueColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim lastColumnLetterPlusOne As String
    Dim lastColumnRow As Long

    lastColumnRow = LastUsedRowNumberOnWorksheet(worksheetName)
    lastColumnLetterPlusOne = LastUsedColumnLetterOnWorksheet(worksheetName)
    lastColumnLetterPlusOne = IncreaseLetterOnce(lastColumnLetterPlusOne)

    If mainWorkbook.Worksheets(worksheetName).Range("A1").Value = "" Then lastColumnLetterPlusOne = "A"

    mainWorkbook.Worksheets(worksheetName).Range(lastColumnLetterPlusOne & "1").Value = columnName
    If setWidth <> 0 Then mainWorkbook.Worksheets(worksheetName).Range(lastColumnLetterPlusOne & "1").ColumnWidth = setWidth

    With mainWorkbook.Worksheets(worksheetName).Range(lastColumnLetterPlusOne & "1")
        .Style = "Header"
        .Value = .Value
    End With

    Call ResetAutoFilterOnWorksheet(worksheetName)
    Call ApplyCellStyleToColumnOnWorksheet("Formula", FindColumnLetterOnWorksheet(columnName, worksheetName), worksheetName)
    Call ResetColumnOnWorksheet(lastColumnLetterPlusOne, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub MoveWorksheetToEnd(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatMoveWorksheetToEnd(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "MoveWorksheetToEnd"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetIsHidden As Boolean: worksheetIsHidden = False
    If mainWorkbook.Worksheets(worksheetName).Visible = xlSheetHidden Then worksheetIsHidden = True

    If worksheetIsHidden = True Then
        mainWorkbook.Worksheets(worksheetName).Visible = xlSheetVisible
    End If

    mainWorkbook.Worksheets(worksheetName).Move After:=Sheets(Worksheets.Count)

    If worksheetIsHidden = True Then
        mainWorkbook.Worksheets(worksheetName).Visible = xlSheetHidden
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RenameWorksheetToValue(ByVal currentWorksheetName As String, ByVal newWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RenameWorksheetToValue"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal newWorksheetName As String", methodName, isRegistered, "Alteration")
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If mainWorkbook.Worksheets(worksheetName).Range("A2").Value <> "" Then mainWorkbook.Worksheets(worksheetName).Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName)).FormulaArray = ";;"

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetMethodSetting(ByVal settingMethod As String, ByVal settingName As String, ByVal settingValue As Long)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetMethodSetting"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal settingMethod As String, ByVal settingName As String, ByVal settingValue As Long", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & settingMethod & """" & ", " & """" & settingName & """" & ", " & settingValue)

    Call ConfigureMethodSetting(settingMethod, settingName, settingValue)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetNewHeaderOnColumnOnWorksheet(ByVal newHeader As String, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetNewHeaderOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal newHeader As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & newHeader & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Range(columnName & "1").Value = newHeader

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetNewWidthOnColumnOnWorksheet(ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatSetNewWidthOnColumnOnWorksheet(newWidth, columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SetNewWidthOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, newWidth & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Range(columnName & "1").ColumnWidth = newWidth

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

    decreasedNumber = ConvertLetterToNumber(letterToDecrease)
    decreasedNumber = decreasedNumber - 1
    DecreaseLetterOnce = ConvertNumberToLetter(decreasedNumber)
End Function

Function IfColumnOnWorksheetExistsReturnValue(ByVal columnName As String, ByVal worksheetName As String, ByVal returnValue As String) As String
    IfColumnOnWorksheetExistsReturnValue = ""
    If ColumnOnWorksheetExists(columnName, worksheetName) Then IfColumnOnWorksheetExistsReturnValue = returnValue
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

    increasedNumber = ConvertLetterToNumber(letterToIncrease)
    increasedNumber = increasedNumber + 1   
    IncreaseLetterOnce = ConvertNumberToLetter(increasedNumber)
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

Sub CreateAboutWorksheet()
    Dim worksheetEntry As Worksheet
    Dim legacyTemplateVersion As Boolean

    If WorksheetExists("About") = False Then
        If NamedRangeExists("ReportDetails") = True Then mainWorkbook.Names("ReportDetails").Delete
        If NamedRangeExists("TemplateDetails") = True Then mainWorkbook.Names("TemplateDetails").Delete
        If NamedRangeExists("RetrievedDate") = True Then mainWorkbook.Names("RetrievedDate").Delete
        If NamedRangeExists("ScriptDuration") = True Then mainWorkbook.Names("ScriptDuration").Delete

        If NamedRangeExists("ReportName") = True Then mainWorkbook.Names("ReportName").Delete
        If NamedRangeExists("TemplateVersion") = True Then mainWorkbook.Names("TemplateVersion").Delete
        If NamedRangeExists("ProgressionStatus") = True Then mainWorkbook.Names("ProgressionStatus").Delete
        If NamedRangeExists("AugmentationModules") = True Then mainWorkbook.Names("AugmentationModules").Delete
        If NamedRangeExists("ReportVision") = True Then mainWorkbook.Names("ReportVision").Delete
        If NamedRangeExists("DependenciesList") = True Then mainWorkbook.Names("DependenciesList").Delete
        If NamedRangeExists("CreationDate") = True Then mainWorkbook.Names("CreationDate").Delete
        If NamedRangeExists("EditionName") = True Then mainWorkbook.Names("EditionName").Delete
        If NamedRangeExists("DurationMilliseconds") = True Then mainWorkbook.Names("DurationMilliseconds").Delete
        If NamedRangeExists("LogSummary") = True Then mainWorkbook.Names("LogSummary").Delete

        mainWorkbook.Worksheets.Add After:=mainWorkbook.ActiveSheet
        mainWorkbook.ActiveSheet.Name = "About"

        If Application.DisplayAlerts = True Then
            Application.DisplayAlerts = False
        End If

        For Each worksheetEntry In mainWorkbook.Worksheets
            If WorksheetIsEmpty(worksheetEntry.Name) Then
                If worksheetEntry.Name <> "About" Then
                    mainWorkbook.Worksheets(worksheetEntry.Name).Delete
                End If
            End If
        Next worksheetEntry

        Application.DisplayAlerts = True
    Else
        If mainWorkbook.Worksheets("About").Visible = xlSheetHidden Then
            mainWorkbook.Worksheets("About").Visible = xlSheetVisible
        End If
    End If

    If NamedRangeExists("ReportDetails") = True Then mainWorkbook.Names("ReportDetails").Delete
    If NamedRangeExists("TemplateDetails") = True Then mainWorkbook.Names("TemplateDetails").Delete
    If NamedRangeExists("RetrievedDate") = True Then mainWorkbook.Names("RetrievedDate").Delete
    If NamedRangeExists("ScriptDuration") = True Then mainWorkbook.Names("ScriptDuration").Delete

    If NamedRangeExists("ReportName") = False Then mainWorkbook.Names.Add Name:="ReportName", RefersToR1C1:="=About!R1C1"
    If NamedRangeExists("TemplateVersion") = False Then mainWorkbook.Names.Add Name:="TemplateVersion", RefersToR1C1:="=About!R2C1"
    If NamedRangeExists("ProgressionStatus") = False Then mainWorkbook.Names.Add Name:="ProgressionStatus", RefersToR1C1:="=About!R3C1"
    If NamedRangeExists("AugmentationModules") = False Then mainWorkbook.Names.Add Name:="AugmentationModules", RefersToR1C1:="=About!R4C1"
    If NamedRangeExists("ReportVision") = False Then mainWorkbook.Names.Add Name:="ReportVision", RefersToR1C1:="=About!R1C2"
    If NamedRangeExists("DependenciesList") = False Then mainWorkbook.Names.Add Name:="DependenciesList", RefersToR1C1:="=About!R3C2"
    If NamedRangeExists("CreationDate") = False Then mainWorkbook.Names.Add Name:="CreationDate", RefersToR1C1:="=About!R1C3"
    If NamedRangeExists("EditionName") = False Then mainWorkbook.Names.Add Name:="EditionName", RefersToR1C1:="=About!R2C3"
    If NamedRangeExists("DurationMilliseconds") = False Then mainWorkbook.Names.Add Name:="DurationMilliseconds", RefersToR1C1:="=About!R3C3"
    If NamedRangeExists("LogSummary") = False Then mainWorkbook.Names.Add Name:="LogSummary", RefersToR1C1:="=About!R4C3"

    If report.Exists("Name") Then
        If mainWorkbook.Worksheets("About").Range("ReportName").Value = "" Then mainWorkbook.Worksheets("About").Range("ReportName").Value = report("Name")
    Else
        If mainWorkbook.Worksheets("About").Range("ReportName").Value = "" Then mainWorkbook.Worksheets("About").Range("ReportName").Value = "N/A"
    End If

    If mainWorkbook.Worksheets("About").Range("TemplateVersion").Value = "Spreadsheet Operations Template (v0.39, 16.02.2024)" Then
        legacyTemplateVersion = True
    End If

    mainWorkbook.Worksheets("About").Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & report("Template Version") & ")"

    If mainWorkbook.Worksheets("About").Range("ProgressionStatus").Value = "" Then mainWorkbook.Worksheets("About").Range("ProgressionStatus").Value = "Progression Status: N/A."
    If mainWorkbook.Worksheets("About").Range("AugmentationModules").Value = "" Then mainWorkbook.Worksheets("About").Range("AugmentationModules").Value = "Augmentation Modules: N/A."

    If report.Exists("Vision") Then
        If mainWorkbook.Worksheets("About").Range("ReportVision").Value = "" Then mainWorkbook.Worksheets("About").Range("ReportVision").Value = report("Vision")
    Else
        If mainWorkbook.Worksheets("About").Range("ReportVision").Value = "" Then mainWorkbook.Worksheets("About").Range("ReportVision").Value = "N/A."
    End If

    If report.Exists("Dependencies") Then
        If mainWorkbook.Worksheets("About").Range("DependenciesList").Value = "" Then mainWorkbook.Worksheets("About").Range("DependenciesList").Value = "Dependencies List: " & report("Dependencies") & "."
    Else
        If mainWorkbook.Worksheets("About").Range("DependenciesList").Value = "" Then mainWorkbook.Worksheets("About").Range("DependenciesList").Value = "Dependencies List: N/A."
    End If

    If report.Exists("Creation Date") Then
        If mainWorkbook.Worksheets("About").Range("CreationDate").Value = "" Then mainWorkbook.Worksheets("About").Range("CreationDate").Value = "Creation Date: " & report("Creation Date") & "."
    End If

    If report.Exists("Edition") Then
        If mainWorkbook.Worksheets("About").Range("EditionName").Value = "" Then mainWorkbook.Worksheets("About").Range("EditionName").Value = "Edition Name: " & report("Edition") & "."
    Else
        If mainWorkbook.Worksheets("About").Range("EditionName").Value = "" Then mainWorkbook.Worksheets("About").Range("EditionName").Value = "Edition Name: N/A."
    End If

    If mainWorkbook.Worksheets("About").Range("DurationMilliseconds").Value = "" Then mainWorkbook.Worksheets("About").Range("DurationMilliseconds").Value = "Duration (Milliseconds): N/A."
    If mainWorkbook.Worksheets("About").Range("LogSummary").Value = "" Then mainWorkbook.Worksheets("About").Range("LogSummary").Value = "Log Summary: N/A."

    If legacyTemplateVersion = True Then
        Dim legacyRetrievedDate As String

        mainWorkbook.Worksheets("About").Range("CreationDate").Value = Replace(mainWorkbook.Worksheets("About").Range("CreationDate").Value, "Retrieved", "Creation")
        legacyRetrievedDate = GetAboutNamedRange("Creation Date")
        Call SetAboutNamedRange(Right(legacyRetrievedDate, 4) & "-" & Mid(legacyRetrievedDate, 4, 2) & "-" & Left(legacyRetrievedDate, 2), "Creation Date")

        mainWorkbook.Worksheets("About").Range("DurationMilliseconds").Value = "Duration (Milliseconds): N/A."
        mainWorkbook.Worksheets("About").Range("LogSummary").Value = "Log Summary: N/A."
    End If

    If mainWorkbook.Worksheets("About").Range("B2").MergeArea.Cells.Count = 1 Then
        mainWorkbook.Worksheets("About").Range("B1:B2").Merge
    End If

    If mainWorkbook.Worksheets("About").Range("B4").MergeArea.Cells.Count = 1 Then
        mainWorkbook.Worksheets("About").Range("B3:B4").Merge
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

    Dim index As Byte
    For index = 0 To UBound(intermissionsArray)
        intermissionState = intermissionsArray(index)

        Select Case intermissionState
            Case "Break Script", "Break", "BS", "B"
                End
            Case "Covert Mode", "Covert", "CM", "C"
                Dim aboutNamedRanges As String: aboutNamedRanges = "AugmentationModules|CreationDate|DependenciesList|DurationMilliseconds|EditionName|LogSummary|ProgressionStatus|ReportDetails|ReportName|ReportVision|RetrievedDate|ScriptDuration|TemplateDetails|TemplateVersion"
                Dim aboutNamedRangesArray() As String: aboutNamedRangesArray = Split(aboutNamedRanges, "|")
                Dim aboutNamedRange As String

                Dim aboutIndex As Byte
                For aboutIndex = 0 To UBound(aboutNamedRangesArray)
                    aboutNamedRange = aboutNamedRangesArray(aboutIndex)

                    If NamedRangeExists(aboutNamedRange) Then
                        mainWorkbook.Names(aboutNamedRange).Delete
                    End If
                Next aboutIndex

                If WorksheetExists("Log") Then
                    mainWorkbook.Worksheets("Log").Delete
                End If
            Case "Duplicate Workbook", "Duplicate", "DW", "D"
                mainWorkbook.SaveCopyAs Left(mainWorkbook.FullName, Len(mainWorkbook.FullName) - 5) & " (" & checkpointName & ")" & ".xlsx"
            Case "End Workbook", "End", "EW", "E"
                mainWorkbook.Close
            Case "Initiate Workbook", "Initiate", "IW", "I"
                Set mainWorkbook = Workbooks.Add
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
                If NamedRangeExists("Progression Status") Then
                    Call SetAboutNamedRange("N/A", "Progression Status", True)
                End If

                If NamedRangeExists("Augmentation Modules") Then
                    Call SetAboutNamedRange("N/A", "Augmentation Modules", True)
                End If
        End Select
    Next index
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

    Select Case aboutNamedRange
        Case "ReportDetails", "Report Details", "ReportName", "ReportName", "Report", "Name"
            If NamedRangeExists("ReportDetails") Then
                mainWorkbook.Worksheets("About").Range("ReportDetails").Value = namedRangeValue
            ElseIf NamedRangeExists("ReportName") Then
                mainWorkbook.Worksheets("About").Range("ReportName").Value = namedRangeValue
            End If
        Case "TemplateDetails", "Template Details", "TemplateVersion", "Template Version", "Template", "Version"
            If NamedRangeExists("TemplateDetails") Then
                mainWorkbook.Worksheets("About").Range("TemplateDetails").Value = "Spreadsheet Operations Template (" & namedRangeValue & ")"
            ElseIf NamedRangeExists("TemplateVersion") Then
                mainWorkbook.Worksheets("About").Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & namedRangeValue & ")"
            End If
        Case "ProgressionStatus", "Progression Status", "Progression"
            currentNamedRangeValue = GetAboutNamedRange("Progression Status")

            If overwrite = True Or currentNamedRangeValue = "N/A" Then
                mainWorkbook.Worksheets("About").Range("ProgressionStatus").Value = "Progression Status: " & namedRangeValue & "."
            Else
                mainWorkbook.Worksheets("About").Range("ProgressionStatus").Value = "Progression Status: " & currentNamedRangeValue & ", " & namedRangeValue & "."
            End If
        Case "AugmentationModules", "Augmentation Modules", "Augmentation"
            currentNamedRangeValue = GetAboutNamedRange("Augmentation Modules")

            If overwrite = True Or currentNamedRangeValue = "N/A" Then
                mainWorkbook.Worksheets("About").Range("AugmentationModules").Value = "Augmentation Modules: " & namedRangeValue & "."
            Else
                mainWorkbook.Worksheets("About").Range("AugmentationModules").Value = "Augmentation Modules: " & currentNamedRangeValue & ", " & namedRangeValue & "."
            End If
        Case "ReportVision", "Report Vision", "Vision"
            mainWorkbook.Worksheets("About").Range("ReportVision").Value = namedRangeValue
        Case "DependenciesList", "Dependencies List", "Dependencies"
            currentNamedRangeValue = GetAboutNamedRange("Dependencies List")

            If overwrite = True Or currentNamedRangeValue = "N/A" Then
                mainWorkbook.Worksheets("About").Range("DependenciesList").Value = "Dependencies List: " & namedRangeValue & "."
            Else
                mainWorkbook.Worksheets("About").Range("DependenciesList").Value = "Dependencies List: " & currentNamedRangeValue & ", " & namedRangeValue & "."
            End If
        Case "CreationDate", "Creation Date", "Creation", "CD", "RetrievedDate", "Retrieved Date", "Retrieved"
            mainWorkbook.Worksheets("About").Range("CreationDate").Value = "Creation Date: " & namedRangeValue & "."
        Case "EditionName", "Edition Name", "Edition"
            mainWorkbook.Worksheets("About").Range("EditionName").Value = "Edition Name: " & namedRangeValue & "."
        Case "ScriptDuration", "Script Duration", "Duration (Milliseconds)", "Duration Milliseconds", "Duration"
            If NamedRangeExists("ScriptDuration") Then
                currentNamedRangeValue = GetAboutNamedRange("Script Duration")
            ElseIf NamedRangeExists("DurationMilliseconds") Then
                currentNamedRangeValue = GetAboutNamedRange("Duration (Milliseconds)")
            End If

            If overwrite = True Or currentNamedRangeValue = "N/A" Then
                mainWorkbook.Worksheets("About").Range("DurationMilliseconds").Value = "Duration (Milliseconds): " & namedRangeValue & "."
            Else
                mainWorkbook.Worksheets("About").Range("DurationMilliseconds").Value = "Duration (Milliseconds): " & (CDbl(currentNamedRangeValue) + CDbl(namedRangeValue)) & "."
            End If
        Case "LogSummary", "Log Summary", "Summary"
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

            mainWorkbook.Worksheets("About").Range("LogSummary").Value = "Log Summary: " & runsSummary & checkpointsSummary & logRowsSummary
        Case Else
            ' Call LogAssertFailure("Core named range value of  """ & coreNamedRange & """" & " is invalid. Valid values: Dependencies List, Report Details, Report Vision, Edition Name, Retrieved Date.")
    End Select
End Sub

' Functions: Background '

Function CheckpointIsNew(ByVal checkpointName As String) As Boolean
    If InputContainsValue(mainWorkbook.Names("ProgressionStatus").RefersToRange.Value, checkpointName) Or InputContainsValue(mainWorkbook.Names("AugmentationModules").RefersToRange.Value, checkpointName) Then
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
                namedRangeValue = mainWorkbook.Worksheets("About").Range("ReportDetails").Value
            ElseIf NamedRangeExists("ReportName") Then
                namedRangeValue = mainWorkbook.Worksheets("About").Range("ReportName").Value
            End If
        Case "TemplateDetails", "Template Details", "TemplateVersion", "Template Version", "Template", "Version"
            If NamedRangeExists("TemplateDetails") Then
                namedRangeValue = mainWorkbook.Worksheets("About").Range("TemplateDetails").Value
            ElseIf NamedRangeExists("TemplateVersion") Then
                namedRangeValue = mainWorkbook.Worksheets("About").Range("TemplateVersion").Value
            End If

            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, "(") + 1)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "ProgressionStatus", "Progression Status", "Progression"
            namedRangeValue = mainWorkbook.Worksheets("About").Range("ProgressionStatus").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "AugmentationModules", "Augmentation Modules", "Augmentation"
            namedRangeValue = mainWorkbook.Worksheets("About").Range("AugmentationModules").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "ReportVision", "Report Vision", "Vision"
            namedRangeValue = mainWorkbook.Worksheets("About").Range("ReportVision").Value
        Case "DependenciesList", "Dependencies List", "Dependencies"
            namedRangeValue = mainWorkbook.Worksheets("About").Range("DependenciesList").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "CreationDate", "Creation Date", "Creation", "CD", "RetrievedDate", "Retrieved Date", "Retrieved"
            namedRangeValue = mainWorkbook.Worksheets("About").Range("CreationDate").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "EditionName", "Edition Name", "Edition"
            namedRangeValue = mainWorkbook.Worksheets("About").Range("EditionName").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "ScriptDuration", "Script Duration", "Duration (Milliseconds)", "Duration Milliseconds", "Duration"
            namedRangeValue = mainWorkbook.Worksheets("About").Range("DurationMilliseconds").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "LogSummary", "Log Summary", "Summary"
            namedRangeValue = mainWorkbook.Worksheets("About").Range("LogSummary").Value
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("", methodName, isRegistered, "Conjuration")
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

Sub CreateBlankWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatCreateBlankWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CreateBlankWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    If worksheetName = "" Then worksheetName = ActiveSheet.Name & "!"
    If WorksheetExists(worksheetName) Then Call LogConclusion("Failed", logConclusionData, "Worksheet """ & worksheetName & """ already exists.")

    mainWorkbook.Worksheets.Add.Name = worksheetName
    
    Call NormalizeLayoutOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub CreateFileListArrayFromDirectory(ByRef fileListArray As Variant, ByVal folderPath As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CreateFileListArrayFromDirectory"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByRef fileListArray As Variant, ByVal folderPath As String", methodName, isRegistered, "Conjuration")
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & fromWorksheetName & """" & ", " & """" & toWorksheetName & """")

    ' Remember to sort the two primary columns prior to using this method, or there will be errors unless they are sorted from before. '

    Dim cellStyleToCopy As String
    Dim widthToCopy As Double
    Dim originalColumnName As String
    originalColumnName = columnName

    Dim validation As String
    Call ValidateWorksheet(toWorksheetName, "toWorksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, "columnName", fromWorksheetName, "fromWorksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim sourceColumnName As String
    Dim matchColumnName As String

    sourceColumnName = mainWorkbook.Worksheets(toWorksheetName).Range("A1").Value
    matchColumnName = mainWorkbook.Worksheets(fromWorksheetName).Range("A1").Value

    Dim deleteDuplicateColumn As Boolean
    deleteDuplicateColumn = False

    If sourceColumnName = "" Then
        Call InsertNewColumnAndSetWidthOnWorksheet("Sorting Column", 7, fromWorksheetName)
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

    If ColumnOnWorksheetExists(columnName, toWorksheetName) = False Then Call InsertNewColumnAndSetWidthOnWorksheet(originalColumnName, widthToCopy, toWorksheetName)

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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal toWorksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & toWorksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(currentWorksheetName, "currentWorksheetName", validation)
    Call ValidateWorksheet(toWorksheetName, "toWorksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal newWorksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & newWorksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(currentWorksheetName, "currentWorksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If currentWorksheetName = newWorksheetName Then newWorksheetName = currentWorksheetName & "!"

    mainWorkbook.Worksheets(currentWorksheetName).Copy After:=Sheets(Worksheets.Count)
    ActiveSheet.Name = newWorksheetName

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ExportWorksheetToBlankWorkbook(ByVal worksheetName As String, ByVal filePath As String) ' Plural Support. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ExportWorksheetToBlankWorkbook"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String, ByVal filePath As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """" & ", " & """" & filePath & """")

    Workbooks.Add
    Set importWorkbook = ActiveWorkbook
    importWorkbook.SaveAs filePath

    Dim filterValueArray() As String: filterValueArray = Split(worksheetName, "|")
    Dim index As Byte

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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetsValues As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & worksheetsValues & """")

    Dim worksheetsArray() As String

    Call OpenGuestWorkbook(filePath)

    If worksheetsValues = "" Then
        Dim excelFileWorksheetsCollection As New Collection

        Dim indexWorksheet As Byte
        For indexWorksheet = 1 To importWorkbook.Worksheets.Count
            If importWorkbook.Worksheets(indexWorksheet).Name <> "About" And importWorkbook.Worksheets(indexWorksheet).Name <> "Log" Then worksheetsValues = worksheetsValues & importWorkbook.Worksheets(indexWorksheet).Name & "|"
        Next indexWorksheet

        worksheetsValues = Left(worksheetsValues, Len(worksheetsValues) - 1)
    End If

    worksheetsArray = Split(worksheetsValues, "|")

    Dim index As Byte
    For index = 0 To UBound(worksheetsArray)
        Dim excelFileWorksheetExists As Boolean

        importWorkbook.Activate
        excelFileWorksheetExists = WorksheetExists(worksheetsArray(index))
            
        If excelFileWorksheetExists = True Then
            mainWorkbook.Activate
            Call CreateBlankWorksheet(worksheetsArray(index))
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & worksheetName & """")

    Call OpenGuestWorkbook(filePath)

    If importWorkbook.Worksheets.Count > 1 Then Call LogConclusion("Failed", logConclusionData, "This method only supports importing one worksheet per workbook, the workbook " & """" & importWorkbook.Name & """" & " contains " & importWorkbook.Worksheets.Count & " worksheets.")

    importWorkbook.Sheets(1).Name = worksheetName
    mainWorkbook.Activate
    If WorksheetExists(worksheetName) Then Call LogConclusion("Failed", logConclusionData, "Worksheet named " & """" & worksheetName & """" & " already exists.")
    Call CreateBlankWorksheet(worksheetName)

    importWorkbook.Sheets(worksheetName).Cells.Copy mainWorkbook.Worksheets(worksheetName).Cells
    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ImportTextFileToWorksheet(ByVal filePath As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ImportTextFileToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & worksheetName & """")

    Call OpenGuestWorkbook(filePath)
    mainWorkbook.Activate
    Call CreateBlankWorksheet(worksheetName)
    importWorkbook.Sheets(1).Cells.Copy mainWorkbook.Worksheets(worksheetName).Cells
    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub OpenGuestWorkbook(ByVal filePath As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "OpenGuestWorkbook"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String", methodName, isRegistered, "Conjuration")
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetName As String, toWorksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & fromWorksheetName & """" & " ," & """" & toWorksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateWorksheet(toWorksheetName, "toWorksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyleName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cellStyleName & """")

    If CellStyleExists(cellStyleName) Then
        mainWorkbook.Styles(cellStyleName).Delete
    Else
        Call LogConclusion("Skipped", logConclusionData)
        return
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DeleteColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteColumnOnWorksheet(columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Columns(columnName & ":" & columnName).EntireColumn.Delete Shift:=xlToLeft

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DeleteRowsOnWorksheet(ByVal numberOfRows As Long, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteRowsOnWorksheet(numberOfRows, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteRowsOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal numberOfRows As Long, ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, numberOfRows & " ," & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Rows("1:" & numberOfRows).Delete Shift:=xlUp

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DeleteWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "DeleteWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Delete

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RemoveBlankLinesOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatRemoveBlankLinesOnWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RemoveBlankLinesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    If WorksheetExists(worksheetName) Then
        Dim lastRowNumberForWorksheet As String: lastRowNumberForWorksheet = LastUsedRowNumberOnWorksheet(worksheetName)

        mainWorkbook.Worksheets(worksheetName).Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        mainWorkbook.Worksheets(worksheetName).Range("A1").Value = "Remove Blank Lines Header( Temporary)"
        mainWorkbook.Worksheets(worksheetName).Range("A2" & ":" & "A" & lastRowNumberForWorksheet).Value = ";;"
        Call RemoveDataBasedOnFormulaOnWorksheet("=OR(B2="""";CHAR(32)=B2)", worksheetName) ' Code 32 (decimal) is a nonprinting spacing character. (https://web.archive.org/web/20221004120012/http://www.columbia.edu/kermit/ascii.html) '
        mainWorkbook.Worksheets(worksheetName).Columns("A:A").Delete Shift:=xlToLeft
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim emptyColumnArray() As String: emptyColumnArray = Split(CreateDelimitedHeaderStringExcludingOnWorksheet("", worksheetName), "|")
    Dim index As Byte

    For index = 0 To UBound(emptyColumnArray)
        If ColumnIsEmptyOnWorksheet(FindColumnLetterOnWorksheet(emptyColumnArray(index), worksheetName), worksheetName) Then Call DeleteColumnOnWorksheet(emptyColumnArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RemoveDataBasedOnFormulaOnWorksheet(ByVal formulaValue As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RemoveDataBasedOnFormulaOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & formulaValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call ApplyFormulaToColumnOnWorksheet(formulaValue, report("Helper Column"), worksheetName)
    Call ApplyFilterToColumnOnWorksheet(True, report("Helper Column"), worksheetName)

    If NumberOfVisibleCells(worksheetName) <> 0 Then
        Call ClearFilterOnWorksheet(worksheetName)
        Call InsertNewColumnAndSetWidthOnWorksheet("Original Order (Temporary)", 13, worksheetName)
        Call ApplyFormulaToColumnOnWorksheet("=ROW(2:2)-1", "Original Order (Temporary)", worksheetName)
        Call ApplyCellStyleToColumnOnWorksheet("Integer", "Original Order (Temporary)", worksheetName)
        Call ApplyFilterToColumnOnWorksheet(True, report("Helper Column"), worksheetName)
    End If

    If NumberOfVisibleCells(worksheetName) <> 0 Then
        If mainWorkbook.Worksheets(worksheetName).FilterMode Then
            ' This approach is used due to the instability of excel when dealing with large datasets. Just deleting the data can crash the application. '

            Call SelectWorksheet(worksheetName)

            ActiveWindow.ActivateNext
            ActiveWindow.WindowState = xlMaximized

            Dim savedFreezePanesMode As String
            savedFreezePanesMode = ActiveWindow.SplitColumn & ", " & ActiveWindow.SplitRow

            mainWorkbook.Worksheets(worksheetName).Rows("1:1").EntireRow.Hidden = True
            mainWorkbook.Worksheets(worksheetName).Range(ColumnRangeTypeOnWorksheet("A", "Data", worksheetName)).SpecialCells(xlCellTypeConstants, 23).ClearContents ' Header hidden before or bug occurs when only one line gets deleted. '
            mainWorkbook.Worksheets(worksheetName).Rows("1:1").EntireRow.Hidden = False
            Call ClearFilterOnWorksheet(worksheetName)
            Call SortColumnByOrderOnWorksheet("A", "Ascending", worksheetName)
            mainWorkbook.Worksheets(worksheetName).Name = "!"
            Call CreateBlankWorksheet(worksheetName)
            mainWorkbook.Worksheets("!").Range("A1:" & LastUsedColumnLetterOnWorksheet("!") & LastUsedRowNumberOnWorksheet("!")).Copy mainWorkbook.Worksheets(worksheetName).Range("A1")
            mainWorkbook.Worksheets("!").Range("A1:" & LastUsedColumnLetterOnWorksheet("!") & "1").Copy
            mainWorkbook.Worksheets(worksheetName).Range("A1").PasteSpecial xlPasteColumnWidths
            Application.CutCopyMode = False
            Call ApplyAutoFilterOnWorksheet(worksheetName)
            Call SortColumnByOrderOnWorksheet("Original Order (Temporary)", "Ascending", worksheetName)
            Call DeleteColumnOnWorksheet("Original Order (Temporary)", worksheetName)

            If savedFreezePanesMode <> "0, 0" Then Call ApplyFreezePanesModeOnWorksheet(savedFreezePanesMode, worksheetName)

            If mainWorkbook.Worksheets("!").Tab.Color <> False Then
                mainWorkbook.Worksheets(worksheetName).Tab.Color = mainWorkbook.Worksheets("!").Tab.Color
            End If

            Call DeleteWorksheet("!")
        Else
            ' Call LogTaskWarning("Worksheet " & """" & worksheetName & """" & " does not have an active filter.", logOrderLocal)
        End If
    End If

    If NumberOfVisibleCells(worksheetName) = 0 Then Call ClearFilterOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SynchronizeWorksheetOnColumnFromWorksheet(ByVal synchronizeWorksheetName As String, ByVal columnName As String, ByVal fromWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SynchronizeWorksheetOnColumnFromWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal synchronizeWorksheetName As String, ByVal columnName As String, ByVal fromWorksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & synchronizeWorksheetName & """" & ", " & """" & columnName & """" & ", " & """" & fromWorksheetName & """")

    Dim synchronizeWorksheetColumnName As String
    Dim fromWorksheetColumnName As String

    synchronizeWorksheetColumnName = columnName
    fromWorksheetColumnName = columnName

    Dim validation As String
    Call ValidateColumnOnWorksheet(synchronizeWorksheetColumnName, "synchronizeWorksheetColumnName", synchronizeWorksheetName, "synchronizeWorksheetName", validation)
    Call ValidateColumnOnWorksheet(fromWorksheetColumnName, "fromWorksheetColumnName", fromWorksheetName, "fromWorksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim synchronizeWorksheetColumnStart As String
    Dim fromWorksheetColumnData As String

    synchronizeWorksheetColumnStart = ColumnStartOnWorksheet(synchronizeWorksheetColumnName, synchronizeWorksheetName)
    fromWorksheetColumnData = ColumnDataOnWorksheet(fromWorksheetColumnName, fromWorksheetName)

    Call RemoveDataBasedOnFormulaOnWorksheet("=ISNA(XLOOKUP(" & synchronizeWorksheetColumnStart & ";" & fromWorksheetColumnData & ";" & fromWorksheetColumnData & "))", synchronizeWorksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Functions: Destruction '

' ************ '
' Elementals   '
' ************ '

Sub ApplyFilterToColumnOnWorksheet(ByVal filterValue As String, ByVal columnName As String, ByVal worksheetName As String) ' Plural Support. '
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFilterToColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filterValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filterValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

If InputContainsValue(filterValue, "|") Then
    Dim filterValueArray() As String: filterValueArray = Split(filterValue, "|")

    mainWorkbook.Worksheets(worksheetName).Range(columnName & "1:" & columnName & "1").AutoFilter Field:=ConvertLetterToNumber(columnName), Criteria1:=filterValueArray, Operator:=xlFilterValues
Else
    mainWorkbook.Worksheets(worksheetName).Range(columnName & "1:" & columnName & "1").AutoFilter Field:=ConvertLetterToNumber(columnName), Criteria1:=filterValue, Operator:=xlAnd
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ClearFilterOnWorksheet(ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ClearFilterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).AutoFilter.Sort.SortFields.Clear
    If mainWorkbook.Worksheets(worksheetName).AutoFilterMode Then mainWorkbook.Worksheets(worksheetName).AutoFilter.ShowAllData

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub CutColumnAndPasteAtColumnOnWorksheet(ByVal cutColumnName As String, ByVal pasteColumnName As String, ByVal worksheetName As String) ' Repeat Support: cutColumnName. '

If InputContainsValue(cutColumnName, "|") Then
    Call RepeatCutColumnAndPasteAtColumnOnWorksheet(cutColumnName, pasteColumnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "CutColumnAndPasteAtColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cutColumnName As String, ByVal pasteColumnName As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cutColumnName & """" & ", " & """" & pasteColumnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(cutColumnName, "cutColumnName", worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(pasteColumnName, "pasteColumnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).Columns(cutColumnName & ":" & cutColumnName).Cut
    mainWorkbook.Worksheets(worksheetName).Columns(pasteColumnName & ":" & pasteColumnName).Insert Shift:=xlToRight

    Call ResetAutoFilterOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub SelectWorksheet(ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SelectWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If mainWorkbook.ActiveSheet.Name = worksheetName Then
        Call LogConclusion("Skipped", logConclusionData)
    Else
        mainWorkbook.Worksheets(worksheetName).Select
        Call LogConclusion("Completed", logConclusionData)
    End If
End Sub

Sub SortColumnByOrderOnWorksheet(ByVal columnName As String, ByVal sortOrder As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatSortColumnByOrderOnWorksheet(columnName, sortOrder, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "SortColumnByOrderOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal sortOrder As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & sortOrder & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If sortOrder <> "Ascending" And sortOrder <> "Descending" Then
        sortOrder = "Ascending"
        ' Call LogTaskWarning("Sorting order " & """" & sortOrder & """" & " not Ascending or Descending, setting to Ascending.", logOrderLocal)
    End If
      
    Call SelectWorksheet(worksheetName) ' Prevents selection bug. '
    If Not mainWorkbook.Worksheets(worksheetName).AutoFilterMode Then Call LogConclusion("Failed", logConclusionData, "Worksheet """ & worksheetName & """ does not have a filter applied.")

    If sortOrder ="Ascending" Then mainWorkbook.Worksheets(worksheetName).AutoFilter.Sort.SortFields.Add Key:=Range(columnName & "1:" & columnName & "1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    If sortOrder ="Descending" Then mainWorkbook.Worksheets(worksheetName).AutoFilter.Sort.SortFields.Add Key:=Range(columnName & "1:" & columnName & "1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    With mainWorkbook.Worksheets(worksheetName).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    mainWorkbook.Worksheets(worksheetName).AutoFilter.Sort.SortFields.Clear

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ManualCodeSection(ByVal sectionName As String, ByVal startOrFinish As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ManualCodeSection"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal sectionName As String, ByVal startOrFinish As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & sectionName & """" & ", " & """" & startOrFinish & """")

    If startOrFinish <> "Start" And startOrFinish <> "Finish" Then Call LogConclusion("Failed", logConclusionData, "Only valid values are ""Start"" and ""Finish"", as opposed to input value of: """ & startOrFinish & """.")

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub TextToColumnsModeOnWorksheet(ByVal modeValue As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "TextToColumnsModeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal modeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & modeValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If modeValue = "" Then modeValue = "Semicolon"

    If modeValue = "Comma" Then mainWorkbook.Worksheets(worksheetName).Range("A1:A" & LastUsedRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Comma:=True
    If modeValue = "Semicolon" Then mainWorkbook.Worksheets(worksheetName).Range("A1:A" & LastUsedRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Semicolon:=True
    If modeValue = "Tab" Then mainWorkbook.Worksheets(worksheetName).Range("A1:A" & LastUsedRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Tab:=True

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Functions: Elementals '

Function ConvertDateTimeToSerial(ByVal columnLetter As String, ByVal startingRow As Long) As String
    ConvertDateTimeToSerial = "=DATEVALUE(" & columnLetter & startingRow & ") + RIGHT(LEFT(" & columnLetter & startingRow & ";13);2)/24 + LEFT(RIGHT(" & columnLetter & startingRow & ";5);2)/1440 + RIGHT(" & columnLetter & startingRow & ";2)/86400"
End Function

Function ConvertLetterToNumber(ByVal letterToConvert As String) As Long
    On Error Resume Next
    ConvertLetterToNumber = mainWorkbook.Worksheets("Log").Range(letterToConvert & 1).Column
    If Err.Number <> 0 Then
        Err.Clear
        Call LogFunctionError("ConvertLetterToNumber(letterToConvert As String) As Long", "Letter " & """" & letterToConvert & """" & " is too high for number conversion.")
    End If
End Function

Function ConvertNumberToLetter(ByVal numberToConvert As Long) As String
    If numberToConvert >= 16385 Then Call LogFunctionError("ConvertNumberToLetter(numberToConvert As Long) As String", "Number " & """" & numberToConvert & """" & " is too high for letter conversion.")
    ConvertNumberToLetter = Split(Cells(1, numberToConvert).Address, "$")(1)
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
    Dim indexHeaders As Byte
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
    Dim indexWorksheets As Byte
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
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "FindColumnLetterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then
        Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")
        Call LogConclusion("Failed", logConclusionData, validation)
    End If

    Dim targetWorksheet As Worksheet
    Dim headerSearchRange As Range
    Dim foundHeaderCell As Range

    Set targetWorksheet = ActiveWorkbook.Worksheets(worksheetName)
    Set headerSearchRange = targetWorksheet.Rows(1)
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

        Dim index As Byte
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

        Dim index As Byte
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
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "LastUsedColumnLetterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then
        Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")
        Call LogConclusion("Failed", logConclusionData, validation)
    End If

    Dim targetWorksheet As Worksheet
    Dim lastUsedColumnNumber As Long
    Dim workingColumnNumber As Long
    Dim resultingColumnLetter As String

    Set targetWorksheet = mainWorkbook.Worksheets(worksheetName)

    lastUsedColumnNumber = targetWorksheet.Cells(1, targetWorksheet.Columns.Count).End(xlToLeft).Column
    workingColumnNumber  = lastUsedColumnNumber

    Do While workingColumnNumber > 0
        workingColumnNumber = workingColumnNumber - 1
        resultingColumnLetter = Chr$(65 + (workingColumnNumber Mod 26)) & resultingColumnLetter
        workingColumnNumber = workingColumnNumber \ 26
    Loop

    LastUsedColumnLetterOnWorksheet = resultingColumnLetter
End Function

Function LastUsedRowNumberOnWorksheet(ByVal worksheetName As String) As Long
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "LastUsedRowNumberOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then
        Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")
        Call LogConclusion("Failed", logConclusionData, validation)
    End If

    Dim targetWorksheet As Worksheet
    Dim lastUsedRowNumber As Long

    Set targetWorksheet = mainWorkbook.Worksheets(worksheetName)

    lastUsedRowNumber = targetWorksheet.Cells(targetWorksheet.Rows.Count, 1).End(xlUp).Row

    If lastUsedRowNumber = 1 And IsEmpty(targetWorksheet.Cells(1, 1).Value) Then
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
        If ColumnLetterValid(columnName) = False Then Call LogFunctionError("ColumnRangeTypeOnWorksheet(ByVal columnName As String, ByVal rangeType As String, ByVal worksheetName As String) As String", "Column letter " & """" & columnName & """" & " on worksheet " & """" & worksheetName & """" & " not valid.")
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If Not mainWorkbook.Worksheets(worksheetName).AutoFilterMode Then
        mainWorkbook.Worksheets(worksheetName).Range("A1:" & LastUsedColumnLetterOnWorksheet(worksheetName) & "1").AutoFilter
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyle As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cellStyle & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

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
    With mainWorkbook.Worksheets(worksheetName).Range(columnLetter & "2:" & columnLetter & "1048576")
        .Style = cellStyle
        .Value = .Value
    End With

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ApplyFreezePanesModeOnWorksheet(ByVal freezePanesMode As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatApplyFreezePanesModeOnWorksheet(freezePanesMode, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyFreezePanesModeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal freezePanesMode As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & freezePanesMode & """" & ", " & """" & worksheetName & """")

    Call SelectWorksheet(worksheetName)

    ActiveWindow.ActivateNext
    ActiveWindow.WindowState = xlMaximized

    Application.WorksheetFunction.Trim(freezePanesMode)
    If InputContainsValue(freezePanesMode, ",") = False Then freezePanesMode = "0,1"

    Dim numericalModeArray() As String: numericalModeArray = Split(freezePanesMode, ",")

    If IsNumeric(numericalModeArray(0)) = False Or IsNumeric(numericalModeArray(1)) = False Then
        freezePanesMode = "0,1"
        numericalModeArray = Split(freezePanesMode, ",")
    End If

    With ActiveWindow
        .SplitColumn = CInt(numericalModeArray(0))
        .SplitRow = CInt(numericalModeArray(1))
    End With
    ActiveWindow.FreezePanes = True

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ConvertColumnToDateFormattingOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName '

If InputContainsValue(columnName, "|") Then
    Call RepeatConvertColumnToDateFormattingOnWorksheet(columnName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ConvertColumnToDateFormattingOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call EnsureHelperColumnOnWorksheet(worksheetName)

    Call ApplyFormulaToColumnOnWorksheet("=IF(" & ColumnStartOnWorksheet(columnName, worksheetName) & "<>"""";INT(" & ColumnStartOnWorksheet(columnName, worksheetName) & ");"""")", report("Helper Column"), worksheetName)
    Call ApplyFormulaToColumnOnWorksheet("=IF(" & ColumnStartOnWorksheet(report("Helper Column"), worksheetName) & "<>"""";" & ColumnStartOnWorksheet(report("Helper Column"), worksheetName) & ";"""")", columnName, worksheetName)
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call EnsureHelperColumnOnWorksheet(worksheetName)

    Dim columnData As String
    Dim helperData As String
    columnData = ColumnRangeTypeOnWorksheet(columnName, "Start", worksheetName)
    helperData = ColumnRangeTypeOnWorksheet(report("Helper Column"), "Start", worksheetName)

    Call ApplyFormulaToColumnOnWorksheet("=IF(NOT(ISBLANK(" & columnData & "));IF(NOT(ISERROR(VALUE(" & columnData & ")));VALUE(" & columnData & ");"""");"""")", report("Helper Column"), worksheetName)
    Call ApplyFormulaToColumnOnWorksheet("=IF(" & helperData & "<>"""";" & helperData & ";"""")", columnName, worksheetName)
    
    Call EnsureHelperColumnDeletedOnWorksheet(worksheetName)

    Call ApplyCellStyleToColumnOnWorksheet("Decimal", columnName, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ColorColumnOfTypeAndElementOnWorksheet(ByVal colorName As String, ByVal columnName As String, ByVal columnType As String, ByVal elementType As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorColumnOfTypeAndElementOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal columnName As String, ByVal columnType As String, ByVal elementType As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & colorName & """" & ", " & """" & columnName & """" & ", " & """" & columnType & """" & ", " & """" & elementType & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim colorRange As String

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

    If elementType = "Background" Then mainWorkbook.Worksheets(worksheetName).Range(colorRange).Interior.Color = colorValue
    If elementType = "Border" Then
        With mainWorkbook.Worksheets(worksheetName).Range(colorRange).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With mainWorkbook.Worksheets(worksheetName).Range(colorRange).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With mainWorkbook.Worksheets(worksheetName).Range(colorRange).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With mainWorkbook.Worksheets(worksheetName).Range(colorRange).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
    End If
    If elementType = "Font" Then mainWorkbook.Worksheets(worksheetName).Range(colorRange).Font.Color = colorValue

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ColorRangeBackgroundAndFontAndBorderOnWorksheet(ByVal rangeToColor As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal borderColorName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorRangeBackgroundAndFontAndBorderOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal rangeToColor As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal borderColorName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & rangeToColor & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & borderColorName & """" & ", " & """" & worksheetName & """")

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

    mainWorkbook.Worksheets(worksheetName).Range(rangeToColor).Interior.Color = backgroundColorValue
    mainWorkbook.Worksheets(worksheetName).Range(rangeToColor).Font.Color = fontColorValue

    With mainWorkbook.Worksheets(worksheetName).Range(rangeToColor).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With mainWorkbook.Worksheets(worksheetName).Range(rangeToColor).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With mainWorkbook.Worksheets(worksheetName).Range(rangeToColor).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With mainWorkbook.Worksheets(worksheetName).Range(rangeToColor).Borders(xlEdgeRight)
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & colorName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim colorValue As Long: colorValue = 2

    colorValue = StyleColor(colorName)
    If colorValue = 2 Then Call LogConclusion("Failed", logConclusionData, "Color name " & """" & colorName & """" & " is not a valid predefined color name.")

    mainWorkbook.Worksheets(worksheetName).Tab.Color = colorValue

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertImageOnWorksheetWithLeftAndTopValues(ByVal imageFilePath As String, ByVal worksheetName As String, ByVal leftValue As Double, ByVal topValue As Double)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "InsertImageOnWorksheetWithLeftAndTopValues"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal imageFilePath As String, ByVal worksheetName As String, ByVal leftValue As Double, ByVal topValue As Double", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & imageFilePath & """" & ", " & """" & worksheetName & """" & ", " & leftValue & ", " & topValue)

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call SelectWorksheet(worksheetName)
    mainWorkbook.Worksheets(worksheetName).Pictures.Insert(imageFilePath).Select
    Selection.ShapeRange.IncrementLeft leftValue
    Selection.ShapeRange.IncrementTop topValue

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub NormalizeLayoutOnWorksheet(ByVal worksheetName As String, Optional ByVal applyAutoFilter As Boolean = True, Optional ByVal freezeTopRow As Boolean = True, Optional ByVal headerStyle As Boolean = True) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatNormalizeLayoutOnWorksheet(worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "NormalizeLayoutOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String, Optional ByVal applyAutoFilter As Boolean = True, Optional ByVal freezeTopRow As Boolean = True, Optional ByVal headerStyle As Boolean = True", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """" & ", " & applyAutoFilter & ", " & freezeTopRow & ", " & headerStyle)

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    With mainWorkbook.Worksheets(worksheetName)
        .Cells.RowHeight = 16.5
        .Rows("1:1").RowHeight = 49.5
    End With

    If applyAutoFilter = True Then
        If mainWorkbook.Worksheets(worksheetName).AutoFilterMode = False Then
            mainWorkbook.Worksheets(worksheetName).Range("A1:" & LastUsedColumnLetterOnWorksheet(worksheetName) & "1").AutoFilter
        End If
    End If

    If freezeTopRow = True Then
        mainWorkbook.Worksheets(worksheetName).Activate
        With ActiveWindow
            If .FreezePanes = False Then
                .SplitColumn = 0
                .SplitRow = 1
                .FreezePanes = True
            End If
        End With
    End If

    If headerStyle = True Then
        With mainWorkbook.Worksheets(worksheetName).Range("A1:" & LastUsedColumnLetterOnWorksheet(worksheetName) & "1")
            .Style = "Header"
            .Value = .Value
        End With
    End If

    Call LogConclusion("Completed", logConclusionData)
End If
    
End Sub

Sub ResetAutoFilterOnWorksheet(ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ResetAutoFilterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Worksheets(worksheetName).AutoFilterMode = False
    mainWorkbook.Worksheets(worksheetName).Range("A1:" & LastUsedColumnLetterOnWorksheet(worksheetName) & "1").AutoFilter

    Call LogConclusion("Completed", logConclusionData)
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

Sub LogCheckpoint(ByVal checkpointType As String, ByVal checkpointName As String, ByVal checkpointStatus As String, ByVal utcTimestamp As String, ByVal qpc As Double, ByVal tickCount As Double)
    Const methodName As String = "LogCheckpoint"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal checkpointType As String, ByVal checkpointName As String, ByVal checkpointStatus As String, ByVal utcTimestamp As String, ByVal qpc As Double, ByVal tickCount As Double", methodName, isRegistered, "Logging")
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

    Call LogBeginning(methodName, CCur(tickCount / 10000), logConclusionData, """" & checkpointType & """" & ", " & """" & checkpointName & """" & ", " & """" & checkpointStatus & """" & ", " & """" & utcTimestamp & """" & ", " & qpc & ", " & tickCount)

    If checkpointStatus = "Beginning" And report("Log Summary Beginning Tick Count") <> 0 Then
        report("Log Summary Conclusion Tick Count") = logConclusionData.tickCount

        Call SetAboutNamedRange((report("Log Summary Conclusion Tick Count") - report("Log Summary Beginning Tick Count")), "Duration (Milliseconds)")
        Call SetAboutNamedRange("1 Run. 0 Checkpoints. 0 Rows.", "Log Summary")

        report("Log Summary Beginning Tick Count")  = 0
        report("Log Summary Conclusion Tick Count") = 0
    End If

    If checkpointStatus = "Beginning" Then
        report("Log Summary Beginning Tick Count") = logConclusionData.tickCount

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
    Application.Goto Reference:=ActiveSheet.Cells.SpecialCells(xlCellTypeVisible).Range("A" & (lastRowCheckpoint - 30)), Scroll:=True
    Range("A" & logConclusionData.operationSequenceNumber & ":B" & logConclusionData.operationSequenceNumber).Select

    If checkpointStatus = "Conclusion" Then
        If checkpointType = "Foundation" Then
            Call SetAboutNamedRange(checkpointName, "Progression Status")
        ElseIf checkpointType = "Augmentation" Then
            Call SetAboutNamedRange(checkpointName, "Augmentation Modules")
        End If

        Call SetAboutNamedRange("0 Runs. 1 Checkpoint. 0 Rows.", "Log Summary")

        With mainWorkbook.Worksheets("About")
            .Range("ProgressionStatus").Font.Bold = False
            .Range("AugmentationModules").Font.Bold = False
            .Range("DependenciesList").Font.Bold = False
            .Range("CreationDate").Font.Bold = False
            .Range("EditionName").Font.Bold = False
            .Range("DurationMilliseconds").Font.Bold = False
            .Range("LogSummary").Font.Bold = False

            If Len(.Range("ProgressionStatus").Value) >= 18 Then
                .Range("ProgressionStatus").Characters(Start:=1, Length:=18).Font.Bold = True
            End If

            If Len(.Range("AugmentationModules").Value) >= 20 Then
                .Range("AugmentationModules").Characters(Start:=1, Length:=20).Font.Bold = True
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
        End With
    End If

    Call LogConclusion("Completed", logConclusionData)

    If checkpointStatus = "Conclusion" Then
        report("Log Summary Conclusion Tick Count") = (logConclusionData.tickCount + mainWorkbook.Worksheets("Log").Cells(logConclusionData.operationSequenceNumber, 6).Value)

        Call SetAboutNamedRange((report("Log Summary Conclusion Tick Count") - report("Log Summary Beginning Tick Count")), "Duration (Milliseconds)")

        With mainWorkbook.Worksheets("About")
            If Len(.Range("DurationMilliseconds").Value) >= 23 Then
                .Range("DurationMilliseconds").Characters(Start:=1, Length:=23).Font.Bold = True
            End If

            If Len(.Range("LogSummary").Value) >= 11 Then
                .Range("LogSummary").Characters(Start:=1, Length:=11).Font.Bold = True
            End If
        End With

        mainWorkbook.Worksheets("Log").Visible = xlSheetHidden

        report("Log Summary Beginning Tick Count")  = 0
        report("Log Summary Conclusion Tick Count") = 0

        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
    End If
End Sub

Sub LogEngine(ByVal reportName As String, ByVal templateVersion As String, ByVal excelVersion As String, ByVal runIdentifier As String, ByVal qpc As Double, ByVal tickCount As Double)
If report("Log Engine Active") = True Then Exit Sub

    Dim settings As Object
    Const methodName As String = "LogEngine"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal reportName As String, ByVal templateVersion As String, ByVal excelVersion As String, ByVal runIdentifier As String, ByVal qpc As Double, ByVal tickCount As Double", methodName, isRegistered, "Logging")

        Call ConfigureMethodSetting(methodName, "Delete Unnecessary Cell Styles", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Configure Core Cell Styles", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Add Custom Cell Styles", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Format About Worksheet", 1, 0, 1)
        Call ConfigureMethodSetting(methodName, "Format Log Worksheet", 1, 0, 1)
    End If

    Set settings = methodRegistry(methodName)("Settings")

    If WorksheetExists("Log") = True Then
        With mainWorkbook.Worksheets("Log")
            If .Range("A1").Value = "Task" And _
            .Range("B1").Value = "Arguments" And _
            .Range("C1").Value = "Category" And _
            .Range("D1").Value = "Status" And _
            .Range("E1").Value = "Active Worksheet" And _
            .Range("F1").Value = "Original Order" And _
            .Range("G1").Value = "Date Time Start" And _
            .Range("H1").Value = "Date Time End" And _
            .Range("I1").Value = "Stopwatch Start" And _
            .Range("J1").Value = "Stopwatch End" Then
                If Application.DisplayAlerts = True Then
                    Application.DisplayAlerts = False
                End If

                mainWorkbook.Worksheets("Log").Delete

                Application.DisplayAlerts = True

                mainWorkbook.Worksheets.Add After:=mainWorkbook.ActiveSheet
                mainWorkbook.ActiveSheet.Name = "Log"
            Else
                report("Operation Sequence Number") = CDbl(LastUsedRowNumberOnWorksheet("Log"))
            End If
        End With
    Else
        mainWorkbook.Worksheets.Add After:=mainWorkbook.ActiveSheet
        mainWorkbook.ActiveSheet.Name = "Log"
    End If

    If mainWorkbook.Worksheets("Log").Range("A1").Value = "" Then mainWorkbook.Worksheets("Log").Range("A1:F1").Value = Array("Method", "Arguments", "Category", "Outcome", "Tick Count", "Duration")
  
    Call LogBeginning(methodName, CCur(tickCount / 10000), logConclusionData, """" & reportName & """" & ", " & """" & templateVersion & """" & ", " & """" & excelVersion & """" & ", " & """" & runIdentifier & """" & ", " & qpc & ", " & tickCount)

    If report("Settings").Count <> 0 Then
        Dim settingArray As Variant
        For Each settingArray In report("Settings")           
            Call SetMethodSetting(settingArray(0), settingArray(1), settingArray(2))
        Next settingArray
    End If

    If settings("Delete Unnecessary Cell Styles")("Value") = True Then
        Dim cellStyleKey As Variant
        For Each cellStyleKey In cellStyles.Keys
            If cellStyleKey <> "Followed Hyperlink" And cellStyleKey <> "Hyperlink" And cellStyleKey <> "Normal" And cellStyleKey <> "Percent" Then
                If CellStyleExists(cellStyles(cellStyleKey)) Then
                    Call DeleteCellStyle(cellStyles(cellStyleKey))
                Else
                    If CellStyleExists(cellStyleKey) Then
                        Call DeleteCellStyle(cellStyleKey)
                    End If
                End If
            End If
        Next cellStyleKey
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

        Dim aboutWorksheet As Worksheet: Set aboutWorksheet = mainWorkbook.Worksheets("About")
        Dim temporaryCell As Range: Set temporaryCell = aboutWorksheet.Range("B1")

        Call aboutWorksheet.Hyperlinks.Add(temporaryCell, "", "'" & aboutWorksheet.Name & "'!A1")

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

        Dim customCellStyleIndex As Byte
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

    If settings("Format About Worksheet")("Value") = True Then
        With mainWorkbook.Worksheets("About")
            .Cells.RowHeight = 16.5
            .Rows("1:4").RowHeight = 33.0
            .Columns("A:B").ColumnWidth = 90.71
            .Columns("C").ColumnWidth = 27.29

            .Range("A1:C1").WrapText = True
            .Range("B2:C2").WrapText = True
            .Range("A3:C4").WrapText = True

            .Range("B3:B4").VerticalAlignment = xlTop

            .Range("A1:C4").Font.Name = "Tahoma"
            .Range("A1:A2").Font.Bold = True
            .Range("A1:C4").Font.Size = 8
            .Range("A1:B2").Font.Size = 12
        End With
    End If

    If settings("Format Log Worksheet")("Value") = True Then
        With mainWorkbook.Worksheets("Log")
            If .Cells.RowHeight <> 16.5 Or _
            .Rows("1:1").RowHeight <> 49.5 Or _
            .AutoFilterMode <> True Then
                Call NormalizeLayoutOnWorksheet("Log")
            End If

            If .Columns("A").ColumnWidth <> 57.57 Or _
            .Columns("B").ColumnWidth <> 161.86 Or _
            .Columns("C").ColumnWidth <> 9.86 Or _
            .Columns("D").ColumnWidth <> 9.86 Or _
            .Columns("E").ColumnWidth <> 12.14 Or _
            .Columns("F").ColumnWidth <> 9.00 Then
                .Columns("A").ColumnWidth = 57.57
                .Columns("B").ColumnWidth = 161.86
                .Columns("C").ColumnWidth = 9.86
                .Columns("D").ColumnWidth = 9.86
                .Columns("E").ColumnWidth = 12.14
                .Columns("F").ColumnWidth = 9.00
            End If
        End With

        If mainWorkbook.Worksheets("Log").Range("E2").Style <> "Integer" Then
            Call ApplyCellStyleToColumnOnWorksheet("Integer", "Tick Count", "Log")
        End If

        If mainWorkbook.Worksheets("Log").Range("F2").Style <> "Integer" Then
            Call ApplyCellStyleToColumnOnWorksheet("Integer", "Duration", "Log")
        End If
    End If

    report("Log Engine Active") = True
    report("Log Summary Beginning Tick Count") = logConclusionData.tickCount

    Call LogConclusion("Completed", logConclusionData)
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

Sub RegisterMethod(ByVal contract As String, ByVal methodName As String, ByRef isRegistered As Boolean, ByVal categoryName As String)
    Dim methodDictionary As Object
    
    If methodRegistry.Exists(methodName) = False Then
        Set methodDictionary = CreateObject("Scripting.Dictionary")
        Set methodRegistry(methodName) = methodDictionary
    Else
        Set methodDictionary = methodRegistry(methodName)
    End If
    
    methodDictionary("Category")    = categoryName
    methodDictionary("Declaration") = methodName & "(" & contract & ") @ " & categoryName
    isRegistered = True

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

' Core: Logging '

Private Sub LogBeginning(ByVal methodName As String, ByVal tickCount As Currency, ByRef logConclusionData As LogEntry, ByVal arguments As String)
    report("Operation Sequence Number") = report("Operation Sequence Number") + 1

    logConclusionData.operationSequenceNumber = report("Operation Sequence Number")
    logConclusionData.methodName = methodName
    logConclusionData.arguments = arguments
    logConclusionData.tickCount = CDbl(tickCount * 10000)

    mainWorkbook.Worksheets("Log").Range("A" & logConclusionData.operationSequenceNumber & ":" & "E" & logConclusionData.operationSequenceNumber).Value = _
        Array(logConclusionData.methodName, logConclusionData.arguments, methodRegistry(methodName)("Category"), "Beginning", logConclusionData.tickCount)
End Sub

Private Sub LogConclusion(ByVal conclusionStatus As String, ByRef logConclusionData As LogEntry, Optional ByVal errorMessage As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Dim duration As Double: duration = CDbl(tickCount * 10000) - logConclusionData.tickCount

    With mainWorkbook.Worksheets("Log")
        .Cells(logConclusionData.operationSequenceNumber, 4).Value = conclusionStatus
        .Cells(logConclusionData.operationSequenceNumber, 6).Value = duration
    End With

    If errorMessage <> "" Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True

        mainWorkbook.Worksheets("Log").Range("A" & (logConclusionData.operationSequenceNumber) & ":" & "D" & (logConclusionData.operationSequenceNumber)).Font.Color = -11526924

        Dim targetCell As Range: Set targetCell = mainWorkbook.Worksheets("Log").Range("D" & logConclusionData.operationSequenceNumber)
        Dim errorMessageNote As Comment: Set errorMessageNote = targetCell.AddComment
    
        If methodRegistry(logConclusionData.methodName).Exists("Contract") Then
            errorMessageNote.Text Text:="Declaration: " & methodRegistry(logConclusionData.methodName)("Declaration") & vbLf & "Parameters: " & methodRegistry(logConclusionData.methodName)("Parameters") _
                & vbLf & "Arguments: " & logConclusionData.Arguments & vbLf & "Error Output: " & errorMessage
        Else
            errorMessageNote.Text Text:="Declaration: " & methodRegistry(logConclusionData.methodName)("Declaration") & vbLf & "Error Output: " & errorMessage
        End If

        With errorMessageNote.Shape
            .TextFrame.Characters.Font.Name = "Segoe UI"
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.AutoSize = True
        End With
    
        errorMessage = methodRegistry(logConclusionData.methodName)("Declaration") & ". " & errorMessage
        Err.Raise 1000, Description:=errorMessage
    End If
End Sub

' ************ '
' Repetition   '
' ************ '

Sub RepeatAnonymizeNumbersOnColumnOnWorksheet(ByVal columnNames, ByVal worksheetName)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatAnonymizeNumbersOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames, ByVal worksheetName", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String: columnNamesArray = Split(columnNames, "|")
    Dim index As Byte
    
    For index = 0 To UBound(columnNamesArray)
        Call AnonymizeNumbersOnColumnOnWorksheet(columnNamesArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatApplyAutoFilterOnWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatApplyAutoFilterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte
    
    For index = 0 To UBound(worksheetNamesArray)
        Call ApplyAutoFilterOnWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatApplyCellStyleToColumnOnWorksheet(ByVal cellStyle As String, ByVal columnNames As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatApplyCellStyleToColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyle As String, ByVal columnNames As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cellStyle & """" & ", " & """" & columnNames & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String
    Dim worksheetNamesArray() As String
    Dim index As Byte

    If InputContainsValue(columnNames, "|") Then
        columnNamesArray = Split(columnNames, "|")
        For index = 0 To UBound(columnNamesArray)
            Call ApplyCellStyleToColumnOnWorksheet(cellStyle, columnNamesArray(index), worksheetNames)
        Next index
    End If

    If InputContainsValue(worksheetNames, "|") Then
        worksheetNamesArray = Split(worksheetNames, "|")
        For index = 0 To UBound(worksheetNamesArray)
            Call ApplyCellStyleToColumnOnWorksheet(cellStyle, columnNames, worksheetNamesArray(index))
        Next index
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatApplyFreezePanesModeOnWorksheet(ByVal freezePanesMode As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatApplyFreezePanesModeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal freezePanesMode As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & freezePanesMode & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call ApplyFreezePanesModeOnWorksheet(freezePanesMode, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatCutColumnAndPasteAtColumnOnWorksheet(ByVal cutColumnNames As String, ByVal pasteColumnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatCutColumnAndPasteAtColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cutColumnNames As String, ByVal pasteColumnName As String, ByVal worksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cutColumnNames & """" & ", " & """" & pasteColumnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatArgument(cutColumnNames, "cutColumnNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim cutColumnNamesArray() As String
    Dim index As Byte

    If InputContainsValue(cutColumnNames, "|") Then
        cutColumnNamesArray = Split(cutColumnNames, "|")
        For index = 0 To UBound(cutColumnNamesArray)
            Call CutColumnAndPasteAtColumnOnWorksheet(cutColumnNamesArray(index), pasteColumnName, worksheetName)
        Next index
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorBorderOuterAndInnerOnWorksheet(ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatColorBorderOuterAndInnerOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & outerColorName & """" & ", " & """" & innerColorName & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call ColorBorderOuterAndInnerOnWorksheet(outerColorName, innerColorName, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorDataBackgroundAndFontOnWorksheet(ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatColorDataBackgroundAndFontOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim indexColumnName As Byte
    Dim indexWorksheetName As Byte
    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")

    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            Call ColorDataBackgroundAndFontOnWorksheet(columnNamesArray(indexColumnName), backgroundColorName, fontColorName, worksheetNamesArray(indexWorksheetName))
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        Call ColorDataBackgroundAndFontOnWorksheet(columnNamesArray(indexColumnName), backgroundColorName, fontColorName, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        Call ColorDataBackgroundAndFontOnWorksheet(columnNames, backgroundColorName, fontColorName, worksheetNamesArray(indexWorksheetName))
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorHeaderBackgroundAndFontOnWorksheet(ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatColorHeaderBackgroundAndFontOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim indexColumnName As Byte
    Dim indexWorksheetName As Byte
    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")

    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            Call ColorHeaderBackgroundAndFontOnWorksheet(columnNamesArray(indexColumnName), backgroundColorName, fontColorName, worksheetNamesArray(indexWorksheetName))
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        Call ColorHeaderBackgroundAndFontOnWorksheet(columnNamesArray(indexColumnName), backgroundColorName, fontColorName, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        Call ColorHeaderBackgroundAndFontOnWorksheet(columnNames, backgroundColorName, fontColorName, worksheetNamesArray(indexWorksheetName))
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorWorksheet(ByVal colorName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatColorWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & colorName & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call ColorWorksheet(colorName, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatConvertColumnToDateFormattingOnWorksheet(ByVal columnNames As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatConvertColumnToDateFormattingOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String: columnNamesArray = Split(columnNames, "|")
    Dim index As Byte

    For index = 0 To UBound(columnNamesArray)
        Call ConvertColumnToDateFormattingOnWorksheet(columnNamesArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatConvertColumnToDecimalFormattingOnWorksheet(ByVal columnNames As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatConvertColumnToDecimalFormattingOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String: columnNamesArray = Split(columnNames, "|")
    Dim index As Byte

    For index = 0 To UBound(columnNamesArray)
        Call ConvertColumnToDecimalFormattingOnWorksheet(columnNamesArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatCreateBlankWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatCreateBlankWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call CreateBlankWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteColumnOnWorksheet(ByVal columnNames As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDeleteColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim indexColumnName As Byte
    Dim indexWorksheetName As Byte
    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")

    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            Call DeleteColumnOnWorksheet(columnNamesArray(indexColumnName), worksheetNamesArray(indexWorksheetName))
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        Call DeleteColumnOnWorksheet(columnNamesArray(indexColumnName), worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        Call DeleteColumnOnWorksheet(columnNames, worksheetNamesArray(indexWorksheetName))
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteRowsOnWorksheet(ByVal numberOfRows As Long, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDeleteRowsOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal numberOfRows As Long, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, numberOfRows & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call DeleteRowsOnWorksheet(numberOfRows, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDeleteWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call DeleteWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDuplicateColumnFromWorksheetToWorksheet(ByVal columnNames As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDuplicateColumnFromWorksheetToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & fromWorksheetName & """" & ", " & """" & toWorksheetName & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String: columnNamesArray = Split(columnNames, "|")
    Dim index As Byte

    For index = 0 To UBound(columnNamesArray)
        Call DuplicateColumnFromWorksheetToWorksheet(columnNamesArray(index), fromWorksheetName, toWorksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatFindAndReplaceOnRangeOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValues As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatFindAndReplaceOnRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValues As String, ByVal worksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & rangeValues & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatArgument(rangeValues, "rangeValues", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim rangeValuesArray() As String: rangeValuesArray = Split(rangeValues, "|")
    Dim index As Byte

    For index = 0 To UBound(rangeValuesArray)
        Call FindAndReplaceOnRangeOnWorksheet(findValue, replaceValue, rangeValuesArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatHideWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatHideWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call HideWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatInsertDataValuesOnWorksheet(ByVal dataValues As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatInsertDataValuesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal dataValues As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & dataValues & """" & ", "  & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(dataValues, "dataValues", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call InsertDataValuesOnWorksheet(dataValues, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatInsertNewColumnAndSetWidthOnWorksheet(ByVal columnNames As String, ByVal setWidth As Double, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatInsertNewColumnAndSetWidthOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal setWidth As Double, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", "  & setWidth & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim indexColumnName As Byte
    Dim indexWorksheetName As Byte
    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")

    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            Call InsertNewColumnAndSetWidthOnWorksheet(columnNamesArray(indexColumnName), setWidth, worksheetNamesArray(indexWorksheetName))
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        Call InsertNewColumnAndSetWidthOnWorksheet(columnNamesArray(indexColumnName), setWidth, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        Call InsertNewColumnAndSetWidthOnWorksheet(columnNames, setWidth, worksheetNamesArray(indexWorksheetName))
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatInsertHeaderValuesOnWorksheet(ByVal headerValues As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatInsertHeaderValuesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal headerValues As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & headerValues & """" & ", "  & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(headerValues, "headerValues", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call InsertHeaderValuesOnWorksheet(headerValues, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatMoveWorksheetToEnd(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatMoveWorksheetToEnd"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call MoveWorksheetToEnd(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatNormalizeLayoutOnWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatNormalizeLayoutOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call NormalizeLayoutOnWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatRemoveBlankLinesOnWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatRemoveBlankLinesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetsArray() As String: worksheetsArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetsArray)
        Call RemoveBlankLinesOnWorksheet(worksheetsArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteCellStyle(ByVal cellStyleNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatDeleteCellStyle"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyleNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & cellStyleNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(cellStyleNames, "cellStyleNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim cellStyleNamesArray() As String: cellStyleNamesArray = Split(cellStyleNames, "|")
    Dim index As Byte

    For index = 0 To UBound(cellStyleNamesArray)
        Call DeleteCellStyle(cellStyleNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatRemoveEmptyColumnsOnWorksheet(ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatRemoveEmptyColumnsOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call RemoveEmptyColumnsOnWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatSetNewWidthOnColumnOnWorksheet(ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatSetNewWidthOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call SetNewWidthOnColumnOnWorksheet(newWidth, columnName, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatSortColumnByOrderOnWorksheet(ByVal columnNames As String, ByVal sortOrder As String, ByVal worksheetNames As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatSortColumnByOrderOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal sortOrder As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnNames & """" & ", " & """" & sortOrder & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatArgument(columnNames, "columnNames", validation)
    Call ValidateRepeatArgument(worksheetNames, "worksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim indexColumnName As Byte
    Dim indexWorksheetName As Byte
    Dim worksheetNamesArray() As String
    Dim columnNamesArray() As String

If InputContainsValue(columnNames, "|") And InputContainsValue(worksheetNames, "|") Then
    worksheetNamesArray = Split(worksheetNames, "|")
    columnNamesArray = Split(columnNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        For indexColumnName = 0 To UBound(columnNamesArray)
            Call SortColumnByOrderOnWorksheet(columnNamesArray(indexColumnName), sortOrder, worksheetNamesArray(indexWorksheetName))
        Next indexColumnName
    Next indexWorksheetName
ElseIf InputContainsValue(columnNames, "|") Then
    columnNamesArray = Split(columnNames, "|")
    For indexColumnName = 0 To UBound(columnNamesArray)
        Call SortColumnByOrderOnWorksheet(columnNamesArray(indexColumnName), sortOrder, worksheetNames)
    Next indexColumnName
Else
    worksheetNamesArray = Split(worksheetNames, "|")
    For indexWorksheetName = 0 To UBound(worksheetNamesArray)
        Call SortColumnByOrderOnWorksheet(columnNames, sortOrder, worksheetNamesArray(indexWorksheetName))
    Next indexWorksheetName
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatTransferDataFromWorksheetToWorksheet(ByVal fromWorksheetNames As String, toWorksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "RepeatTransferDataFromWorksheetToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetNames As String, toWorksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & fromWorksheetNames & """" & ", " & """" & toWorksheetName & """")

    Dim validation As String
    Call ValidateRepeatArgument(fromWorksheetNames, "fromWorksheetNames", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(fromWorksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call TransferDataFromWorksheetToWorksheet(worksheetNamesArray(index), toWorksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

' No logging. '

Sub RepeatAssertMinimumRowsOfDataOnWorksheet(ByVal minimumRowsOfData As Long, ByVal worksheetNames As String)
    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte
    
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim formulaColumn As String
    Dim columnReference As String
    formulaColumn = "Formula Column"
    columnReference = ColumnStartOnWorksheet(columnName, worksheetName)

    ' https://web.archive.org/web/20201028212900/https://heroes.thelazy.net/index.php/Creature
    Call InsertNewColumnAndSetWidthOnWorksheet(formulaColumn, 13.57, worksheetName)
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal consecutiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & consecutiveGroupName & """" & ", " & """" & formulaValue & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call ApplyFormulaToColumnOnWorksheet(formulaValue, report("Helper Column"), worksheetName)
    Call ApplyFilterToColumnOnWorksheet(True, report("Helper Column"), worksheetName)

    Call FindAndReplaceOnRangeOnWorksheet(";;", ";" & consecutiveGroupName & ";;", columnName, worksheetName)
    Call ClearFilterOnWorksheet(worksheetName)

    Call SelectWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet(ByVal definitiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal definitiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & definitiveGroupName & """" & ", " & """" & formulaValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If formulaValue = "" Then formulaValue = "=True"

    Call ApplyFormulaToColumnOnWorksheet(formulaValue, report("Helper Column"), worksheetName)
    Call ApplyFilterToColumnOnWorksheet(True, report("Helper Column"), worksheetName)
    Call FindAndReplaceOnRangeOnWorksheet(";;", definitiveGroupName, columnName, worksheetName)
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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & outerColorName & """" & ", " & """" & innerColorName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If outerColorName = "" Then outerColorName = "Light Grey"
    If innerColorName = "" Then innerColorName = "Light Grey"

    Dim lastColumnLetterPlusOne As String
    lastColumnLetterPlusOne = LastUsedColumnLetterOnWorksheet(worksheetName)
    lastColumnLetterPlusOne = IncreaseLetterOnce(lastColumnLetterPlusOne)
    lastColumnLetterPlusOne = IncreaseLetterOnce(lastColumnLetterPlusOne)

    ' Outer Border. '
    With mainWorkbook.Worksheets(worksheetName).Cells.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = StyleColor(outerColorName)
        .Weight = xlThin
    End With
    With mainWorkbook.Worksheets(worksheetName).Cells.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = StyleColor(outerColorName)
        .Weight = xlThin
    End With

    Dim lastRowNumberPlusOne As Long
    lastRowNumberPlusOne = LastUsedRowNumberOnWorksheet(worksheetName) + 1

    mainWorkbook.Worksheets(worksheetName).Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    mainWorkbook.Worksheets(worksheetName).Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Inner Border. '
    With mainWorkbook.Worksheets(worksheetName).Range("A1:" & lastColumnLetterPlusOne & lastRowNumberPlusOne + 1).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = StyleColor(innerColorName)
        .Weight = xlThin
    End With
    With mainWorkbook.Worksheets(worksheetName).Range("A1:" & lastColumnLetterPlusOne & lastRowNumberPlusOne + 1).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = StyleColor(innerColorName)
        .Weight = xlThin
    End With

    ' Clean up of bug or weird behaviour which doesn't apply the border everywhere. '
    lastColumnLetterPlusOne = DecreaseLetterOnce(lastColumnLetterPlusOne)
    mainWorkbook.Worksheets(worksheetName).Columns("A:A").Delete Shift:=xlToLeft
    mainWorkbook.Worksheets(worksheetName).Rows("1:1").Delete Shift:=xlUp
    mainWorkbook.Worksheets(worksheetName).Columns(lastColumnLetterPlusOne & ":" & lastColumnLetterPlusOne).Delete Shift:=xlToLeft
    mainWorkbook.Worksheets(worksheetName).Columns(lastColumnLetterPlusOne & ":" & lastColumnLetterPlusOne).Delete Shift:=xlToLeft
    mainWorkbook.Worksheets(worksheetName).Rows(lastRowNumberPlusOne & ":" & lastRowNumberPlusOne).Delete Shift:=xlUp
    mainWorkbook.Worksheets(worksheetName).Rows(lastRowNumberPlusOne & ":" & lastRowNumberPlusOne).Delete Shift:=xlUp

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ColorDataBackgroundAndFontOnWorksheet(ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatColorHeaderBackgroundAndFontOnWorksheet(columnName, backgroundColorName, fontColorName, worksheetName)
Else
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ColorDataBackgroundAndFontOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

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
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & columnName & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateColumnOnWorksheet(columnName, "columnName", worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If backgroundColorName = "" Then backgroundColorName = "White"
    If fontColorName = "" Then fontColorName = "Black"

    Call ColorColumnOfTypeAndElementOnWorksheet(backgroundColorName, columnName, "Header", "Background", worksheetName)
    Call ColorColumnOfTypeAndElementOnWorksheet(fontColorName, columnName, "Header", "Font", worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub FindAndReplaceExactOnRangeOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValue As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "FindAndReplaceExactOnRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & rangeValue & """" & ", " & """" & worksheetName & """")

    Call ApplyFilterToColumnOnWorksheet(findValue, rangeValue, worksheetName)
    Call FindAndReplaceOnRangeOnWorksheet(findValue, replaceValue, rangeValue, worksheetName)
    Call ClearFilterOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ImportExcelFileWithDependencyAndRenameWorksheet(ByVal filePath As String, ByVal dependency As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "ImportExcelFileWithDependencyAndRenameWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal dependency As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & filePath & """" & " ," & """" & dependency & """" & " ," & """" & worksheetName & """")

    Call ImportExcelFileAndRenameWorksheet(filePath, worksheetName)
    Call SetAboutNamedRange(dependency, "Dependencies List")

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub MoveDataFromWorksheetToWorksheet(ByVal fromWorksheetName As String, ByVal worksheetName As String)
    Dim tickCount As Currency: tickCount = GetTickCount64()
    Const methodName As String = "MoveDataFromWorksheetToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, tickCount, logConclusionData, """" & fromWorksheetName & """" & ", " & """" & worksheetName & """")
  
    Dim validation As String
    Call ValidateWorksheet(fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call EnsureHelperColumnDeletedOnWorksheet(fromWorksheetName)
    Call EnsureHelperColumnDeletedOnWorksheet(worksheetName)
   
    Dim numberOfColumnsFromWorksheetName As Long: numberOfColumnsFromWorksheetName = ConvertLetterToNumber(LastUsedColumnLetterOnWorksheet(fromWorksheetName))
    Dim numberOfColumnsWorksheetName As Long: numberOfColumnsWorksheetName = ConvertLetterToNumber(LastUsedColumnLetterOnWorksheet(worksheetName))

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
    If ColumnOnWorksheetExists(report("Helper Column"), worksheetName) Then
        Call ResetColumnOnWorksheet(report("Helper Column"), worksheetName)
    Else
        Call InsertNewColumnAndSetWidthOnWorksheet(report("Helper Column"), 7.00, worksheetName)
    End If
End Sub

Sub EnsureHelperColumnDeletedOnWorksheet(ByVal worksheetName As String)
    If ColumnOnWorksheetExists(report("Helper Column"), worksheetName) Then
        Call DeleteColumnOnWorksheet(report("Helper Column"), worksheetName)
    End If
End Sub

Sub EnsureHelperColumnsDeletedOnMainWorkbook()
    Dim worksheetNumber As Long
    Dim index As Long

    worksheetNumber = mainWorkbook.Worksheets.count

    For index = 1 To worksheetNumber
        Call EnsureHelperColumnDeletedOnWorksheet(mainWorkbook.Worksheets(index).Name)
    Next
End Sub

Sub ValidateColumnOnWorksheet(ByVal columnName As String, ByVal columnParameterName As String, ByVal worksheetName As String, ByVal worksheetParameterName As String, ByRef validation As String)
    Dim validationMessage As String
    Dim validationLength As Long

    validationLength  = Len(validationMessage)
    Call ValidateWorksheet(worksheetName, worksheetParameterName, validationMessage)

    If validationMessage <> "" And validationLength <> Len(validationMessage) Then
        validationMessage = "Parameter """ & worksheetParameterName & """ failed validation. " & validationMessage

        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If

        Exit Sub
    End If

    If Len(columnName) = 0 Then
        validationMessage = "Column name can't be blank."
    ElseIf Len(Trim$(columnName)) = 0 Then
        validationMessage = "The column name cannot consist only of spaces."
    ElseIf InStr(columnName, "*") > 0 Then
        validationMessage = "Column name can't contain an asterisk."
    ElseIf InStr(columnName, "?") > 0 Then
        validationMessage = "Column name can't contain a question mark."
    End If

    If validationMessage = "" Then
        Dim worksheetToValidate As Worksheet
        Set worksheetToValidate = mainWorkbook.Worksheets(worksheetName)

        Dim headerRowRange As Range
        Set headerRowRange = worksheetToValidate.Rows(1)

        Dim matchResult As Variant
        matchResult = Application.Match(columnName, headerRowRange, 0)

        If IsError(matchResult) Then
            validationMessage = "The column name doesn't exist on the worksheet."
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

Sub ValidateRangeOnWorksheet(ByRef rangeValue As String, ByVal worksheetName As String, ByRef validation As String)
    If Len(rangeValue) <= 3 Then
        If ColumnLetterValid(rangeValue) Then
            rangeValue = ColumnRangeTypeOnWorksheet(rangeValue, "Data", worksheetName)
        Else
            ' Call LogConclusion("Failed", logConclusionData, "Column letter " & """" & rangeValue & """" & " on worksheet """ & worksheetName & """ is not a valid column.")
        End If
    Exit Sub
    End If

    Dim validRange As Boolean
    validRange = True

    If CBool(InStr(rangeValue, ":")) = False Then validRange = False ' A range always contains the : character.'
    If Len(rangeValue) > 21 = True Then validRange = False ' Max: XFD1048576:XFD1048576. '
    If IsNumeric(Left(rangeValue, 1)) = True Then validRange = False ' Last value must be a number. '
    If IsNumeric(Right(rangeValue, 1)) = False Then validRange = False ' First value must be a character. '

    If validRange = False Then
        If ColumnOnWorksheetExists(rangeValue, worksheetName) = False Then Call LogConclusion("Failed", logConclusionData, "Column " & """" & rangeValue & """" & " on worksheet """ & worksheetName & """ not found.")
        rangeValue = ColumnRangeTypeOnWorksheet(rangeValue, "Data", worksheetName)
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

Sub ValidateUniqueColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String, ByRef validation As String)
    ' If ColumnOnWorksheetExists(columnName, worksheetName) Then Call LogConclusion("Failed", logConclusionData, "Column " & """" & columnName & """" & " on worksheet " & """" & worksheetName & """" & " already exists.")
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
        Dim characterIndex As Byte
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

Function ColumnIsEmptyOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As Boolean
    If WorksheetFunction.CountA(Sheets(worksheetName).Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName))) = 0 Then
        ColumnIsEmptyOnWorksheet = True
        Exit Function
    End If

    Dim sumZeroRange As Range

    Set sumZeroRange = Worksheets(worksheetName).Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName))
    If WorksheetFunction.Sum(sumZeroRange) = 0 Then
        If Application.Count(Sheets(worksheetName).Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName))) = Application.CountA(Sheets(worksheetName).Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName))) Then ColumnIsEmptyOnWorksheet = True
    End If
End Function

Function ColumnLetterValid(ByVal columnLetter As String) As Boolean
    ColumnLetterValid = columnLetter Like WorksheetFunction.Rept("[a-zA-Z]", Len(columnLetter))

    On Error Resume Next

    Dim convertLetterToNumber As Long
    convertLetterToNumber = mainWorkbook.Worksheets("Log").Range(columnLetter & 1).Column

    If Err.Number <> 0 Then
        Err.Clear
        ColumnLetterValid = False
    End If
End Function

Function ColumnOnWorksheetExists(ByVal columnName As String, ByVal worksheetName As String) As Boolean
    Dim columnRange As Range

    If WorksheetExists(worksheetName) Then Set columnRange = mainWorkbook.Worksheets(worksheetName).Range("A1:" & LastUsedColumnLetterOnWorksheet(worksheetName) & "1").Find(columnName, LookIn:=xlValues, Lookat:=xlWhole)

    If columnRange Is Nothing Then
        ColumnOnWorksheetExists = False
    Else
        ColumnOnWorksheetExists = True
    End If

    Application.FindFormat.Clear
End Function

Function InputContainsValue(ByVal inputText As String, ByVal searchValue As String) As Boolean
    If Len(searchValue) = 0 Then
        InputContainsValue = False
        Exit Function
    End If

    InputContainsValue = InStr(1, inputText, searchValue, vbTextCompare) > 0
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
    Dim targetWorksheet As Worksheet

    worksheetFound = WorksheetExists(worksheetName)

    If worksheetFound = False Then
        WorksheetIsEmpty = Null
        Exit Function
    End If

    Set targetWorksheet = mainWorkbook.Worksheets(worksheetName)
    
    WorksheetIsEmpty = Application.WorksheetFunction.CountA(targetWorksheet.Cells) = 0 And targetWorksheet.Shapes.Count = 0
End Function

' ************ '
' ************ '

Sub Run()
    Set mainWorkbook = ActiveWorkbook

    Set methodRegistry = CreateObject("Scripting.Dictionary")
    Set report = CreateObject("Scripting.Dictionary")

    report("Checkpoint Type")   = "Foundation"
    report("Helper Column")     = "Helper Column"
    report("Log Engine Active") = False
    report("Log Summary Beginning Tick Count")  = 0#
    report("Log Summary Conclusion Tick Count") = 0#
    report("Operation Sequence Number") = 1&
    report("Original Workbook") = mainWorkbook.FullName
    report("Creation Date")     = Format$(Date, "yyyy-mm-dd")
    report("Template Version")  = "v0.40, 2026-08-06"
    Set report("Settings") = New Collection

    Call Startup()
    Call Master()
End Sub

' Script is finished when "About" is placed last. '
' When saving the document, choose yes at prompt. '