' ************ '
' Alteration   '
' ************ '

Sub ApplyActiveFormulaToRangeOnWorksheet(ByVal activeFormulaValue As String, ByVal rangeValue As String, ByVal worksheetName As String)
    Const methodName As String = "ApplyActiveFormulaToRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal activeFormulaValue As String, ByVal rangeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & activeFormulaValue & """" & ", " & """" & rangeValue & """" & ", " & """" & worksheetName & """")
    
    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If rangeValue = report("Helper Column") Then Call EnsureHelperColumnOnWorksheet(worksheetName)
    Call ValidateRangeOnWorksheet(rangeValue, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If Sheets(worksheetName).Range(FirstColumnInRange(rangeValue) & FirstRowInRange(rangeValue)).Style <> "Formula" Then Call ApplyCellStyleToRangeOnWorksheet("Formula", rangeValue, worksheetName)
    Sheets(worksheetName).Range(FirstColumnInRange(rangeValue) & FirstRowInRange(RangeValue)).FormulaLocal = activeFormulaValue

    Call SelectWorksheet(worksheetName)

    If Sheets(worksheetName).Range("A2").Value = "" Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & worksheetName & """" & " doesn't appear to have any valid data, can't apply formula.")

    If Range(rangeValue).Rows.Count <> 1 Then
        Sheets(worksheetName).Range(rangeValue).FillDown
    End If
    
    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyFormulaToCellWithCellStyleOnWorksheet(ByVal formulaValue As String, ByVal cellValue As String, ByVal cellStyle As String, ByVal worksheetName As String)
    Const methodName As String = "ApplyFormulaToCellWithCellStyleOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal cellValue As String, ByVal cellStyle As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & formulaValue & """" & ", " & """" & cellValue & """" & ", " & """" & cellStyle & """" & ", " & """" & worksheetName & """")
    
    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If CellStyleExists(cellStyle) = False Then
        Call LogConclusion("Failed", logConclusionData, "Cell style """ & cellStyle & """ does not exist.")
    End If

    If Sheets(worksheetName).Range(cellValue).Style <> "Formula" Then
        With Sheets(worksheetName).Range(cellValue)
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
    Sheets(worksheetName).Range(cellValue).Formula2 = formulaValue
    If Err.Number <> 0 Then
        Err.Clear
        Sheets(worksheetName).Range(cellValue).FormulaLocal = formulaValue
    End If
    
    If Sheets(worksheetName).Range(cellValueLetter & cellValueNumber) <> "" Then
        Sheets(worksheetName).Range(Sheets(worksheetName).Range(cellValue), Sheets(worksheetName).Range(cellValue).End(xlDown)).Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    Else
        Sheets(worksheetName).Range(cellValue).Copy
        Sheets(worksheetName).Range(cellValue).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End If

    With Sheets(worksheetName).Range(cellValue)
        .Style = cellStyle
        .Value = .Value
    End With

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyFormulaToRangeOnWorksheet(ByVal formulaValue As String, ByVal rangeValue As String, ByVal worksheetName As String)
    Const methodName As String = "ApplyFormulaToRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal rangeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & formulaValue & """" & ", " & """" & rangeValue & """" & ", " & """" & worksheetName & """")
    
    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If rangeValue = report("Helper Column") Then Call EnsureHelperColumnOnWorksheet(worksheetName)
    Call ValidateRangeOnWorksheet(rangeValue, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If Sheets(worksheetName).Range(FirstColumnInRange(rangeValue) & FirstRowInRange(rangeValue)).Style <> "Formula" Then Call ApplyCellStyleToRangeOnWorksheet("Formula", rangeValue, worksheetName)
    Sheets(worksheetName).Range(FirstColumnInRange(rangeValue) & FirstRowInRange(RangeValue)).FormulaLocal = formulaValue

    Call SelectWorksheet(worksheetName)

    If Sheets(worksheetName).Range("A2").Value = "" Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & worksheetName & """" & " doesn't appear to have any valid data, can't apply formula.")

    If Range(rangeValue).Rows.Count <> 1 Then
        Sheets(worksheetName).Range(rangeValue).FillDown
    End If
    
    ' To avoid bugs with number formatting, i.e. leading zeroes. '
    Sheets(worksheetName).Range(rangeValue).Copy
    Sheets(worksheetName).Range(rangeValue).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyHyperlinkToCellOnWorksheet(ByVal hyperlinkValue As String, ByVal cellValue As String, ByVal worksheetName As String)
    Const methodName As String = "ApplyHyperlinkToCellOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal hyperlinkValue As String, ByVal cellValue As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & hyperlinkValue & """" & ", " & """" & cellValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim currentCellValue As String: currentCellValue = Sheets(worksheetName).Range(cellValue).Value

    If InputContainsValue(hyperlinkValue, "http") Or InputContainsValue(hyperlinkValue, "https") Then
        Sheets(worksheetName).Hyperlinks.Add Anchor:=Sheets(worksheetName).Range(cellValue), Address:="", SubAddress:=hyperlinkValue, TextToDisplay:=currentCellValue
    Else
        Sheets(worksheetName).Hyperlinks.Add Anchor:=Sheets(worksheetName).Range(cellValue), Address:="", SubAddress:="'" & hyperlinkValue & "'!A1", TextToDisplay:=currentCellValue
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub FindAndReplaceOnRangeOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValue As String, ByVal worksheetName As String)  ' Repeat Support: rangeValue. '

If InputContainsValue(rangeValue, "|") Then
    Call RepeatFindAndReplaceOnRangeOnWorksheet(findValue, replaceValue, rangeValue, worksheetName)
Else
    Const methodName As String = "FindAndReplaceOnRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & rangeValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateRangeOnWorksheet(rangeValue, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).Range(rangeValue).Replace What:=findValue, Replacement:=replaceValue, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub HideWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatHideWorksheet(worksheetName)
Else
    Const methodName As String = "HideWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).Visible = xlSheetHidden

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertDataValuesOnWorksheet(ByVal dataValues As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatInsertDataValuesOnWorksheet(dataValues, worksheetName)
Else
    Const methodName As String = "InsertDataValuesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal dataValues As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & dataValues & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim delimitedStringArray() As String: delimitedStringArray = Split(dataValues, "|")
    Dim currentLastDataRow As Long: currentLastDataRow = LastRowNumberOnWorksheet(worksheetName) + 1
    Dim index As Long

    For index = 0 To UBound(delimitedStringArray)
        mainWorkbook.Sheets(worksheetName).Range(ConvertNumberToLetter(index + 1) & currentLastDataRow).Value = delimitedStringArray(index)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertHeaderValuesOnWorksheet(ByVal headerValues As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatInsertHeaderValuesOnWorksheet(headerValues, worksheetName)
Else
    Const methodName As String = "InsertHeaderValuesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal headerValues As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & headerValues & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Sheets(worksheetName).Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    Dim delimitedStringArray() As String: delimitedStringArray = Split(headerValues, "|")
    Dim index As Long

    For index = 0 To UBound(delimitedStringArray)
        mainWorkbook.Sheets(worksheetName).Range(ConvertNumberToLetter(index + 1) & "1").Value = delimitedStringArray(index)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertNewColumnAndSetWidthOnWorksheet(ByVal columnName As String, ByVal setWidth As Double, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatInsertNewColumnAndSetWidthOnWorksheet(columnName, setWidth, worksheetName)
Else
    Const methodName As String = "InsertNewColumnAndSetWidthOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal setWidth As Double, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & setWidth & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateUniqueColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim lastColumnLetterPlusOne As String
    Dim lastColumnRow As Long

    lastColumnRow = LastRowNumberOnWorksheet(worksheetName)
    lastColumnLetterPlusOne = LastColumnLetterOnWorksheet(worksheetName)
    lastColumnLetterPlusOne = IncreaseLetterOnce(lastColumnLetterPlusOne)

    If Sheets(worksheetName).Range("A1").Value = "" Then lastColumnLetterPlusOne = "A"

    Sheets(worksheetName).Range(lastColumnLetterPlusOne & "1").Value = columnName
    If setWidth <> 0 Then Sheets(worksheetName).Range(lastColumnLetterPlusOne & "1").ColumnWidth = setWidth

    With Sheets(worksheetName).Range(lastColumnLetterPlusOne & "1")
        .Style = "Header"
        .Value = .Value
    End With

    Call ResetAutoFilterOnWorksheet(worksheetName)
    Call ApplyCellStyleToRangeOnWorksheet("Formula", FindColumnLetterOnWorksheet(columnName, worksheetName), worksheetName)
    Call ResetColumnOnWorksheet(lastColumnLetterPlusOne, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub MoveWorksheetToEnd(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatMoveWorksheetToEnd(worksheetName)
Else
    Const methodName As String = "MoveWorksheetToEnd"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetIsHidden As Boolean: worksheetIsHidden = False
    If Sheets(worksheetName).Visible = xlSheetHidden Then worksheetIsHidden = True

    If worksheetIsHidden = True Then
        Sheets(worksheetName).Visible = xlSheetVisible
    End If

    Sheets(worksheetName).Move After:=Sheets(Worksheets.Count)

    If worksheetIsHidden = True Then
        Sheets(worksheetName).Visible = xlSheetHidden
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RenameWorksheetToValue(ByVal currentWorksheetName As String, ByVal newWorksheetName As String)
    Const methodName As String = "RenameWorksheetToValue"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal newWorksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & newWorksheetName & """")
    
    If WorksheetExists(currentWorksheetName) = False Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & currentWorksheetName & """" & " doesn't exist.")
    If WorksheetExists(newWorksheetName) Then Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & newWorksheetName & """" & " already exists.")
    If Len(newWorksheetName) > 31 Then Call LogConclusion("Failed", logConclusionData, "Worksheet name " & """" & newWorksheetName & """" & " is too long, max is 31 characters.")

    Sheets(currentWorksheetName).Name = newWorksheetName

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ResetColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String)
    Const methodName As String = "ResetColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If Sheets(worksheetName).Range("A2").Value <> "" Then Sheets(worksheetName).Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName)).FormulaArray = ";;"

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetMethodSetting(ByVal settingMethod As String, ByVal settingName As String, ByVal settingValue As Long)
    If report("Log Engine Active") = False Then
        Call report("Pre-Log Settings").Add(Array(settingMethod, settingName, settingValue))
        Exit Sub
    End If

    Const methodName As String = "SetMethodSetting"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal settingMethod As String, ByVal settingName As String, ByVal settingValue As Long", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & settingMethod & """" & ", " & """" & settingName & """" & ", " & """" & settingValue & """")
End Sub

Sub SetNewHeaderOnColumnOnWorksheet(ByVal newHeader As String, ByVal columnName As String, ByVal worksheetName As String)
    Const methodName As String = "SetNewHeaderOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal newHeader As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & newHeader & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).Range(columnName & "1").Value = newHeader

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub SetNewWidthOnColumnOnWorksheet(ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatSetNewWidthOnColumnOnWorksheet(newWidth, columnName, worksheetName)
Else
    Const methodName As String = "SetNewWidthOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Alteration")
    End If

    Call LogBeginning(methodName, logConclusionData, newWidth & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If Len(columnName) >= 4 Then Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).Range(columnName & "1").ColumnWidth = newWidth

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

Sub ConfigureAbout(ByVal reportDetailsInput As String, ByVal reportVisionInput As String, ByVal dependenciesListInput As String, ByVal retrievedDateInput As String)
    If Range("ReportDetails").Font.Color <> 1973278 Then
        With mainWorkbook.Styles("Normal")
            .IncludeNumber = True
            .IncludeFont = True
            .IncludeAlignment = True
            .IncludeBorder = False
            .IncludePatterns = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ReadingOrder = xlContext
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
        End With
        mainWorkbook.Styles("Normal").NumberFormat = "@"

        With mainWorkbook.Styles("Normal").Font
            .Name = "Arial"
            .Size = 10
            .Color = 0
        End With

        If CellStyleExists("Hyperlink") Then mainWorkbook.Styles("Hyperlink").Delete
        mainWorkbook.Styles.Add Name:="Hyperlink"

        With mainWorkbook.Styles("Hyperlink")
            .IncludeNumber = True
            .IncludeFont = False
            .IncludeAlignment = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ReadingOrder = xlContext
            .WrapText = True
        End With
        mainWorkbook.Styles("Hyperlink").NumberFormat = "@"
    End If

    If mainWorkbook.Sheets("About").Range("A1").Value = "" Then
        mainWorkbook.Sheets("About").Cells.RowHeight = 16.5
        mainWorkbook.Sheets("About").Rows("1:4").RowHeight = 33.0
        mainWorkbook.Sheets("About").Columns("A:B").ColumnWidth = 90.71
        mainWorkbook.Sheets("About").Columns("C:C").ColumnWidth = 27.29

        mainWorkbook.Sheets("About").Range("A1:C4").Font.Name = "Tahoma"
        mainWorkbook.Sheets("About").Range("A1:C4").Font.Bold = True
        mainWorkbook.Sheets("About").Range("A1:C4").Font.Size = 8
        mainWorkbook.Sheets("About").Range("A1:B2").Font.Size = 12

        mainWorkbook.Sheets("About").Range("A3:C4").WrapText = True
        mainWorkbook.Sheets("About").Range("B2:C2").WrapText = True
        mainWorkbook.Sheets("About").Range("A1:C1").WrapText = True
        mainWorkbook.Sheets("About").Range("B3:B4").VerticalAlignment = xlTop
        mainWorkbook.Sheets("About").Range("B1:B2").Merge
        mainWorkbook.Sheets("About").Range("B3:B4").Merge
        mainWorkbook.Sheets("About").Range("A1:C4").Font.Color = 1973278
        mainWorkbook.Sheets("About").Columns("D:D").ColumnWidth = 4

        mainWorkbook.Sheets("About").Cells.Borders.LineStyle = xlLineStyleNone
        mainWorkbook.Sheets("About").Cells.Interior.Pattern = xlNone
        mainWorkbook.Sheets("About").Cells.Interior.TintAndShade = 0
        mainWorkbook.Sheets("About").Cells.Interior.PatternTintAndShade = 0

        Sheets("About").Rows("5:1048576").RowHeight = 16.5
        Sheets("About").Columns("D:XFD").Cells.Style = "Normal"
        Sheets("About").Rows("5:1048576").Cells.Style = "Normal"

        Sheets("About").Columns("D:D").Delete
        Sheets("About").Rows("5:5").Delete

        With mainWorkbook.Styles("Normal")
            .IncludeFont = False
        End With
    End If
 
    If mainWorkbook.Sheets("About").Range("ReportDetails").Value = "" And reportDetailsInput <> "" Then mainWorkbook.Sheets("About").Range("ReportDetails").Value = reportDetailsInput
    If mainWorkbook.Sheets("About").Range("ReportVision").Value = "" And reportVisionInput <> "" Then mainWorkbook.Sheets("About").Range("ReportVision").Value = reportVisionInput
    If (mainWorkbook.Sheets("About").Range("DependenciesList").Value = "" Or mainWorkbook.Sheets("About").Range("DependenciesList").Value = "Dependencies List: N/A.") And dependenciesListInput <> "" Then mainWorkbook.Sheets("About").Range("DependenciesList").Value = "Dependencies List: " & dependenciesListInput & "."
    If mainWorkbook.Sheets("About").Range("RetrievedDate").Value = "" And retrievedDateInput <> "" Then mainWorkbook.Sheets("About").Range("RetrievedDate").Value = "Retrieved Date: " & retrievedDateInput & "."

    If reportDetailsInput = "" Then mainWorkbook.Sheets("About").Range("ReportDetails").Value = "Standalone Report"
    If reportVisionInput = "" Then mainWorkbook.Sheets("About").Range("ReportVision").Value = "Standalone Mode."
    If dependenciesListInput = "" Then mainWorkbook.Sheets("About").Range("DependenciesList").Value = "Dependencies List: N/A."
    If retrievedDateInput = "" Then mainWorkbook.Sheets("About").Range("RetrievedDate").Value = "Retrieved Date: " & Format(Now, "DD.MM.YYYY") & "."
End Sub

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
        Set methodSettingsDictionary = methodRegistry(methodName)("Settings")
    End If

    If methodSettingsDictionary.Exists(settingName) = False Then
        Set methodSubSettingDictionary = CreateObject("Scripting.Dictionary")
        Set methodSettingsDictionary(settingName) = methodSubSettingDictionary
    Else
        Set methodSubSettingDictionary = methodRegistry(methodName)(settingName)
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
        If methodRegistry(methodName)("Settings")(settingName)("Value") > methodRegistry(methodName)("Settings")(settingName)("Ceiling") Then
            methodRegistry(methodName)("Settings")(settingName)("Value") = methodRegistry(methodName)("Settings")(settingName)("Ceiling")
        ElseIf methodRegistry(methodName)("Settings")(settingName)("Value") < methodRegistry(methodName)("Settings")(settingName)("Floor") Then
            methodRegistry(methodName)("Settings")(settingName)("Value") = methodRegistry(methodName)("Settings")(settingName)("Floor")
        End If
    End If

    If floor = 0 And ceiling = 1 And VarType(methodSubSettingDictionary("Value")) = vbLong Then
        methodSubSettingDictionary("Default") = CBool(methodSubSettingDictionary("Default"))
        methodSubSettingDictionary("Floor") = CBool(methodSubSettingDictionary("Floor"))
        methodSubSettingDictionary("Ceiling") = CBool(methodSubSettingDictionary("Ceiling"))
        methodSubSettingDictionary("Value") = CBool(methodSubSettingDictionary("Value"))
    End If
End Sub

Sub ConfigureLogging()
    If WorksheetExists("Log") = False Then
        Call LogCheckpoint("Foundation", "Launch", "Beginning")

        If mainWorkbook.Sheets("Log").Range("A1").Style <> "Header" Then
            With Sheets("Log").Range("A1:" & LastColumnLetterOnWorksheet("Log") & "1")
                .Style = "Header"
                .Value = .Value
            End With
        End If

        If CellStyleExists("Bad") And CellStyleExists("Good") Then
            Call RemoveCellStyle("Bad|Good|Neutral|Calculation|Check Cell|Explanatory Text|Input|Linked Cell|Note|Output|Warning Text|Heading 1|Heading 2|Heading 3|Heading 4|Title|Total|20% - Accent1|20% - Accent2|20% - Accent3|20% - Accent4|20% - Accent5|20% - Accent6|40% - Accent1|40% - Accent2|40% - Accent3|40% - Accent4|40% - Accent5|40% - Accent6|60% - Accent1|60% - Accent2|60% - Accent3|60% - Accent4|60% - Accent5|60% - Accent6|Accent1|Accent2|Accent3|Accent4|Accent5|Accent6|Comma|Comma [0]|Currency|Currency [0]")
        End If

        Call ApplyCellStyleToRangeOnWorksheet("Integer", ColumnRangeTypeOnWorksheet("Original Order", "Full", "Log"), "Log") ' Original Order '
        Call ApplyCellStyleToRangeOnWorksheet("Date Time", ColumnRangeTypeOnWorksheet("Date Time Start", "Full", "Log"), "Log") ' Date Time Start '
        Call ApplyCellStyleToRangeOnWorksheet("Date Time", ColumnRangeTypeOnWorksheet("Date Time End", "Full", "Log"), "Log") ' Date Time End '
        Call ApplyCellStyleToRangeOnWorksheet("Decimal", ColumnRangeTypeOnWorksheet("Stopwatch Start", "Full", "Log"), "Log") 'Stopwatch Start '
        Call ApplyCellStyleToRangeOnWorksheet("Decimal", ColumnRangeTypeOnWorksheet("Stopwatch End", "Full", "Log"), "Log") ' Stopwatch End '
        Call ApplyCellStyleToRangeOnWorksheet("Formula", ColumnRangeTypeOnWorksheet("K", "Full", "Log"), "Log") ' Custom: Date Time Calculation (Formula) '

        Call LogCheckpoint("Foundation", "Launch", "Conclusion")
    End If
End Sub

Sub ConfigureStyles()
    Dim newCellStyles() As String: newCellStyles = Split("Header|Date|Date Time|Decimal|Followed Hyperlink|Integer|Formula", "|")
    Dim index As Byte

    For index = 0 To UBound(newCellStyles)
        If CellStyleExists(newCellStyles(index)) = False Then
            mainWorkbook.Styles.Add Name:=newCellStyles(index)

            With mainWorkbook.Styles(newCellStyles(index))
                .IncludeAlignment = True
                .IncludeNumber = True
                .IncludePatterns = False
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .ReadingOrder = xlContext
            End With

            If newCellStyles(index) = "Decimal" Then
                mainWorkbook.Styles("Decimal").NumberFormat = "General"
                mainWorkbook.Styles("Decimal").NumberFormat = "0.00"
            End If

            If newCellStyles(index) = "Date" Then
                mainWorkbook.Styles("Date").NumberFormat = "dd/mm/yyyy"
            End If

            If newCellStyles(index) = "Date Time" Then
                mainWorkbook.Styles("Date Time").NumberFormat = "dd/mm/yyyy HH:mm:ss"
            End If

            If newCellStyles(index) = "Followed Hyperlink" Then
                With mainWorkbook.Styles("Followed Hyperlink")
                    .IncludeFont = False
                End With
                mainWorkbook.Styles("Followed Hyperlink").NumberFormat = "@"
                With mainWorkbook.Styles("Followed Hyperlink").Font
                    .ColorIndex = xlAutomatic
                End With
            End If

            If newCellStyles(index) = "Header" Then
                With mainWorkbook.Styles("Header")
                    .IncludeFont = True
                    .WrapText = True
                End With
                mainWorkbook.Styles("Header").NumberFormat = "@"
                With mainWorkbook.Styles("Header").Font
                    .Bold = True
                    .ColorIndex = xlAutomatic
                End With
            End If

            If newCellStyles(index) = "Integer" Then
                With mainWorkbook.Styles("Integer")
                    .IncludeFont = True
                End With
                mainWorkbook.Styles("Integer").NumberFormat = "0"
                With mainWorkbook.Styles("Integer").Font
                    .Name = "Consolas"
                    .ColorIndex = xlAutomatic
                End With
            End If

            If newCellStyles(index) = "Formula" Then
                mainWorkbook.Styles("Formula").NumberFormat = "General"
            End If
        End If
    Next index

    If mainWorkbook.Styles("Percent").NumberFormat = "0%" Then
        With mainWorkbook.Styles("Percent")
            .IncludeNumber = True
            .IncludeAlignment = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ReadingOrder = xlContext
        End With
        mainWorkbook.Styles("Percent").NumberFormat = "0.00 %"

        With mainWorkbook.Styles("Percent").Font
            .Name = "Arial"
            .Size = 10
        End With
        With mainWorkbook.Styles("Percent")
            .IncludeFont = False
        End With
    End If
End Sub

Sub CreateAboutWorksheet()
    Dim worksheetEntry As Worksheet
    Dim legacyTemplateVersion As Boolean

    If WorksheetExists("About") = False Then
        mainWorkbook.Sheets.Add After:=mainWorkbook.ActiveSheet
        mainWorkbook.ActiveSheet.Name = "About"

        If Application.DisplayAlerts = True Then
            Application.DisplayAlerts = False
        End If

        For Each worksheetEntry In mainWorkbook.Worksheets
            If WorksheetIsEmpty(worksheetEntry.Name) Then
                If worksheetEntry.Name <> "About" Then
                    mainWorkbook.Sheets(worksheetEntry.Name).Delete
                End If
            End If
        Next worksheetEntry

        Application.DisplayAlerts = True
    Else
        If mainWorkbook.Sheets("About").Visible = False Then
            mainWorkbook.Sheets("About").Visible = True
        End If
    End If

    If NamedRangeExists("ReportDetails") = True Then mainWorkbook.Names("ReportDetails").Delete
    If NamedRangeExists("TemplateDetails") = True Then mainWorkbook.Names("TemplateDetails").Delete
    If NamedRangeExists("ScriptDuration") = True Then mainWorkbook.Names("ScriptDuration").Delete

    If NamedRangeExists("ReportName") = False Then mainWorkbook.Names.Add Name:="ReportName", RefersToR1C1:="=About!R1C1"
    If NamedRangeExists("TemplateVersion") = False Then mainWorkbook.Names.Add Name:="TemplateVersion", RefersToR1C1:="=About!R2C1"
    If NamedRangeExists("ProgressionStatus") = False Then mainWorkbook.Names.Add Name:="ProgressionStatus", RefersToR1C1:="=About!R3C1"
    If NamedRangeExists("AugmentationModules") = False Then mainWorkbook.Names.Add Name:="AugmentationModules", RefersToR1C1:="=About!R4C1"
    If NamedRangeExists("ReportVision") = False Then mainWorkbook.Names.Add Name:="ReportVision", RefersToR1C1:="=About!R1C2"
    If NamedRangeExists("DependenciesList") = False Then mainWorkbook.Names.Add Name:="DependenciesList", RefersToR1C1:="=About!R3C2"
    If NamedRangeExists("RetrievedDate") = False Then mainWorkbook.Names.Add Name:="RetrievedDate", RefersToR1C1:="=About!R1C3"
    If NamedRangeExists("EditionName") = False Then mainWorkbook.Names.Add Name:="EditionName", RefersToR1C1:="=About!R2C3"
    If NamedRangeExists("DurationMilliseconds") = False Then mainWorkbook.Names.Add Name:="DurationMilliseconds", RefersToR1C1:="=About!R3C3"
    If NamedRangeExists("LogSummary") = False Then mainWorkbook.Names.Add Name:="LogSummary", RefersToR1C1:="=About!R4C3"

    If report.Exists("Name") Then
        If mainWorkbook.Sheets("About").Range("ReportName").Value = "" Then mainWorkbook.Sheets("About").Range("ReportName").Value = report("Name")
    Else
        If mainWorkbook.Sheets("About").Range("ReportName").Value = "" Then mainWorkbook.Sheets("About").Range("ReportName").Value = "N/A"
    End If

    If mainWorkbook.Sheets("About").Range("TemplateVersion").Value = "Spreadsheet Operations Template (v0.39, 16.02.2024)" Then
        legacyTemplateVersion = True
    End If

    mainWorkbook.Sheets("About").Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & report("Template Version") & ")"

    If mainWorkbook.Sheets("About").Range("ProgressionStatus").Value = "" Then mainWorkbook.Sheets("About").Range("ProgressionStatus").Value = "Progression Status: N/A."
    If mainWorkbook.Sheets("About").Range("AugmentationModules").Value = "" Then mainWorkbook.Sheets("About").Range("AugmentationModules").Value = "Augmentation Modules: N/A."

    If report.Exists("Vision") Then
        If mainWorkbook.Sheets("About").Range("ReportVision").Value = "" Then mainWorkbook.Sheets("About").Range("ReportVision").Value = report("Vision")
    Else
        If mainWorkbook.Sheets("About").Range("ReportVision").Value = "" Then mainWorkbook.Sheets("About").Range("ReportVision").Value = "N/A."
    End If

    If report.Exists("Dependencies") Then
        If mainWorkbook.Sheets("About").Range("DependenciesList").Value = "" Then mainWorkbook.Sheets("About").Range("DependenciesList").Value = "Dependencies List: " & report("Dependencies") & "."
    Else
        If mainWorkbook.Sheets("About").Range("DependenciesList").Value = "" Then mainWorkbook.Sheets("About").Range("DependenciesList").Value = "Dependencies List: N/A."
    End If

    If report.Exists("Retrieved Date") Then
        If mainWorkbook.Sheets("About").Range("RetrievedDate").Value = "" Then mainWorkbook.Sheets("About").Range("RetrievedDate").Value = "Retrieved Date: " & report("Retrieved Date") & "."
    End If

    If report.Exists("Edition") Then
        If mainWorkbook.Sheets("About").Range("EditionName").Value = "" Then mainWorkbook.Sheets("About").Range("EditionName").Value = "Edition Name: " & report("Edition") & "."
    Else
        If mainWorkbook.Sheets("About").Range("EditionName").Value = "" Then mainWorkbook.Sheets("About").Range("EditionName").Value = "Edition Name: N/A."
    End If

    If mainWorkbook.Sheets("About").Range("DurationMilliseconds").Value = "" Then mainWorkbook.Sheets("About").Range("DurationMilliseconds").Value = "Duration (Milliseconds): N/A."
    If mainWorkbook.Sheets("About").Range("LogSummary").Value = "" Then mainWorkbook.Sheets("About").Range("LogSummary").Value = "Log Summary: N/A."

    If legacyTemplateVersion = True Then
        Dim legacyRetrievedDate As String

        legacyRetrievedDate = GetAboutNamedRange("Retrieved Date")
        Call SetAboutNamedRange(Right(legacyRetrievedDate, 4) & "-" & Mid(legacyRetrievedDate, 4, 2) & "-" & Left(legacyRetrievedDate, 2), "Retrieved Date")

        mainWorkbook.Sheets("About").Range("DurationMilliseconds").Value = "Duration (Milliseconds): N/A."
        mainWorkbook.Sheets("About").Range("LogSummary").Value = "Log Summary: N/A."
    End If

    With mainWorkbook.Sheets("About")
        .Cells.RowHeight = 16.5
        .Rows("1:4").RowHeight = 33.0
        .Columns("A:B").ColumnWidth = 90.71
        .Columns("C").ColumnWidth = 27.29
        .Columns("D").ColumnWidth = 4.0

        .Range("A1:C1").WrapText = True
        .Range("B2:C2").WrapText = True
        .Range("A3:C4").WrapText = True

        .Range("B3:B4").VerticalAlignment = xlTop
    End With

    If mainWorkbook.Worksheets("About").Range("B2").MergeArea.Cells.Count = 1 Then
        mainWorkbook.Worksheets("About").Range("B1:B2").Merge
    End If

    If mainWorkbook.Worksheets("About").Range("B4").MergeArea.Cells.Count = 1 Then
        mainWorkbook.Worksheets("About").Range("B3:B4").Merge
    End If
End Sub

Sub Intermission(ByVal intermissionStates As String, ByVal checkpointName As String)
    If intermissionStates = "" Then Exit Sub

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' https://learn.microsoft.com/en-us/office/vba/api/excel.range.replace
    mainWorkbook.Sheets("About").Range("A1").Replace What:="", Replacement:="", LookAt:=xlPart

    Dim intermissionState As String
    Dim intermissionsArray() As String: intermissionsArray = Split(intermissionStates, "|")

    Dim index As Byte
    For index = 0 To UBound(intermissionsArray)
        intermissionState = intermissionsArray(index)

        Select Case intermissionState
            Case "Break Script", "Break", "BS", "B"
                End
            Case "Covert Mode", "Covert", "CM", "C"
                Dim aboutNamedRanges As String: aboutNamedRanges = "AugmentationModules|DependenciesList|DurationMilliseconds|EditionName|LogSummary|ProgressionStatus|ReportDetails|ReportName|ReportVision|RetrievedDate|ScriptDuration|TemplateDetails|TemplateVersion"
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
                    mainWorkbook.Sheets("Log").Visible = True
                    mainWorkbook.Sheets("Log").Delete
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
                    If mainWorkbook.Sheets(indexResetView).Name <> "Log" Then
                        worksheetIsHidden = False
                        If mainWorkbook.Sheets(indexResetView).Visible = xlSheetHidden Then worksheetIsHidden = True

                        If worksheetIsHidden = True Then
                            mainWorkbook.Sheets(indexResetView).Visible = xlSheetVisible
                        End If

                        mainWorkbook.Sheets(indexResetView).Select
                        Application.Goto Reference:=ActiveSheet.Cells.SpecialCells(xlCellTypeVisible).Range("A1"), Scroll:=True

                        If worksheetIsHidden = True Then
                            mainWorkbook.Sheets(indexResetView).Visible = xlSheetHidden
                        End If
                    End If
                Next indexResetView

                mainWorkbook.Activate
                mainWorkbook.Sheets("About").Select
                mainWorkbook.Sheets("About").Activate
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

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
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
                mainWorkbook.Sheets("About").Range("ReportDetails").Value = namedRangeValue
            ElseIf NamedRangeExists("ReportName") Then
                mainWorkbook.Sheets("About").Range("ReportName").Value = namedRangeValue
            End If
        Case "TemplateDetails", "Template Details", "TemplateVersion", "Template Version", "Template", "Version"
            If NamedRangeExists("TemplateDetails") Then
                mainWorkbook.Sheets("About").Range("TemplateDetails").Value = "Spreadsheet Operations Template (" & namedRangeValue & ")"
            ElseIf NamedRangeExists("TemplateVersion") Then
                mainWorkbook.Sheets("About").Range("TemplateVersion").Value = "Spreadsheet Operations Template (" & namedRangeValue & ")"
            End If
        Case "ProgressionStatus", "Progression Status", "Progression"
            currentNamedRangeValue = GetAboutNamedRange("Progression Status")

            If overwrite = True Or currentNamedRangeValue = "N/A" Then
                mainWorkbook.Sheets("About").Range("ProgressionStatus").Value = "Progression Status: " & namedRangeValue & "."
            Else
                mainWorkbook.Sheets("About").Range("ProgressionStatus").Value = "Progression Status: " & currentNamedRangeValue & ", " & namedRangeValue & "."
            End If
        Case "AugmentationModules", "Augmentation Modules", "Augmentation"
            currentNamedRangeValue = GetAboutNamedRange("Augmentation Modules")

            If overwrite = True Or currentNamedRangeValue = "N/A" Then
                mainWorkbook.Sheets("About").Range("AugmentationModules").Value = "Augmentation Modules: " & namedRangeValue & "."
            Else
                mainWorkbook.Sheets("About").Range("AugmentationModules").Value = "Augmentation Modules: " & currentNamedRangeValue & ", " & namedRangeValue & "."
            End If
        Case "ReportVision", "Report Vision", "Vision"
            mainWorkbook.Sheets("About").Range("ReportVision").Value = namedRangeValue
        Case "DependenciesList", "Dependencies List", "Dependencies"
            currentNamedRangeValue = GetAboutNamedRange("Dependencies List")

            If overwrite = True Or currentNamedRangeValue = "N/A" Then
                mainWorkbook.Sheets("About").Range("DependenciesList").Value = "Dependencies List: " & namedRangeValue & "."
            Else
                mainWorkbook.Sheets("About").Range("DependenciesList").Value = "Dependencies List: " & currentNamedRangeValue & ", " & namedRangeValue & "."
            End If
        Case "RetrievedDate", "Retrieved Date", "Retrieved"
            mainWorkbook.Sheets("About").Range("RetrievedDate").Value = "Retrieved Date: " & namedRangeValue & "."
        Case "EditionName", "Edition Name", "Edition"
            mainWorkbook.Sheets("About").Range("EditionName").Value = "Edition Name: " & namedRangeValue & "."
        Case "ScriptDuration", "Script Duration", "Duration (Milliseconds)", "Duration Milliseconds", "Duration"
            If NamedRangeExists("ScriptDuration") Then
                currentNamedRangeValue = GetAboutNamedRange("Script Duration")
            ElseIf NamedRangeExists("DurationMilliseconds") Then
                currentNamedRangeValue = GetAboutNamedRange("Duration (Milliseconds)")
            End If

            If overwrite = True Or currentNamedRangeValue = "N/A" Then
                mainWorkbook.Sheets("About").Range("DurationMilliseconds").Value = "Duration (Milliseconds): " & namedRangeValue & "."
            Else
                mainWorkbook.Sheets("About").Range("DurationMilliseconds").Value = "Duration (Milliseconds): " & (CDbl(currentNamedRangeValue) + CDbl(namedRangeValue)) & "."
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

            mainWorkbook.Sheets("About").Range("LogSummary").Value = "Log Summary: " & runsSummary & checkpointsSummary & logRowsSummary
        Case Else
            ' Call LogAssertFailure("Core named range value of  """ & coreNamedRange & """" & " is invalid. Valid values: Dependencies List, Report Details, Report Vision, Edition Name, Retrieved Date.")
    End Select
End Sub

Sub StartLoggingEngine(ByVal usernameValue As String, ByVal versionValue As String, ByVal filePath As String)
    stopwatchTimer = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    If WorksheetExists("Log") Then
        Sheets("Log").Visible = True
    Else
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Name = "Log"

        Dim logColumnNames() As String: logColumnNames = Split("Task|Arguments|Category|Status|Active Worksheet|Original Order|Date Time Start|Date Time End|Stopwatch Start|Stopwatch End", "|")
        Dim logColumnWidths() As String: logColumnWidths = Split("81,86|69,86|10,71|9,00|26,71|7,57|17,00|17,00|9,86|9,86", "|")
        Dim index As Byte

        For index = 0 To UBound(logColumnNames)
            Range(ConvertNumberToLetter(index + 1) & "1").Value = logColumnNames(index)
            Range(ConvertNumberToLetter(index + 1) & "1").ColumnWidth =  CDbl(logColumnWidths(index))
        Next index
    End If

    If Sheets("Log").Range("A2").Value <> "" Then
        ' logOrder = LastRowNumberOnWorksheet("Log") + 1
    Else
        ' logOrder = 2
    End If

    ' genesisLog = logOrder

    If Not Sheets("Log").AutoFilterMode Then
        Sheets("Log").Range("A1:" & LastColumnLetterOnWorksheet("Log") & "1").AutoFilter

        Sheets("Log").Cells.Borders.LineStyle = xlLineStyleNone
        Sheets("Log").Cells.Interior.Pattern = xlNone
        Sheets("Log").Cells.Interior.TintAndShade = 0
        Sheets("Log").Cells.Interior.PatternTintAndShade = 0
        Sheets("Log").Cells.RowHeight = 16.5
        Sheets("Log").Rows("1:1").RowHeight = 49.5
        Sheets("Log").Cells.Style = "Normal"

        mainWorkbook.Activate
        mainWorkbook.Sheets("Log").Select

        ActiveWindow.ActivateNext
        ActiveWindow.WindowState = xlMaximized

        Dim numericalModeArray() As String: numericalModeArray = Split("0,1", ",")

        With ActiveWindow
            .SplitColumn = CInt(numericalModeArray(0))
            .SplitRow = CInt(numericalModeArray(1))
        End With
        ActiveWindow.FreezePanes = True
    Else
        mainWorkbook.Activate
        mainWorkbook.Sheets("Log").Select

        ActiveWindow.ActivateNext
        ActiveWindow.WindowState = xlMaximized
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ResetLoggingEngine(ByVal checkpointType As String, ByVal checkpointName As String)
    Dim lastRowReset As Long

    Sheets("Log").Select
    lastRowReset = LastRowNumberOnWorksheet("Log")

    ' Custom: Date Time Calculation (Formula) '

    Dim dateTimeConversion() As String: dateTimeConversion = Split("G|H", "|")
    Dim index As Long 

    Call LogConclusion("Completed", logConclusionData)
    For index = 0 To UBound(dateTimeConversion)
        Range("K" & genesisLog).FormulaLocal = ConvertDateTimeToSerial(dateTimeConversion(index), genesisLog)
        Sheets("Log").Range("K" & genesisLog & ":K" & lastRowReset).FillDown
        Sheets("Log").Range(dateTimeConversion(index) & genesisLog & ":" & dateTimeConversion(index) & lastRowReset).Value = Sheets("Log").Range("K" & genesisLog & ":K" & lastRowReset).Value
    Next index
    
    Sheets("Log").Columns("K:K").EntireColumn.ClearContents

    Dim lastRowCheckpoint As Long
    lastRowCheckpoint = LastRowNumberOnWorksheet("Log")
    Range("LogSummary").Value = "Log Summary: " & lastRowCheckpoint -1 & " rows."

    If Range("ScriptDuration").Value <> "Script Duration: N/A." Then
        Range("ScriptDuration").Value = Left(Range("ScriptDuration").Value, InStr(1, Range("ScriptDuration").Value, " seconds.") - 1)
        Range("ScriptDuration").Value = Right(Range("ScriptDuration").Value, Len(Range("ScriptDuration").Value) - 17)
        Dim currentDuration As Double
        currentDuration = CDbl(Range("ScriptDuration").Value)
        Range("ScriptDuration").Value = "Script Duration: " & currentDuration + Round(Timer - stopwatchTimer, 2) & " seconds."
    Else
        Range("ScriptDuration").Value = "Script Duration: " & Round(Timer - stopwatchTimer, 2) & " seconds."
    End If

    Worksheets("Log").visible = xlSheetHidden
    Sheets("About").Select

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
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
                namedRangeValue = mainWorkbook.Sheets("About").Range("ReportDetails").Value
            ElseIf NamedRangeExists("ReportName") Then
                namedRangeValue = mainWorkbook.Sheets("About").Range("ReportName").Value
            End If
        Case "TemplateDetails", "Template Details", "TemplateVersion", "Template Version", "Template", "Version"
            If NamedRangeExists("TemplateDetails") Then
                namedRangeValue = mainWorkbook.Sheets("About").Range("TemplateDetails").Value
            ElseIf NamedRangeExists("TemplateVersion") Then
                namedRangeValue = mainWorkbook.Sheets("About").Range("TemplateVersion").Value
            End If

            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, "(") + 1)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "ProgressionStatus", "Progression Status", "Progression"
            namedRangeValue = mainWorkbook.Sheets("About").Range("ProgressionStatus").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "AugmentationModules", "Augmentation Modules", "Augmentation"
            namedRangeValue = mainWorkbook.Sheets("About").Range("AugmentationModules").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "ReportVision", "Report Vision", "Vision"
            namedRangeValue = mainWorkbook.Sheets("About").Range("ReportVision").Value
        Case "DependenciesList", "Dependencies List", "Dependencies"
            namedRangeValue = mainWorkbook.Sheets("About").Range("DependenciesList").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "RetrievedDate", "Retrieved Date", "Retrieved"
            namedRangeValue = mainWorkbook.Sheets("About").Range("RetrievedDate").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "EditionName", "Edition Name", "Edition"
            namedRangeValue = mainWorkbook.Sheets("About").Range("EditionName").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "ScriptDuration", "Script Duration", "Duration (Milliseconds)", "Duration Milliseconds", "Duration"
            namedRangeValue = mainWorkbook.Sheets("About").Range("DurationMilliseconds").Value
            namedRangeValue = Mid(namedRangeValue, InStr(namedRangeValue, ":") + 2)
            namedRangeValue = Left(namedRangeValue, Len(namedRangeValue) - 1)
        Case "LogSummary", "Log Summary", "Summary"
            namedRangeValue = mainWorkbook.Sheets("About").Range("LogSummary").Value
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
    Const methodName As String = "CloseGuestWorkbook"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, "")

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
    Const methodName As String = "CreateBlankWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    If worksheetName = "" Then worksheetName = ActiveSheet.Name & "!"
    If WorksheetExists(worksheetName) Then Call LogConclusion("Failed", logConclusionData, "Worksheet """ & worksheetName & """ already exists.")

    Sheets.Add.Name = worksheetName
    
    Call NormalizeLayoutOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub CreateFileListArrayFromDirectory(ByRef fileListArray As Variant, ByVal folderPath As String)
    Const methodName As String = "CreateFileListArrayFromDirectory"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByRef fileListArray As Variant, ByVal folderPath As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, "[" & "fileListArray" & "]" & " ," & """" & folderPath & """")

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
    Const methodName As String = "DuplicateColumnFromWorksheetToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & fromWorksheetName & """" & ", " & """" & toWorksheetName & """")

    ' Remember to sort the two primary columns prior to using this method, or there will be errors unless they are sorted from before. '

    Dim cellStyleToCopy As String
    Dim widthToCopy As Double
    Dim originalColumnName As String
    originalColumnName = columnName

    Dim validation As String
    Call ValidateWorksheet(fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateWorksheet(toWorksheetName, "toWorksheetName", validation)
    If ColumnOnWorksheetExists(columnName, fromWorksheetName) = False Then Call LogConclusion("Failed", logConclusionData, "Column " & """" & columnName & """" & " on worksheet " & """" & fromWorksheetName & """" & " does not exist.")
    Call ValidateColumnOnWorksheet(columnName, fromWorksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim sourceColumnName As String
    Dim matchColumnName As String

    sourceColumnName = mainWorkbook.Sheets(toWorksheetName).Range("A1").Value
    matchColumnName = mainWorkbook.Sheets(fromWorksheetName).Range("A1").Value

    Dim deleteDuplicateColumn As Boolean
    deleteDuplicateColumn = False

    If sourceColumnName = "" Then
        Call InsertNewColumnAndSetWidthOnWorksheet("Sorting Column", 7, fromWorksheetName)
        Call ApplyFormulaToRangeOnWorksheet("=ROW()-1", "Sorting Column", fromWorksheetName)
        Call SortColumnByOrderOnWorksheet(matchColumnName, "Ascending", fromWorksheetName)
        Sheets(fromWorksheetName).Range("A1:A" & LastRowNumberOnWorksheet(fromWorksheetName)).Copy
        Sheets(toWorksheetName).Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        mainWorkbook.Sheets(toWorksheetName).Range("A1").Value = "Temporary"
        sourceColumnName = "Temporary"
        deleteDuplicateColumn = True
    End If

    cellStyleToCopy = mainWorkbook.Worksheets(fromWorksheetName).Range(columnName & "2").Style
    widthToCopy = mainWorkbook.Worksheets(fromWorksheetName).Range(columnName & "2").ColumnWidth

    If ColumnOnWorksheetExists(columnName, toWorksheetName) = False Then Call InsertNewColumnAndSetWidthOnWorksheet(originalColumnName, widthToCopy, toWorksheetName)

    Dim lookupValue As String
    lookupValue = "XLOOKUP(" & ColumnStartOnWorksheet(sourceColumnName, toWorksheetName) & ";" & ColumnDataOnWorksheet(matchColumnName, fromWorksheetName) & ";" & ColumnDataOnWorksheet(originalColumnName, fromWorksheetName)

    Call ApplyFormulaToRangeOnWorksheet("=IF(" & lookupValue & ";"""";0;2)<>"""";" & lookupValue & ";"""";0;2);"""")", originalColumnName, toWorksheetName)
    ' [search_mode]: 2 => Perform a binary search that relies on lookup_array being sorted in ascending order. If not sorted, invalid results will be returned. '
    Call ApplyCellStyleToRangeOnWorksheet(cellStyleToCopy, originalColumnName, toWorksheetName)

    If deleteDuplicateColumn = True Then Call DeleteColumnOnWorksheet("Temporary", toWorksheetName)

    Sheets(fromWorksheetName).Range(FindColumnLetterOnWorksheet(originalColumnName, fromWorksheetName) & "1").Copy
    Sheets(toWorksheetName).Range(FindColumnLetterOnWorksheet(originalColumnName, toWorksheetName) & "1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    If deleteDuplicateColumn = True Then
        Call SortColumnByOrderOnWorksheet("Sorting Column", "Ascending", fromWorksheetName)
        Call DeleteColumnOnWorksheet("Sorting Column", fromWorksheetName)
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DuplicateHeadersFromWorksheetToName(ByVal currentWorksheetName As String, ByVal toWorksheetName As String)
    Const methodName As String = "DuplicateHeadersFromWorksheetToName"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal toWorksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & toWorksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(currentWorksheetName, "currentWorksheetName", validation)
    Call ValidateWorksheet(toWorksheetName, "toWorksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If Sheets(toWorksheetName).Range("A1").Value <> "" Then Call LogConclusion("Failed", logConclusionData, "Worksheet """ & toWorksheetName & """ is not blank.")

    Sheets(currentWorksheetName).Range("A1:" & LastColumnLetterOnWorksheet(currentWorksheetName) & "1").Copy
    Sheets(toWorksheetName).Range("A1").PasteSpecial xlPasteAllUsingSourceTheme
    Sheets(toWorksheetName).Range("A1").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    Call ApplyAutoFilterOnWorksheet(toWorksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub DuplicateWorksheetToName(ByVal currentWorksheetName As String, ByVal newWorksheetName As String)
    Const methodName As String = "DuplicateWorksheetToName"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal currentWorksheetName As String, ByVal newWorksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & currentWorksheetName & """" & ", " & """" & newWorksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(currentWorksheetName, "currentWorksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If currentWorksheetName = newWorksheetName Then newWorksheetName = currentWorksheetName & "!"

    Sheets(currentWorksheetName).Copy After:=Sheets(Worksheets.Count)
    ActiveSheet.Name = newWorksheetName

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ExportWorksheetToBlankWorkbook(ByVal worksheetName As String, ByVal filePath As String) ' Plural Support. '
    Const methodName As String = "ExportWorksheetToBlankWorkbook"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String, ByVal filePath As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """" & ", " & """" & filePath & """")

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

        mainWorkbook.Sheets(filterValueArray(index)).Copy Before:=importWorkbook.Sheets(1)
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
    Const methodName As String = "ImportExcelFileWorksheets"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetsValues As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & filePath & """" & " ," & """" & worksheetsValues & """")

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
            importWorkbook.Sheets(worksheetsArray(index)).Cells.Copy mainWorkbook.Sheets(worksheetsArray(index)).Cells
        Else
            Call LogConclusion("Failed", logConclusionData, "Worksheet " & """" & worksheetsArray(index) & """" & " not found in the workbook " & """" & importWorkbook.Name & """" & ".")
        End If
    Next index

    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ImportExcelFileAndRenameWorksheet(ByVal filePath As String, ByVal worksheetName As String)
    Const methodName As String = "ImportExcelFileAndRenameWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & filePath & """" & " ," & """" & worksheetName & """")

    Call OpenGuestWorkbook(filePath)

    If importWorkbook.Worksheets.Count > 1 Then Call LogConclusion("Failed", logConclusionData, "This method only supports importing one worksheet per workbook, the workbook " & """" & importWorkbook.Name & """" & " contains " & importWorkbook.Worksheets.Count & " worksheets.")

    importWorkbook.Sheets(1).Name = worksheetName
    mainWorkbook.Activate
    If WorksheetExists(worksheetName) Then Call LogConclusion("Failed", logConclusionData, "Worksheet named " & """" & worksheetName & """" & " already exists.")
    Call CreateBlankWorksheet(worksheetName)

    importWorkbook.Sheets(worksheetName).Cells.Copy mainWorkbook.Sheets(worksheetName).Cells
    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ImportTextFileToWorksheet(ByVal filePath As String, ByVal worksheetName As String)
    Const methodName As String = "ImportTextFileToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal worksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & filePath & """" & " ," & """" & worksheetName & """")

    Call OpenGuestWorkbook(filePath)
    mainWorkbook.Activate
    Call CreateBlankWorksheet(worksheetName)
    importWorkbook.Sheets(1).Cells.Copy mainWorkbook.Sheets(worksheetName).Cells
    Call CloseGuestWorkbook()

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub OpenGuestWorkbook(ByVal filePath As String)
    Const methodName As String = "OpenGuestWorkbook"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & filePath & """")

    Call AssertFilePathNotEmptyForVariable(filePath, "importWorkbookFilePath")

    Set importWorkbook = Workbooks.Open(filePath)
    importWorkbook.Activate

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub TransferDataFromWorksheetToWorksheet(ByVal fromWorksheetName As String, toWorksheetName As String)

If InputContainsValue(fromWorksheetName, "|") Then
    Call RepeatTransferDataFromWorksheetToWorksheet(fromWorksheetName, toWorksheetName)
Else
    Const methodName As String = "TransferDataFromWorksheetToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetName As String, toWorksheetName As String", methodName, isRegistered, "Conjuration")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & fromWorksheetName & """" & " ," & """" & toWorksheetName & """")

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

Sub DeleteColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteColumnOnWorksheet(columnName, worksheetName)
Else
    Const methodName As String = "DeleteColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).Columns(columnName & ":" & columnName).EntireColumn.Delete Shift:=xlToLeft

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DeleteRowsOnWorksheet(ByVal numberOfRows As Long, ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteRowsOnWorksheet(numberOfRows, worksheetName)
Else
    Const methodName As String = "DeleteRowsOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal numberOfRows As Long, ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, logConclusionData, numberOfRows & " ," & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    mainWorkbook.Sheets(worksheetName).Rows("1:" & numberOfRows).Delete Shift:=xlUp

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub DeleteWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatDeleteWorksheet(worksheetName)
Else
    Const methodName As String = "DeleteWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).Delete

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RemoveBlankLinesOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatRemoveBlankLinesOnWorksheet(worksheetName)
Else
    Const methodName As String = "RemoveBlankLinesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    If WorksheetExists(worksheetName) Then
        Dim lastRowNumberForWorksheet As String: lastRowNumberForWorksheet = LastRowNumberOnWorksheet(worksheetName)

        mainWorkbook.Sheets(worksheetName).Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        mainWorkbook.Sheets(worksheetName).Range("A1").Value = "Remove Blank Lines Header( Temporary)"
        mainWorkbook.Sheets(worksheetName).Range("A2" & ":" & "A" & lastRowNumberForWorksheet).Value = ";;"
        Call RemoveDataBasedOnFormulaOnWorksheet("=OR(B2="""";CHAR(32)=B2)", worksheetName) ' Code 32 (decimal) is a nonprinting spacing character. (https://web.archive.org/web/20221004120012/http://www.columbia.edu/kermit/ascii.html) '
        mainWorkbook.Sheets(worksheetName).Columns("A:A").Delete Shift:=xlToLeft
    Else
        ' Call LogTaskWarning("Worksheet " & """" & worksheetName & """" & " does not exist.", logOrderLocal)
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RemoveCellStyle(ByVal cellStyleName As String) ' Repeat Support: cellStyleName. '

If InputContainsValue(cellStyleName, "|") Then
    Call RepeatRemoveCellStyle(cellStyleName)
Else
    Const methodName As String = "RemoveCellStyle"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyleName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & cellStyleName & """")

    If CellStyleExists(cellStyleName) Then
        mainWorkbook.Styles(cellStyleName).Delete
    Else
        ' Call LogTaskWarning("Cell style " & """" & cellStyleName & """" & " does not exist.", logOrderLocal)
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub RemoveEmptyColumnsOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatRemoveEmptyColumnsOnWorksheet(worksheetName)
Else
    Const methodName As String = "RemoveEmptyColumnsOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

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
    Const methodName As String = "RemoveDataBasedOnFormulaOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal formulaValue As String, ByVal worksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & formulaValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call ApplyFormulaToRangeOnWorksheet(formulaValue, report("Helper Column"), worksheetName)
    Call ApplyFilterToColumnOnWorksheet(True, report("Helper Column"), worksheetName)

    If NumberOfVisibleCells(worksheetName) <> 0 Then
        Call ClearFilterOnWorksheet(worksheetName)
        Call InsertNewColumnAndSetWidthOnWorksheet("Original Order (Temporary)", 13, worksheetName)
        Call ApplyFormulaToRangeOnWorksheet("=ROW(2:2)-1", "Original Order (Temporary)", worksheetName)
        Call ApplyCellStyleToRangeOnWorksheet("Integer", "Original Order (Temporary)", worksheetName)
        Call ApplyFilterToColumnOnWorksheet(True, report("Helper Column"), worksheetName)
    End If

    If NumberOfVisibleCells(worksheetName) <> 0 Then
        If Sheets(worksheetName).FilterMode Then
            ' This approach is used due to the instability of excel when dealing with large datasets. Just deleting the data can crash the application. '

            Call SelectWorksheet(worksheetName)

            ActiveWindow.ActivateNext
            ActiveWindow.WindowState = xlMaximized

            Dim savedFreezePanesMode As String
            savedFreezePanesMode = ActiveWindow.SplitColumn & ", " & ActiveWindow.SplitRow

            Sheets(worksheetName).Rows("1:1").EntireRow.Hidden = True
            Sheets(worksheetName).Range(ColumnRangeTypeOnWorksheet("A", "Data", worksheetName)).SpecialCells(xlCellTypeConstants, 23).ClearContents ' Header hidden before or bug occurs when only one line gets deleted. '
            Sheets(worksheetName).Rows("1:1").EntireRow.Hidden = False
            Call ClearFilterOnWorksheet(worksheetName)
            Call SortColumnByOrderOnWorksheet("A", "Ascending", worksheetName)
            Sheets(worksheetName).Name = "!"
            Call CreateBlankWorksheet(worksheetName)
            Sheets("!").Range("A1:" & LastColumnLetterOnWorksheet("!") & LastRowNumberOnWorksheet("!")).Copy Sheets(worksheetName).Range("A1")
            Sheets("!").Range("A1:" & LastColumnLetterOnWorksheet("!") & "1").Copy
            Sheets(worksheetName).Range("A1").PasteSpecial xlPasteColumnWidths
            Application.CutCopyMode = False
            Call ApplyAutoFilterOnWorksheet(worksheetName)
            Call SortColumnByOrderOnWorksheet("Original Order (Temporary)", "Ascending", worksheetName)
            Call DeleteColumnOnWorksheet("Original Order (Temporary)", worksheetName)

            If savedFreezePanesMode <> "0, 0" Then Call ApplyFreezePanesModeOnWorksheet(savedFreezePanesMode, worksheetName)

            If Sheets("!").Tab.Color <> False Then
                mainWorkbook.Sheets(worksheetName).Tab.Color = Sheets("!").Tab.Color
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
    Const methodName As String = "SynchronizeWorksheetOnColumnFromWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal synchronizeWorksheetName As String, ByVal columnName As String, ByVal fromWorksheetName As String", methodName, isRegistered, "Destruction")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & synchronizeWorksheetName & """" & ", " & """" & columnName & """" & ", " & """" & fromWorksheetName & """")

    Dim synchronizeWorksheetColumnName As String
    Dim fromWorksheetColumnName As String

    synchronizeWorksheetColumnName = columnName
    fromWorksheetColumnName = columnName

    Dim validation As String
    Call ValidateWorksheet(synchronizeWorksheetName, "synchronizeWorksheetName", validation)
    Call ValidateWorksheet(fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateColumnOnWorksheet(synchronizeWorksheetColumnName, synchronizeWorksheetName, validation)
    Call ValidateColumnOnWorksheet(fromWorksheetColumnName, fromWorksheetName, validation)
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
    Const methodName As String = "ApplyFilterToColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filterValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & filterValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

If InputContainsValue(filterValue, "|") Then
    Dim filterValueArray() As String: filterValueArray = Split(filterValue, "|")

    Sheets(worksheetName).Range(columnName & "1:" & columnName & "1").AutoFilter Field:=ConvertLetterToNumber(columnName), Criteria1:=filterValueArray, Operator:=xlFilterValues
Else
    Sheets(worksheetName).Range(columnName & "1:" & columnName & "1").AutoFilter Field:=ConvertLetterToNumber(columnName), Criteria1:=filterValue, Operator:=xlAnd
End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ClearFilterOnWorksheet(ByVal worksheetName As String)
    Const methodName As String = "ClearFilterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).AutoFilter.Sort.SortFields.Clear
    If Sheets(worksheetName).AutoFilterMode Then Sheets(worksheetName).AutoFilter.ShowAllData

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub CutColumnAndPasteAtColumnOnWorksheet(ByVal cutColumnName As String, ByVal pasteColumnName As String, ByVal worksheetName As String) ' Repeat Support: cutColumnName. '

If InputContainsValue(cutColumnName, "|") Then
    Call RepeatCutColumnAndPasteAtColumnOnWorksheet(cutColumnName, pasteColumnName, worksheetName)
Else
    Const methodName As String = "CutColumnAndPasteAtColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cutColumnName As String, ByVal pasteColumnName As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & cutColumnName & """" & ", " & """" & pasteColumnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(cutColumnName, worksheetName, validation)
    Call ValidateColumnOnWorksheet(pasteColumnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).Columns(cutColumnName & ":" & cutColumnName).Cut
    Sheets(worksheetName).Columns(pasteColumnName & ":" & pasteColumnName).Insert Shift:=xlToRight

    Call ResetAutoFilterOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub SelectWorksheet(ByVal worksheetName As String)
    Const methodName As String = "SelectWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If mainWorkbook.ActiveSheet.Name = worksheetName Then
        Call LogConclusion("Skipped", logConclusionData)
    Else
        mainWorkbook.Sheets(worksheetName).Select
        Call LogConclusion("Completed", logConclusionData)
    End If
End Sub

Sub SortColumnByOrderOnWorksheet(ByVal columnName As String, ByVal sortOrder As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatSortColumnByOrderOnWorksheet(columnName, sortOrder, worksheetName)
Else
    Const methodName As String = "SortColumnByOrderOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal sortOrder As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & sortOrder & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If sortOrder <> "Ascending" And sortOrder <> "Descending" Then
        sortOrder = "Ascending"
        ' Call LogTaskWarning("Sorting order " & """" & sortOrder & """" & " not Ascending or Descending, setting to Ascending.", logOrderLocal)
    End If
      
    Call SelectWorksheet(worksheetName) ' Prevents selection bug. '
    If Not Sheets(worksheetName).AutoFilterMode Then Call LogConclusion("Failed", logConclusionData, "Worksheet """ & worksheetName & """ does not have a filter applied.")

    If sortOrder ="Ascending" Then Sheets(worksheetName).AutoFilter.Sort.SortFields.Add Key:=Range(columnName & "1:" & columnName & "1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    If sortOrder ="Descending" Then Sheets(worksheetName).AutoFilter.Sort.SortFields.Add Key:=Range(columnName & "1:" & columnName & "1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

    With Sheets(worksheetName).AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Sheets(worksheetName).AutoFilter.Sort.SortFields.Clear

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ManualCodeSection(ByVal sectionName As String, ByVal startOrFinish As String)
    Const methodName As String = "ManualCodeSection"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal sectionName As String, ByVal startOrFinish As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & sectionName & """" & ", " & """" & startOrFinish & """")

    If startOrFinish <> "Start" And startOrFinish <> "Finish" Then Call LogConclusion("Failed", logConclusionData, "Only valid values are ""Start"" and ""Finish"", as opposed to input value of: """ & startOrFinish & """.")

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub TextToColumnsModeOnWorksheet(ByVal modeValue As String, ByVal worksheetName As String)
    Const methodName As String = "TextToColumnsModeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal modeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Elementals")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & modeValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If modeValue = "" Then modeValue = "Semicolon"

    If modeValue = "Comma" Then Sheets(worksheetName).Range("A1:A" & LastRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Comma:=True
    If modeValue = "Semicolon" Then Sheets(worksheetName).Range("A1:A" & LastRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Semicolon:=True
    If modeValue = "Tab" Then Sheets(worksheetName).Range("A1:A" & LastRowNumberOnWorksheet(worksheetName)).TextToColumns Destination:=Sheets(worksheetName).Range("A1"), DataType:=xlDelimited, Tab:=True

    Call LogConclusion("Completed", logConclusionData)
End Sub

' Functions: Elementals '

Function ConvertDateTimeToSerial(ByVal columnLetter As String, ByVal startingRow As Long) As String
    ConvertDateTimeToSerial = "=DATEVALUE(" & columnLetter & startingRow & ") + RIGHT(LEFT(" & columnLetter & startingRow & ";13);2)/24 + LEFT(RIGHT(" & columnLetter & startingRow & ";5);2)/1440 + RIGHT(" & columnLetter & startingRow & ";2)/86400"
End Function

Function ConvertLetterToNumber(ByVal letterToConvert As String) As Long
    On Error Resume Next
    ConvertLetterToNumber = Sheets("Log").Range(letterToConvert & 1).Column
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

    Set worksheetColumnRange = Sheets(worksheetName).Range(ColumnRangeTypeOnWorksheet(columnName, "Data", worksheetName))
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

    Set worksheetHeaderRange = mainWorkbook.Sheets(worksheetName).Range("A1:" & LastColumnLetterOnWorksheet(worksheetName) & "1")
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
                If mainWorkbook.Sheets(index).Name = exclusionWorksheetsArray(indexWorksheets) Then
                    includeWorksheet = False
                End If
             Next indexWorksheets
        End If

        If includeWorksheet = True Then
            CreateDelimitedWorksheetStringExcluding = CreateDelimitedWorksheetStringExcluding & mainWorkbook.Sheets(index).Name & "|"
        End If
    Next index

    CreateDelimitedWorksheetStringExcluding = Left(CreateDelimitedWorksheetStringExcluding, Len(CreateDelimitedWorksheetStringExcluding) - 1)
End Function

Function FindColumnLetterOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) As String
    Dim columnRange As Range
    
    If worksheetName = "" Then worksheetName = ActiveSheet.Name

    Set columnRange = Sheets(worksheetName).Range("A1:" & LastColumnLetterOnWorksheet(worksheetName) & "1").Find(columnName, LookIn:=xlValues, Lookat:=xlWhole)
    If columnRange Is Nothing Then Call LogFunctionError("FindColumnLetterOnWorksheet(columnName As String, worksheetName As String) As String", "Column " & """" & columnName & """" & " on worksheet " & """" & worksheetName & """" & " not found.")
    Application.FindFormat.Clear
    FindColumnLetterOnWorksheet = ConvertNumberToLetter(columnRange.Column)
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

Function LastBlankRowNumberOnWorksheet(ByVal worksheetName As String) As Long
    LastBlankRowNumberOnWorksheet = Sheets(worksheetName).Cells(Rows.Count, "A").End(xlUp).Row + 1
End Function

Function LastColumnLetterOnWorksheet(ByVal worksheetName As String) As String
    Dim lastColumnNumber As Long
    lastColumnNumber = LastColumnNumberOnWorksheet(worksheetName)
    LastColumnLetterOnWorksheet = ConvertNumberToLetter(lastColumnNumber)
End Function

Function LastColumnNumberOnWorksheet(ByVal worksheetName As String) As Long
    LastColumnNumberOnWorksheet = Sheets(worksheetName).Cells(1, Columns.Count).End(xlToLeft).Column
End Function

Function LastRowNumberOnWorksheet(ByVal worksheetName As String) As Long
    LastRowNumberOnWorksheet = Sheets(worksheetName).Cells(Rows.Count, "A").End(xlUp).Row
End Function

Function NumberOfVisibleCells(worksheetName As String) As Long
    NumberOfVisibleCells = Sheets(worksheetName).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1
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
    If rangeType = "Data" Then ColumnRangeTypeOnWorksheet = columnName & "$2:" & columnName & "$" & LastRowNumberOnWorksheet(worksheetName)
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
    GetPrimaryColumnHeaderNameOnWorksheet = mainWorkbook.Sheets(worksheetName).Range("A1").Value
End Function

' ************ '
' Formatting   '
' ************ '

Sub ApplyAutoFilterOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatApplyAutoFilterOnWorksheet(worksheetName)
Else
    Const methodName As String = "ApplyAutoFilterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If Not Sheets(worksheetName).AutoFilterMode Then
        Sheets(worksheetName).Range("A1:" & LastColumnLetterOnWorksheet(worksheetName) & "1").AutoFilter
    End If

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ApplyCellStyleToRangeOnWorksheet(ByVal cellStyle As String, ByVal rangeValue As String, ByVal worksheetName As String) ' Repeat Support: rangeValue, worksheetName. '

If InputContainsValue(rangeValue, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatApplyCellStyleToRangeOnWorksheet(cellStyle, rangeValue, worksheetName)
Else
    Const methodName As String = "ApplyCellStyleToRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyle As String, ByVal rangeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & cellStyle & """" & ", " & """" & rangeValue & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If Sheets(worksheetName).Range("A2").Value = "" Then
        rangeValue = FindColumnLetterOnWorksheet(rangeValue, worksheetName) & "2:" & FindColumnLetterOnWorksheet(rangeValue, worksheetName) & "1048576"
    Else
        Call ValidateRangeOnWorksheet(rangeValue, worksheetName, validation)
    End If

    If CellStyleExists(cellStyle) = False Then
        Call LogConclusion("Failed", logConclusionData, "Cell style """ & cellStyle & """ does not exist.")
    End If

    With Sheets(worksheetName).Range(rangeValue)
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
    Const methodName As String = "ApplyFreezePanesModeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal freezePanesMode As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & freezePanesMode & """" & ", " & """" & worksheetName & """")

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

Sub ApplyHeaderStyleOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatApplyHeaderStyleOnWorksheet(worksheetName)
Else
    Const methodName As String = "ApplyHeaderStyleOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    With Sheets(worksheetName).Range("A1:" & LastColumnLetterOnWorksheet(worksheetName) & "1")
        .Style = "Header"
        .Value = .Value
    End With

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ConvertColumnToDateFormattingOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName '

If InputContainsValue(columnName, "|") Then
    Call RepeatConvertColumnToDateFormattingOnWorksheet(columnName, worksheetName)
Else
    Const methodName As String = "ConvertColumnToDateFormattingOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call EnsureHelperColumnOnWorksheet(worksheetName)

    Call ApplyFormulaToRangeOnWorksheet("=IF(" & ColumnStartOnWorksheet(columnName, worksheetName) & "<>"""";INT(" & ColumnStartOnWorksheet(columnName, worksheetName) & ");"""")", report("Helper Column"), worksheetName)
    Call ApplyFormulaToRangeOnWorksheet("=IF(" & ColumnStartOnWorksheet(report("Helper Column"), worksheetName) & "<>"""";" & ColumnStartOnWorksheet(report("Helper Column"), worksheetName) & ";"""")", columnName, worksheetName)
    Call ApplyCellStyleToRangeOnWorksheet("Date", columnName, worksheetName)

    Call EnsureHelperColumnDeletedOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ConvertColumnToDecimalFormattingOnWorksheet(ByVal columnName As String, ByVal worksheetName As String) ' Repeat Support: columnName. '

If InputContainsValue(columnName, "|") Then
    Call RepeatConvertColumnToDecimalFormattingOnWorksheet(columnName, worksheetName)
Else
    Const methodName As String = "ConvertColumnToDecimalFormattingOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call EnsureHelperColumnOnWorksheet(worksheetName)

    Dim columnData As String
    Dim helperData As String
    columnData = ColumnRangeTypeOnWorksheet(columnName, "Start", worksheetName)
    helperData = ColumnRangeTypeOnWorksheet(report("Helper Column"), "Start", worksheetName)

    Call ApplyFormulaToRangeOnWorksheet("=IF(NOT(ISBLANK(" & columnData & "));IF(NOT(ISERROR(VALUE(" & columnData & ")));VALUE(" & columnData & ");"""");"""")", report("Helper Column"), worksheetName)
    Call ApplyFormulaToRangeOnWorksheet("=IF(" & helperData & "<>"""";" & helperData & ";"""")", columnName, worksheetName)
    
    Call EnsureHelperColumnDeletedOnWorksheet(worksheetName)

    Call ApplyCellStyleToRangeOnWorksheet("Decimal", columnName, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ColorColumnOfTypeAndElementOnWorksheet(ByVal colorName As String, ByVal columnName As String, ByVal columnType As String, ByVal elementType As String, ByVal worksheetName As String)
    Const methodName As String = "ColorColumnOfTypeAndElementOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal columnName As String, ByVal columnType As String, ByVal elementType As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & colorName & """" & ", " & """" & columnName & """" & ", " & """" & columnType & """" & ", " & """" & elementType & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim colorRange As String

    If columnType = "Data" Then colorRange = columnName & "$2:" & columnName & "$" & LastRowNumberOnWorksheet(worksheetName)
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

    If elementType = "Background" Then Sheets(worksheetName).Range(colorRange).Interior.Color = colorValue
    If elementType = "Border" Then
        With Sheets(worksheetName).Range(colorRange).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With Sheets(worksheetName).Range(colorRange).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With Sheets(worksheetName).Range(colorRange).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
        With Sheets(worksheetName).Range(colorRange).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Color = colorValue
            .Weight = xlThin
        End With
    End If
    If elementType = "Font" Then Sheets(worksheetName).Range(colorRange).Font.Color = colorValue

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ColorRangeBackgroundAndFontAndBorderOnWorksheet(ByVal rangeToColor As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal borderColorName As String, ByVal worksheetName As String)
    Const methodName As String = "ColorRangeBackgroundAndFontAndBorderOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal rangeToColor As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal borderColorName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & rangeToColor & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & borderColorName & """" & ", " & """" & worksheetName & """")

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

    mainWorkbook.Sheets(worksheetName).Range(rangeToColor).Interior.Color = backgroundColorValue
    mainWorkbook.Sheets(worksheetName).Range(rangeToColor).Font.Color = fontColorValue

    With mainWorkbook.Sheets(worksheetName).Range(rangeToColor).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With mainWorkbook.Sheets(worksheetName).Range(rangeToColor).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With mainWorkbook.Sheets(worksheetName).Range(rangeToColor).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = borderColorValue
        .Weight = xlThin
    End With
    With mainWorkbook.Sheets(worksheetName).Range(rangeToColor).Borders(xlEdgeRight)
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
    Const methodName As String = "ColorWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & colorName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim colorValue As Long: colorValue = 2

    colorValue = StyleColor(colorName)
    If colorValue = 2 Then Call LogConclusion("Failed", logConclusionData, "Color name " & """" & colorName & """" & " is not a valid predefined color name.")

    mainWorkbook.Sheets(worksheetName).Tab.Color = colorValue

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub InsertImageOnWorksheetWithLeftAndTopValues(ByVal imageFilePath As String, ByVal worksheetName As String, ByVal leftValue As Double, ByVal topValue As Double)
    Const methodName As String = "InsertImageOnWorksheetWithLeftAndTopValues"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal imageFilePath As String, ByVal worksheetName As String, ByVal leftValue As Double, ByVal topValue As Double", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & imageFilePath & """" & ", " & """" & worksheetName & """" & ", " & leftValue & ", " & topValue)

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call SelectWorksheet(worksheetName)
    mainWorkbook.Worksheets(worksheetName).Pictures.Insert(imageFilePath).Select
    Selection.ShapeRange.IncrementLeft leftValue
    Selection.ShapeRange.IncrementTop topValue

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub NormalizeLayoutOnWorksheet(ByVal worksheetName As String) ' Repeat Support: worksheetName. '

If InputContainsValue(worksheetName, "|") Then
    Call RepeatNormalizeLayoutOnWorksheet(worksheetName)
Else
    Const methodName As String = "NormalizeLayoutOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim lastColumnPlusOne As String
    Dim lastRowPlusOne As Long

    lastColumnPlusOne = LastColumnLetterOnWorksheet(worksheetName)
    If lastColumnPlusOne <> "XFD" Then lastColumnPlusOne = IncreaseLetterOnce(lastColumnPlusOne)
    lastRowPlusOne = LastRowNumberOnWorksheet(worksheetName)
    If lastRowPlusOne <> 1048576 Then lastRowPlusOne = lastRowPlusOne + 1

    Sheets(worksheetName).Cells.Borders.LineStyle = xlLineStyleNone
    Sheets(worksheetName).Cells.Interior.Pattern = xlNone
    Sheets(worksheetName).Cells.Interior.TintAndShade = 0
    Sheets(worksheetName).Cells.Interior.PatternTintAndShade = 0

    Sheets(worksheetName).Cells.RowHeight = 16.5
    Sheets(worksheetName).Rows("1:1").RowHeight = 49.5
    Sheets(worksheetName).Cells.Style = "Normal"

    Sheets(worksheetName).Cells.Font.Name = "Arial"
    Sheets(worksheetName).Cells.Font.Size = 10
    Sheets(worksheetName).Cells.HorizontalAlignment = xlCenter
    Sheets(worksheetName).Cells.VerticalAlignment = xlCenter

    Sheets(worksheetName).Columns(lastColumnPlusOne & ":" & lastColumnPlusOne).Delete
    Sheets(worksheetName).Rows(lastRowPlusOne & ":" & lastRowPlusOne).Delete

    Call LogConclusion("Completed", logConclusionData)
End If
    
End Sub

Sub ResetAutoFilterOnWorksheet(ByVal worksheetName As String)
    Const methodName As String = "ResetAutoFilterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetName As String", methodName, isRegistered, "Formatting")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Sheets(worksheetName).AutoFilterMode = False
    Sheets(worksheetName).Range("A1:" & LastColumnLetterOnWorksheet(worksheetName) & "1").AutoFilter

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

Sub LogCheckpoint(ByVal checkpointType As String, ByVal checkpointName As String, ByVal checkpointStatus As String, ByVal queryPerformanceCounter As Double)
    Const methodName As String = "LogCheckpoint"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal checkpointType As String, ByVal checkpointName As String, ByVal checkpointStatus As String, ByVal queryPerformanceCounter As Double", methodName, isRegistered, "Logging")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & checkpointType & """" & ", " & """" & checkpointName & """" & ", " & """" & checkpointStatus & """" & ", " & queryPerformanceCounter)

    If report("Log Engine Active") = False Then
        report("Log Engine Active") = True

        report("Log Summary Conclusion Tick Count") = logConclusionData.tickCount

        Call SetAboutNamedRange((report("Log Summary Conclusion Tick Count") - report("Log Summary Beginning Tick Count")), "Duration (Milliseconds)")
        Call SetAboutNamedRange("1 Runs. 0 Checkpoints. 0 Rows.", "Log Summary")

        report("Log Summary Beginning Tick Count")  = 0
        report("Log Summary Conclusion Tick Count") = 0
    End If

    If checkpointStatus = "Beginning" Then
        report("Log Summary Beginning Tick Count") = logConclusionData.tickCount
    End If

    ' If checkPointStatus = "Beginning" Then Call StartLoggingEngine(Environ("Username"), templateVersion, mainWorkbook.FullName)

    ' If checkPointStatus = "Conclusion" Then
    '     Call MoveWorksheetToEnd("About|Log")
    '     Call EnsureHelperColumnsDeletedOnMainWorkbook()
    ' End If

    ' If checkpointType <> "Foundation" And checkpointType <> "Augmentation" Then
    '     Call LogConclusion("Failed", logConclusionData, "Invalid checkpoint type: """ & checkpointType & """.")
    ' End If

    Dim checkpointColorCode As Long

    If checkpointStatus = "Beginning" And checkpointType ="Foundation" Then checkpointColorCode = -16737281
    If checkpointStatus = "Conclusion" And checkpointType ="Foundation" Then checkpointColorCode = -10040320

    If checkpointStatus = "Beginning" And checkpointType ="Augmentation" Then checkpointColorCode = -13056
    If checkpointStatus = "Conclusion" And checkpointType ="Augmentation" Then checkpointColorCode = -1897831

    mainWorkbook.Sheets("Log").Range("A" & logConclusionData.operationSequenceNumber & ":" & "B" & logConclusionData.operationSequenceNumber).Font.Color = checkpointColorCode

    Dim lastRowCheckpoint As Long
    lastRowCheckpoint = LastRowNumberOnWorksheet("Log")
    If lastRowCheckpoint <= 30 Then lastRowCheckpoint = 31

    Sheets("Log").Select
    Application.Goto Reference:=ActiveSheet.Cells.SpecialCells(xlCellTypeVisible).Range("A" & (lastRowCheckpoint - 30)), Scroll:=True
    Range("A" & logConclusionData.operationSequenceNumber & ":B" & logConclusionData.operationSequenceNumber).Select

    ' If checkPointStatus = "Conclusion" And checkpointType ="Foundation" Then
    '     If Range("ProgressionStatus").Value <> "Progression Status: N/A." Then
    '         Range("ProgressionStatus").Value = Left(Range("ProgressionStatus").Value, Len(Range("ProgressionStatus").Value) - 1)
    '         Range("ProgressionStatus").Value = Range("ProgressionStatus").Value & ", " & checkpointName & "."
    '     Else
    '         Range("ProgressionStatus").Value = Left(Range("ProgressionStatus").Value, Len(Range("ProgressionStatus").Value) - 4)
    '         Range("ProgressionStatus").Value = Range("ProgressionStatus").Value & checkpointName & "."
    '     End If
    ' End If

    ' If checkPointStatus = "Conclusion" And checkpointType ="Augmentation" Then
    '     If Range("AugmentationModules").Value <> "Augmentation Modules: N/A." Then
    '         Range("AugmentationModules").Value = Left(Range("AugmentationModules").Value, Len(Range("AugmentationModules").Value) - 1)
    '         Range("AugmentationModules").Value = Range("AugmentationModules").Value & ", " & checkpointName & "."
    '     Else
    '         Range("AugmentationModules").Value = Left(Range("AugmentationModules").Value, Len(Range("AugmentationModules").Value) - 4)
    '         Range("AugmentationModules").Value = Range("AugmentationModules").Value & checkpointName & "."
    '     End If
    ' End If

    ' If (checkpointName <> "Launch") Then
    '     Application.ScreenUpdating = True
    '     Application.DisplayAlerts = True
    '     Application.EnableEvents = True
    '     Application.Wait(Now + TimeValue("0:00:2"))
    '     Application.ScreenUpdating = False
    '     Application.DisplayAlerts = False
    '     Application.EnableEvents = False
    ' End If

    ' If checkPointStatus = "Conclusion" Then Call ResetLoggingEngine(checkpointType, checkpointName)

    Call LogConclusion("Completed", logConclusionData)

    If checkpointStatus = "Conclusion" Then
        report("Log Summary Conclusion Tick Count") = (logConclusionData.tickCount + mainWorkbook.Sheets("Log").Cells(logConclusionData.operationSequenceNumber, 6).Value)

        Call SetAboutNamedRange((report("Log Summary Conclusion Tick Count") - report("Log Summary Beginning Tick Count")), "Duration (Milliseconds)")
        Call SetAboutNamedRange("0 Runs. 1 Checkpoints. 0 Rows.", "Log Summary")

        report("Log Summary Beginning Tick Count")  = 0
        report("Log Summary Conclusion Tick Count") = 0
    End If
End Sub

Sub LogEngine(ByVal reportName As String, ByVal templateVersion As String, ByVal qpcMidpointTimestamp As Double, ByVal qpcFrequency As Double, ByVal tickCount As Double, ByVal utcTimestampPrecise As String)
If report("Log Engine Active") = True Then Exit Sub

    Const methodName As String = "LogEngine"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal reportName As String, ByVal templateVersion As String, ByVal qpcMidpointTimestamp As Double, ByVal qpcFrequency As Double, ByVal tickCount As Double, ByVal utcTimestampPrecise As String", methodName, isRegistered, "Logging")
    End If

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

                mainWorkbook.Sheets("Log").Delete

                Application.DisplayAlerts = True

                mainWorkbook.Sheets.Add After:=mainWorkbook.ActiveSheet
                mainWorkbook.ActiveSheet.Name = "Log"
            Else
                Dim logWorksheet As Worksheet: Set logWorksheet = mainWorkbook.Worksheets("Log")
                Dim lastRow As Long: lastRow = logWorksheet.Cells(logWorksheet.Rows.Count, 1).End(xlUp).Row
                report("Operation Sequence Number") = lastRow&
            End If
        End With
    Else
        mainWorkbook.Sheets.Add After:=mainWorkbook.ActiveSheet
        mainWorkbook.ActiveSheet.Name = "Log"
    End If

    If mainWorkbook.Sheets("Log").Range("A1").Value = "" Then mainWorkbook.Sheets("Log").Range("A1:F1").Value = Array("Method", "Arguments", "Category", "Outcome", "Tick Count", "Duration")
  
    Call LogBeginning(methodName, logConclusionData, """" & reportName & """" & ", " & """" & templateVersion & """" & ", " & qpcMidpointTimestamp & ", " & qpcFrequency & ", " & tickCount & ", " & """" & utcTimestampPrecise & """")
    report("Log Summary Beginning Tick Count") = logConclusionData.tickCount

    Call ConfigureMethodSetting(methodName, "Configure New Cell Styles", 1, 0, 1)

    If methodRegistry(methodName)("Settings")("Configure New Cell Styles")("Value") = True Then
        Dim newCellStyles() As String: newCellStyles = Split("Date|Date Time|Decimal|Followed Hyperlink|Formula|Header|Integer", "|")

        Dim index As Byte
        For index = 0 To UBound(newCellStyles)
            If CellStyleExists(newCellStyles(index)) = False Then
                mainWorkbook.Styles.Add Name:=newCellStyles(index)

                With mainWorkbook.Styles(newCellStyles(index))
                    .IncludeAlignment = True
                    .IncludeNumber = True
                    .IncludePatterns = False
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .ReadingOrder = xlContext
                End With

                If newCellStyles(index) = "Decimal" Then
                    mainWorkbook.Styles("Decimal").NumberFormat = "General"
                    mainWorkbook.Styles("Decimal").NumberFormat = "0.00"
                End If

                If newCellStyles(index) = "Date" Then
                    mainWorkbook.Styles("Date").NumberFormat = "dd/mm/yyyy"
                End If

                If newCellStyles(index) = "Date Time" Then
                    mainWorkbook.Styles("Date Time").NumberFormat = "dd/mm/yyyy HH:mm:ss"
                End If

                If newCellStyles(index) = "Followed Hyperlink" Then
                    With mainWorkbook.Styles("Followed Hyperlink")
                        .IncludeFont = False
                    End With
                    mainWorkbook.Styles("Followed Hyperlink").NumberFormat = "@"
                    With mainWorkbook.Styles("Followed Hyperlink").Font
                        .ColorIndex = xlAutomatic
                    End With
                End If

                If newCellStyles(index) = "Header" Then
                    With mainWorkbook.Styles("Header")
                        .IncludeFont = True
                        .WrapText = True
                    End With
                    mainWorkbook.Styles("Header").NumberFormat = "@"
                    With mainWorkbook.Styles("Header").Font
                        .Bold = True
                        .ColorIndex = xlAutomatic
                    End With
                End If

                If newCellStyles(index) = "Integer" Then
                    With mainWorkbook.Styles("Integer")
                        .IncludeFont = True
                    End With
                    mainWorkbook.Styles("Integer").NumberFormat = "0"
                    With mainWorkbook.Styles("Integer").Font
                        .Name = "Consolas"
                        .ColorIndex = xlAutomatic
                    End With
                End If

                If newCellStyles(index) = "Formula" Then
                    mainWorkbook.Styles("Formula").NumberFormat = "General"
                End If
            End If
        Next index
    End If

    If mainWorkbook.Styles("Percent").NumberFormat = "0%" Then
        With mainWorkbook.Styles("Percent")
            .IncludeNumber = True
            .IncludeAlignment = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .ReadingOrder = xlContext
        End With
        mainWorkbook.Styles("Percent").NumberFormat = "0.00 %"

        With mainWorkbook.Styles("Percent").Font
            .Name = "Arial"
            .Size = 10
        End With
        With mainWorkbook.Styles("Percent")
            .IncludeFont = False
        End With
    End If

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

Sub LogConclusion(ByVal conclusionStatus As String, ByRef logConclusionData As LogEntry, Optional ByVal errorMessage As String)
    Dim tickCountNow As Currency: tickCountNow = GetTickCount64()
    Dim duration As Double: duration = CDbl(tickCountNow * 10000) - logConclusionData.tickCount

    With mainWorkbook.Sheets("Log")
        .Cells(logConclusionData.operationSequenceNumber, 4).Value = conclusionStatus
        .Cells(logConclusionData.operationSequenceNumber, 6).Value = duration
    End With

    If errorMessage <> "" Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True

        mainWorkbook.Sheets("Log").Range("A" & (logConclusionData.operationSequenceNumber) & ":" & "D" & (logConclusionData.operationSequenceNumber)).Font.Color = -11526924

        Dim targetCell As Range: Set targetCell = mainWorkbook.Sheets("Log").Range("D" & logConclusionData.operationSequenceNumber)
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

Sub LogTaskWarning(ByVal warningMessage As String, ByVal taskLogOrder As Long)
    mainWorkbook.Sheets("Log").Range("D" & (taskLogOrder)).Value = "Warning"
    mainWorkbook.Sheets("Log").Range("H" & (taskLogOrder)).Value = Format(Now, "dd/mm/yyyy HH:mm:ss")
    mainWorkbook.Sheets("Log").Range("J" & (taskLogOrder)).Value = Round(Timer - stopwatchTimer, 2)

    If mainWorkbook.Sheets("Log").Range("D" & (taskLogOrder)).Comment Is Nothing Then
        mainWorkbook.Sheets("Log").Range("D" & (taskLogOrder)).AddComment (warningMessage)
    Else
        mainWorkbook.Sheets("Log").Range("D" & (taskLogOrder)).Comment.Text Text:=Sheets("Log").Range("D" & (taskLogOrder)).Comment.Text & " " & warningMessage
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

' Functions: Logging '

Private Sub LogBeginning(ByVal methodName As String, ByRef logConclusionData As LogEntry, ByVal arguments As String)
    Dim tickCountNow As Currency: tickCountNow = GetTickCount64()
    report("Operation Sequence Number") = report("Operation Sequence Number") + 1

    logConclusionData.operationSequenceNumber = report("Operation Sequence Number")
    logConclusionData.methodName = methodName
    logConclusionData.arguments = arguments
    logConclusionData.tickCount = CDbl(tickCountNow * 10000)

    mainWorkbook.Sheets("Log").Range("A" & logConclusionData.operationSequenceNumber & ":" & "E" & logConclusionData.operationSequenceNumber).Value = _
        Array(logConclusionData.methodName, logConclusionData.arguments, methodRegistry(methodName)("Category"), "Beginning", logConclusionData.tickCount)
End Sub

' ************ '
' Repetition   '
' ************ '

Sub RepeatAnonymizeNumbersOnColumnOnWorksheet(ByVal columnNames, ByVal worksheetName)
    Const methodName As String = "RepeatAnonymizeNumbersOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames, ByVal worksheetName", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String: columnNamesArray = Split(columnNames, "|")
    Dim index As Byte
    
    For index = 0 To UBound(columnNamesArray)
        Call AnonymizeNumbersOnColumnOnWorksheet(columnNamesArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatApplyAutoFilterOnWorksheet(ByVal worksheetNames As String)
    Const methodName As String = "RepeatApplyAutoFilterOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte
    
    For index = 0 To UBound(worksheetNamesArray)
        Call ApplyAutoFilterOnWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatApplyCellStyleToRangeOnWorksheet(ByVal cellStyle As String, ByVal rangeValues As String, ByVal worksheetNames As String)
    Const methodName As String = "RepeatApplyCellStyleToRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyle As String, ByVal rangeValues As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & cellStyle & """" & ", " & """" & rangeValues & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(rangeValues, validation)
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim rangeValuesArray() As String
    Dim worksheetNamesArray() As String
    Dim index As Byte

    If InputContainsValue(rangeValues, "|") Then
        rangeValuesArray = Split(rangeValues, "|")
        For index = 0 To UBound(rangeValuesArray)
            Call ApplyCellStyleToRangeOnWorksheet(cellStyle, rangeValuesArray(index), worksheetNames)
        Next index
    End If

    If InputContainsValue(worksheetNames, "|") Then
        worksheetNamesArray = Split(worksheetNames, "|")
        For index = 0 To UBound(worksheetNamesArray)
            Call ApplyCellStyleToRangeOnWorksheet(cellStyle, rangeValues, worksheetNamesArray(index))
        Next index
    End If

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatApplyFreezePanesModeOnWorksheet(ByVal freezePanesMode As String, ByVal worksheetNames As String)
    Const methodName As String = "RepeatApplyFreezePanesModeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal freezePanesMode As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & freezePanesMode & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call ApplyFreezePanesModeOnWorksheet(freezePanesMode, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatCutColumnAndPasteAtColumnOnWorksheet(ByVal cutColumnNames As String, ByVal pasteColumnName As String, ByVal worksheetName As String)
    Const methodName As String = "RepeatCutColumnAndPasteAtColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cutColumnNames As String, ByVal pasteColumnName As String, ByVal worksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & cutColumnNames & """" & ", " & """" & pasteColumnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(cutColumnNames, validation)
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

Sub RepeatApplyHeaderStyleOnWorksheet(ByVal worksheetNames As String)
    Const methodName As String = "RepeatApplyHeaderStyleOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte
    
    For index = 0 To UBound(worksheetNamesArray)
        Call ApplyHeaderStyleOnWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorBorderOuterAndInnerOnWorksheet(ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetNames As String)
    Const methodName As String = "RepeatColorBorderOuterAndInnerOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & outerColorName & """" & ", " & """" & innerColorName & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call ColorBorderOuterAndInnerOnWorksheet(outerColorName, innerColorName, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatColorDataBackgroundAndFontOnWorksheet(ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String)
    Const methodName As String = "RepeatColorDataBackgroundAndFontOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    Call ValidateRepeatInputParameter(worksheetNames, validation)
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
    Const methodName As String = "RepeatColorHeaderBackgroundAndFontOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    Call ValidateRepeatInputParameter(worksheetNames, validation)
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
    Const methodName As String = "RepeatColorWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal colorName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & colorName & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call ColorWorksheet(colorName, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatConvertColumnToDateFormattingOnWorksheet(ByVal columnNames As String, ByVal worksheetName As String)
    Const methodName As String = "RepeatConvertColumnToDateFormattingOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String: columnNamesArray = Split(columnNames, "|")
    Dim index As Byte

    For index = 0 To UBound(columnNamesArray)
        Call ConvertColumnToDateFormattingOnWorksheet(columnNamesArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatConvertColumnToDecimalFormattingOnWorksheet(ByVal columnNames As String, ByVal worksheetName As String)
    Const methodName As String = "RepeatConvertColumnToDecimalFormattingOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String: columnNamesArray = Split(columnNames, "|")
    Dim index As Byte

    For index = 0 To UBound(columnNamesArray)
        Call ConvertColumnToDecimalFormattingOnWorksheet(columnNamesArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatCreateBlankWorksheet(ByVal worksheetNames As String)
    Const methodName As String = "RepeatCreateBlankWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call CreateBlankWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteColumnOnWorksheet(ByVal columnNames As String, ByVal worksheetNames As String)
    Const methodName As String = "RepeatDeleteColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    Call ValidateRepeatInputParameter(worksheetNames, validation)
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
    Const methodName As String = "RepeatDeleteRowsOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal numberOfRows As Long, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, numberOfRows & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call DeleteRowsOnWorksheet(numberOfRows, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDeleteWorksheet(ByVal worksheetNames As String)
    Const methodName As String = "RepeatDeleteWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call DeleteWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatDuplicateColumnFromWorksheetToWorksheet(ByVal columnNames As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String)
    Const methodName As String = "RepeatDuplicateColumnFromWorksheetToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal fromWorksheetName As String, ByVal toWorksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", " & """" & fromWorksheetName & """" & ", " & """" & toWorksheetName & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim columnNamesArray() As String: columnNamesArray = Split(columnNames, "|")
    Dim index As Byte

    For index = 0 To UBound(columnNamesArray)
        Call DuplicateColumnFromWorksheetToWorksheet(columnNamesArray(index), fromWorksheetName, toWorksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatFindAndReplaceOnRangeOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValues As String, ByVal worksheetName As String)
    Const methodName As String = "RepeatFindAndReplaceOnRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValues As String, ByVal worksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & rangeValues & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(rangeValues, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim rangeValuesArray() As String: rangeValuesArray = Split(rangeValues, "|")
    Dim index As Byte

    For index = 0 To UBound(rangeValuesArray)
        Call FindAndReplaceOnRangeOnWorksheet(findValue, replaceValue, rangeValuesArray(index), worksheetName)
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatHideWorksheet(ByVal worksheetNames As String)
    Const methodName As String = "RepeatHideWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call HideWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatInsertDataValuesOnWorksheet(ByVal dataValues As String, ByVal worksheetNames As String)
    Const methodName As String = "RepeatInsertDataValuesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal dataValues As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & dataValues & """" & ", "  & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(dataValues, validation)
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call InsertDataValuesOnWorksheet(dataValues, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatInsertNewColumnAndSetWidthOnWorksheet(ByVal columnNames As String, ByVal setWidth As Double, ByVal worksheetNames As String)
    Const methodName As String = "RepeatInsertNewColumnAndSetWidthOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal setWidth As Double, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", "  & setWidth & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    Call ValidateRepeatInputParameter(worksheetNames, validation)
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
    Const methodName As String = "RepeatInsertHeaderValuesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal headerValues As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & headerValues & """" & ", "  & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(headerValues, validation)
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call InsertHeaderValuesOnWorksheet(headerValues, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatMoveWorksheetToEnd(ByVal worksheetNames As String)
    Const methodName As String = "RepeatMoveWorksheetToEnd"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call MoveWorksheetToEnd(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatNormalizeLayoutOnWorksheet(ByVal worksheetNames As String)
    Const methodName As String = "RepeatNormalizeLayoutOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call NormalizeLayoutOnWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatRemoveBlankLinesOnWorksheet(ByVal worksheetNames As String)
    Const methodName As String = "RepeatRemoveBlankLinesOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetsArray() As String: worksheetsArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetsArray)
        Call RemoveBlankLinesOnWorksheet(worksheetsArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatRemoveCellStyle(ByVal cellStyleNames As String)
    Const methodName As String = "RepeatRemoveCellStyle"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal cellStyleNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & cellStyleNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(cellStyleNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim cellStyleNamesArray() As String: cellStyleNamesArray = Split(cellStyleNames, "|")
    Dim index As Byte

    For index = 0 To UBound(cellStyleNamesArray)
        Call RemoveCellStyle(cellStyleNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatRemoveEmptyColumnsOnWorksheet(ByVal worksheetNames As String)
    Const methodName As String = "RepeatRemoveEmptyColumnsOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call RemoveEmptyColumnsOnWorksheet(worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatSetNewWidthOnColumnOnWorksheet(ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetNames As String)
    Const methodName As String = "RepeatSetNewWidthOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal newWidth As Double, ByVal columnName As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(worksheetNames, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Dim worksheetNamesArray() As String: worksheetNamesArray = Split(worksheetNames, "|")
    Dim index As Byte

    For index = 0 To UBound(worksheetNamesArray)
        Call SetNewWidthOnColumnOnWorksheet(newWidth, columnName, worksheetNamesArray(index))
    Next index

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub RepeatSortColumnByOrderOnWorksheet(ByVal columnNames As String, ByVal sortOrder As String, ByVal worksheetNames As String)
    Const methodName As String = "RepeatSortColumnByOrderOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnNames As String, ByVal sortOrder As String, ByVal worksheetNames As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnNames & """" & ", " & """" & sortOrder & """" & ", " & """" & worksheetNames & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(columnNames, validation)
    Call ValidateRepeatInputParameter(worksheetNames, validation)
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
    Const methodName As String = "RepeatTransferDataFromWorksheetToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetNames As String, toWorksheetName As String", methodName, isRegistered, "Repetition")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & fromWorksheetNames & """" & ", " & """" & toWorksheetName & """")

    Dim validation As String
    Call ValidateRepeatInputParameter(fromWorksheetNames, validation)
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
    Const methodName As String = "AnonymizeNumbersOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
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

    Call ApplyFormulaToRangeOnWorksheet("=SUBSTITUTE(TEXT(" & ColumnStartOnWorksheet(formulaColumn, worksheetName) & ";""@"");""!"";"""")", columnName, worksheetName)
    Call ApplyCellStyleToRangeOnWorksheet("Normal", columnName, worksheetName)
    Call DeleteColumnOnWorksheet(formulaColumn, worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ApplyConsecutiveGroupBasedOnFormulaOnColumnOnWorksheet(ByVal consecutiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String)
    Const methodName As String = "ApplyConsecutiveGroupBasedOnFormulaOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal consecutiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & consecutiveGroupName & """" & ", " & """" & formulaValue & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call ApplyFormulaToRangeOnWorksheet(formulaValue, report("Helper Column"), worksheetName)
    Call ApplyFilterToColumnOnWorksheet(True, report("Helper Column"), worksheetName)

    Call FindAndReplaceOnRangeOnWorksheet(";;", ";" & consecutiveGroupName & ";;", columnName, worksheetName)
    Call ClearFilterOnWorksheet(worksheetName)

    Call SelectWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet(ByVal definitiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String)
    Const methodName As String = "ApplyDefinitiveGroupBasedOnFormulaOnColumnOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal definitiveGroupName As String, ByVal formulaValue As String, ByVal columnName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & definitiveGroupName & """" & ", " & """" & formulaValue & """" & ", " & """" & columnName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If formulaValue = "" Then formulaValue = "=True"

    Call ApplyFormulaToRangeOnWorksheet(formulaValue, report("Helper Column"), worksheetName)
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
    Const methodName As String = "ColorBorderOuterAndInnerOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal outerColorName As String, ByVal innerColorName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & outerColorName & """" & ", " & """" & innerColorName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If outerColorName = "" Then outerColorName = "Light Grey"
    If innerColorName = "" Then innerColorName = "Light Grey"

    Dim lastColumnLetterPlusOne As String
    lastColumnLetterPlusOne = LastColumnLetterOnWorksheet(worksheetName)
    lastColumnLetterPlusOne = IncreaseLetterOnce(lastColumnLetterPlusOne)
    lastColumnLetterPlusOne = IncreaseLetterOnce(lastColumnLetterPlusOne)

    ' Outer Border. '
    With Sheets(worksheetName).Cells.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = StyleColor(outerColorName)
        .Weight = xlThin
    End With
    With Sheets(worksheetName).Cells.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = StyleColor(outerColorName)
        .Weight = xlThin
    End With

    Dim lastRowNumberPlusOne As Long
    lastRowNumberPlusOne = LastRowNumberOnWorksheet(worksheetName) + 1

    Sheets(worksheetName).Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Sheets(worksheetName).Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Inner Border. '
    With Sheets(worksheetName).Range("A1:" & lastColumnLetterPlusOne & lastRowNumberPlusOne + 1).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = StyleColor(innerColorName)
        .Weight = xlThin
    End With
    With Sheets(worksheetName).Range("A1:" & lastColumnLetterPlusOne & lastRowNumberPlusOne + 1).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = StyleColor(innerColorName)
        .Weight = xlThin
    End With

    ' Clean up of bug or weird behaviour which doesn't apply the border everywhere. '
    lastColumnLetterPlusOne = DecreaseLetterOnce(lastColumnLetterPlusOne)
    Sheets(worksheetName).Columns("A:A").Delete Shift:=xlToLeft
    Sheets(worksheetName).Rows("1:1").Delete Shift:=xlUp
    Sheets(worksheetName).Columns(lastColumnLetterPlusOne & ":" & lastColumnLetterPlusOne).Delete Shift:=xlToLeft
    Sheets(worksheetName).Columns(lastColumnLetterPlusOne & ":" & lastColumnLetterPlusOne).Delete Shift:=xlToLeft
    Sheets(worksheetName).Rows(lastRowNumberPlusOne & ":" & lastRowNumberPlusOne).Delete Shift:=xlUp
    Sheets(worksheetName).Rows(lastRowNumberPlusOne & ":" & lastRowNumberPlusOne).Delete Shift:=xlUp

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub ColorDataBackgroundAndFontOnWorksheet(ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String) ' Repeat Support: columnName, worksheetName, columnName/worksheetName. '

If InputContainsValue(columnName, "|") Or InputContainsValue(worksheetName, "|") Then
    Call RepeatColorHeaderBackgroundAndFontOnWorksheet(columnName, backgroundColorName, fontColorName, worksheetName)
Else
    Const methodName As String = "ColorDataBackgroundAndFontOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
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
    Const methodName As String = "ColorHeaderBackgroundAndFontOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal columnName As String, ByVal backgroundColorName As String, ByVal fontColorName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & columnName & """" & ", " & """" & backgroundColorName & """" & ", " & """" & fontColorName & """" & ", " & """" & worksheetName & """")

    Dim validation As String
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    Call ValidateColumnOnWorksheet(columnName, worksheetName, validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    If backgroundColorName = "" Then backgroundColorName = "White"
    If fontColorName = "" Then fontColorName = "Black"

    Call ColorColumnOfTypeAndElementOnWorksheet(backgroundColorName, columnName, "Header", "Background", worksheetName)
    Call ColorColumnOfTypeAndElementOnWorksheet(fontColorName, columnName, "Header", "Font", worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End If

End Sub

Sub FindAndReplaceExactOnRangeOnWorksheet(ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValue As String, ByVal worksheetName As String)
    Const methodName As String = "FindAndReplaceExactOnRangeOnWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal findValue As String, ByVal replaceValue As String, ByVal rangeValue As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & findValue & """" & ", " & """" & replaceValue & """" & ", " & """" & rangeValue & """" & ", " & """" & worksheetName & """")

    Call ApplyFilterToColumnOnWorksheet(findValue, rangeValue, worksheetName)
    Call FindAndReplaceOnRangeOnWorksheet(findValue, replaceValue, rangeValue, worksheetName)
    Call ClearFilterOnWorksheet(worksheetName)

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub ImportExcelFileWithDependencyAndRenameWorksheet(ByVal filePath As String, ByVal dependency As String, ByVal worksheetName As String)
    Const methodName As String = "ImportExcelFileWithDependencyAndRenameWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal filePath As String, ByVal dependency As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & filePath & """" & " ," & """" & dependency & """" & " ," & """" & worksheetName & """")

    Call ImportExcelFileAndRenameWorksheet(filePath, worksheetName)
    Call SetAboutNamedRange(dependency, "Dependencies List")

    Call LogConclusion("Completed", logConclusionData)
End Sub

Sub MoveDataFromWorksheetToWorksheet(ByVal fromWorksheetName As String, ByVal worksheetName As String)
    Const methodName As String = "MoveDataFromWorksheetToWorksheet"
    Static isRegistered As Boolean
    Dim logConclusionData As LogEntry

    If isRegistered = False Then
        Call RegisterMethod("ByVal fromWorksheetName As String, ByVal worksheetName As String", methodName, isRegistered, "Sequencing")
    End If

    Call LogBeginning(methodName, logConclusionData, """" & fromWorksheetName & """" & ", " & """" & worksheetName & """")
  
    Dim validation As String
    Call ValidateWorksheet(fromWorksheetName, "fromWorksheetName", validation)
    Call ValidateWorksheet(worksheetName, "worksheetName", validation)
    If validation <> "" Then Call LogConclusion("Failed", logConclusionData, validation)

    Call EnsureHelperColumnDeletedOnWorksheet(fromWorksheetName)
    Call EnsureHelperColumnDeletedOnWorksheet(worksheetName)
   
    Dim numberOfColumnsFromWorksheetName As Long: numberOfColumnsFromWorksheetName = ConvertLetterToNumber(LastColumnLetterOnWorksheet(fromWorksheetName))
    Dim numberOfColumnsWorksheetName As Long: numberOfColumnsWorksheetName = ConvertLetterToNumber(LastColumnLetterOnWorksheet(worksheetName))

    If numberOfColumnsFromWorksheetName <> numberOfColumnsWorksheetName Then Call LogConclusion("Failed", logConclusionData, "Number of columns in worksheet """ & fromWorksheetName & """ doesn't match up with the worksheet """ & worksheetName & """.")

    Dim pasteRow As Long: pasteRow = LastRowNumberOnWorksheet(worksheetName) + 1

    If NumberOfVisibleCells(fromWorksheetName) <> 0 Then
        Sheets(fromWorksheetName).Range("A2:" & LastColumnLetterOnWorksheet(fromWorksheetName) & LastRowNumberOnWorksheet(fromWorksheetName)).Copy
        Sheets(worksheetName).Range("A" & pasteRow).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
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
    rowsOfDataFound = LastRowNumberOnWorksheet(worksheetName) - 1

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

Sub ValidateColumnOnWorksheet(ByRef columnValue As String, ByVal worksheetName As String, ByRef validation As String)
    If Len(columnValue) <= 3 Then
        If ColumnLetterValid(columnValue) Then
            Exit Sub
        Else
            ' Call LogConclusion("Failed", logConclusionData, "Column letter " & """" & columnValue & """" & " on worksheet """ & worksheetName & """ is not a valid column.")
        End If
    End If

    If ColumnOnWorksheetExists(columnValue, worksheetName) = False Then Call LogConclusion("Failed", logConclusionData, "Column " & """" & columnValue & """" & " on worksheet """ & worksheetName & """ not found.")
    columnValue = FindColumnLetterOnWorksheet(columnValue, worksheetName)
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

Sub ValidateRepeatInputParameter(ByVal inputParameter As String, ByRef validation As String)
    ' If InStr(inputParameter, "||") <> 0 Then Call LogConclusion("Failed", logConclusionData, "Input parameter contains multiple instances of ""||"" next to each other: " & """" & inputParameter & """" & ".")
    ' If Right(inputParameter, 1) = "|" Then Call LogConclusion("Failed", logConclusionData, "Input parameter contains instance of ""|"" at the very end: " & """" & inputParameter & """" & ".")
    ' If Left(inputParameter, 1) = "|" Then Call LogConclusion("Failed", logConclusionData, "Input parameter contains instance of ""|"" at the very start: " & """" & inputParameter & """" & ".")
End Sub

Sub ValidateUniqueColumnOnWorksheet(ByVal columnName As String, ByVal worksheetName As String, ByRef validation As String)
    ' If ColumnOnWorksheetExists(columnName, worksheetName) Then Call LogConclusion("Failed", logConclusionData, "Column " & """" & columnName & """" & " on worksheet " & """" & worksheetName & """" & " already exists.")
End Sub

Sub ValidateWorksheet(ByVal worksheetName As String, ByVal parameterName As String, ByRef validation As String, Optional ByVal validWorksheet As Boolean)
    Dim validationMessage As String
    Dim worksheetNameExists As Boolean: worksheetNameExists = WorksheetExists(worksheetName)

    If worksheetNameExists = True And validWorksheet = True Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. Worksheet name already exists."
    ElseIf Len(worksheetName) >= 27 Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. Worksheet name is too long, unable to process further."
    ElseIf Len(worksheetName) = 0 Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. Worksheet name can't be blank."
    ElseIf Left$(worksheetName, 1) = "'" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. Worksheet name can't start with the apostrophe character (')."
    ElseIf Right$(worksheetName, 1) = "'" Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. Worksheet name can't end with the apostrophe character (')."
    ElseIf StrComp(worksheetName, "History", vbTextCompare) = 0 Then
        validationMessage = "Parameter """ & parameterName & """ failed validation. Worksheet name is reserved."
    End If

    If validationMessage = "" Then
        Dim forbiddenCharacters As String: forbiddenCharacters = "/\?*:[]"
        Dim characterIndex As Byte
        Dim currentForbiddenCharacter As String
    
        For characterIndex = 1 To Len(forbiddenCharacters)
            currentForbiddenCharacter = Mid$(forbiddenCharacters, characterIndex, 1)
            
            If InStr(1, worksheetName, currentForbiddenCharacter, vbBinaryCompare) > 0 Then
                validationMessage = "Parameter """ & parameterName & """ failed validation. Worksheet name has a forbidden character: " & currentForbiddenCharacter & "."
                Exit For
            End If
        Next characterIndex

        If validationMessage = "" And worksheetNameExists = False And validWorksheet = False Then
            validationMessage = "Parameter """ & parameterName & """ failed validation. Worksheet name not found."
        End If
    End If

    If validationMessage <> "" Then
        If validation = "" Then
            validation = validationMessage
        Else
            validation = validation & " " & validationMessage
        End If
    End If
End Sub

' Functions: Validation '

Function CellStyleExists(ByVal styleName As String) As Boolean
    On Error Resume Next
        CellStyleExists = Len(mainWorkbook.Styles(styleName).Name) > 0
    On Error GoTo 0
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
    convertLetterToNumber = Sheets("Log").Range(columnLetter & 1).Column

    If Err.Number <> 0 Then
        Err.Clear
        ColumnLetterValid = False
    End If
End Function

Function ColumnOnWorksheetExists(ByVal columnName As String, ByVal worksheetName As String) As Boolean
    Dim columnRange As Range

    If WorksheetExists(worksheetName) Then Set columnRange = Sheets(worksheetName).Range("A1:" & LastColumnLetterOnWorksheet(worksheetName) & "1").Find(columnName, LookIn:=xlValues, Lookat:=xlWhole)

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

    report("Template Version")   = "v0.40, 2026-08-06"
    report("Retrieved Date")     = Format$(Date, "yyyy-mm-dd")
    report("Checkpoint Counter") = 0&
    report("Checkpoint Type")    = "Foundation"
    report("Helper Column")      = "Helper Column"
    report("Log Engine Counter") = 0&
    report("Log Engine Active")  = False
    report("Original Workbook")  = mainWorkbook.FullName
    report("Operation Sequence Number") = 1&
    Set report("Pre-Log Settings")      = New Collection
    report("Log Summary Beginning Tick Count")  = 0#
    report("Log Summary Conclusion Tick Count") = 0#

    Call Startup()
    Call Master()
End Sub

' Script is finished when "About" is placed last. '
' When saving the document, choose yes at prompt. '