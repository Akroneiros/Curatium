#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Logging Library.ahk

AssignSpreadsheetOperationsTemplateCombined(version := "") {
    static methodName := RegisterMethod("AssignSpreadsheetOperationsTemplateCombined(version As String [Optional])", A_LineFile, A_LineNumber + 7)
    overlayValue := "Assign Spreadsheet Operations Template Code"
    if version = "" {
        overlayValue := overlayValue . " ([Latest])"
    } else {
        overlayValue := overlayValue . " (" . version . ")"
    }
    logValuesForConclusion := LogInformationBeginning(overlayValue, methodName, [version])

    spreadsheetOperationsTemplateDirectory := system["Shared Directory"] . "Spreadsheet Operations Template\"
    versionManifestFilePath := spreadsheetOperationsTemplateDirectory . "Version Manifest.ini"
    version := StrReplace(version, "v", "")

    if version = "" {
        latestVersion := ""
        latestDate    := ""

        sectionList := IniRead(versionManifestFilePath)
        loop parse sectionList, "`n", "`r" {
            candidateVersion := A_LoopField
            if candidateVersion = "" {
                continue
            }

            candidateDate := IniRead(versionManifestFilePath, candidateVersion, "ReleaseDate", "")
            if latestDate = "" || StrCompare(candidateDate, latestDate) > 0 {
                latestVersion := candidateVersion
                latestDate    := candidateDate
            }
        }
        version := latestVersion
    }

    releaseDate := IniRead(versionManifestFilePath, version, "ReleaseDate", "")
    introHash   := IniRead(versionManifestFilePath, version, "IntroSHA-256", "")
    outroHash   := IniRead(versionManifestFilePath, version, "OutroSHA-256", "")

    try {
        if releaseDate = "" {
            throw Error("Version not found: " . version)
        }
    } catch as versionNotFoundError {
        LogInformationConclusion("Failed", logValuesForConclusion, versionNotFoundError)
    }

    templateCombined := Map(
        "Version",       version,
        "Release Date",  releaseDate,
        "Intro SHA-256", introHash,
        "Outro SHA-256", outroHash
    )

    templateCombined["Intro Code"] := ReadFileOnHashMatch(spreadsheetOperationsTemplateDirectory . "Spreadsheet Operations Template (v" version ", " releaseDate ") Intro.vba", templateCombined["Intro SHA-256"])
    templateCombined["Outro Code"] := ReadFileOnHashMatch(spreadsheetOperationsTemplateDirectory . "Spreadsheet Operations Template (v" version ", " releaseDate ") Outro.vba", templateCombined["Outro SHA-256"])

    LogInformationConclusion("Completed", logValuesForConclusion)
    return templateCombined
}

ModifyScreenCoordinates(horizontalValue, verticalValue, coordinatePair) {
    static methodName := RegisterMethod("ModifyScreenCoordinates(horizontalValue As Integer, verticalValue As Integer, coordinatePair As String [Type: Coordinate Pair])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Modify Screen Coordinates (" . horizontalValue . "x" . verticalValue . ", " . coordinatePair . ")", methodName, [horizontalValue, verticalValue, coordinatePair])

    widthDisplayResolution  := A_ScreenWidth
    heightDisplayResolution := A_ScreenHeight

    coordinates := StrSplit(Trim(coordinatePair), "x")
    originalX := coordinates[1] + 0
    originalY := coordinates[2] + 0

    newX := originalX + horizontalValue
    newY := originalY + verticalValue
    modifiedCoordinatePair := Format("{}x{}", newX, newY)

    try {
        if newX < 0 || newX > widthDisplayResolution - 1 {
            throw Error("X out of bounds. Tried " . newX . " (valid 0 to " . (widthDisplayResolution - 1) . ").")
        }
    } catch as xOutOfBoundsError {
        LogInformationConclusion("Failed", logValuesForConclusion, xOutOfBoundsError)
    }

    try {
        if newY < 0 || newY > heightDisplayResolution - 1 {
            throw Error("Y out of bounds. Tried " . newY . " (valid 0 to " . (heightDisplayResolution - 1) . ").")
        }
    } catch as yOutOfBoundsError {
        LogInformationConclusion("Failed", logValuesForConclusion, yOutOfBoundsError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
    return modifiedCoordinatePair
}

PasteCode(code, commentPrefix) {
    static commentPrefixWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "'",  "--", "#", "%", "//", ";")
    static methodName := RegisterMethod("PasteCode(code As String [Type: Code], commentPrefix As String [Whitelist: " . commentPrefixWhitelist . "]", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Paste Code (Length: " . StrLen(code) . ")", methodName, [code, commentPrefix])

    sentinel := commentPrefix . " == AutoHotkey Paste Sentinel == " . commentPrefix
    code := code . "`r`n" . sentinel
    
    attempts    := 0
    maxAttempts := 4
    sleepAmount := 360
    success     := false

    while attempts < maxAttempts {
        attempts++
        sleepAmount := sleepAmount + (attempts * 40)

        if attempts >= 2 {
            logValuesForConclusion["Context"] := "Retrying, attempt " attempts " of " maxAttempts ". Sleep amount is currently " . sleepAmount . " milliseconds."
        }

        SendEvent("^a") ; CTRL+A (Select All)
        Sleep(sleepAmount/2)
        SendEvent("^a") ; CTRL+A (Select All)
        Sleep(sleepAmount/2)
        SendEvent("{Delete}") ; Delete (Delete)
        Sleep(sleepAmount/2)

        A_Clipboard := code ; Load combined code into clipboard.
        if !ClipWait(1 * attempts) { ; Clipboard not ready, go to next attempt.
            continue
        }
        SendEvent("^v") ; CTRL+V (Paste)
        Sleep(sleepAmount + sleepAmount)

        ; Verify the paste by reading the sentinel line.
        SendEvent("+{Home}") ; SHIFT+HOME (Select the whole last line)
        Sleep(sleepAmount/2)
        SendEvent("^c") ; CTRL+C (Copy)
        Sleep(sleepAmount)

        if A_Clipboard !== sentinel {
            continue ; Sentinel content not copied, go to next attempt.
        }

        SendEvent("{Delete}")
        Sleep(sleepAmount/2)
        SendEvent("{Backspace}")
        Sleep(sleepAmount/2)

        success := true
        if attempts >= 2 {
            logValuesForConclusion["Context"] := "Succeeded on attempt " attempts " of " maxAttempts ". Sleep amount is currently " . sleepAmount . " milliseconds."
        }
        break
    }

    if !success {
        try {
            throw Error("Paste of code failed.")
        } catch as pasteOfCodeFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, pasteOfCodeFailedError)
        }
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

PastePath(savePath) {
    static methodName := RegisterMethod("PastePath(savePath As String [Type: Absolute Save Path])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Paste Path (" . savePath . ")", methodName, [savePath])

    attempts    := 0
    maxAttempts := 4
    sleepAmount := 200
    success     := false

    while attempts < maxAttempts {
        attempts++
        sleepAmount := sleepAmount + (attempts * 20)

        if attempts >= 2 {
            logValuesForConclusion["Context"] := "Retrying, attempt " attempts " of " maxAttempts ". Sleep amount is currently " . sleepAmount . " milliseconds."
        }
        
        SendEvent("{End}") ; END (End of Line)
        Sleep(sleepAmount)
        SendEvent("+{Home}") ; SHIFT+HOME (Select the full line)
        Sleep(sleepAmount/2)
        SendEvent("{Delete}") ; Delete (Delete)
        Sleep(sleepAmount/2)
        if attempts != maxAttempts {
            SendText(savePath)
            Sleep(sleepAmount/2)
        } else {
            for character in StrSplit(savePath) {
                SendEvent("{Raw}" . character)
                Sleep(102)
            }
        }

        ; Verify the paste by reading the sentinel line.
        SendEvent("+{Home}") ; SHIFT+HOME (Select the whole last line)
        Sleep(sleepAmount)
        SendEvent("^c") ; CTRL+C (Copy)
        Sleep(sleepAmount)

        if A_Clipboard !== savePath {
            continue ; Clipboard content does not match Save Path, go to next attempt.
        }

        SendEvent("{End}") ; END (End of Line)

        success := true
        if attempts >= 2 {
            logValuesForConclusion["Context"] := "Succeeded on attempt " attempts " of " maxAttempts ". Sleep amount is currently " . sleepAmount . " milliseconds."
        }
        break
    }

    if !success {
        try {
            throw Error("Paste of path failed.")
        } catch as pasteOfPathFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, pasteOfPathFailedError)
        }
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

PasteSearch(searchValue) {
    static methodName := RegisterMethod("PasteSearch(searchValue As String [Type: Search Open])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Paste Search (" . searchValue . ")", methodName, [searchValue])

    attempts    := 0
    maxAttempts := 4
    sleepAmount := 200
    success     := false

    while attempts < maxAttempts {
        attempts++
        sleepAmount := sleepAmount + (attempts * 20)

        if attempts >= 2 {
            logValuesForConclusion["Context"] := "Retrying, attempt " attempts " of " maxAttempts ". Sleep amount is currently " . sleepAmount . " milliseconds."
        }

        SendEvent("{End}") ; END (End of Line)
        Sleep(sleepAmount)
        SendEvent("+{Home}") ; SHIFT+HOME (Select the full line)
        Sleep(sleepAmount/2)
        SendEvent("{Delete}") ; Delete (Delete)
        Sleep(sleepAmount/2)
        if attempts != maxAttempts {
            SendText(searchValue)
            Sleep(sleepAmount + sleepAmount)
        } else {
            for character in StrSplit(searchValue) {
                SendEvent("{Raw}" . character)
                Sleep(102)
            }
        }

        ; Verify the paste by reading the sentinel line.
        SendEvent("+{Home}") ; SHIFT+HOME (Select the whole last line)
        Sleep(sleepAmount)
        SendEvent("^c") ; CTRL+C (Copy)
        Sleep(sleepAmount)

        if A_Clipboard !== searchValue {
            continue ; Clipboard content does not match Save Path, go to next attempt.
        }

        SendEvent("{End}") ; END (End of Line)

        success := true
        if attempts >= 2 {
            logValuesForConclusion["Context"] := "Succeeded on attempt " attempts " of " maxAttempts ". Sleep amount is currently " . sleepAmount . " milliseconds."
        }
        break
    }

    if !success {
        try {
            throw Error("Paste of search failed.")
        } catch as pasteOfPathFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, pasteOfPathFailedError)
        }
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

PerformMouseActionAtCoordinates(mouseAction, coordinatePair) {
    static mouseActionWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}"', "Double", "Left", "Middle", "Move", "Move Smooth", "Right", "Wheel Down", "Wheel Up")
    static methodName := RegisterMethod("PerformMouseActionAtCoordinates(mouseAction As String [Whitelist: " . mouseActionWhitelist . "], coordinatePair As String [Type: Coordinate Pair])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Perform Mouse Action at Coordinates (" . mouseAction . " @ " . coordinatePair . ")", methodName, [mouseAction, coordinatePair])

    widthDisplayResolution  := A_ScreenWidth
    heightDisplayResolution := A_ScreenHeight

    coordinates := StrSplit(coordinatePair, "x")
    x := coordinates[1] + 0
    y := coordinates[2] + 0

    overlayVisibility := OverLayIsVisible()
    if overlayVisibility {
        OverlayChangeVisibility()
    }

    modeBeforeAction := A_CoordModeMouse
    CoordMode("Mouse", "Screen")
    
    switch StrLower(mouseAction) {
        case "double":
            Click("left", x, y, 2)
        case "left":
            Click("left", x, y)
        case "middle":
            Click("middle", x, y)
        case "move":
            MouseMove(x, y, 0)
        case "move smooth":
            originalSendMode := A_SendMode
            SendMode("Event")
            MouseGetPos(&currentMouseX, &currentMouseY)
            MouseMove(x, y, ComputeMouseSpeed(currentMouseX . "x" . currentMouseY, coordinatePair))
            SendMode(originalSendMode)
        case "right":
            Click("right", x, y)
        case "wheel down":
            Click("WheelDown", x, y)
        case "wheel up":
            Click("WheelUp", x, y)
    }

    CoordMode("Mouse", modeBeforeAction)

    if overlayVisibility {
        OverlayChangeVisibility()
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

PerformMouseDragBetweenCoordinates(startCoordinatePair, endCoordinatePair, mouseButton := "Left", modifierKeys := "") {
    static mouseActionWhitelist := Format('"{1}", "{2}"', "Left", "Right")
    static methodName := RegisterMethod("PerformMouseDragBetweenCoordinates(startCoordinatePair As String [Type: Coordinate Pair], endCoordinatePair As String [Type: Coordinate Pair], mouseButton As String [Whitelist: " . mouseActionWhitelist . "], modifierKeys As String [Optional])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("PerformMouseDrag (" . mouseButton . ", " . startCoordinatePair . " to " . endCoordinatePair . ")", methodName, [startCoordinatePair, endCoordinatePair, mouseButton, modifierKeys])

    modeBeforeAction := A_CoordModeMouse
    CoordMode("Mouse", "Screen")

    widthDisplayResolution  := A_ScreenWidth
    heightDisplayResolution := A_ScreenHeight

    startCoordinates := StrSplit(startCoordinatePair, "x")
    startX := startCoordinates[1] + 0
    startY := startCoordinates[2] + 0

    endCoordinates := StrSplit(endCoordinatePair, "x")
    endX := endCoordinates[1] + 0
    endY := endCoordinates[2] + 0

    ; Parse modifiers robustly (supports + , space, tab as separators).
    normalizedModifierList := []
    seenModifierMap := Map()

    if modifierKeys != "" {
        ; Tokenize on any of: + , space, tab.
        loop parse modifierKeys, "+, " . "`t" {
            rawToken := A_LoopField
            if rawToken = "" {
                continue
            }

            tokenLowercase := StrLower(Trim(rawToken))

            ; Map synonyms to canonical AHK Send key names.
            if tokenLowercase = "shift" {
                canonical := "Shift"
            } else if tokenLowercase = "lshift" || tokenLowercase = "leftshift" {
                canonical := "LShift"
            } else if tokenLowercase = "rshift" || tokenLowercase = "rightshift" {
                canonical := "RShift"
            } else if tokenLowercase = "ctrl" || tokenLowercase = "control" || tokenLowercase = "ctl" {
                canonical := "Ctrl"
            } else if tokenLowercase = "lctrl" || tokenLowercase = "lcontrol" || tokenLowercase = "leftctrl" {
                canonical := "LCtrl"
            } else if tokenLowercase = "rctrl" || tokenLowercase = "rcontrol" || tokenLowercase = "rightctrl" {
                canonical := "RCtrl"
            } else if tokenLowercase = "alt" {
                canonical := "Alt"
            } else if tokenLowercase = "lalt" || tokenLowercase = "leftalt" {
                canonical := "LAlt"
            } else if tokenLowercase = "ralt" || tokenLowercase = "rightalt" || tokenLowercase = "altgr" {
                canonical := "RAlt"
            } else if tokenLowercase = "win" || tokenLowercase = "windows" || tokenLowercase = "meta" || tokenLowercase = "super" {
                canonical := "LWin"
            } else if tokenLowercase = "lwin" || tokenLowercase = "leftwin" || tokenLowercase = "winleft" {
                canonical := "LWin"
            } else if tokenLowercase = "rwin" || tokenLowercase = "rightwin" || tokenLowercase = "winright" {
                canonical := "RWin"
            } else {
                try {
                    Throw Error("Unsupported modifier: " . rawToken)
                } catch as unsupportedModifierError {
                    LogInformationConclusion("Failed", logValuesForConclusion, unsupportedModifierError)
                }
            }

            if !seenModifierMap.Has(canonical) {
                normalizedModifierList.Push(canonical)
                seenModifierMap[canonical] := true
            }
        }
    }

    ; Press modifiers down.
    for index, modifierName in normalizedModifierList {
        Send "{" modifierName " down}"
    }

    originalSendMode := A_SendMode
    SendMode("Event")
    MouseMove(startX, startY, 0)
    Sleep(16)
    MouseClickDrag(StrLower(mouseButton), startX, startY, endX, endY, ComputeMouseSpeed(startCoordinatePair, endCoordinatePair))
    SendMode(originalSendMode)

    ; Release modifiers in reverse order.
    loop normalizedModifierList.Length {
        reverseIndex := normalizedModifierList.Length - A_Index + 1
        Send "{" normalizedModifierList[reverseIndex] " up}"
    }

    CoordMode("Mouse", modeBeforeAction)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

ComputeMouseSpeed(startCoordinatePair, endCoordinatePair) {
    static methodName := RegisterMethod("ComputeMouseSpeed(startCoordinatePair As String [Type: Coordinate Pair], endCoordinatePair As String [Type: Coordinate Pair])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [startCoordinatePair, endCoordinatePair])

    startCoordinates := StrSplit(startCoordinatePair, "x")
    startX := startCoordinates[1] + 0
    startY := startCoordinates[2] + 0

    endCoordinates := StrSplit(endCoordinatePair, "x")
    endX := endCoordinates[1] + 0
    endY := endCoordinates[2] + 0

    deltaX := endX - startX
    deltaY := endY - startY
    movementDistancePixels := Sqrt(deltaX*deltaX + deltaY*deltaY)

    screenWidthPixels  := A_ScreenWidth
    screenHeightPixels := A_ScreenHeight
    screenDiagonalPixels := Sqrt(screenWidthPixels*screenWidthPixels + screenHeightPixels*screenHeightPixels)

    nearDistanceThresholdPixels := 100
    farDistanceThresholdPixels  := Round(screenDiagonalPixels * 0.60)
    if farDistanceThresholdPixels <= nearDistanceThresholdPixels {
        farDistanceThresholdPixels := nearDistanceThresholdPixels + 1
    }

    nearSpeed := 80
    farSpeed  := 20
    computedMouseSpeed := unset
    if movementDistancePixels <= nearDistanceThresholdPixels {
        computedMouseSpeed := nearSpeed
    } else if movementDistancePixels >= farDistanceThresholdPixels {
        computedMouseSpeed := farSpeed
    } else {
        ratio := (movementDistancePixels - nearDistanceThresholdPixels)  / (farDistanceThresholdPixels - nearDistanceThresholdPixels)
        ratio := ratio ** 1.5
        computedMouseSpeed := Round(nearSpeed + (farSpeed - nearSpeed) * ratio)
    }

    return computedMouseSpeed
}

ConvertArrayIntoCsvString(array) {
    static methodName := RegisterMethod("ConvertArrayIntoCsvString(array As Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [array])

    static newLine := "`r`n"

    result := ""
    for index, value in array {
        if index > 1 {
            result .= newLine
        }

        result .= value
    }

    return result
}

ConvertHexStringToBase64(hexString, removePadding := true) {
    static methodName := RegisterMethod("ConvertHexStringToBase64(hexString As String [Type: Hexadecimal String], removePadding As Boolean [Optional: true])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [hexString])

    byteCount := StrLen(hexString) // 2
    binaryBuffer := Buffer(byteCount)
    loop byteCount {
        characterIndex := (A_Index - 1) * 2 + 1
        byteHexadecimalPair := SubStr(hexString, characterIndex, 2)
        computedByteValue := ("0x" . byteHexadecimalPair) + 0
        NumPut("UChar", computedByteValue, binaryBuffer, A_Index - 1)
    }

    static CRYPT_STRING_BASE64 := 0x1
    static CRYPT_STRING_NOCRLF := 0x40000000
    static encodingFlags := CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF

    requiredCharacterCount := 0
    sizeProbeRetrievedSuccessfully := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", binaryBuffer.Ptr, "UInt", binaryBuffer.Size, "UInt", encodingFlags, "Ptr", 0, "UInt*", &requiredCharacterCount, "Int")
    if !sizeProbeRetrievedSuccessfully {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve size probe. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    outputUtf16Buffer := Buffer(requiredCharacterCount * 2)
    encodingSuccessful := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", binaryBuffer.Ptr, "UInt", binaryBuffer.Size, "UInt", encodingFlags, "Ptr", outputUtf16Buffer.Ptr, "UInt*", &requiredCharacterCount, "Int")
    if !encodingSuccessful {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to encode. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    base64 := StrGet(outputUtf16Buffer.Ptr, "UTF-16")
    if removePadding {
        base64 := RegExReplace(base64, "=+$")
    }   
        
    return base64
}

ExtractRowFromArrayOfMapsOnHeaderCondition(rowsAsMaps, headerName, targetValue) {
    static methodName := RegisterMethod("ExtractRowFromArrayOfMapsOnHeaderCondition(rowsAsMaps As Object, headerName As String, targetValue As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [rowsAsMaps, headerName, targetValue])

    foundRow := unset
    for rowMap in rowsAsMaps {
        if !rowMap.Has(headerName) {
            continue
        }

        if rowMap[headerName] = targetValue {
            foundRow := rowMap
            break
        }
    }

    if !IsSet(foundRow) {
        LogHelperError(logValuesForConclusion, A_LineNumber, "No row found where '" . headerName . "' = '" . targetValue . "'.")
    }

    return foundRow
}

ExtractValuesFromArrayDimension(array, dimension) {
    static methodName := RegisterMethod("ExtractValuesFromArrayDimension(array As Object, dimension As Integer)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [array, dimension])

    arrayDimension := []

    for outerIndex, innerArray in array {
        arrayDimension.Push(innerArray[dimension])
    }

    return arrayDimension
}

ExtractUniqueValuesFromSubMaps(parentMapOfMaps, subMapKeyName) {
    static methodName := RegisterMethod("ExtractUniqueValuesFromSubMaps(parentMapOfMaps As Object, subMapKeyName As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [parentMapOfMaps, subMapKeyName])

    uniqueValues := []

    for outerMapKey, innerMap in parentMapOfMaps {
        if !innerMap.Has(subMapKeyName) {
            continue
        }

        currentValue := innerMap[subMapKeyName]

        if currentValue = "" {
            continue
        }

        valueExists := false
        for index, existingValue in uniqueValues {
            if existingValue = currentValue {
                valueExists := true
                break
            }
        }

        if !valueExists {
            uniqueValues.Push(currentValue)
        }
    }

    return uniqueValues
}

GetBase64FromFile(filePath) {
    static methodName := RegisterMethod("GetBase64FromFile(filePath As String [Type: Absolute Path])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [filePath])

    fileContentBuffer := FileRead(filePath, "RAW")

    static CRYPT_STRING_BASE64 := 0x1
    static CRYPT_STRING_NOCRLF := 0x40000000
    static base64Flags := CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF

    requiredCharacters := 0
    sizeProbeRetrievedSuccessfully := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", fileContentBuffer.Ptr, "UInt", fileContentBuffer.Size, "UInt", base64Flags, "Ptr", 0, "UInt*", &requiredCharacters, "Int")
    if !sizeProbeRetrievedSuccessfully {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve size probe. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    outputUtf16Buffer := Buffer(requiredCharacters * 2, 0)
    encodingSuccessful := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", fileContentBuffer.Ptr, "UInt", fileContentBuffer.Size, "UInt", base64Flags, "Ptr", outputUtf16Buffer.Ptr, "UInt*", &requiredCharacters, "Int")
    if !encodingSuccessful {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to encode. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    base64Output := StrGet(outputUtf16Buffer.Ptr, "UTF-16")

    return base64Output
}

IfStringIsNotEmptyReturnValue(stringValue, returnValue) {
    static methodName := RegisterMethod("IfStringIsNotEmptyReturnValue(stringValue As String [Optional], returnValue As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [stringValue, returnValue])

    If stringValue = "" {
        returnValue := ""
    }

    return returnValue
}

RemoveDuplicatesFromArray(array) {
    static methodName := RegisterMethod("RemoveDuplicatesFromArray(array As Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [array])

    seen := Map()
    index := array.Length

    while index >= 1 {
        currentValue := array[index]
        if seen.Has(currentValue) {
            array.RemoveAt(index)
        } else {
            seen[currentValue] := true
        }

        index -= 1
    }

    return array
}