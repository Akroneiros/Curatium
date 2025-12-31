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
    static methodName := RegisterMethod("ModifyScreenCoordinates(horizontalValue As Integer, verticalValue As Integer, coordinatePair As String [Constraint: Coordinate Pair])", A_LineFile, A_LineNumber + 1)
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

PasteText(text, commentPrefix := "") {
    static commentPrefixWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "'",  "--", "#", "%", "//", ";")
    static methodName := RegisterMethod("PasteText(text As String [Constraint: Summary], commentPrefix As String [Optional] [Whitelist: " . commentPrefixWhitelist . "]", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Paste Text)", methodName, [text, commentPrefix])

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Max Attempts", 4)
        SetMethodSetting(methodName, "Short Delay", 120)
        SetMethodSetting(methodName, "Medium Delay", 260)
        SetMethodSetting(methodName, "Short Accumulation", 20)
        SetMethodSetting(methodName, "Medium Accumulation", 40)

        defaultMethodSettingsSet := true
    }

    settings    := methodRegistry[methodName]["Settings"]
    maxAttempts := settings.Get("Max Attempts")
    shortDelay  := settings.Get("Short Delay")
    mediumDelay := settings.Get("Medium Delay")
    shortAccumulation  := settings.Get("Short Accumulation")
    mediumAccumulation := settings.Get("Medium Accumulation")

    rows := StrSplit(text, "`n").Length

    pasteSentinel := commentPrefix . " == AutoHotkey Paste Sentinel == " . commentPrefix
    if rows != 1 {
        text := text . "`r`n" . pasteSentinel
    }
    
    attempts := 0
    success  := false

    while attempts < maxAttempts {
        if attempts >= 2 {
            logValuesForConclusion["Context"] := "Failed on attempt " attempts " of " maxAttempts ". Short delay is currently " . shortDelay . " milliseconds. Medium delay is currently " . mediumDelay . " milliseconds. " . 
                "Short Accumulation is currently " . shortAccumulation . " milliseconds. Medium Accumulation is currently " . mediumAccumulation . " milliseconods."
        }

        attempts++
        mediumDelay := mediumDelay + (attempts * mediumAccumulation)
        shortDelay  := shortDelay + (attempts * shortAccumulation)

        if rows = 1 {
            SendInput("{End}") ; End of Line.
            Sleep(mediumDelay)
            KeyboardShortcut("SHIFT", "HOME") ; Select the full line.
            Sleep(shortDelay)
            SendInput("{Delete}") ; Delete
        } else {
            KeyboardShortcut("CTRL", "A") ; Select All
            Sleep(shortDelay)
            SendInput("{Delete}") ; Delete
        }

        Sleep(shortDelay)

        if rows = 1 { 
            if attempts != maxAttempts {
                SendText(text)
                Sleep(shortDelay)
            } else {
                for character in StrSplit(text) {
                    SendEvent("{Raw}" . character)
                    Sleep(102)
                }
            }

            KeyboardShortcut("SHIFT", "HOME") ; Select the whole last line
            Sleep(mediumDelay)
            KeyboardShortcut("CTRL", "C") ; Copy
            ClipWait()
            Sleep(mediumDelay)

            if A_Clipboard !== text {
                continue ; Clipboard does not match, go to next attempt.
            }
        } else {
            A_Clipboard := text ; Load combined text into clipboard.
            Sleep(shortDelay)
            KeyboardShortcut("CTRL", "V") ; Paste
            Sleep(mediumDelay + mediumDelay)
            KeyboardShortcut("SHIFT", "HOME") ; Select the whole last line which should be the sentintel.
            Sleep(shortDelay)
            KeyboardShortcut("SHIFT", "LEFT") ; Select one character more to the left.
            Sleep(shortDelay)
            KeyboardShortcut("CTRL", "X") ; Cut
            ClipWait()
            Sleep(mediumDelay)
            clipboardSentinel := StrReplace(StrReplace(A_Clipboard, "`r", ""), "`n", "")

            if clipboardSentinel !== pasteSentinel {
                continue ; Paste Sentinel not copied, go to next attempt.
            }
        }

        success := true
        if attempts >= 2 {
            logValuesForConclusion["Context"] := "Succeeded on attempt " attempts " of " maxAttempts ". Short delay is currently " . shortDelay . " milliseconds. Medium delay is currently " . mediumDelay . " milliseconds. " . 
                "Short Accumulation is currently " . shortAccumulation . " milliseconds. Medium Accumulation is currently " . mediumAccumulation . " milliseconods."
        }
        break
    }

    if !success {
        try {
            throw Error("Paste of text failed.")
        } catch as pasteOfTextFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, pasteOfTextFailedError)
        }
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

PerformMouseActionAtCoordinates(mouseAction, coordinatePair) {
    static mouseActionWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}"', "Double", "Left", "Middle", "Move", "Move Smooth", "Right", "Wheel Down", "Wheel Up")
    static methodName := RegisterMethod("PerformMouseActionAtCoordinates(mouseAction As String [Whitelist: " . mouseActionWhitelist . "], coordinatePair As String [Constraint: Coordinate Pair])", A_LineFile, A_LineNumber + 1)
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
    static methodName := RegisterMethod("PerformMouseDragBetweenCoordinates(startCoordinatePair As String [Constraint: Coordinate Pair], endCoordinatePair As String [Constraint: Coordinate Pair], mouseButton As String [Whitelist: " . mouseActionWhitelist . "], modifierKeys As String [Optional])", A_LineFile, A_LineNumber + 1)
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

SetAutoHotkeyThreadPriority(threadPriority) {
    static threadPriorityWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}"', "Lowest", "Below Normal", "Normal", "Above Normal", "Highest")
    static methodName := RegisterMethod("SetAutoHotkeyThreadPriority(threadPriority As String [Whitelist: " . threadPriorityWhitelist . "])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Set AutoHotkey Thread Priority (" . threadPriority . ")", methodName, [threadPriority])

    switch threadPriority {
        case "Lowest":
            threadPriority := -2
        case "Below Normal":
            threadPriority := -1
        case "Normal":
            threadPriority := 0
        case "Above Normal":
            threadPriority := 1
        case "Highest":
            threadPriority := 2
    }

    autoHotkeyThreadHandle := DllCall("GetCurrentThread", "Ptr")
    DllCall("SetThreadPriority", "Ptr", autoHotkeyThreadHandle, "Int", threadPriority)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

ValidateDisplayScaling() {
    static methodName := RegisterMethod("ValidateDisplayScaling()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Validate Display Scaling", methodName)

    validateDisplayResolution := ValidateDataUsingSpecification(system["Display Resolution"], "String", "Display Resolution")

    try {
        if validateDisplayResolution != "" {
            throw Error(validateDisplayResolution)
        }
    } catch as invalidDisplayResolutionError {
        LogInformationConclusion("Failed", logValuesForConclusion, invalidDisplayResolutionError)
    }

    validateDpiScale := ValidateDataUsingSpecification(system["DPI Scale"], "String", "DPI Scale")

    try {
        if validateDpiScale != "" {
            throw Error(validateDpiScale)
        }
    } catch as invalidDpiScaleError {
         LogInformationConclusion("Failed", logValuesForConclusion, invalidDpiScaleError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

CombineCode(introCode, mainCode, outroCode := "") {
    static methodName := RegisterMethod("CombineCode(introCode As String, mainCode As String, outroCode As String [Optional])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [introCode, mainCode, outroCode])

    combinedCode := introCode . "`r`n`r`n" . mainCode

    if outroCode != "" {
        combinedCode := combinedCode . "`r`n`r`n" . outroCode
    } 

    return combinedCode
}

ComputeMouseSpeed(startCoordinatePair, endCoordinatePair) {
    static methodName := RegisterMethod("ComputeMouseSpeed(startCoordinatePair As String [Constraint: Coordinate Pair], endCoordinatePair As String [Constraint: Coordinate Pair])", A_LineFile, A_LineNumber + 1)
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
    static methodName := RegisterMethod("ConvertHexStringToBase64(hexString As String [Constraint: Hexadecimal String], removePadding As Boolean [Optional: true])", A_LineFile, A_LineNumber + 1)
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

GetAutoHotkeyThreadPriority() {
    static methodName := RegisterMethod("GetAutoHotkeyThreadPriority()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    autoHotkeyThreadHandle := DllCall("GetCurrentThread", "Ptr")
    threadPriority := DllCall("GetThreadPriority", "Ptr", autoHotkeyThreadHandle, "Int")

    return threadPriority
}

GetBase64FromFile(filePath) {
    static methodName := RegisterMethod("GetBase64FromFile(filePath As String [Constraint: Absolute Path])", A_LineFile, A_LineNumber + 1)
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

KeyboardShortcut(modifier, key) {
    if StrLen(key) > 1 {
        key := "{" . key . "}"
    } else {
        key := StrLower(key)
    }

    switch modifier, false {
        case "ALT":
            SendInput("{Alt down}" . key . "{Alt up}")
        case "CTRL", "CONTROL":
            SendInput("{Ctrl down}" . key . "{Ctrl up}")
        case "SHIFT":
            SendInput("{Shift down}" . key . "{Shift up}")
        case "WIN", "WINDOWS":
            SendInput("{LWin down}" . key . "{LWin up}")
    }
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

SetMethodSetting(methodName, settingName, settingValue, override := false) {
    global methodRegistry

    if !methodRegistry[methodName]["Settings"].Has(settingName) || override {
        methodRegistry[methodName]["Settings"][settingName] := settingValue
    }
}

ValidateDataUsingSpecification(dataValue, dataType, dataConstraint := "", whitelist := []) {
    validation := ""

    static hexadecimalAllowedCharactersMessage     := "Only 0–9, A–F, and a–f are allowed."
    static windowsInvalidFilenameCharactersPattern := '[\\/:*?"<>|]'
    static windowsInvalidFilenameCharactersList    := '\ / : * ? " < > |'
    static windowsReservedDeviceNamesPattern       := "i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$"

    static resolutions := unset
    static scales      := unset

    switch dataType {
        case "Array":
            if Type(dataValue) != "Array" {
                validation := "Value must be an Array."
            }
        case "Boolean":
            if !(Type(dataValue) = "Integer" && (dataValue = 0 || dataValue = 1)) {
                validation := "Boolean must be an Integer with value 0 or 1."
            }
        case "Integer":
            if Type(dataValue) != "Integer" {
                validation := "Value must be an Integer."
            } else {
                switch dataConstraint {
                    case "Byte":
                        if dataValue < 0 || dataValue > 255 {
                            validation := dataConstraint . " out of range (0–255)."
                        }
                    case "Day":
                        if dataValue < 1 || dataValue > 31 {
                            validation := dataConstraint . " must be between 1 and 31."
                        }
                    case "Hour":
                        if dataValue < 0 || dataValue > 23 {
                            validation := dataConstraint . " must be between 0 and 23."
                        }
                    case "Minute":
                        if dataValue < 0 || dataValue > 59 {
                            validation := dataConstraint . " must be between 0 and 59."
                        }
                    case "Month":
                        if dataValue < 1 || dataValue > 12 {
                            validation := dataConstraint . " must be between 1 and 12."
                        }
                    case "Second":
                        if dataValue < 0 || dataValue > 59 {
                            validation := dataConstraint . " must be between 0 and 59."
                        }
                    case "Year":
                        if dataValue < 1601 || dataValue > 9999 {
                            validation := dataConstraint . " must be between 1601 and 9999."
                        }
                }
            }
        case "Object":
            ; No validation, but allowed Data Type.
        case "String":
            if whitelist.Length != 0 {
                valueIsWhitelisted := false

                for index, whitelistEntry in whitelist {
                    if dataValue = whitelistEntry {
                        valueIsWhitelisted := true
                        break
                    }
                }

                if !valueIsWhitelisted {
                    validation := "Whitelisted values did not match value."
                }
            } else if Type(dataValue) != "String" {
                validation := "Value must be a String."
            } else {
                switch dataConstraint {
                    case "Absolute Path", "Absolute Save Path":
                        isDrive := RegExMatch(dataValue, "^[A-Za-z]:\\")
                        isUNC   := RegExMatch(dataValue, "^\\\\{2}[^\\\/]+\\[^\\\/]+\\")

                        if !(isDrive || isUNC) {
                            validation := dataConstraint . " must start with a drive (C:\) or UNC path (\\server\share\)."
                        } else if !DirExist(ExtractDirectory(dataValue)) {
                            validation := dataConstraint . " Directory doesn't exist."
                        } else if !FileExist(dataValue) && dataConstraint = "Absolute Path" {
                            validation := dataConstraint . " File doesn't exist."
                        }
                    case "Base64":
                        if !RegExMatch(dataValue, "^[A-Za-z0-9+/]*={0,2}$") {
                            validation := dataConstraint . " invalid characters. Only A–Z, a–z, 0–9, +, /, and = allowed."
                        } else if Mod(StrLen(dataValue), 4) != 0 {
                            validation := dataConstraint . " invalid length. Length must be a multiple of 4."
                        } else if RegExMatch(dataValue, "=[^=]") {
                            validation := dataConstraint . " invalid padding. The character = can only appear at the end."
                        }
                    case "Coordinate Pair":
                        widthDisplayResolution  := A_ScreenWidth
                        heightDisplayResolution := A_ScreenHeight

                        if !RegExMatch(dataValue, "^(?<x>\d+)x(?<y>\d+)$", &matchObject) {
                            validation := dataConstraint . " is not formatted correctly."
                        } else if (
                            (x := matchObject["x"] + 0), (y := matchObject["y"] + 0), (x < 0 || x >= widthDisplayResolution || y < 0 || y >= heightDisplayResolution)
                        ) {
                            if x < 0 || x >= widthDisplayResolution {
                                validation := dataConstraint . " has X out of bounds. Valid 0 to " . (widthDisplayResolution - 1) . "."
                            } else {
                                validation := dataConstraint . " has Y out of bounds. Valid 0 to " . (heightDisplayResolution - 1) . "."
                            }
                        }
                    case "Directory":
                        isDrive := RegExMatch(dataValue, "^[A-Za-z]:\\")
                        isUNC   := RegExMatch(dataValue, "^\\\\{2}[^\\\/]+\\[^\\\/]+\\")

                        if !(isDrive || isUNC) {
                            validation := dataConstraint . " path must start with a drive (C:\) or UNC path (\\server\share\)."
                        } else if !DirExist(dataValue) {
                            validation := dataConstraint . " doesn't exist."
                        } else if SubStr(dataValue, -1) != "\" {
                            validation := dataConstraint . " path must end with a backslash \."
                        }
                    case "Display Resolution":
                        if !IsSet(resolutions) {
                            resolutions := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Resolutions (2025-09-20).csv")
                        }

                        validation := dataConstraint . " is invalid."
                        for resolution in resolutions {
                            if resolution["Resolution"] = dataValue {
                                validation := ""
                                break
                            }
                        }
                    case "DPI Scale":
                        if !IsSet(scales) {
                            scales := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Scales (2025-09-20).csv")
                        }

                        validation := dataConstraint . " is invalid."
                        for scale in scales {
                            if scale["Scale"] = dataValue {
                                validation := ""
                                break
                            }
                        }
                    case "Filename":
                        if RegExMatch(dataValue, "[\x00-\x1F]") {
                            validation := dataConstraint . " contains control characters (ASCII 0–31)."
                        } else if RegExMatch(dataValue, windowsInvalidFilenameCharactersPattern) {
                            validation := dataConstraint . " contains invalid characters (" . windowsInvalidFilenameCharactersList . ")."
                        } else if dataValue = "." || dataValue = ".." {
                            validation := dataConstraint . " is reserved."
                        } else if RegExMatch(dataValue, windowsReservedDeviceNamesPattern) {
                            validation := dataConstraint . " uses a reserved device name (CON, PRN, AUX, NUL, COM1–COM9, LPT1–LPT9)."
                        } else if Trim(dataValue, " .") = "" {
                            validation := dataConstraint . " cannot consist only of spaces or periods."
                        } else if RegExMatch(dataValue, "[\. ]$") {
                            validation := dataConstraint . " cannot end with a space or period."
                        }
                    case "Hexadecimal String":
                        if Mod(StrLen(dataValue), 2) != 0 {
                            validation := dataConstraint . " must contain an even number of characters (two characters per byte)."
                        } else {
                            totalCharacterCount := StrLen(dataValue)
                            loop totalCharacterCount {
                                currentCharacter := SubStr(dataValue, A_Index, 1)
                                currentCharacterCode := Ord(currentCharacter)

                                isHexadecimalDigit := (currentCharacterCode >= 48 && currentCharacterCode <= 57) || (currentCharacterCode >= 65 && currentCharacterCode <= 70) || (currentCharacterCode >= 97 && currentCharacterCode <= 102)

                                if !isHexadecimalDigit {
                                    validation := dataConstraint . " contains an invalid character '" . currentCharacter . "' at position " . A_Index . ". " . hexadecimalAllowedCharactersMessage
                                    break
                                }
                            }
                        }
                    case "ISO Date", "ISO Date Time", "Raw Date Time":
                        if dataConstraint = "ISO Date" && !RegExMatch(dataValue, "^\d{4}-\d{2}-\d{2}$") {
                            validation := dataConstraint . " is invalid (must be YYYY-MM-DD)."
                        } else if dataConstraint = "ISO Date Time" && !RegExMatch(dataValue, "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$") {
                            validation := dataConstraint . " is invalid (must be YYYY-MM-DD HH:MM:SS)."
                        } else if dataConstraint = "Raw Date Time" && !RegExMatch(dataValue, "^\d{14}$") {
                            validation := dataConstraint . " is invalid (must be YYYYMMDDHHMMSS)."
                        } else {
                            if dataConstraint = "Raw Date Time" {
                                dataValue :=
                                    SubStr(dataValue, 1, 4)  . "-" . SubStr(dataValue, 5, 2)  . "-" .  SubStr(dataValue, 7, 2) .
                                    " " . SubStr(dataValue, 9, 2)  . ":" . SubStr(dataValue, 11, 2) . ":" . SubStr(dataValue, 13, 2)
                            }

                            dateTimeparts := StrSplit(dataValue, " ")
                            dateParts     := StrSplit(dateTimeparts[1], "-")
                            year   := dateParts[1] + 0
                            month  := dateParts[2] + 0
                            day    := dateParts[3] + 0

                            if validation = "" {
                                validation := ValidateDataUsingSpecification(year, "Integer", "Year")
                            }

                            if validation = "" {
                                validation := ValidateDataUsingSpecification(month, "Integer", "Month")
                            }

                            if validation = "" {
                                validation := ValidateDataUsingSpecification(day, "Integer", "Day")
                            }

                            if validation = "" {
                                isLeap := (Mod(year, 400) = 0) || (Mod(year, 4) = 0 && Mod(year, 100) != 0)
                                daysInMonth := [31, (isLeap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

                                if day > daysInMonth[month] {
                                    validation := "Day is invalid for this month and year."
                                }
                            }

                            if dataConstraint != "ISO Date" && validation = "" {
                                timeParts := StrSplit(dateTimeparts[2], ":")
                                hour   := timeParts[1] + 0
                                minute := timeParts[2] + 0
                                second := timeParts[3] + 0

                                if validation = "" {
                                    validation := ValidateDataUsingSpecification(hour, "Integer", "Hour")
                                }

                                if validation = "" {
                                    validation := ValidateDataUsingSpecification(minute, "Integer", "Minute")
                                }

                                if validation = "" {
                                    validation := ValidateDataUsingSpecification(second, "Integer", "Second")
                                }
                            }

                            if validation != "" {
                                validation := dataConstraint . " " . validation
                            }
                        }
                    case "Locator":
                        if InStr(dataValue, "|") {
                            validation := dataConstraint . " contains the character |."
                        }
                    case "Percent Range":
                        dataValue := StrReplace(dataValue, ",", ".")

                        if !RegExMatch(dataValue, "^\d{1,3}(?:\.\d)?-\d{1,3}(?:\.\d)?$") {
                            validation := dataConstraint . " must be two numbers (0–100) with up to one decimal, separated by -."
                        } else {
                            parts  := StrSplit(dataValue, "-")
                            first  := parts[1] + 0
                            second := parts[2] + 0

                            if first = 100 {
                                validation := dataConstraint . " first value maximum allowed is 99.9."
                            } else if first < 0 || first > 100 || second < 0 || second > 100 {
                                validation := dataConstraint . " values must be between 0 and 100."
                            } else if first >= second {
                                validation := dataConstraint . " first value must be lower than second."
                            }
                        }
                    case "SHA-256":
                        if StrLen(dataValue) != 64 {
                            validation := dataConstraint . " expected length is 64 but instead got length of: " . StrLen(dataValue) "."
                        } else if !RegExMatch(dataValue, "^[0-9a-fA-F]+$") {
                            validation := dataConstraint . " must be hex digits only. " . hexadecimalAllowedCharactersMessage
                        }
                }
            }
        default:
            validation := "Data Type is invalid: " . dataType . "."
    }

    return validation
}