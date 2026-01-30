#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include ..\jsongo_AHKv2 (2025-02-26)\jsongo.v2.ahk
#Include Logging Library.ahk

global methodRegistry := Map()

AssignSpreadsheetOperationsTemplateCombined(version := "") {
    static methodName := RegisterMethod("version As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 7)
    overlayValue := "Assign Spreadsheet Operations Template Code"
    if version = "" {
        overlayValue := overlayValue . " ([Latest])"
    } else {
        overlayValue := overlayValue . " (" . version . ")"
    }
    logValuesForConclusion := LogBeginning(methodName, [version], overlayValue)

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

    if releaseDate = "" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Version not found: " . version)
    }

    templateCombined := Map(
        "Version",       version,
        "Release Date",  releaseDate,
        "Intro SHA-256", introHash,
        "Outro SHA-256", outroHash
    )

    templateCombined["Intro Code"] := ReadFileOnHashMatch(spreadsheetOperationsTemplateDirectory . "Spreadsheet Operations Template (v" version ", " releaseDate ") Intro.vba", templateCombined["Intro SHA-256"])
    templateCombined["Outro Code"] := ReadFileOnHashMatch(spreadsheetOperationsTemplateDirectory . "Spreadsheet Operations Template (v" version ", " releaseDate ") Outro.vba", templateCombined["Outro SHA-256"])

    LogConclusion("Completed", logValuesForConclusion)
    return templateCombined
}

ModifyScreenCoordinates(horizontalValue, verticalValue, coordinatePair) {
    static methodName := RegisterMethod("horizontalValue As Integer, verticalValue As Integer, coordinatePair As String [Constraint: Coordinate Pair]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [horizontalValue, verticalValue, coordinatePair], "Modify Screen Coordinates (" . horizontalValue . "x" . verticalValue . ", " . coordinatePair . ")")

    widthDisplayResolution  := A_ScreenWidth
    heightDisplayResolution := A_ScreenHeight

    coordinates := StrSplit(Trim(coordinatePair), "x")
    originalX := coordinates[1] + 0
    originalY := coordinates[2] + 0

    newX := originalX + horizontalValue
    newY := originalY + verticalValue
    modifiedCoordinatePair := Format("{}x{}", newX, newY)

    if newX < 0 || newX > widthDisplayResolution - 1 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "X out of bounds. Tried " . newX . " (valid 0 to " . (widthDisplayResolution - 1) . ").")
    }

    if newY < 0 || newY > heightDisplayResolution - 1 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Y out of bounds. Tried " . newY . " (valid 0 to " . (heightDisplayResolution - 1) . ").")
    }

    LogConclusion("Completed", logValuesForConclusion)
    return modifiedCoordinatePair
}

PasteText(text, commentPrefix := "") {
    static commentPrefixWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "'", "--", "#", "%", "//", ";")
    static methodName := RegisterMethod("text As String [Constraint: Summary], commentPrefix As String [Optional] [Whitelist: " . commentPrefixWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [text, commentPrefix], "Paste Text")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Max Attempts", 4, false)
        SetMethodSetting(methodName, "Short Delay", 120, false)
        SetMethodSetting(methodName, "Medium Delay", 260, false)
        SetMethodSetting(methodName, "Short Accumulation", 20, false)
        SetMethodSetting(methodName, "Medium Accumulation", 40, false)

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
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Paste of text failed.")
    }

    LogConclusion("Completed", logValuesForConclusion)
}

PerformMouseActionAtCoordinates(mouseAction, coordinatePair) {
    static mouseActionWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}"', "Double", "Left", "Middle", "Move", "Move Smooth", "Right", "Wheel Down", "Wheel Up")
    static methodName := RegisterMethod("mouseAction As String [Whitelist: " . mouseActionWhitelist . "], coordinatePair As String [Constraint: Coordinate Pair]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [mouseAction, coordinatePair], "Perform Mouse Action at Coordinates (" . mouseAction . " @ " . coordinatePair . ")")

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

    LogConclusion("Completed", logValuesForConclusion)
}

PerformMouseDragBetweenCoordinates(startCoordinatePair, endCoordinatePair, mouseButton := "Left", modifierKeys := "") {
    static mouseActionWhitelist := Format('"{1}", "{2}"', "Left", "Right")
    static methodName := RegisterMethod("startCoordinatePair As String [Constraint: Coordinate Pair], endCoordinatePair As String [Constraint: Coordinate Pair], mouseButton As String [Whitelist: " . mouseActionWhitelist . "], modifierKeys As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [startCoordinatePair, endCoordinatePair, mouseButton, modifierKeys], "PerformMouseDrag (" . mouseButton . ", " . startCoordinatePair . " to " . endCoordinatePair . ")")

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
                LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Unsupported modifier: " . rawToken)
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

    LogConclusion("Completed", logValuesForConclusion)
}

SetAutoHotkeyThreadPriority(threadPriority) {
    static threadPriorityWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}"', "Lowest", "Below Normal", "Normal", "Above Normal", "Highest")
    static methodName := RegisterMethod("threadPriority As String [Whitelist: " . threadPriorityWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [threadPriority], "Set AutoHotkey Thread Priority (" . threadPriority . ")")

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

    LogConclusion("Completed", logValuesForConclusion)
}

ValidateConfiguration(configurationPath) {
    static methodName := RegisterMethod("configurationPath As String [Constraint: Absolute Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Validate Configuration")

    global system

    jsongo.silent_error := false
    try {
        system["Configuration"] := jsongo.Parse(FileRead(configurationPath))
    } catch {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to load Configuration File due to invalid JSON.")
    }

    if Type(system["Configuration"]) != "Map" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Configuration is unexpectedly not a Map.")
    }

    LogConclusion("Completed", logValuesForConclusion)
}

ValidateDisplayScaling() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Validate Display Scaling")

    validateDisplayResolution := ValidateDataUsingSpecification(system["Display Resolution"], "String", "Display Resolution")

    if validateDisplayResolution != "" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, validateDisplayResolution)
    }

    validateDpiScale := ValidateDataUsingSpecification(system["DPI Scale"], "String", "DPI Scale")

    if validateDpiScale != "" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, validateDpiScale)
    }

    LogConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Core Methods                 ;
; **************************** ;

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
        case "Variant":
            if whitelist.Length != 0 {
                validation := "Whitelist not supported for Data Type of Variant."
            } else if Type(dataValue) != "Integer" && Type(dataValue) != "String" {
                validation := "Value must be an Integer or a String."
            }
        default:
            validation := "Data Type is invalid: " . dataType . "."
    }

    return validation
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

ActivateWindow(windowTitle, customErrorMessage := "") {
    static methodName := RegisterMethod("windowTitle As String, customErrorMessage As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [windowTitle, customErrorMessage])

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Seconds to Attempt", 60, false)
        SetMethodSetting(methodName, "Short Delay", 128, false)

        defaultMethodSettingsSet := true
    }

    settings         := methodRegistry[methodName]["Settings"]
    secondsToAttempt := settings.Get("Seconds to Attempt")
    shortDelay       := settings.Get("Short Delay")

    windowHandle := WinWait(windowTitle, , secondsToAttempt)
    if !windowHandle {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to find window after trying for " . secondsToAttempt . " seconds." . IfStringIsNotEmptyReturnValue(customErrorMessage, " " . customErrorMessage))
    }

    totalSleep := Round((Round(shortDelay / 2)) + shortDelay * (2 + 3 + 4 + 5 + 6 + 7 + 8 + 9 + 10))
    loop 10 {
        loopDelay := shortDelay * A_Index
        if A_Index = 1 {
            loopDelay := Round(loopDelay/2)
        }
        Sleep(loopDelay)

        try {
            WinActivate("ahk_id " . windowHandle)
            break
        } catch {
            if A_Index = 10 {
                LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to activate window after trying for " . totalSleep . " milliseconds." . IfStringIsNotEmptyReturnValue(customErrorMessage, " " . customErrorMessage))
            }
        }
    }

    windowHandle := WinWaitActive("ahk_id " . windowHandle, , secondsToAttempt)
    if !windowHandle {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to activate window after waiting for " . secondsToAttempt . " seconds." . IfStringIsNotEmptyReturnValue(customErrorMessage, " " . customErrorMessage))
    }

    return windowHandle
}

CombineCode(introCode, mainCode, outroCode := "") {
    static methodName := RegisterMethod("introCode As String, mainCode As String, outroCode As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [introCode, mainCode, outroCode])

    combinedCode := introCode . "`r`n`r`n" . mainCode

    if outroCode != "" {
        combinedCode := combinedCode . "`r`n`r`n" . outroCode
    } 

    return combinedCode
}

ComputeMouseSpeed(startCoordinatePair, endCoordinatePair) {
    static methodName := RegisterMethod("startCoordinatePair As String [Constraint: Coordinate Pair], endCoordinatePair As String [Constraint: Coordinate Pair]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [startCoordinatePair, endCoordinatePair])

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
    static methodName := RegisterMethod("array As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [array])

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
    static methodName := RegisterMethod("hexString As String [Constraint: Hexadecimal String], removePadding As Boolean [Optional: true]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [hexString])

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
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to retrieve size probe. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    outputUtf16Buffer := Buffer(requiredCharacterCount * 2)
    encodingSuccessful := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", binaryBuffer.Ptr, "UInt", binaryBuffer.Size, "UInt", encodingFlags, "Ptr", outputUtf16Buffer.Ptr, "UInt*", &requiredCharacterCount, "Int")
    if !encodingSuccessful {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to encode. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    base64 := StrGet(outputUtf16Buffer.Ptr, "UTF-16")
    if removePadding {
        base64 := RegExReplace(base64, "=+$")
    }   
        
    return base64
}

ExtractRowFromArrayOfMapsOnHeaderCondition(rowsAsMaps, headerName, targetValue) {
    static methodName := RegisterMethod("rowsAsMaps As Object, headerName As String, targetValue As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [rowsAsMaps, headerName, targetValue])

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
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "No row found where '" . headerName . "' = '" . targetValue . "'.")
    }

    return foundRow
}

ExtractValuesFromArrayDimension(array, dimension) {
    static methodName := RegisterMethod("array As Object, dimension As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [array, dimension])

    arrayDimension := []

    for outerIndex, innerArray in array {
        arrayDimension.Push(innerArray[dimension])
    }

    return arrayDimension
}

ExtractUniqueValuesFromSubMaps(parentMapOfMaps, subMapKeyName) {
    static methodName := RegisterMethod("parentMapOfMaps As Object, subMapKeyName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [parentMapOfMaps, subMapKeyName])

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
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath])

    fileContentBuffer := FileRead(filePath, "RAW")

    static CRYPT_STRING_BASE64 := 0x1
    static CRYPT_STRING_NOCRLF := 0x40000000
    static base64Flags := CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF

    requiredCharacters := 0
    sizeProbeRetrievedSuccessfully := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", fileContentBuffer.Ptr, "UInt", fileContentBuffer.Size, "UInt", base64Flags, "Ptr", 0, "UInt*", &requiredCharacters, "Int")
    if !sizeProbeRetrievedSuccessfully {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to retrieve size probe. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    outputUtf16Buffer := Buffer(requiredCharacters * 2, 0)
    encodingSuccessful := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", fileContentBuffer.Ptr, "UInt", fileContentBuffer.Size, "UInt", base64Flags, "Ptr", outputUtf16Buffer.Ptr, "UInt*", &requiredCharacters, "Int")
    if !encodingSuccessful {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to encode. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    base64Output := StrGet(outputUtf16Buffer.Ptr, "UTF-16")

    return base64Output
}

IfStringIsNotEmptyReturnValue(stringValue, returnValue) {
    static methodName := RegisterMethod("stringValue As String [Optional], returnValue As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [stringValue, returnValue])

    If stringValue = "" {
        returnValue := ""
    }

    return returnValue
}

KeyboardShortcut(modifier, key) {
    static modifierWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "ALT", "CTRL", "CONTROL", "SHIFT", "WIN", "WINDOWS")
    static methodName := RegisterMethod("modifier As String [Whitelist: " . modifierWhitelist . "], key As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [modifier, key])

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Tiny Delay", 32, false)

        defaultMethodSettingsSet := true
    }

    settings  := methodRegistry[methodName]["Settings"]
    tinyDelay := settings.Get("Tiny Delay")

    if StrLen(key) > 1 {
        key := "{" . key . "}"
    } else {
        key := StrLower(key)
    }

    switch modifier, false {
        case "ALT":
            SendInput("{Alt down}")
        case "CTRL", "CONTROL":
            SendInput("{Ctrl down}")
        case "SHIFT":
            SendInput("{Shift down}")
        case "WIN", "WINDOWS":
            SendInput("{LWin down}")
    }

    Sleep(tinyDelay + tinyDelay)
    SendInput(key)
    Sleep(tinyDelay)

    switch modifier, false {
        case "ALT":
            SendInput("{Alt up}")
        case "CTRL", "CONTROL":
            SendInput("{Ctrl up}")
        case "SHIFT":
            SendInput("{Shift up}")
        case "WIN", "WINDOWS":
            SendInput("{LWin up}")
    }
}

RemoveDuplicatesFromArray(array) {
    static methodName := RegisterMethod("array As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [array])

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

SetMethodSetting(settingMethod, settingName, settingValue, override := true) {
    static methodName := RegisterMethod("settingMethod As String, settingName As String, settingValue As Variant, override As Boolean [Optional: true]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [settingMethod, settingName, settingValue, override])

    global methodRegistry

    if !methodRegistry.Has(settingMethod) {
        methodRegistry[settingMethod] := Map()
        methodRegistry[settingMethod]["Symbol"] := ""
    }

    if !methodRegistry[settingMethod].Has("Settings") {
        methodRegistry[settingMethod]["Settings"] := Map()
    }

    if !methodRegistry[settingMethod]["Settings"].Has(settingName) || override {
        methodRegistry[settingMethod]["Settings"][settingName] := settingValue
    }
}

; **************************** ;
; System Methods               ;
; **************************** ;

GetInternationalFormatting() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    internationalRegistryKeyPath := "HKEY_CURRENT_USER\Control Panel\International"
    registryValueData            := ""
    internationalFormattingMap   := Map()

    excludedRegistryValueNames := [
        "iCountry",
        "iLocale",
        "iPaperSize",
        "Locale",
        "LocaleName",
        "sIntlCurrency",
        "sLanguage"
    ]

    loop reg, internationalRegistryKeyPath, "V" {
        registryValueName := A_LoopRegName
        skipValue := false

        for excludedRegistryValueName in excludedRegistryValueNames {
            if excludedRegistryValueName = registryValueName {
                skipValue := true
                break
            }
        }

        if skipValue {
            continue
        }

        try {
        	registryValueData := RegRead(internationalRegistryKeyPath, registryValueName)
        }

        internationalFormattingMap[registryValueName] := registryValueData
    }

    internationalFormatting := internationalFormattingMap
    
    return internationalFormatting
}

GetOperatingSystem() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    currentVersionRegistryKey := "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion"

    family := "Unknown Windows"
    edition := "Unknown Edition"
    architectureTag := A_Is64bitOS ? "x64" : "x86"
    currentBuildNumber := ""
    version := "Unknown Version"
    updateBuildRevisionNumber := ""
    releaseDisplay := ""

    try {
        currentBuildNumber := RegRead(currentVersionRegistryKey, "CurrentBuildNumber") + 0
        version := "Build " . currentBuildNumber
    }
        
    switch true {
        case currentBuildNumber >= 22000:
            family := "Windows 11"
        case currentBuildNumber >= 10240:
            family := "Windows 10"
        case currentBuildNumber >= 9600:
            family := "Windows 8.1"
        case currentBuildNumber >= 9200:
            family := "Windows 8"
        case currentBuildNumber >= 7600:
            family := "Windows 7"
    }
   
    try {
        edition := RegRead(currentVersionRegistryKey, "EditionID")
    }

    try {
        updateBuildRevisionNumber := RegRead(currentVersionRegistryKey, "UBR")

        if updateBuildRevisionNumber != "" && currentBuildNumber >= 9200 {
            version := version . "." . updateBuildRevisionNumber
        }
    }

    if family = "Windows 7" {
        servicePackVersion := ""

        try {
            servicePackVersion := RegRead(currentVersionRegistryKey, "CSDVersion")
        }

        releaseDisplay := (servicePackVersion != "" ? " (" . servicePackVersion . ")" : "")
    } else {
        try {
            releaseDisplay := RegRead(currentVersionRegistryKey, "DisplayVersion")
        }

        if releaseDisplay = "" {
            try {
                releaseDisplay := RegRead(currentVersionRegistryKey, "ReleaseId")
            }
        }

        if releaseDisplay != "" {
            releaseDisplay := " (" . releaseDisplay . ")"
        }
    }

    operatingSystem := "Microsoft " . family . " " . edition . " " . "(" . architectureTag . ")" . " " . version . releaseDisplay

    return operatingSystem
}

GetWindowsInstallationDateUtcTimestamp() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    registryKeySystemSetup               := "HKEY_LOCAL_MACHINE\System\Setup"
    registryKeyCurrentVersionNonWow      := "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion"
    registryKeyCurrentVersionWow6432Node := "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows NT\CurrentVersion"

    installDateSeconds := unset

    ; Prefer the oldest DWORD InstallDate under SYSTEM\Setup\Source OS (...).
    oldestSeconds := unset
    loop reg, registryKeySystemSetup, "K" {
        subKeyName := A_LoopRegName
        if !RegExMatch(subKeyName, "^Source OS") {
            continue
        }

        fullKey := registryKeySystemSetup . "\" . subKeyName
        try {
            candidate := RegRead(fullKey, "InstallDate")
            candidate := candidate & 0xFFFFFFFF ; Treat as unsigned DWORD.
        }

        if candidate > 0 && (!IsSet(oldestSeconds) || candidate < oldestSeconds) {
            oldestSeconds := candidate
        }
    }

    if IsSet(oldestSeconds) {
        installDateSeconds := oldestSeconds
    }

    ; Fallback: CurrentVersion DWORD InstallDate (prefer non-WOW, then WOW6432Node).
    if !IsSet(installDateSeconds) {
        try {
            installDateSeconds := RegRead(registryKeyCurrentVersionNonWow, "InstallDate")
        }
    }

    if !IsSet(installDateSeconds) {
        try {
            installDateSeconds := RegRead(registryKeyCurrentVersionWow6432Node, "InstallDate")
        }
    }

    if !IsSet(installDateSeconds) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "InstallDate (DWORD) not found in SYSTEM\Setup snapshots or CurrentVersion.")
    }

    installDateSeconds := installDateSeconds & 0xFFFFFFFF
    if installDateSeconds <= 0 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Invalid InstallDate seconds: " . installDateSeconds)
    }

    utcTimestamp := ConvertUnixTimeToUtcTimestamp(installDateSeconds)

    return utcTimestamp
}

GetComputerIdentifier() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    computerIdentifier := "N/A"

    machineGuid := ""
    try {
        machineGuid := RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Cryptography", "MachineGuid")
    }

    systemUniversallyUniqueIdentifier := ""
    try {
        windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
        windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\CIMV2")
        windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3

        win32ComputerSystemProductQuery := "
        (
            SELECT
                UUID
            FROM
                Win32_ComputerSystemProduct
        )"

        queryResults := windowsManagementInstrumentationService.ExecQuery(win32ComputerSystemProductQuery)
        for record in queryResults {
            systemUniversallyUniqueIdentifier := record.UUID
            break
        }
    }

    if machineGuid = "" && systemUniversallyUniqueIdentifier != "" {
        computerIdentifier := systemUniversallyUniqueIdentifier
    } else if systemUniversallyUniqueIdentifier = "" && machineGuid != "" {
        computerIdentifier := machineGuid
    } else if machineGuid != "" && systemUniversallyUniqueIdentifier != "" {
        computerIdentifier := machineGuid . " + " . systemUniversallyUniqueIdentifier
    }

    return computerIdentifier
}

GetTimeZoneKeyName() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    timeZoneKeyName := "Unknown"

    dynamicTimeZoneInformationBuffer := Buffer(432, 0)
    callResult := DllCall("Kernel32\GetDynamicTimeZoneInformation", "Ptr", dynamicTimeZoneInformationBuffer, "UInt")
    if callResult != 0xFFFFFFFF {
        extractedKey := StrGet(dynamicTimeZoneInformationBuffer.Ptr + 172, 128, "UTF-16")

        if extractedKey != "" {
            timeZoneKeyName := extractedKey
        }
    }

    if timeZoneKeyName = "Unknown" {
        try {
            regValue := RegRead("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation", "TimeZoneKeyName")
        }

        if regValue != "" {
            timeZoneKeyName := regValue
        }
    }

    if timeZoneKeyName = "Unknown" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to retrieve time zone information.")
    }

    return timeZoneKeyName
}

GetRegionFormat() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    LOCALE_NAME_MAX_LENGTH := 85
    BYTES_PER_WIDE_CHAR    := 2
    localeNameBuffer       := Buffer(LOCALE_NAME_MAX_LENGTH * BYTES_PER_WIDE_CHAR, 0)

    regionFormat := "Unknown"

    regionFormatFromRegistry := ""
    try {
        regionFormatFromRegistry := RegRead("HKEY_CURRENT_USER\Control Panel\International", "LocaleName", "")
    }

    if regionFormatFromRegistry != "" {
        regionFormat := regionFormatFromRegistry
    } else {
        wasLocaleResolved := DllCall("Kernel32\GetUserDefaultLocaleName", "Ptr", localeNameBuffer.Ptr, "Int", LOCALE_NAME_MAX_LENGTH, "Int")
        if wasLocaleResolved {
            resolvedLocaleName := StrGet(localeNameBuffer)

            if resolvedLocaleName != "" {
                regionFormat := resolvedLocaleName
            }
        }
    }

    return regionFormat
}

GetInputLanguage() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    LOCALE_NAME_MAX_LENGTH := 85
    BYTES_PER_WIDE_CHAR    := 2
    localeNameUtf16Buffer  := Buffer(LOCALE_NAME_MAX_LENGTH * BYTES_PER_WIDE_CHAR, 0)

    inputLanguageName := "Unknown Language"

    keyboardLayoutHandle := DllCall("User32\GetKeyboardLayout", "UInt", 0, "Ptr")
    if keyboardLayoutHandle {
        languageIdentifier := keyboardLayoutHandle & 0xFFFF
        localeIdentifier   := languageIdentifier

        wasLcidToLocaleNameSuccessful := DllCall("Kernel32\LCIDToLocaleName", "UInt", localeIdentifier, "Ptr", localeNameUtf16Buffer.Ptr, "Int", LOCALE_NAME_MAX_LENGTH, "UInt", 0, "Int")

        if wasLcidToLocaleNameSuccessful {
            resolvedLocaleName := StrGet(localeNameUtf16Buffer)

            if resolvedLocaleName != "" {
                inputLanguageName := resolvedLocaleName
            }
        }
    }

    return inputLanguageName
}

GetActiveKeyboardLayout() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    KEYBOARD_LAYOUT_ID_LENGTH_CHARACTERS := 9
    LOCALE_NAME_MAX_LENGTH               := 85
    MAX_PATH_CHARACTERS                  := 260
    BYTES_PER_WIDE_CHAR                  := 2

    keyboardLayoutIdBuffer    := Buffer(KEYBOARD_LAYOUT_ID_LENGTH_CHARACTERS * BYTES_PER_WIDE_CHAR, 0)
    localeNameBuffer          := Buffer(LOCALE_NAME_MAX_LENGTH * BYTES_PER_WIDE_CHAR, 0)
    indirectDisplayNameBuffer := Buffer(MAX_PATH_CHARACTERS * BYTES_PER_WIDE_CHAR, 0)
    immDescBuffer             := Buffer(MAX_PATH_CHARACTERS * BYTES_PER_WIDE_CHAR, 0)

    resultLocaleName            := "Unknown Language"
    resultKeyboardLayoutName    := "Unknown Layout"
    resultKeyboardLayoutId      := "????????"
    currentKeyboardLayoutHandle := 0

    wasKeyboardLayoutIdResolved := DllCall("User32\GetKeyboardLayoutNameW", "Ptr", keyboardLayoutIdBuffer.Ptr, "Int")
    if wasKeyboardLayoutIdResolved {
        resolvedKeyboardLayoutIdText := StrGet(keyboardLayoutIdBuffer)

        if resolvedKeyboardLayoutIdText != "" {
            resultKeyboardLayoutId := resolvedKeyboardLayoutIdText
        }
    }

    if resultKeyboardLayoutId != "????????" {
        languageIdentifierHex := SubStr(resultKeyboardLayoutId, 5)
        if languageIdentifierHex != "" {
            languageIdentifier := ("0x" . languageIdentifierHex) + 0
            wasLocaleNameResolved := DllCall("Kernel32\LCIDToLocaleName", "UInt", languageIdentifier, "Ptr", localeNameBuffer.Ptr, "Int", LOCALE_NAME_MAX_LENGTH, "UInt", 0, "Int")
            if wasLocaleNameResolved {
                resolvedLocaleName := StrGet(localeNameBuffer)

                if resolvedLocaleName != "" {
                    resultLocaleName := resolvedLocaleName
                } else {
                    resultLocaleName := languageIdentifierHex
                }
            } else {
                resultLocaleName := languageIdentifierHex
            }
        }
    }

    if resultKeyboardLayoutId != "????????" {
        baseRegKey := "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Keyboard Layouts\" . resultKeyboardLayoutId

        layoutText := ""
        try {
            layoutText := RegRead(baseRegKey, "Layout Text")
        }

        if layoutText = "" {
            try {
                layoutText := RegRead(baseRegKey, "Layout Display Name")
            }

            if layoutText != "" && SubStr(layoutText, 1, 1) = "@" {
                wasIndirectLoaded := (DllCall("Shlwapi\SHLoadIndirectString", "WStr", layoutText, "Ptr", indirectDisplayNameBuffer.Ptr, "Int", MAX_PATH_CHARACTERS, "Ptr", 0, "Int") = 0)

                if wasIndirectLoaded {
                    maybeResolved := StrGet(indirectDisplayNameBuffer)
                    if maybeResolved != "" {
                        layoutText := maybeResolved
                    }
                }
            }
        }

        if layoutText != "" {
            resultKeyboardLayoutName := layoutText
        }
    } else {
        currentKeyboardLayoutHandle := DllCall("User32\GetKeyboardLayout", "UInt", 0, "Ptr")
        if currentKeyboardLayoutHandle {
            languageIdentifierFromHandle := currentKeyboardLayoutHandle & 0xFFFF
            wasLocaleNameResolved := DllCall("Kernel32\LCIDToLocaleName", "UInt", languageIdentifierFromHandle, "Ptr", localeNameBuffer.Ptr, "Int", LOCALE_NAME_MAX_LENGTH, "UInt", 0, "Int")
            if wasLocaleNameResolved {
                maybeLocale := StrGet(localeNameBuffer)
                if maybeLocale != "" {
                    resultLocaleName := maybeLocale
                }
            }
        }
    }

    if resultKeyboardLayoutName = "Unknown Layout" {
        if !currentKeyboardLayoutHandle {
            currentKeyboardLayoutHandle := DllCall("User32\GetKeyboardLayout", "UInt", 0, "Ptr")
        }
        if currentKeyboardLayoutHandle {
            immDescriptionLength := DllCall("Imm32\ImmGetDescriptionW", "Ptr", currentKeyboardLayoutHandle, "Ptr", immDescBuffer.Ptr, "UInt", MAX_PATH_CHARACTERS, "UInt")
            if immDescriptionLength > 0 {
                maybeImmName := StrGet(immDescBuffer)
                if maybeImmName != "" {
                    resultKeyboardLayoutName := maybeImmName
                }
            }
        }
    }

    keyboardLayoutDescription := resultLocaleName . " - " . resultKeyboardLayoutName . " (" . resultKeyboardLayoutId . ")"

    return keyboardLayoutDescription
}

GetMotherboard() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    rawManufacturer := ""
    rawProduct := ""

    try {
        windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
        windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\CIMV2")
        windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3

        win32BaseBoardQuery := "
        (
            SELECT
                Manufacturer,
                Product
            FROM
                Win32_BaseBoard
        )"

        queryResults := windowsManagementInstrumentationService.ExecQuery(win32BaseBoardQuery)
        for record in queryResults {
            try {
                rawManufacturer := Trim(record.Manufacturer . "")
            }

            try {
                rawProduct := Trim(record.Product . "")
            }
            
            break
        }
    }

    if rawManufacturer = "" {
        rawManufacturer := "Unknown Manufacturer"
    }

    if rawProduct = "" {
        rawProduct := "Unknown Product"
    }

    motherboard := Trim(rawManufacturer . " " . rawProduct)

    return motherboard
}

GetCpu() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    modelName := ""
    defaultModelName := "Unknown CPU"
    registryPath := "HKEY_LOCAL_MACHINE\Hardware\Description\System\CentralProcessor\0"
    registryValueName := "ProcessorNameString"

    rawName := ""
    cleanedName := ""
    try {
        rawName := RegRead(registryPath, registryValueName)
        cleanedName := Trim(RegExReplace(rawName, "\s+", " "))
    }

    if cleanedName != "" {
        modelName := cleanedName
    }

    if modelName = "" {
        modelName := defaultModelName
    }

    return modelName
}

GetMemorySizeAndType() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    systemManagementBiosType17MemoryDeviceTypes := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "System Management BIOS Type 17 Memory Device - Type.csv")

    ramValues := Map()
    for systemManagementBiosType17MemoryDeviceType in systemManagementBiosType17MemoryDeviceTypes {
        ramValues[systemManagementBiosType17MemoryDeviceType["Value"] + 0] := systemManagementBiosType17MemoryDeviceType["Meaning"]
    }

    memoryTypeDetailFlagCounts := Map()
    memoryTypeCodeCounts := Map()
    partNumberStrings := []
    installedMemoryTypeDisplay := ""
    resolvedLegacySubtype := ""
    resolvedLegacyCount := -1

    installedKilobytes := 0
    retrievedTheAmountOfRamPhysicallyInstalledSuccessfully := DllCall("Kernel32\GetPhysicallyInstalledSystemMemory", "UInt64*", &installedKilobytes, "Int")

    if !retrievedTheAmountOfRamPhysicallyInstalledSuccessfully {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to retrieve the amount of RAM that is physically installed on the computer. [Kernel32\GetPhysicallyInstalledSystemMemory" . ", System Error Code: " . A_LastError . "]")
    }

    installedMemorySizeInGigabytes := (installedKilobytes > 0) ? (installedKilobytes // 1048576) : 0
    installedMemorySizeDisplay := installedMemorySizeInGigabytes ? (installedMemorySizeInGigabytes . " GB") : "Unknown Size"

    try {
        windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
        windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\CIMV2")
        windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3

        win32PhysicalMemoryQuery := "
        (
            SELECT
                SMBIOSMemoryType,
                PartNumber
            FROM
                Win32_PhysicalMemory
        )"

        queryResults := windowsManagementInstrumentationService.ExecQuery(win32PhysicalMemoryQuery)
        for record in queryResults {
            systemManagementBiosMemoryTypeCode := ""
            try {
                systemManagementBiosMemoryTypeCode := record.SMBIOSMemoryType + 0
            }

            if systemManagementBiosMemoryTypeCode >= 3 {
                if !memoryTypeCodeCounts.Has(systemManagementBiosMemoryTypeCode) {
                    memoryTypeCodeCounts[systemManagementBiosMemoryTypeCode] := 0
                }

                memoryTypeCodeCounts[systemManagementBiosMemoryTypeCode] += 1
            }

            partNumberValue := ""
            try {
                partNumberValue := Trim(record.PartNumber . "")
            }

            if partNumberValue != "" {
                partNumberStrings.Push(partNumberValue)
            }
        }
    }

    if memoryTypeCodeCounts.Count > 0 {
        mostCommonMemoryTypeCode := ""
        mostCommonMemoryTypeCount := -1
        for memoryTypeCode, memoryTypeCount in memoryTypeCodeCounts {
            if memoryTypeCount > mostCommonMemoryTypeCount {
                mostCommonMemoryTypeCount := memoryTypeCount
                mostCommonMemoryTypeCode := memoryTypeCode
            }
        }

        if ramValues.Has(mostCommonMemoryTypeCode) {
            installedMemoryTypeDisplay := ramValues[mostCommonMemoryTypeCode]
        } else {
            installedMemoryTypeDisplay := "Unknown Type (code " . mostCommonMemoryTypeCode . ")"
        }
    } else {
        ; No usable rows from Win32_PhysicalMemory. Parse raw SMBIOS (Type 17) to obtain MemoryType.
        try {
            windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\WMI")

            mssmBiosRawSMBiosTablesQuery := "
            (
                SELECT
                    SMBiosData
                FROM
                    MSSMBios_RawSMBiosTables
            )"

            rawTablesResults := windowsManagementInstrumentationService.ExecQuery(mssmBiosRawSMBiosTablesQuery)

            rawSmbiosByteArray := ""
            for rawTablesRecord in rawTablesResults {
                rawSmbiosByteArray := rawTablesRecord.SMBiosData
                break
            }

            if IsSet(rawSmbiosByteArray) && rawSmbiosByteArray != "" {
                totalByteCount := rawSmbiosByteArray.MaxIndex() + 1
                if totalByteCount > 0 {
                    rawSmbiosBuffer := Buffer(totalByteCount, 0)
                    copyIndex := 0
                    while copyIndex < totalByteCount {
                        NumPut("UChar", rawSmbiosByteArray[copyIndex], rawSmbiosBuffer, copyIndex)
                        copyIndex += 1
                    }

                    parseOffset := 0
                    while parseOffset + 4 <= totalByteCount {
                        structureType   := NumGet(rawSmbiosBuffer, parseOffset + 0, "UChar")
                        structureLength := NumGet(rawSmbiosBuffer, parseOffset + 1, "UChar")
                        if structureLength < 4 {
                            break
                        }

                        nextStructureOffset := parseOffset + structureLength
                        while nextStructureOffset + 1 < totalByteCount {
                            byteA := NumGet(rawSmbiosBuffer, nextStructureOffset + 0, "UChar")
                            byteB := NumGet(rawSmbiosBuffer, nextStructureOffset + 1, "UChar")
                            nextStructureOffset += 1
                            if byteA = 0 && byteB = 0 {
                                nextStructureOffset += 1
                                break
                            }
                        }

                        if structureType = 17 { ; Memory Device.
                            if structureLength >= 0x13 {
                                systemManagementBiosMemoryTypeCodeFromRaw := NumGet(rawSmbiosBuffer, parseOffset + 0x12, "UChar")
                                if systemManagementBiosMemoryTypeCodeFromRaw >= 3 {
                                    if !memoryTypeCodeCounts.Has(systemManagementBiosMemoryTypeCodeFromRaw) {
                                        memoryTypeCodeCounts[systemManagementBiosMemoryTypeCodeFromRaw] := 0
                                    }
                                    memoryTypeCodeCounts[systemManagementBiosMemoryTypeCodeFromRaw] += 1
                                }
                            }

                            ; Always read TypeDetail if present (even when MemoryType is Unknown/Other/DRAM).
                            if structureLength >= 0x15 {
                                memoryTypeDetailWordFromRaw := NumGet(rawSmbiosBuffer, parseOffset + 0x13, "UShort")

                                if memoryTypeDetailWordFromRaw & 0x0100 {
                                    if !memoryTypeDetailFlagCounts.Has("EDO") {
                                        memoryTypeDetailFlagCounts["EDO"] := 0
                                    }
                                    memoryTypeDetailFlagCounts["EDO"] += 1
                                }
                                if memoryTypeDetailWordFromRaw & 0x0004 {
                                    if !memoryTypeDetailFlagCounts.Has("Fast-paged") {
                                        memoryTypeDetailFlagCounts["Fast-paged"] := 0
                                    }
                                    memoryTypeDetailFlagCounts["Fast-paged"] += 1
                                }
                                if memoryTypeDetailWordFromRaw & 0x0040 {
                                    if !memoryTypeDetailFlagCounts.Has("Synchronous DRAM") {
                                        memoryTypeDetailFlagCounts["Synchronous DRAM"] := 0
                                    }
                                    memoryTypeDetailFlagCounts["Synchronous DRAM"] += 1
                                }
                            }
                        }

                        parseOffset := nextStructureOffset
                    }
                }
            }
        }

        ; If raw parsing populated counts, resolve the display now.
        if memoryTypeCodeCounts.Count > 0 {
            mostCommonMemoryTypeCode := ""
            mostCommonMemoryTypeCount := -1
            for memoryTypeCode, memoryTypeCount in memoryTypeCodeCounts {
                if memoryTypeCount > mostCommonMemoryTypeCount {
                    mostCommonMemoryTypeCount := memoryTypeCount
                    mostCommonMemoryTypeCode := memoryTypeCode
                }
            }

            if ramValues.Has(mostCommonMemoryTypeCode) {
                installedMemoryTypeDisplay := ramValues[mostCommonMemoryTypeCode]
            } else {
                installedMemoryTypeDisplay := "Unknown Type (code " . mostCommonMemoryTypeCode . ")"
            }
        }
    }

    ; Resolve the most common legacy subtype from TypeDetail counts (EDO vs Fast-paged vs Synchronous DRAM).
    for legacyLabel, legacyCount in memoryTypeDetailFlagCounts {
        ; Only consider the three legacy subtypes we actually surface as base types.
        if legacyLabel = "EDO" || legacyLabel = "Fast-paged" || legacyLabel = "Synchronous DRAM" {
            if legacyCount > resolvedLegacyCount {
                resolvedLegacyCount := legacyCount
                resolvedLegacySubtype := legacyLabel
            }
        }
    }

    if installedMemoryTypeDisplay = "" || InStr(installedMemoryTypeDisplay, "Unknown") || installedMemoryTypeDisplay = "DRAM" || installedMemoryTypeDisplay = "Other" {
        combinedPartNumbers := ""
        for partNumberItem in partNumberStrings {
            combinedPartNumbers .= partNumberItem . " "
        }
        combinedPartNumbersLower := StrLower(combinedPartNumbers)

        switch true {
            case (installedMemoryTypeDisplay = "" || InStr(installedMemoryTypeDisplay, "Unknown") || installedMemoryTypeDisplay = "DRAM" || installedMemoryTypeDisplay = "Other") && (resolvedLegacySubtype != ""):
                installedMemoryTypeDisplay := resolvedLegacySubtype
            case InStr(combinedPartNumbersLower, "lpddr5x"):
                installedMemoryTypeDisplay := "LPDDR5X"
            case InStr(combinedPartNumbersLower, "lpddr5"):
                installedMemoryTypeDisplay := "LPDDR5"
            case InStr(combinedPartNumbersLower, "ddr5") || InStr(combinedPartNumbersLower, "pc5"):
                installedMemoryTypeDisplay := "DDR5 SDRAM"
            case InStr(combinedPartNumbersLower, "lpddr4x"):
                installedMemoryTypeDisplay := "LPDDR4X"
            case InStr(combinedPartNumbersLower, "lpddr4"):
                installedMemoryTypeDisplay := "LPDDR4"
            case InStr(combinedPartNumbersLower, "ddr4") || InStr(combinedPartNumbersLower, "pc4"):
                installedMemoryTypeDisplay := "DDR4 SDRAM"
            case InStr(combinedPartNumbersLower, "ddr3") || InStr(combinedPartNumbersLower, "pc3"):
                installedMemoryTypeDisplay := "DDR3 SDRAM"
            case InStr(combinedPartNumbersLower, "ddr2") || InStr(combinedPartNumbersLower, "pc2"):
                installedMemoryTypeDisplay := "DDR2 SDRAM"
            case InStr(combinedPartNumbersLower, "ddr "):
                installedMemoryTypeDisplay := "DDR SDRAM"
            default:
                if installedMemoryTypeDisplay = "" {
                    installedMemoryTypeDisplay := "Unknown Type"
                }
        }
    }

    memorySizeAndType := installedMemorySizeDisplay . " " . installedMemoryTypeDisplay

    return memorySizeAndType
}

GetSystemDisk() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    EXPLORER_SIZE_MAX_CHARACTERS       := 64
    BYTES_PER_WIDE_CHARACTER           := 2
    diskCapacityUtf16Buffer            := Buffer(EXPLORER_SIZE_MAX_CHARACTERS * BYTES_PER_WIDE_CHARACTER, 0)
    systemPartitionCapacityUtf16Buffer := Buffer(EXPLORER_SIZE_MAX_CHARACTERS * BYTES_PER_WIDE_CHARACTER, 0)

    diskModelText := "Unknown Disk"
    diskCapacityText := "Unknown"
    systemPartitionCapacityText := "Unknown"

    try {
        logicalDriveLetter := SubStr(A_WinDir, 1, 2)
        systemPartitionByteCount := ""

        windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
        windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\CIMV2")
        windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3

        win32LogicalDiskQuery := "
        (
        SELECT
            Size
        FROM
            Win32_LogicalDisk
        WHERE
            DeviceID='
        )" . logicalDriveLetter . "
        (
        '
        )"

        for logicalDisk in windowsManagementInstrumentationService.ExecQuery(win32LogicalDiskQuery) {
            systemPartitionByteCount := logicalDisk.Size + 0
            break
        }

        partitionObjectQuery := "
        (
        ASSOCIATORS OF
            {Win32_LogicalDisk.DeviceID='
        )" . logicalDriveLetter . "
        (
        '}
        WHERE
            AssocClass=Win32_LogicalDiskToPartition
        )"

        selectedPartitionDeviceId := ""
        for partitionObject in windowsManagementInstrumentationService.ExecQuery(partitionObjectQuery) {
            selectedPartitionDeviceId := partitionObject.DeviceID
            break
        }

        diskDriveObjectQuery := "
        (
        ASSOCIATORS OF
            {Win32_DiskPartition.DeviceID='
        )" . selectedPartitionDeviceId . "
        (
        '}
        WHERE
            AssocClass=Win32_DiskDriveToDiskPartition
        )"

        physicalDiskByteCount := ""
        if selectedPartitionDeviceId != "" {
            for diskDriveObject in windowsManagementInstrumentationService.ExecQuery(diskDriveObjectQuery) {
                if Trim(diskDriveObject.Model) != "" {
                    diskModelText := Trim(diskDriveObject.Model)
                }

                if diskDriveObject.Size != "" && diskDriveObject.Size >= 0 {
                    physicalDiskByteCount := diskDriveObject.Size + 0
                }

                break
            }
        }

        if physicalDiskByteCount != "" && physicalDiskByteCount >= 0 {
            DllCall("Shlwapi\StrFormatByteSizeW", "Int64", physicalDiskByteCount, "Ptr", diskCapacityUtf16Buffer.Ptr, "UInt", EXPLORER_SIZE_MAX_CHARACTERS)
            diskCapacityText := StrGet(diskCapacityUtf16Buffer)
        }

        if systemPartitionByteCount != "" && systemPartitionByteCount >= 0 {
            DllCall("Shlwapi\StrFormatByteSizeW", "Int64", systemPartitionByteCount, "Ptr", systemPartitionCapacityUtf16Buffer.Ptr, "UInt", EXPLORER_SIZE_MAX_CHARACTERS)
            systemPartitionCapacityText := StrGet(systemPartitionCapacityUtf16Buffer)
        }
    }

    if diskCapacityText = systemPartitionCapacityText {
        diskCapacityAndSystemPartitionCapacity := " (" . diskCapacityText . ")"
    } else {
        diskCapacityAndSystemPartitionCapacity := " (" . diskCapacityText . " / " . systemPartitionCapacityText . ")"
    }

    modelWithDiskCapacityAndSystemPartitionCapacity := diskModelText . diskCapacityAndSystemPartitionCapacity
    
    return modelWithDiskCapacityAndSystemPartitionCapacity
}

GetActiveDisplayGpu() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    ; Prefer Win32 (User32) enumeration - authoritative primary GPU.
    DISPLAY_DEVICE_ACTIVE_FLAG := 0x00000001
    DISPLAY_DEVICE_PRIMARY_DEVICE_FLAG := 0x00000004
    DISPLAY_DEVICEW_STRUCTURE_SIZE_IN_BYTES := 840

    primaryAdapterFriendlyName := ""
    firstActiveAdapterFriendlyName := ""

    displayDeviceBuffer := Buffer(DISPLAY_DEVICEW_STRUCTURE_SIZE_IN_BYTES, 0)
    displayDeviceIndex  := 0
    loop {
        ; Reinitialize the struct for this iteration and set cb (size).
        DllCall("Msvcrt\memset", "Ptr", displayDeviceBuffer.Ptr, "Int", 0, "UPtr", DISPLAY_DEVICEW_STRUCTURE_SIZE_IN_BYTES, "Int")
        NumPut("UInt", DISPLAY_DEVICEW_STRUCTURE_SIZE_IN_BYTES, displayDeviceBuffer, 0)

        enumerationSuccessful := DllCall("User32\EnumDisplayDevicesW", "Ptr", 0, "UInt", displayDeviceIndex, "Ptr", displayDeviceBuffer.Ptr, "UInt", 0, "Int")
        if enumerationSuccessful = 0 {
            break
        }

        displayDeviceStateFlags   := NumGet(displayDeviceBuffer, 68 + 256, "UInt")
        displayDeviceFriendlyName := StrGet(displayDeviceBuffer.Ptr + 68, "UTF-16")

        if InStr(displayDeviceFriendlyName, "Microsoft Basic Display") || InStr(displayDeviceFriendlyName, "Remote Display") || InStr(displayDeviceFriendlyName, "RDP")
        {
            displayDeviceIndex += 1
            continue
        }

        if displayDeviceStateFlags & DISPLAY_DEVICE_ACTIVE_FLAG {
            if firstActiveAdapterFriendlyName = "" {
                firstActiveAdapterFriendlyName := displayDeviceFriendlyName
            }
            if displayDeviceStateFlags & DISPLAY_DEVICE_PRIMARY_DEVICE_FLAG {
                primaryAdapterFriendlyName := displayDeviceFriendlyName
                break
            }
        }

        displayDeviceIndex += 1
    }

    ; Fallback only if no value retrieved from Win32: WMI heuristic.
    activeModelNameFromWmi := ""
    firstModelNameFromWmi := ""

    if primaryAdapterFriendlyName = "" && firstActiveAdapterFriendlyName = "" {
        try {
            windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
            windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\CIMV2")
            windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3

            win32VideoControllerQuery := "
            (
                SELECT
                    Name,
                    ConfigManagerErrorCode,
                    CurrentHorizontalResolution,
                    CurrentVerticalResolution
                FROM
                    Win32_VideoController
                WHERE
                    ConfigManagerErrorCode = 0
            )"

            queryResults := windowsManagementInstrumentationService.ExecQuery(win32VideoControllerQuery)
            for record in queryResults {
                controllerName := record.Name

                if firstModelNameFromWmi = "" {
                    firstModelNameFromWmi := controllerName
                }

                ; Ignore virtual or placeholder adapters here as well.
                if InStr(controllerName, "Microsoft Basic Display") || InStr(controllerName, "Remote Display") || InStr(controllerName, "RDP") {
                    continue
                }

                if record.CurrentHorizontalResolution > 0 && record.CurrentVerticalResolution > 0 {
                    activeModelNameFromWmi := controllerName
                    break
                }
            }
        }
    }

    modelName := "Unknown GPU"

    if primaryAdapterFriendlyName != "" {
        modelName := primaryAdapterFriendlyName
    } else if firstActiveAdapterFriendlyName != "" {
        modelName := firstActiveAdapterFriendlyName
    } else if activeModelNameFromWmi != "" {
        modelName := activeModelNameFromWmi
    } else if firstModelNameFromWmi != "" {
        modelName := firstModelNameFromWmi
    }

    return modelName
}

GetActiveMonitor() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    DISPLAY_DEVICEW_SIZE  := 840
    OFFSET_DeviceString   := 68
    OFFSET_StateFlags     := 324
    OFFSET_DeviceID       := 328
    DISPLAY_DEVICE_ACTIVE := 0x00000001


    unifiedExtensibleFirmwareInterfacePlugaAndPlayIdOfficialRegistry   := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Unified Extensible Firmware Interface Plug and Play ID Official Registry.csv")
    unifiedExtensibleFirmwareInterfacePlugaAndPlayIdUnofficialRegistry := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Unified Extensible Firmware Interface Plug and Play ID Unofficial Registry.csv")

    plugAndPlayManufacturers := Map()
    for manufacturer in unifiedExtensibleFirmwareInterfacePlugaAndPlayIdOfficialRegistry {
        plugAndPlayManufacturers[manufacturer["Vendor ID"]] := manufacturer["Vendor Name"]
    }

    for manufacturer in unifiedExtensibleFirmwareInterfacePlugaAndPlayIdUnofficialRegistry {
        plugAndPlayManufacturers[manufacturer["Vendor ID"]] := manufacturer["Vendor Name"]
    }

    monitorNameResult := "Unknown Monitor"

    primaryDisplayDeviceName := ""
    primaryMonitorIndex := MonitorGetPrimary()
    if primaryMonitorIndex > 0 {
        primaryDisplayDeviceName := MonitorGetName(primaryMonitorIndex)
    }

    monitorDeviceInstanceId := ""
    monitorFriendlyDeviceString := ""
    if primaryDisplayDeviceName != "" {
        enumerationIndex := 0
        displayDeviceBuffer := Buffer(DISPLAY_DEVICEW_SIZE, 0)
        loop {
            NumPut("UInt", DISPLAY_DEVICEW_SIZE, displayDeviceBuffer, 0)
            enumerationCallSucceeded := DllCall("User32\EnumDisplayDevicesW", "WStr", primaryDisplayDeviceName, "UInt", enumerationIndex, "Ptr", displayDeviceBuffer, "UInt", 0, "Int")
            if enumerationCallSucceeded = 0 {
                break
            }
            enumerationIndex += 1

            stateFlags := NumGet(displayDeviceBuffer, OFFSET_StateFlags, "UInt")
            if stateFlags & DISPLAY_DEVICE_ACTIVE {
                monitorDeviceInstanceId     := StrGet(displayDeviceBuffer.Ptr + OFFSET_DeviceID,     128, "UTF-16")
                monitorFriendlyDeviceString := StrGet(displayDeviceBuffer.Ptr + OFFSET_DeviceString, 128, "UTF-16")
                break
            }
        }
    }

    vendorCode := ""
    productCode := ""
    if monitorDeviceInstanceId != "" && RegExMatch(monitorDeviceInstanceId, "i)^(?:MONITOR|DISPLAY)\\([A-Z]{3})([0-9A-F]{4})", &matchPnP) {
        vendorCode := matchPnP[1]
        productCode := matchPnP[2]
    }

    if monitorNameResult = "Unknown Monitor" {
        bestBrand := ""
        bestModel := ""
        bestIsPrimaryMatch := false

        try {
            windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
            windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\WMI")
            windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3

            wmiMonitorIDQuery := "
            (
                SELECT
                    InstanceName,
                    ManufacturerName,
                    UserFriendlyName,
                    Active
                FROM
                    WmiMonitorID
                WHERE
                    Active=True
            )"

            queryResults := windowsManagementInstrumentationService.ExecQuery(wmiMonitorIDQuery)
            for record in queryResults {
                instanceName := record.InstanceName
                manufacturerArray := record.ManufacturerName
                candidateBrandCode := ""
                for codePoint in manufacturerArray {
                    if codePoint = 0 {
                        break
                    }
                    candidateBrandCode .= Chr(codePoint)
                }
                candidateBrandCode := StrUpper(Trim(candidateBrandCode))
                candidateBrand := plugAndPlayManufacturers.Has(candidateBrandCode) ? plugAndPlayManufacturers[candidateBrandCode] : candidateBrandCode

                userFriendlyArray := record.UserFriendlyName
                candidateModel := ""
                for codePoint in userFriendlyArray {
                    if codePoint = 0 {
                        break
                    }
                    candidateModel .= Chr(codePoint)
                }
                candidateModel := Trim(candidateModel)

                if candidateModel = "" || RegExMatch(candidateModel, "i)^\s*Generic\b.*\bPnP\b") {
                    continue
                }

                instanceMatchesPrimary := vendorCode != "" && productCode != "" && RegExMatch(instanceName, "i)" . vendorCode . productCode)

                if instanceMatchesPrimary {
                    bestBrand := candidateBrand
                    bestModel := candidateModel
                    bestIsPrimaryMatch := true
                    break
                } else if !bestIsPrimaryMatch && bestModel = "" {
                    bestBrand := candidateBrand
                    bestModel := candidateModel
                }
            }
        }
        if bestModel != "" {
            if bestBrand != "" && !RegExMatch(bestModel, "i)^\Q" . bestBrand . "\E\b") {
                monitorNameResult := bestBrand . " " . bestModel
            } else {
                monitorNameResult := bestModel
            }
        }
    }

    if monitorNameResult = "Unknown Monitor" && vendorCode != "" && productCode != "" {
        for registryClass in ["DISPLAY", "MONITOR"] {
            registryBasePath := "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Enum\" . registryClass . "\" . vendorCode . productCode
            try {
                loop reg, registryBasePath, "K" {
                    instanceKeyName := A_LoopRegName
                    parametersPath := registryBasePath . "\" . instanceKeyName . "\Device Parameters"
                    friendlyNameCandidate := ""
                    try {
                        friendlyNameCandidate := RegRead(parametersPath, "FriendlyName", "")
                    }
                    if friendlyNameCandidate != "" && !RegExMatch(friendlyNameCandidate, "i)^\s*Generic\b.*\bPnP\b") {
                        monitorNameResult := friendlyNameCandidate
                        break 2
                    }
                }
            }
        }
    }

    if monitorNameResult = "Unknown Monitor" && vendorCode != "" && productCode != "" {
        for registryClass in ["DISPLAY", "MONITOR"] {
            registryBasePath := "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Enum\" . registryClass . "\" . vendorCode . productCode
            try {
                loop reg, registryBasePath, "K" {
                    instanceKeyName := A_LoopRegName
                    parametersPath := registryBasePath . "\" . instanceKeyName . "\Device Parameters"
                    edidBuffer := ""
                    try {
                        edidBuffer := RegRead(parametersPath, "EDID")
                    }
                    if IsObject(edidBuffer) && edidBuffer.Size >= 128 {
                        descriptorStart := 54
                        descriptorLength := 18
                        loop 4 {
                            descriptorOffset := descriptorStart + (A_Index - 1) * descriptorLength
                            byte0 := NumGet(edidBuffer, descriptorOffset + 0, "UChar")
                            byte1 := NumGet(edidBuffer, descriptorOffset + 1, "UChar")
                            tag   := NumGet(edidBuffer, descriptorOffset + 3, "UChar")
                            if byte0 = 0x00 && byte1 = 0x00 && tag = 0xFC {
                                modelFromEdid := ""
                                loop 13 {
                                    edidNameAsciiByte := NumGet(edidBuffer, descriptorOffset + 5 + (A_Index - 1), "UChar")
                                    if edidNameAsciiByte = 0x00 || edidNameAsciiByte = 0x0A {
                                        break
                                    }
                                    modelFromEdid .= Chr(edidNameAsciiByte)
                                }
                                modelFromEdid := Trim(modelFromEdid)
                                if modelFromEdid != "" {
                                    brandFromMap := plugAndPlayManufacturers.Has(vendorCode) ? plugAndPlayManufacturers[vendorCode] : vendorCode
                                    monitorNameResult := (brandFromMap != "" ? brandFromMap . " " : "") . modelFromEdid
                                }
                                break
                            }
                        }
                        if monitorNameResult != "Unknown Monitor" {
                            break 2
                        }
                    }
                }
            }
        }
    }

    if monitorNameResult = "Unknown Monitor" && vendorCode != "" && productCode != "" {
        derivedBrand := plugAndPlayManufacturers.Has(vendorCode) ? plugAndPlayManufacturers[vendorCode] : vendorCode
        monitorNameResult := derivedBrand . " " . StrUpper(productCode)
    }

    if monitorNameResult = "Unknown Monitor" && monitorFriendlyDeviceString != "" && !RegExMatch(monitorFriendlyDeviceString, "i)^\s*Generic\b.*\bPnP\b") {
        monitorNameResult := monitorFriendlyDeviceString
    }

    return monitorNameResult
}

GetBios() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    biosVersion := ""
    biosDateIso := ""
    uefiIsEnabled := false

    try {
        windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
        windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\CIMV2")
        windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3

        win32BiosQuery := "
        (
            SELECT
                SMBIOSBIOSVersion,
                Version,
                BIOSVersion,
                ReleaseDate,
                BiosCharacteristics
            FROM
                Win32_BIOS
        )"

        queryResults := windowsManagementInstrumentationService.ExecQuery(win32BiosQuery)
        for record in queryResults {
            biosVersionCandidate := Trim(record.SMBIOSBIOSVersion . "")
            if biosVersionCandidate = "" {
                biosVersionCandidate := Trim(record.Version . "")
            }
            if biosVersionCandidate = "" {
                biosVersionField := record.BIOSVersion
                if IsObject(biosVersionField) {
                    for index, versionEntry in biosVersionField {
                        if Trim(versionEntry . "") != "" {
                            biosVersionCandidate := Trim(versionEntry . "")
                            break
                        }
                    }
                } else {
                    candidateString := Trim(biosVersionField . "")
                    if candidateString != "" {
                        biosVersionCandidate := candidateString
                    }
                }
            }
            biosVersion := biosVersionCandidate

            ; UEFI-capable via WMI: 75 = "UEFI specification supported".
            biosCharacteristics := record.BiosCharacteristics
            if IsObject(biosCharacteristics) {
                for index, characteristicCode in biosCharacteristics {
                    if characteristicCode + 0 = 75 {
                        uefiIsEnabled := true
                        break
                    }
                }
            }

            ; BIOS release date (CIM_DATETIME to YYYY-MM-DD).
            rawBiosReleaseDate := Trim(record.ReleaseDate . "")
            if StrLen(rawBiosReleaseDate) >= 8 {
                biosReleaseDateYear  := SubStr(rawBiosReleaseDate, 1, 4)
                biosReleaseDateMonth := SubStr(rawBiosReleaseDate, 5, 2)
                biosReleaseDateDay   := SubStr(rawBiosReleaseDate, 7, 2)
                biosDateIso := biosReleaseDateYear . "-" . biosReleaseDateMonth . "-" . biosReleaseDateDay
            }

            break
        }
    }

    ; Fallback UEFI-capable via RAW SMBIOS. Type 0 (BIOS Information) → Extension Byte 2 at offset 0x13, bit 3 (0x08).
    if !uefiIsEnabled {
        try {
            rawSystemManagementBiosSignature := 0x52534D42

            requiredBufferSize := DllCall("Kernel32\GetSystemFirmwareTable", "UInt", rawSystemManagementBiosSignature, "UInt", 0, "Ptr", 0, "UInt", 0, "UInt")

            if requiredBufferSize > 0 {
                systemManagementBiosBuffer := Buffer(requiredBufferSize, 0)

                bytesReturned := DllCall("Kernel32\GetSystemFirmwareTable", "UInt", rawSystemManagementBiosSignature, "UInt", 0, "Ptr", systemManagementBiosBuffer.Ptr, "UInt", requiredBufferSize, "UInt")

                if bytesReturned >= 8 {
                    smbiosDataLength  := NumGet(systemManagementBiosBuffer, 4, "UInt")
                    smbiosDataPointer := systemManagementBiosBuffer.Ptr + 8
                    smbiosEndPointer  := smbiosDataPointer + smbiosDataLength

                    currentStructurePointer := smbiosDataPointer
                    while currentStructurePointer + 4 <= smbiosEndPointer {
                        structureType   := NumGet(currentStructurePointer, 0, "UChar")
                        structureLength := NumGet(currentStructurePointer, 1, "UChar")
                        if structureLength < 4 || currentStructurePointer + structureLength > smbiosEndPointer {
                            break
                        }

                        if structureType = 0 { ; BIOS Information.
                            ; Extension Byte 2 is at offset 0x13 (19) when present. Bit 3 (0x08) == "UEFI specification supported".
                            if structureLength >= 0x14 {
                                biosCharacteristicsExtensionByte2 := NumGet(currentStructurePointer, 0x13, "UChar")
                                if biosCharacteristicsExtensionByte2 & 0x08 {
                                    uefiIsEnabled := true
                                }
                            }
                        }

                        ; Advance to next structure (formatted area + string-set until double NUL).
                        nextPointer := currentStructurePointer + structureLength
                        while nextPointer < smbiosEndPointer {
                            if NumGet(nextPointer, 0, "UChar") = 0 {
                                if nextPointer + 1 <= smbiosEndPointer && NumGet(nextPointer + 1, 0, "UChar") = 0 {
                                    nextPointer += 2
                                    break
                                }
                            }
                            nextPointer += 1
                        }
                        currentStructurePointer := nextPointer
                        if (structureType = 127) { ; End-of-table.
                            break
                        }
                    }
                }
            }
        }
    }

    if biosVersion = "" {
        biosVersion := "Unknown"
    }

    if biosDateIso = "" {
        biosDateIso := "Unknown"
    }

    uefiOrBios := uefiIsEnabled ? "UEFI" : "BIOS"

    biosSummary := biosVersion . " (" . biosDateIso . ") " . uefiOrBios

    return biosSummary
}

GetQueryPerformanceCounterFrequency() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    queryPerformanceCounterFrequencyBuffer := Buffer(8, 0)
    queryPerformanceCounterFrequencyRetrievedSuccessfully := DllCall("QueryPerformanceFrequency", "Ptr", queryPerformanceCounterFrequencyBuffer.Ptr, "Int")
    if !queryPerformanceCounterFrequencyRetrievedSuccessfully {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to retrieve the frequency of the performance counter. [QueryPerformanceFrequency" . ", System Error Code: " . A_LastError . "]")
    }

    queryPerformanceCounterFrequency := NumGet(queryPerformanceCounterFrequencyBuffer, 0, "Int64")

    return queryPerformanceCounterFrequency
}

GetActiveMonitorRefreshRateHz() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    ENUM_CURRENT_SETTINGS := -1
    DEVMODEW_BYTES := 220
    OFFSET_dmSize := 68
    OFFSET_dmFields := 76
    OFFSET_dmDisplayFrequency := 120
    DM_DISPLAYFREQUENCY := 0x00400000

    GDI_VREFRESH_INDEX := 116

    refreshRateHertzResult := 0

    primaryDisplayDeviceName := ""
    primaryMonitorIndex := MonitorGetPrimary()
    if primaryMonitorIndex > 0 {
        primaryDisplayDeviceName := MonitorGetName(primaryMonitorIndex)
    }

    if primaryDisplayDeviceName != "" {
        deviceModeBuffer := Buffer(DEVMODEW_BYTES, 0)
        NumPut("UShort", DEVMODEW_BYTES, deviceModeBuffer, OFFSET_dmSize)
        NumPut("UShort", 0, deviceModeBuffer, 70)

        enumCallSucceeded := DllCall("User32\EnumDisplaySettingsW", "WStr", primaryDisplayDeviceName, "Int", ENUM_CURRENT_SETTINGS, "Ptr", deviceModeBuffer, "Int")

        if enumCallSucceeded {
            deviceModeFields := NumGet(deviceModeBuffer, OFFSET_dmFields, "UInt")
            if deviceModeFields & DM_DISPLAYFREQUENCY {
                candidateFrequencyFromDevMode := NumGet(deviceModeBuffer, OFFSET_dmDisplayFrequency, "UInt")
                if candidateFrequencyFromDevMode >= 20 && candidateFrequencyFromDevMode <= 1000 {
                    refreshRateHertzResult := candidateFrequencyFromDevMode
                }
            }
        }
    }

    ; Fallback: GDI CreateDCW("DISPLAY") + GetDeviceCaps(VREFRESH) for the primary monitor.
    if refreshRateHertzResult = 0 {
        primaryDisplayDeviceContextHandle := DllCall("Gdi32\CreateDCW", "WStr", "DISPLAY", "Ptr", 0, "Ptr", 0, "Ptr", 0, "Ptr")

        if primaryDisplayDeviceContextHandle {
            candidateFrequencyFromGdi := DllCall("Gdi32\GetDeviceCaps", "Ptr", primaryDisplayDeviceContextHandle, "Int", GDI_VREFRESH_INDEX, "Int")
            DllCall("Gdi32\DeleteDC", "Ptr", primaryDisplayDeviceContextHandle)

            if candidateFrequencyFromGdi >= 20 && candidateFrequencyFromGdi <= 1000 {
                refreshRateHertzResult := candidateFrequencyFromGdi
            }
        }
    }

    return refreshRateHertzResult
}

GetWindowsColorMode() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    registryPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"

    hasAppsUseLightTheme := false
    hasSystemUsesLightTheme := false
    appsUseLightThemeFlag := 0
    systemUsesLightThemeFlag := 0

    try {
        value := RegRead(registryPath, "AppsUseLightTheme")
        hasAppsUseLightTheme := true
        appsUseLightThemeFlag := (value + 0) ? 1 : 0
    }

    try {
        value := RegRead(registryPath, "SystemUsesLightTheme")
        hasSystemUsesLightTheme := true
        systemUsesLightThemeFlag := (value + 0) ? 1 : 0
    }

    presentFlagCount := (hasAppsUseLightTheme ? 1 : 0) + (hasSystemUsesLightTheme ? 1 : 0)
    colorMode := ""

    switch presentFlagCount
    {
        case 2:
            if appsUseLightThemeFlag = systemUsesLightThemeFlag {
                if appsUseLightThemeFlag = 1 {
                    colorMode := "Light"
                } else {
                    colorMode := "Dark"
                }
            } else {
                colorMode := "Custom"
            }
        case 1:
            onlyFlag := hasAppsUseLightTheme ? appsUseLightThemeFlag : systemUsesLightThemeFlag
            if onlyFlag = 1 {
                colorMode := "Light"
            } else {
                colorMode := "Dark"
            }
        default:
            ; Keys do not exist, treat as Light (closest equivalent).
            colorMode := "Light"
    }

    return colorMode
}