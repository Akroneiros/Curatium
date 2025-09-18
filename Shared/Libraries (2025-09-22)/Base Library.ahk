#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Logging Library.ahk

AssignSpreadsheetOperationsTemplateCombined(version := "") {
    static methodName := RegisterMethod("AssignSpreadsheetOperationsTemplateCombined(version As String [Optional])" . LibraryTag(A_LineFile), A_LineNumber + 7)
    overlayValue := "Assign Spreadsheet Operations Template Code"
    if version = "" {
        overlayValue := overlayValue . " ([Latest])"
    } else {
        overlayValue := overlayValue . " (" . version . ")"
    }
    logValuesForConclusion := LogInformationBeginning(overlayValue, methodName, [version])

    sharedFolderPath := A_ScriptDir . "\Shared\Spreadsheet Operations Template\"
    manifestFilePath := sharedFolderPath . "Version Manifest.ini"
    version := StrReplace(version, "v", "")

    if version = "" {
        latestVersion := ""
        latestDate    := ""

        sectionList := IniRead(manifestFilePath)
        Loop Parse sectionList, "`n", "`r" {
            candidateVersion := A_LoopField
            if candidateVersion = "" {
                continue
            }

            candidateDate := IniRead(manifestFilePath, candidateVersion, "ReleaseDate", "")
            if latestDate = "" || StrCompare(candidateDate, latestDate) > 0 {
                latestVersion := candidateVersion
                latestDate    := candidateDate
            }
        }
        version := latestVersion
    }

    releaseDate := IniRead(manifestFilePath, version, "ReleaseDate", "")
    introHash   := IniRead(manifestFilePath, version, "IntroSHA-256", "")
    outroHash   := IniRead(manifestFilePath, version, "OutroSHA-256", "")

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
        "Outro SHA-256", outroHash,
        "Intro Code",    "",
        "Outro Code",    ""
    )

    templateCombined["Intro Code"] := ReadFileOnHashMatch(sharedFolderPath . "Spreadsheet Operations Template (v" version ", " releaseDate ") Intro.vba", templateCombined["Intro SHA-256"])
    templateCombined["Outro Code"] := ReadFileOnHashMatch(sharedFolderPath . "Spreadsheet Operations Template (v" version ", " releaseDate ") Outro.vba", templateCombined["Outro SHA-256"])

    LogInformationConclusion("Completed", logValuesForConclusion)
    return templateCombined
}

AssignHeroAliases() {
    static methodName := RegisterMethod("AssignHeroAliases()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Assign Variations", methodName)

    static variations := Map(
        "a", "Adelaide",  ; Castle,     Cleric
        "b", "Bron",      ; Fortress,   Beastmaster
        "c", "Christian", ; Castle,     Knight
        "d", "Darkstorn", ; Dungeon,    Warlock
        "e", "Elleshar",  ; Rampart,    Druid
        "f", "Fafner",    ; Tower,      Alchemist
        "g", "Gundula",   ; Stronghold, Battle Mage
        "h", "Halon",     ; Tower,      Wizard
        "i", "Isra",      ; Necropolis, Death Knight
        "j", "Jabarkas",  ; Stronghold, Barbarian
        "k", "Kyrre",     ; Rampart,    Ranger
        "l", "Lorelei",   ; Dungeon,    Overlord
        "m", "Mirlanda",  ; Fortress,   Witch
        "n", "Nagash",    ; Necropolis, Necromancer
        "o", "Olema",     ; Inferno,    Heretic
        "p", "Pyre"       ; Inferno,    Demoniac
    )

    LogInformationConclusion("Completed", logValuesForConclusion)
    return variations
}

ModifyScreenCoordinates(horizontalValue, verticalValue, coordinatePair) {
    static methodName := RegisterMethod("ModifyScreenCoordinates(horizontalValue As String [Type: Screen Delta], verticalValue As String [Type: Screen Delta], coordinatePair As String [Pattern: ^\d+x\d+$])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Modify Screen Coordinates (" . horizontalValue . "x" . verticalValue . ", " . coordinatePair . ")", methodName, [horizontalValue, verticalValue, coordinatePair])

    widthDisplayResolution  := A_ScreenWidth
    heightDisplayResolution := A_ScreenHeight

    coordinates := StrSplit(Trim(coordinatePair), "x")
    originalX := coordinates[1] + 0
    originalY := coordinates[2] + 0

    newX := originalX + (horizontalValue + 0)
    newY := originalY + (verticalValue + 0)
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
    static methodName := RegisterMethod("PasteCode(code As String [Type: Code], commentPrefix As String [Whitelist: " . commentPrefixWhitelist . "]" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Paste Code (Length: " . StrLen(code) . ")", methodName, [code, commentPrefix])

    sentinel := commentPrefix . " == AutoHotkey Paste Sentinel == " . commentPrefix
    code := code . "`r`n" . sentinel
    
    attempts    := 0
    maxAttempts := 4
    sleepAmount := 360
    success     := false

    while (attempts < maxAttempts) {
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

    if success = false {
        try {
            throw Error("Paste of code failed.")
        } catch as pasteOfCodeFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, pasteOfCodeFailedError)
        }
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

PastePath(savePath) {
    static methodName := RegisterMethod("PastePath(savePath As String [Type: Absolute Save Path])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Paste Path (" . savePath . ")", methodName, [savePath])

    attempts    := 0
    maxAttempts := 4
    sleepAmount := 200
    success     := false

    while (attempts < maxAttempts) {
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
        SendText(savePath)
        Sleep(sleepAmount + sleepAmount)

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

    if success = false {
        try {
            throw Error("Paste of path failed.")
        } catch as pasteOfPathFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, pasteOfPathFailedError)
        }
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

PasteSearch(searchValue) {
    static methodName := RegisterMethod("PasteSearch(searchValue As String [Type: Search Open])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Paste Search (" . searchValue . ")", methodName, [searchValue])

    attempts    := 0
    maxAttempts := 4
    sleepAmount := 200
    success     := false

    while (attempts < maxAttempts) {
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
        SendText(searchValue)
        Sleep(sleepAmount + sleepAmount)

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

    if success = false {
        try {
            throw Error("Paste of search failed.")
        } catch as pasteOfPathFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, pasteOfPathFailedError)
        }
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

PerformMouseActionAtCoordinates(mouseAction, coordinatePair) {
    static mouseActionWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}"', "Double", "Left", "Middle", "Move", "Right")
    static methodName := RegisterMethod("PerformMouseActionAtCoordinates(mouseAction As String [Whitelist: " . mouseActionWhitelist . "], coordinatePair As String [Pattern: ^\d+x\d+$])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Perform Mouse Action at Coordinates (" . mouseAction . " @ " . coordinatePair . ")", methodName, [mouseAction, coordinatePair])

    widthDisplayResolution  := A_ScreenWidth
    heightDisplayResolution := A_ScreenHeight

    coordinates := StrSplit(Trim(coordinatePair), "x")
    x := coordinates[1] + 0
    y := coordinates[2] + 0

    try {
        if x < 0 || x > widthDisplayResolution - 1 {
            throw Error("X out of bounds. Tried " . x . " (valid 0 to " . (widthDisplayResolution - 1) . ").")
        }
    } catch as xOutOfBoundsError {
        LogInformationConclusion("Failed", logValuesForConclusion, xOutOfBoundsError)
    }

    try {
        if y < 0 || y > heightDisplayResolution - 1 {
            throw Error("Y out of bounds. Tried " . y . " (valid 0 to " . (heightDisplayResolution - 1) . ").")
        }
    } catch as yOutOfBoundsError {
        LogInformationConclusion("Failed", logValuesForConclusion, yOutOfBoundsError)
    }

    mouseAction := StrLower(Trim(mouseAction))

    overlayVisibility := OverLayIsVisible()
    if overlayVisibility = True {
        OverlayChangeVisibility()
    }

    modeBeforeAction := A_CoordModeMouse
    CoordMode("Mouse", "Screen")
    
    switch mouseAction {
        case "double":
            Click("left", x, y, 2)
        case "left":
            Click("left", x, y)
        case "middle":
            Click("middle", x, y)
        case "move":
            MouseMove(x, y)
        case "right":
            Click("right", x, y)
        default:
    }

    CoordMode("Mouse", modeBeforeAction)

    if overlayVisibility = True {
        OverlayChangeVisibility()
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

ExtractValuesFromArrayDimension(array, dimension) {
    arrayDimension := []

    for outerIndex, innerArray in array {
        arrayDimension.Push(innerArray[dimension])
    }

    return arrayDimension
}

ExtractUniqueValuesFromSubMaps(parentMapOfMaps, subMapKeyName) {
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

IfStringIsNotEmptyReturnValue(stringValue, returnValue) {
    If stringValue = "" {
        returnValue := ""
    }

    return returnValue
}