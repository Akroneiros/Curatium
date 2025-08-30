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

    templateCombined["Intro Code"] := ReadFileOnHashMatch(sharedFolderPath . "Spreadsheet Operations Template (v" version ", " releaseDate ") Intro.txt", templateCombined["Intro SHA-256"])
    templateCombined["Outro Code"] := ReadFileOnHashMatch(sharedFolderPath . "Spreadsheet Operations Template (v" version ", " releaseDate ") Outro.txt", templateCombined["Outro SHA-256"])

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

; ******************** ;
; Helper Methods       ;
; ******************** ;

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