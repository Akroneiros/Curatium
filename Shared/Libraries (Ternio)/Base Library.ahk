#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include ..\jsongo_AHKv2 (2025-02-26)\jsongo.v2.ahk
#Include Application Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

global applicationRegistry := Map()
global imageRegistry := Map()
global methodRegistry := Map(
    "LogEngine", Map(
        "Settings", Map(
            "Start Milliseconds Treshold", Map(
                "Value",   512,
                "Default", 512,
                "Floor",   32,
                "Ceiling", 998,
                "Delta",   0
            ),
            "Telemetry Timestamp Duration in Milliseconds", Map(
                "Value",   192,
                "Default", 192,
                "Floor",   16,
                "Ceiling", 1000,
                "Delta",   0
            )
        )
    )
)
global overlay := Map(
    "GUI", Gui("+AlwaysOnTop -Caption +ToolWindow +E0x80000 +E0x20 +DPIScale -SysMenu -Border"),
    "Counter", 0,
    "Lines", Map(),
    "Order", [],
    "Status", Map(
        "Beginning", "... Beginning " . "▶️",
        "Skipped",   "... Skipped " .   "➡️",
        "Completed", "... Completed " . "✔️",
        "Failed",    "... Failed " .    "✖️"
    )
)
global symbolLedger := Map(
    "Context", Map(),
    "Error", Map(),
    "Method", Map(),
    "Overlay", Map(),
    "Reference", Map(),
    "Whitelist", Map()
)
global system := Map(
    "Configuration", Map(),
    "Constants", Map(),
    "Directories", Map(),
    "Environment", Map(),
    "Hardware", Map(),
    "Logging", Map(
        "Counters", Map(
            "Context", 0,
            "Error", 0,
            "Method", 0,
            "Overlay", 0,
            "Reference", 0,
            "Whitelist", 0,
            "Operation Sequence Number", 0,
            "Run Telemetry Order", 0
        ),
        "Log Engine State", "Pending",
        "Log Entries", Map(
            "Execution Log", [],
            "Operation Log", [],
            "Run Telemetry", [],
            "Symbol Ledger", []
        ),
        "Log to Array", true
    ),
    "Mappings", Map(),
    "Paths", Map(),
    "Runtime", Map(),
    "Telemetry", Map()
)

ActivateWindow(windowSearchResults, maximizeWindow := false) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("windowSearchResults As Map, maximizeWindow As Boolean [Optional: False]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [windowSearchResults, maximizeWindow], "Activate Window")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Seconds to Attempt", 60, 1, 3600)
        ConfigureMethodSetting(methodName, "Short Delay", 128, 64, 1280)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    secondsToAttempt := settings["Seconds to Attempt"].Get("Value")
    shortDelay       := settings["Short Delay"].Get("Value")

    windowSearchTerm := "Window search term: " . windowSearchResults["Window Title"] . "."
    totalSleep       := Round((Round(shortDelay / 2)) + shortDelay * (2 + 3 + 4 + 5 + 6 + 7 + 8 + 9 + 10))

    if windowSearchResults["Success"] = false {
        errorMessage := "Failed to find window after trying for " . windowSearchResults["Seconds to Attempt"] . " seconds. " . windowSearchTerm

        if windowSearchResults["Custom Error Message"] != "" {
            errorMessage := windowSearchResults["Custom Error Message"]
        }

        LogConclusion("Failed", logConclusionData, A_LineNumber, errorMessage)
    }

    Loop 10 {
        loopDelay := shortDelay * A_Index
        if A_Index = 1 {
            loopDelay := Round(loopDelay/2)
        }
        Sleep(loopDelay)

        try {
            WinActivate(windowSearchResults["Window Title"])
            break
        } catch {
            if A_Index = 10 {
                LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to activate window after trying for " . totalSleep . " milliseconds. " . windowSearchTerm)
            }
        }
    }

    matchingWindowFound := WinWaitActive(windowSearchResults["Window Title"], , secondsToAttempt)
    if !matchingWindowFound {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to activate window after waiting for " . secondsToAttempt . " seconds. " . windowSearchTerm)
    }

    if maximizeWindow {
        WinMaximize(windowSearchResults["Window Title"])
        Sleep(shortDelay)
    }

    logConclusionData["Context"] := windowSearchTerm

    LogConclusion("Completed", logConclusionData)
}

AssignSpreadsheetOperationsTemplateCombined(version := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    overlayValue      := "Assign Spreadsheet Operations Template Code" . (version = "" ? " ([Latest])" : " (" . version . ")")
    static methodName := RegisterMethod("version As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [version], overlayValue)

    versionManifestFilePath := system["Directories"]["Spreadsheet Operations Template"] . "Version Manifest.ini"
    version := StrReplace(version, "v", "")

    if version = "" {
        latestVersion := ""
        latestDate    := ""

        sectionList := IniRead(versionManifestFilePath)
        Loop Parse sectionList, "`n", "`r" {
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
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Version not found: " . version)
    }

    templateCombined := Map(
        "Version",       version,
        "Release Date",  releaseDate,
        "Intro SHA-256", introHash,
        "Outro SHA-256", outroHash
    )

    templateCombined["Intro Code"] := ReadFileOnHashMatch(system["Directories"]["Spreadsheet Operations Template"] . "Spreadsheet Operations Template (v" version ", " releaseDate ") Intro.vba", templateCombined["Intro SHA-256"])
    templateCombined["Outro Code"] := ReadFileOnHashMatch(system["Directories"]["Spreadsheet Operations Template"] . "Spreadsheet Operations Template (v" version ", " releaseDate ") Outro.vba", templateCombined["Outro SHA-256"])

    LogConclusion("Completed", logConclusionData)
    return templateCombined
}

PasteText(text, commentPrefix := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static commentPrefixWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "'", "--", "#", "%", "//", ";")
    static methodName := RegisterMethod("text As String, commentPrefix As String [Optional] [Whitelist: " . commentPrefixWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [text, commentPrefix], "Paste Text")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Max Attempts", 4, 1, 16, 1)
        ConfigureMethodSetting(methodName, "Clipboard Timeout in Seconds", 4, 1, 16, 1)
        ConfigureMethodSetting(methodName, "Short Delay", 192, 64, 1024, 24)
        ConfigureMethodSetting(methodName, "Medium Delay", 416, 128, 2080, 48)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    maxAttempts               := settings["Max Attempts"].Get("Value")
    clipboardTimeoutInSeconds := settings["Clipboard Timeout in Seconds"].Get("Value")
    shortDelay                := settings["Short Delay"].Get("Value")
    mediumDelay               := settings["Medium Delay"].Get("Value")

    rows := StrSplit(text, "`n").Length

    pasteSentinel := commentPrefix . " == AutoHotkey Paste Sentinel == " . commentPrefix
    if rows != 1 {
        text := text . "`r`n" . pasteSentinel
    }
    
    attempts          := 0
    loopWasSuccessful := false

    while attempts < maxAttempts {
        attempts++

        if attempts >= 2 {
            shortDelay  := shortDelay + (attempts * methodRegistry[methodName]["Settings"]["Short Delay"]["Delta"])
            mediumDelay := mediumDelay + (attempts * methodRegistry[methodName]["Settings"]["Medium Delay"]["Delta"])

            logConclusionData["Context"] := "Failed on attempt " . attempts . " of " . maxAttempts . ". Clipboard Timeout in Seconds was " . clipboardTimeoutInSeconds .
                ". Short delay was " . shortDelay . " milliseconds. Medium delay was " . mediumDelay . " milliseconds."
            
            IncreaseMethodSetting("KeyboardShortcut", "Tiny Delay")
        }

        A_Clipboard := "" ; Clear clipboard.

        if rows = 1 {
            if attempts != 1 {
                SendInput("{End}") ; End of Line.
                Sleep(shortDelay)
                KeyboardShortcut("SHIFT", "HOME") ; Select the full line.
                Sleep(shortDelay)
                SendInput("{Delete}") ; Delete
            }
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
            Sleep(shortDelay)
            KeyboardShortcut("CTRL", "C") ; Copy

            clipboardHasData := ClipWait(clipboardTimeoutInSeconds)
            if !clipboardHasData {
                continue ; Clipboard doesn't have data, go to next attempt.
            }

            Sleep(mediumDelay)

            if A_Clipboard !== text {
                continue ; Clipboard does not match, go to next attempt.
            }
        } else {
            A_Clipboard := text ; Load combined text into clipboard.
            Sleep(shortDelay + mediumDelay)
            KeyboardShortcut("CTRL", "V") ; Paste
            Sleep(shortDelay + mediumDelay)
            A_Clipboard := "" ; Clear clipboard.
            KeyboardShortcut("CTRL", "A") ; Select All
            Sleep(mediumDelay)
            KeyboardShortcut("CTRL", "C") ; Copy

            clipboardHasData := ClipWait(clipboardTimeoutInSeconds)
            if !clipboardHasData {
                continue ; Clipboard doesn't have data, go to next attempt.
            }

            Sleep(mediumDelay)

            Loop 2 {
                if SubStr(A_Clipboard, -1) = SubStr(text, -1) {
                    break
                }

                A_Clipboard := "" ; Clear clipboard.
                KeyboardShortcut("SHIFT", "LEFT") ; Contract selection by one character to the left.
                Sleep(shortDelay)
                KeyboardShortcut("CTRL", "C") ; Copy

                clipboardHasData := ClipWait(clipboardTimeoutInSeconds)
                if !clipboardHasData {
                    continue ; Clipboard doesn't have data, go to next attempt.
                }

                Sleep(mediumDelay)
            }

            linesInClipboard := StrSplit(A_Clipboard, ["`r`n", "`n"])
            linesInText      := StrSplit(text, ["`r`n", "`n"])

            if linesInClipboard.Length = 0 {
                continue ; Copy somehow failed, go to next attempt.
            }

            if linesInClipboard[1] != linesInText[1] || linesInClipboard[linesInClipboard.Length] != linesInText[linesInText.Length] || linesInClipboard.Length != linesInText.Length {
                continue ; Clipboard doesn't match with the expected length or comparison of first and last line.
            }

            A_Clipboard := "" ; Clear clipboard.
            SendInput("{Right}") ; End of line for the last row in the selection.
            Sleep(shortDelay)
            KeyboardShortcut("SHIFT", "HOME") ; Select the whole last line which should be the sentintel.
            Sleep(shortDelay)
            KeyboardShortcut("SHIFT", "LEFT") ; Select one character more to the left.
            Sleep(shortDelay)
            KeyboardShortcut("CTRL", "X") ; Cut

            clipboardHasData := ClipWait(clipboardTimeoutInSeconds)
            if !clipboardHasData {
                continue ; Clipboard doesn't have data, go to next attempt.
            }

            Sleep(mediumDelay)

            clipboardSentinel := StrReplace(StrReplace(A_Clipboard, "`r", ""), "`n", "")
            if clipboardSentinel !== pasteSentinel {
                continue ; Paste Sentinel not copied, go to next attempt.
            }
        }

        loopWasSuccessful := true
        if attempts >= 2 {
            logConclusionData["Context"] := "Succeeded on attempt " . attempts . " of " . maxAttempts . ". Clipboard Timeout in Seconds is " . clipboardTimeoutInSeconds .
                ". Short delay is " . shortDelay . " milliseconds. Medium delay is " . mediumDelay . " milliseconds."

            IncreaseMethodSetting(methodName, "Short Delay")
            IncreaseMethodSetting(methodName, "Medium Delay")
        }
        
        break
    }

    if !loopWasSuccessful {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to paste text in " . maxAttempts . " attempts.")
    }

    Sleep(shortDelay)

    LogConclusion("Completed", logConclusionData)
}

PerformMouseActionAtCoordinates(mouseAction, coordinatePair) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static mouseActionWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}"', "Double", "Left", "Middle", "Move", "Move Smooth", "Right", "Wheel Down", "Wheel Up")
    static methodName := RegisterMethod("mouseAction As String [Whitelist: " . mouseActionWhitelist . "], coordinatePair As String [Constraint: Coordinate Pair]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [mouseAction, coordinatePair], "Perform Mouse Action at Coordinates (" . mouseAction . " @ " . coordinatePair . ")")

    coordinates := StrSplit(coordinatePair, "x")
    x := coordinates[1] + 0
    y := coordinates[2] + 0

    overlayVisibility := OverLayIsVisible()
    if overlayVisibility {
        OverlayChangeVisibility()
    }

    modeBeforeAction := A_CoordModeMouse
    CoordMode("Mouse", "Screen")
    
    switch mouseAction {
        case "Double":
            Click("left", x, y, 2)
        case "Left":
            Click("left", x, y)
        case "Middle":
            Click("middle", x, y)
        case "Move":
            MouseMove(x, y, 0)
        case "Move Smooth":
            originalSendMode := A_SendMode
            SendMode("Event")
            MouseGetPos(&currentMouseX, &currentMouseY)
            MouseMove(x, y, ComputeMouseMoveSpeed(currentMouseX . "x" . currentMouseY, coordinatePair))
            SendMode(originalSendMode)
        case "Right":
            Click("right", x, y)
        case "Wheel Down":
            Click("WheelDown", x, y)
        case "Wheel Up":
            Click("WheelUp", x, y)
    }

    CoordMode("Mouse", modeBeforeAction)

    if overlayVisibility {
        OverlayChangeVisibility()
    }

    LogConclusion("Completed", logConclusionData)
}

PerformMouseDragBetweenCoordinates(startCoordinatePair, endCoordinatePair, mouseButton := "Left", modifierKeys := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static mouseActionWhitelist := Format('"{1}", "{2}"', "Left", "Right")
    static methodName := RegisterMethod("startCoordinatePair As String [Constraint: Coordinate Pair], endCoordinatePair As String [Constraint: Coordinate Pair], mouseButton As String [Whitelist: " . mouseActionWhitelist . "], modifierKeys As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [startCoordinatePair, endCoordinatePair, mouseButton, modifierKeys], "PerformMouseDrag (" . mouseButton . ", " . startCoordinatePair . " to " . endCoordinatePair . ")")

    modeBeforeAction := A_CoordModeMouse
    CoordMode("Mouse", "Screen")

    startCoordinates := StrSplit(startCoordinatePair, "x")
    startX := startCoordinates[1] + 0
    startY := startCoordinates[2] + 0

    endCoordinates := StrSplit(endCoordinatePair, "x")
    endX := endCoordinates[1] + 0
    endY := endCoordinates[2] + 0

    normalizedModifierList := []
    seenModifierMap        := Map()

    if modifierKeys != "" {
        Loop Parse modifierKeys, "+, " . "`t" {
            rawToken := A_LoopField
            if rawToken = "" {
                continue
            }

            tokenLowercase := StrLower(Trim(rawToken))
            switch tokenLowercase {
                case "shift":
                    canonical := "Shift"
                case "lshift", "leftshift":
                    canonical := "LShift"
                case "rshift", "rightshift":
                    canonical := "RShift"
                case "ctrl", "control", "ctl":
                    canonical := "Ctrl"
                case "lctrl", "lcontrol", "leftctrl":
                    canonical := "LCtrl"
                case "rctrl", "rcontrol", "rightctrl":
                    canonical := "RCtrl"
                case "alt":
                    canonical := "Alt"
                case "lalt", "leftalt":
                    canonical := "LAlt"
                case "ralt", "rightalt", "altgr":
                    canonical := "RAlt"
                case "win", "windows", "meta", "super", "lwin", "leftwin", "winleft":
                    canonical := "LWin"
                case "rwin", "rightwin", "winright":
                    canonical := "RWin"
                default:
                    LogConclusion("Failed", logConclusionData, A_LineNumber, "Unsupported modifier: " . rawToken)
            }


            if !seenModifierMap.Has(canonical) {
                normalizedModifierList.Push(canonical)
                seenModifierMap[canonical] := true
            }
        }
    }

    for modifierName in normalizedModifierList {
        Send("{" . modifierName . " down}")
    }

    originalSendMode := A_SendMode
    SendMode("Event")
    MouseMove(startX, startY, 0)
    Sleep(16)
    MouseClickDrag(StrLower(mouseButton), startX, startY, endX, endY, ComputeMouseMoveSpeed(startCoordinatePair, endCoordinatePair))
    SendMode(originalSendMode)

    Loop normalizedModifierList.Length {
        reverseIndex := normalizedModifierList.Length - A_Index + 1
        Send("{" . normalizedModifierList[reverseIndex] . " up}")
    }

    CoordMode("Mouse", modeBeforeAction)

    LogConclusion("Completed", logConclusionData)
}

ValidateConfiguration(configuration, applications) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("configuration As Map, applications As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [configuration, applications], "Validate Configuration")

    rootSettings := [["Application Executable Directory Candidates", "Array"], ["Application Whitelist", "Array"], ["Candidate Base Directories", "Array"], ["Settings", "Map"]]
    for setting in rootSettings {
        if !configuration.Has(setting[1]) {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Root missing " . setting[1] . ".")
        }

        if Type(configuration[setting[1]]) != setting[2] {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Root for " . setting[1] . " did not return the data type of " . setting[2] . ".")
        }
    }

    for applicationWhitelist in configuration["Application Whitelist"] {
        if Type(applicationWhitelist) != "String" {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Root entry for Application Whitelist did not return the data type of String.")
        }

        applicationFound := false
        for application in applications {
            if applicationWhitelist == application["Name"] {
                applicationFound := true
                break
            }
        }

        if !applicationFound {
            LogConclusion("Failed", logConclusionData, A_LineNumber, 'Configuration Root entry for Application Whitelist refers to the application "' . applicationWhitelist . '"' . " which doesn't exist.")
        }
    }

    for applicationExecutableDirectoryCandidate in configuration["Application Executable Directory Candidates"] {
        if Type(applicationExecutableDirectoryCandidate) != "Array" {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Root entry for Application Executable Directory Candidate did not return the data type of Array.")
        }

        if applicationExecutableDirectoryCandidate.Length != 3 {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Root entry for Application Executable Directory Candidate did not have the expected three elements.")
        }

        for index, entry in applicationExecutableDirectoryCandidate {
            if Type(entry) != "String" {
                LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Root entry for Application Executable Directory Candidate did not return the data type of String.")
            }

            if index = 1 {
                applicationFound := false
                for application in applications {
                    if entry == application["Name"] {
                        applicationFound := true
                        break
                    }
                }

                if !applicationFound {
                    LogConclusion("Failed", logConclusionData, A_LineNumber, 'Configuration Root entry for Application Executable Directory Candidate refers to the application "' . entry . '"' . " which doesn't exist.")
                }
            }
        }
    }

    for candidateBaseDirectory in configuration["Candidate Base Directories"] {
        if Type(candidateBaseDirectory) != "String" {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Root entry for Candidate Base Directories did not return the data type of String.")
        }

        validation := ValidateDataUsingSpecification(candidateBaseDirectory, "String", "Valid Directory")
        if validation != "" {
            LogConclusion("Failed", logConclusionData, A_LineNumber, 'Configuration Root entry for Candidate Base Directories of "' . candidateBaseDirectory . '" is not a Valid Directory.')
        }
    }

    subSettings := [["Image Variant Preset", "String", "Single Line"], ["Application Image Override Directory", "String", "Directory"], ["Computer Alias", "String", "Single Line"]]
    for setting in subSettings {
        if !configuration["Settings"].Has(setting[1]) {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Settings entry for " . setting[1] . " missing.")
        }

        if Type(configuration["Settings"][setting[1]]) != setting[2] {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Settings entry for " . setting[1] . " did not return the data type of " . setting[2] . ".")
        }

        if configuration["Settings"][setting[1]] != "" {
            settingValidation := ValidateDataUsingSpecification(configuration["Settings"][setting[1]], setting[2], setting[3])
            if settingValidation != "" {
                LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Settings entry for " . setting[1] . " failed validation. " . settingValidation)
            }
        }
    }

    if configuration["Settings"]["Image Variant Preset"] !== "Heroes" && configuration["Settings"]["Image Variant Preset"] !== "Middle-earth" && configuration["Settings"]["Image Variant Preset"] !== "NATO Phonetic Alphabet" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Settings entry for Image Variant Preset failed validation. Only three values are allowed: Heroes, Middle-earth or NATO Phonetic Alphabet.")
    }

    if configuration["Settings"]["Application Image Override Directory"] != "" {
        filesInDirectory := GetFilesFromDirectory(configuration["Settings"]["Application Image Override Directory"])
        if filesInDirectory.Length != 0 {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Configuration Settings entry for Application Image Override Directory failed validation. Expected no files in directory but found " . filesInDirectory.Length . ".")
        }

        applicationFolders := GetFoldersFromDirectory(configuration["Settings"]["Application Image Override Directory"])
        for applicationFolder in applicationFolders {
            SplitPath(RTrim(applicationFolder, "\"), &applicationName)

            applicationFound := false
            for application in applications {
                if applicationName == application["Name"] {
                    applicationFound := true
                    break
                }
            }

            if !applicationFound {
                LogConclusion("Failed", logConclusionData, A_LineNumber, 'Configuration Settings entry for Application Image Override Directory refers to the application "' . applicationName . '"' . " which doesn't exist.")
            }
        }
    }

    LogConclusion("Completed", logConclusionData)
}

ValidateDisplayScaling() {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [], "Validate Display Scaling")

    validateDisplayResolution := ValidateDataUsingSpecification(system["Environment"]["Display Resolution"], "String", "Display Resolution")

    if validateDisplayResolution != "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, validateDisplayResolution)
    }

    validateDpiScale := ValidateDataUsingSpecification(system["Environment"]["DPI Scale"], "String", "DPI Scale")

    if validateDpiScale != "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, validateDpiScale)
    }

    LogConclusion("Completed", logConclusionData)
}

; **************************** ;
; Core Methods                 ;
; **************************** ;

AssignNewOverlayKey() {
    global overlay

    overlay["Counter"]++
    overlayKey := overlay["Counter"]

    return overlayKey
}

IncrementCounter(counterName) {
    global system

    system["Logging"]["Counters"][counterName]++
    counter := system["Logging"]["Counters"][counterName]

    return counter
}

ParseMethodWithDeclaration(methodWithDeclaration) {
    atParts     := StrSplit(methodWithDeclaration, "@", , 2)
    signature   := RTrim(atParts[1])
    RegExMatch(atParts[2], "^\s*(.*?)\s*<\s*(\d+)\s*>\s*$", &regularExpressionMatch)
    library := Trim(regularExpressionMatch[1])
    lineNumberForValidation := regularExpressionMatch[2] + 0

    methodParts := StrSplit(signature, "(", , 2)
    methodName  := methodParts[1]
    contract    := RTrim(methodParts[2], ")")

    parameters := ""
    dataTypes  := ""
    metadata   := ""

    parameterContracts   := []
    parameterParts       := []
    currentParameterText := ""

    if contract != "" {
        squareBracketDepth           := 0
        inQuotedString               := false
        removeLeadingSpaceAfterComma := false

        Loop Parse contract {
            currentCharacter := A_LoopField

            if currentCharacter = '"' {
                inQuotedString := !inQuotedString
                currentParameterText .= currentCharacter
                continue
            }

            if !inQuotedString {
                if currentCharacter = "[" {
                    squareBracketDepth += 1
                    currentParameterText .= currentCharacter
                    continue
                }
                if currentCharacter = "]" && squareBracketDepth > 0 {
                    squareBracketDepth -= 1
                    currentParameterText .= currentCharacter
                    continue
                }

                if currentCharacter = "," && squareBracketDepth = 0 {
                    parameterParts.Push(Trim(currentParameterText))
                    currentParameterText := ""
                    removeLeadingSpaceAfterComma := true
                    continue
                }
            }

            if removeLeadingSpaceAfterComma && currentCharacter = " " {
                removeLeadingSpaceAfterComma := false
                continue
            }
            removeLeadingSpaceAfterComma := false

            currentParameterText .= currentCharacter
        }

        parameterParts.Push(Trim(currentParameterText))

        for index, parameterClause in parameterParts {
            RegExMatch(parameterClause, "^[A-Za-z_][A-Za-z0-9_]*", &matchObject)
            parameterName := matchObject[0]
            metadataValue := Trim(SubStr(parameterClause, StrLen(parameterName) + 1))
            metadataValue := RegExReplace(metadataValue, "^(?i)As\s+", "")

            dataTypesValue := StrSplit(metadataValue, " ")[1]
            metadataValue := SubStr(metadataValue, StrLen(dataTypesValue) + 2)

            dataConstraint := ""
            optionalValue  := ""
            whitelist      := []

            for metadataBlock in StrSplit(metadataValue, "]", true) {
                if metadataBlock = "" {
                    continue
                }

                blockContent := Trim(metadataBlock, "[ `t")
                if blockContent = "" {
                    continue
                }

                colonPosition := InStr(blockContent, ":")
                conceptName   := colonPosition ? Trim(SubStr(blockContent, 1, colonPosition - 1)) : Trim(blockContent)
                conceptValue  := colonPosition ? Trim(SubStr(blockContent, colonPosition + 1))    : ""

                switch StrLower(conceptName) {
                    case "optional":
                        if conceptValue = "" {
                           optionalValue := conceptName
                        } else {
                            optionalValue := conceptValue
                        }
                    case "constraint":
                        dataConstraint := conceptValue
                    case "whitelist":
                        if optionalValue = "Optional" {
                            whitelist.Push("")
                        }

                        whitelistValues := StrSplit(conceptValue, '", "')
                        for entry in whitelistValues {
                            entry := Trim(entry, '"')
                            whitelist.Push(entry)
                        }
                }
            }

            parameterContracts.Push(Map(
                "Parameter Name",  parameterName,
                "Data Type",       dataTypesValue,
                "Data Constraint", dataConstraint,
                "Optional",        optionalValue,
                "Whitelist",       whitelist
            ))

            if index < parameterParts.Length {
                parameters := parameters . parameterName . ", "
                dataTypes  := dataTypes . dataTypesValue . ", "
                metadata   := metadata . metadataValue . ", "
            } else {
                parameters := parameters . parameterName
                dataTypes  := dataTypes . dataTypesValue
                metadata   := metadata . metadataValue
            }
        }
    }

    methodWithDeclarationParsed := Map(
        "Declaration",         methodWithDeclaration,
        "Signature",           signature,
        "Library",             library,
        "Contract",            contract,
        "Parameters",          parameters,
        "Data Types",          dataTypes,
        "Metadata",            metadata,
        "Validation Line",     lineNumberForValidation,
        "Parameter Contracts", parameterContracts
    )

    return methodWithDeclarationParsed
}

RegisterMethod(declaration, methodName, sourceFilePath, validationLineNumber) {
    global methodRegistry

    SplitPath(sourceFilePath, , , , &filenameWithoutExtension)
    libraryTag := " @ " . filenameWithoutExtension
    validationLineNumber := " " . "<" . validationLineNumber . ">"
    methodWithDeclaration := methodName . "(" . declaration . ")" . libraryTag . validationLineNumber

    symbol := RegisterSymbol(methodWithDeclaration, "Method")

    if !methodRegistry.Has(methodName) {
        methodRegistry[methodName] := Map()
    }

    if methodRegistry[methodName].Count < 10 {
        methodWithDeclarationParsed := ParseMethodWithDeclaration(methodWithDeclaration)

        methodRegistry[methodName]["Declaration"]         := methodWithDeclarationParsed["Declaration"]
        methodRegistry[methodName]["Signature"]           := methodWithDeclarationParsed["Signature"]
        methodRegistry[methodName]["Library"]             := methodWithDeclarationParsed["Library"]
        methodRegistry[methodName]["Contract"]            := methodWithDeclarationParsed["Contract"]
        methodRegistry[methodName]["Parameters"]          := methodWithDeclarationParsed["Parameters"]
        methodRegistry[methodName]["Data Types"]          := methodWithDeclarationParsed["Data Types"]
        methodRegistry[methodName]["Metadata"]            := methodWithDeclarationParsed["Metadata"]
        methodRegistry[methodName]["Validation Line"]     := methodWithDeclarationParsed["Validation Line"]
        methodRegistry[methodName]["Parameter Contracts"] := methodWithDeclarationParsed["Parameter Contracts"]

        if !methodRegistry[methodName].Has("Overlay Log") {
            methodRegistry[methodName]["Overlay Log"] := false
        }

        methodRegistry[methodName]["Symbol"] := symbol
        
        if !methodRegistry[methodName].Has("Settings") {
            methodRegistry[methodName]["Settings"] := Map()
        }

        for parameterContract in methodRegistry[methodName]["Parameter Contracts"] {
            if parameterContract["Whitelist"].Length != 0 {
                for whitelistValue in parameterContract["Whitelist"] {
                    RegisterSymbol(whitelistValue, "Whitelist")
                }
            }
        }
    }

    return methodName
}

ValidateDataUsingSpecification(dataValue, dataType, dataConstraint := "", whitelist := []) {
    static hexadecimalAllowedCharactersMessage     := "Only 0–9, A–F, and a–f are allowed."
    static windowsInvalidFilenameCharactersPattern := '[\\/:*?"<>|]'
    static windowsInvalidFilenameCharactersList    := '\ / : * ? " < > |'
    static windowsReservedDeviceNamesPattern       := "i)^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(\..*)?$"

    static base52CharacterSet := GetBaseCharacterSet(52)["Digit Map"]
    static base62CharacterSet := GetBaseCharacterSet(62)["Digit Map"]
    static base66CharacterSet := GetBaseCharacterSet(66)["Digit Map"]
    static base86CharacterSet := GetBaseCharacterSet(86)["Digit Map"]
    static base92CharacterSet := GetBaseCharacterSet(92)["Digit Map"]
    static base94CharacterSet := GetBaseCharacterSet(94)["Digit Map"]

    validation := ""

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
        case "Map":
            if Type(dataValue) != "Map" {
                validation := "Value must be a Map."
            }
        case "Object":
            ; No validation, but allowed Data Type.
        case "String":
            if Type(dataValue) != "String" {
                validation := "Value must be a String."
            } else if whitelist.Length != 0 {
                valueIsWhitelisted := false

                for whitelistEntry in whitelist {
                    if dataValue == whitelistEntry {
                        valueIsWhitelisted := true
                        break
                    }
                }

                if !valueIsWhitelisted {
                    validation := "Value did not match whitelisted values."
                }
            } else {
                switch dataConstraint {
                    case "Base52", "Base62", "Base66", "Base86", "Base92", "Base94":
                        baseCharacterSet := unset
                        
                        switch dataConstraint {
                            case "Base52": baseCharacterSet := base52CharacterSet
                            case "Base62": baseCharacterSet := base62CharacterSet
                            case "Base66": baseCharacterSet := base66CharacterSet
                            case "Base86": baseCharacterSet := base86CharacterSet
                            case "Base92": baseCharacterSet := base92CharacterSet
                            case "Base94": baseCharacterSet := base94CharacterSet
                        }

                        Loop Parse dataValue {
                            if !baseCharacterSet.Has(A_LoopField) {
                                validation := dataConstraint . " has invalid character: " . A_LoopField
                                break
                            }
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
                        } else if ((x := matchObject["x"] + 0), (y := matchObject["y"] + 0), (x < 0 || x >= widthDisplayResolution || y < 0 || y >= heightDisplayResolution)) {
                            if x < 0 || x >= widthDisplayResolution {
                                validation := dataConstraint . " has X out of bounds. Valid 0 to " . (widthDisplayResolution - 1) . "."
                            } else {
                                validation := dataConstraint . " has Y out of bounds. Valid 0 to " . (heightDisplayResolution - 1) . "."
                            }
                        }
                    case "Directory", "Valid Directory":
                        isDrive := RegExMatch(dataValue, "^[A-Za-z]:\\")
                        isUNC   := RegExMatch(dataValue, "^\\\\[^\\]+\\[^\\]+\\")

                        if !(isDrive || isUNC) {
                            validation := dataConstraint . " must start with a drive (C:\) or UNC path (\\server\share\)."
                        } else if InStr(dataValue, ":", , 3) {
                            validation := dataConstraint . " must not contain any colon except the drive letter."
                        } else if SubStr(dataValue, StrLen(dataValue)) != "\" {
                            validation := dataConstraint . " must end with a backslash \."
                        } else if RegExMatch(dataValue, '[<|?*""]') {
                            validation := dataConstraint . " contains invalid character."
                        } else if !DirExist(dataValue) && dataConstraint = "Directory" {
                            validation := dataConstraint . " doesn't exist."
                        }
                    case "Display Resolution":
                        validation := dataConstraint . " is invalid."
                        for resolution in system["Constants"]["Resolutions"] {
                            if resolution["Resolution"] = dataValue {
                                validation := ""
                                break
                            }
                        }
                    case "DPI Scale":
                        validation := dataConstraint . " is invalid."
                        for scale in system["Constants"]["Scales"] {
                            if scale["Scale"] = dataValue {
                                validation := ""
                                break
                            }
                        }
                    case "Drive Letter", "Valid Drive Letter":
                        if StrLen(dataValue) >= 4 || StrLen(dataValue) = 2 {
                            validation := dataConstraint . " can't be valid due to wrong length. Length of one or three is acceptable."
                        } else if StrLen(dataValue) = 3 {
                            if SubStr(dataValue, 2, 2) != ":\" {
                                validation := dataConstraint . " when three characters can only end with :\."
                            }

                            dataValue := SubStr(dataValue, 1, 1)
                        }

                        if validation = "" && !RegExMatch(dataValue, "^[A-Za-z]$") {
                            validation := dataConstraint . " has invalid letter."
                        }

                        if dataConstraint = "Drive Letter" {
                            if StrLen(dataValue) = 1 {
                                dataValue := dataValue . ":\"
                            }

                            if !DirExist(dataValue) {
                                validation := dataConstraint . " doesn't exist."
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
                            Loop totalCharacterCount {
                                currentCharacter     := SubStr(dataValue, A_Index, 1)
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
                            validation := dataConstraint . " expected length is 64 but got " . StrLen(dataValue) "."
                        } else if !RegExMatch(dataValue, "^[0-9a-fA-F]+$") {
                            validation := dataConstraint . " must be hex digits only. " . hexadecimalAllowedCharactersMessage
                        }
                    case "Single Line":
                        if RegExMatch(dataValue, "[\r\n\v\f\x{0085}\x{2028}\x{2029}]") {
                            validation := dataConstraint . " must be a single line (no line breaks allowed)."
                        }
                    case "Path", "Valid Path":
                        SplitPath(dataValue, &filename, &directoryPath)
                        directoryPath := directoryPath . "\"

                        validation := ValidateDataUsingSpecification(filename, "String", "Filename")

                        if validation = "" {
                            if dataConstraint = "Path" {
                                validation := ValidateDataUsingSpecification(directoryPath, "String", "Directory")
                            }
                            
                            if dataConstraint = "Valid Path" {
                                validation := ValidateDataUsingSpecification(directoryPath, "String", "Valid Directory")
                            }
                        }

                        if validation != "" {
                            validation := dataConstraint . " " . validation
                        }

                        if validation = "" && dataConstraint = "Path" {
                            if !FileExist(dataValue) {
                                validation := dataConstraint . " invalid as the file doesn't exist."
                            }
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
            validation := "Data Type of " . dataType . " not recognized."
    }

    return validation
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

CombineCode(introCode, mainCode, outroCode := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("introCode As String, mainCode As String, outroCode As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [introCode, mainCode, outroCode])

    combinedCode := introCode . "`r`n`r`n" . mainCode

    if outroCode != "" {
        combinedCode := combinedCode . "`r`n`r`n" . outroCode
    } 

    return combinedCode
}

ComputeMouseMoveSpeed(startCoordinatePair, endCoordinatePair) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("startCoordinatePair As String [Constraint: Coordinate Pair], endCoordinatePair As String [Constraint: Coordinate Pair]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [startCoordinatePair, endCoordinatePair])

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
    computedMouseMoveSpeed := unset
    if movementDistancePixels <= nearDistanceThresholdPixels {
        computedMouseMoveSpeed := nearSpeed
    } else if movementDistancePixels >= farDistanceThresholdPixels {
        computedMouseMoveSpeed := farSpeed
    } else {
        ratio := (movementDistancePixels - nearDistanceThresholdPixels)  / (farDistanceThresholdPixels - nearDistanceThresholdPixels)
        ratio := ratio ** 1.5
        computedMouseMoveSpeed := Round(nearSpeed + (farSpeed - nearSpeed) * ratio)
    }

    return computedMouseMoveSpeed
}

ConvertArrayToLineSeparatedString(array) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("array As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [array])

    static newLine := "`r`n"

    lineSeparatedString := ""
    for index, value in array {
        if index > 1 {
            lineSeparatedString .= newLine
        }

        lineSeparatedString .= value
    }

    return lineSeparatedString
}

ConvertHexStringToBase64(hexString, removePadding := true) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("hexString As String [Constraint: Hexadecimal String], removePadding As Boolean [Optional: true]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [hexString])

    static CRYPT_STRING_BASE64 := 0x1
    static CRYPT_STRING_NOCRLF := 0x40000000
    static encodingFlags       := CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF

    byteCount    := StrLen(hexString) // 2
    binaryBuffer := Buffer(byteCount)
    Loop byteCount {
        characterIndex      := (A_Index - 1) * 2 + 1
        byteHexadecimalPair := SubStr(hexString, characterIndex, 2)
        computedByteValue   := ("0x" . byteHexadecimalPair) + 0
        NumPut("UChar", computedByteValue, binaryBuffer, A_Index - 1)
    }

    requiredCharacterCount := 0
    sizeProbeRetrievedSuccessfully := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", binaryBuffer.Ptr, "UInt", binaryBuffer.Size, "UInt", encodingFlags, "Ptr", 0, "UInt*", &requiredCharacterCount, "Int")
    if !sizeProbeRetrievedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve size probe. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    outputUtf16Buffer  := Buffer(requiredCharacterCount * 2)
    encodingSuccessful := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", binaryBuffer.Ptr, "UInt", binaryBuffer.Size, "UInt", encodingFlags, "Ptr", outputUtf16Buffer.Ptr, "UInt*", &requiredCharacterCount, "Int")
    if !encodingSuccessful {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to encode. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    base64 := StrGet(outputUtf16Buffer.Ptr, , "UTF-16")
    if removePadding {
        base64 := RegExReplace(base64, "=+$")
    }   
        
    return base64
}

ExtractRowFromArrayOfMapsOnHeaderCondition(rowsAsMaps, headerName, targetValue) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("rowsAsMaps As Object, headerName As String, targetValue As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [rowsAsMaps, headerName, targetValue])

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
        LogConclusion("Failed", logConclusionData, A_LineNumber, "No row found where '" . headerName . "' = '" . targetValue . "'.")
    }

    return foundRow
}

ExtractValuesFromArrayDimension(array, dimension) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("array As Array, dimension As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [array, dimension])

    arrayDimension := []

    for subArray in array {
        arrayDimension.Push(subArray[dimension])
    }

    return arrayDimension
}

ExtractUniqueValuesFromSubMaps(parentMapOfMaps, subMapKeyName) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("parentMapOfMaps As Object, subMapKeyName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [parentMapOfMaps, subMapKeyName])

    uniqueValues := []

    for outerKey, innerValue in parentMapOfMaps {
        if !innerValue.Has(subMapKeyName) {
            continue
        }

        currentValue := innerValue[subMapKeyName]

        if currentValue = "" {
            continue
        }

        valueExists := false
        for existingValue in uniqueValues {
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
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String [Constraint: Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath])

    static CRYPT_STRING_BASE64 := 0x1
    static CRYPT_STRING_NOCRLF := 0x40000000
    static base64Flags         := CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF

    fileContentBuffer  := FileRead(filePath, "RAW")
    requiredCharacters := 0

    sizeProbeRetrievedSuccessfully := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", fileContentBuffer.Ptr, "UInt", fileContentBuffer.Size, "UInt", base64Flags, "Ptr", 0, "UInt*", &requiredCharacters, "Int")
    if !sizeProbeRetrievedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve size probe. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    outputUtf16Buffer  := Buffer(requiredCharacters * 2, 0)
    encodingSuccessful := DllCall("Crypt32\CryptBinaryToStringW", "Ptr", fileContentBuffer.Ptr, "UInt", fileContentBuffer.Size, "UInt", base64Flags, "Ptr", outputUtf16Buffer.Ptr, "UInt*", &requiredCharacters, "Int")
    if !encodingSuccessful {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to encode. [Crypt32\CryptBinaryToStringW" . ", System Error Code: " . A_LastError . "]")
    }

    base64Output := StrGet(outputUtf16Buffer.Ptr, "UTF-16")

    return base64Output
}

GetTextHash(text, algorithm) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static algorithmWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}"', "MD2", "MD4", "MD5", "SHA-1", "SHA-256", "SHA-384", "SHA-512")
    static methodName := RegisterMethod("text as String, algorithm As String [Whitelist: " . algorithmWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [text, algorithm])

    switch algorithm {
        case "MD2": algorithm := "MD2"
        case "MD4": algorithm := "MD4"
        case "MD5": algorithm := "MD5"
        case "SHA-1": algorithm := "SHA1"
        case "SHA-256": algorithm := "SHA256"
        case "SHA-384": algorithm := "SHA384"
        case "SHA-512": algorithm := "SHA512"
    }

    try {
        textHash := Hash.String(algorithm, text)
    } catch as textHashError {
        LogConclusion("Failed", logConclusionData, textHashError.Line, textHashError.Message)
    }

    return textHash
}

KeyboardShortcut(primaryModifier, key, secondaryModifier := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static modifierWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "ALT", "CTRL", "CONTROL", "SHIFT", "WIN", "WINDOWS")
    static methodName := RegisterMethod("primaryModifier As String [Whitelist: " . modifierWhitelist . "], key As String, secondaryModifier As String [Optional] [Whitelist: " . modifierWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [primaryModifier, key, secondaryModifier])

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Tiny Delay", 64, 16, 256, 32)
        ConfigureMethodSetting(methodName, "Legacy Threshold", 128, 16, 256)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]
    
    tinyDelay       := settings["Tiny Delay"].Get("Value")
    legacyThreshold := settings["Legacy Threshold"].Get("Value")

    if StrLen(key) > 1 {
        key := "{" . key . "}"
    } else {
        key := StrLower(key)
    }

    SendMode("Input")

    if tinyDelay >= legacyThreshold {
        SendMode("Event")
        SetKeyDelay(0, tinyDelay)
    }

    switch primaryModifier, false {
        case "ALT":
            Send("{Alt down}")
        case "CTRL", "CONTROL":
            Send("{Ctrl down}")
        case "SHIFT":
            Send("{Shift down}")
        case "WIN", "WINDOWS":
            Send("{LWin down}")
    }    

    Sleep(tinyDelay)

    if secondaryModifier != "" {
        switch secondaryModifier, false {
            case "ALT":
                Send("{Alt down}")
            case "CTRL", "CONTROL":
                Send("{Ctrl down}")
            case "SHIFT":
                Send("{Shift down}")
            case "WIN", "WINDOWS":
                Send("{LWin down}")
        }

        Sleep(tinyDelay)
    }

    Send(key)
    Sleep(tinyDelay)

    if secondaryModifier != "" {
        switch secondaryModifier, false {
            case "ALT":
                Send("{Alt up}")
            case "CTRL", "CONTROL":
                Send("{Ctrl up}")
            case "SHIFT":
                Send("{Shift up}")
            case "WIN", "WINDOWS":
                Send("{LWin up}")
        }

        Sleep(tinyDelay)
    }

    switch primaryModifier, false {
        case "ALT":
            Send("{Alt up}")
        case "CTRL", "CONTROL":
            Send("{Ctrl up}")
        case "SHIFT":
            Send("{Shift up}")
        case "WIN", "WINDOWS":
            Send("{LWin up}")
    }

    Sleep(tinyDelay)
}

ModifyScreenCoordinates(horizontalValue, verticalValue, coordinatePair) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("horizontalValue As Integer, verticalValue As Integer, coordinatePair As String [Constraint: Coordinate Pair]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [horizontalValue, verticalValue, coordinatePair])

    widthDisplayResolution  := A_ScreenWidth
    heightDisplayResolution := A_ScreenHeight

    coordinates := StrSplit(Trim(coordinatePair), "x")
    originalX := coordinates[1] + 0
    originalY := coordinates[2] + 0

    newX := originalX + horizontalValue
    newY := originalY + verticalValue
    modifiedCoordinatePair := Format("{}x{}", newX, newY)

    if newX < 0 || newX > widthDisplayResolution - 1 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "X out of bounds. Tried " . newX . " (valid 0 to " . (widthDisplayResolution - 1) . ").")
    }

    if newY < 0 || newY > heightDisplayResolution - 1 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Y out of bounds. Tried " . newY . " (valid 0 to " . (heightDisplayResolution - 1) . ").")
    }

    return modifiedCoordinatePair
}

RemoveDuplicatesFromArray(array) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("array As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [array])

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

SearchForWindow(windowTitle, secondsToAttempt, customErrorMessage := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("windowTitle As String, secondsToAttempt As Integer, customErrorMessage As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [windowTitle, secondsToAttempt, customErrorMessage])

    windowSearchResults := Map(
        "Window Title", windowTitle,
        "Seconds to Attempt", secondsToAttempt,
        "Custom Error Message", customErrorMessage,
        "Success", false
    )

    windowHandle := WinWait(windowTitle, , secondsToAttempt)
    if windowHandle {
        windowSearchResults["Window Handle"] := windowHandle
        windowSearchResults["Success"]       := true
    }

    return windowSearchResults
}

; **************************** ;
; Settings                     ;
; **************************** ;

ConfigureMethodSetting(settingMethod, settingName, settingValue, floor, ceiling, delta := 0) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("settingMethod As String, settingName As String, settingValue As Integer, floor As Integer, ceiling As Integer, delta As Integer [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [settingMethod, settingName, settingValue, floor, ceiling, delta])

    global methodRegistry

    if !methodRegistry.Has(settingMethod) {
        methodRegistry[settingMethod] := Map()
        methodRegistry[settingMethod]["Symbol"] := ""
    }

    if !methodRegistry[settingMethod].Has("Settings") {
        methodRegistry[settingMethod]["Settings"] := Map()
    }

    if !methodRegistry[settingMethod]["Settings"].Has(settingName) {
        methodRegistry[settingMethod]["Settings"][settingName] := Map()
    }

    methodRegistry[settingMethod]["Settings"][settingName]["Default"] := settingValue

    methodRegistry[settingMethod]["Settings"][settingName]["Floor"] := floor

    if ceiling != 0 {
        methodRegistry[settingMethod]["Settings"][settingName]["Ceiling"] := ceiling
    }

    methodRegistry[settingMethod]["Settings"][settingName]["Delta"] := delta

    if !methodRegistry[settingMethod]["Settings"][settingName].Has("Value") {
        methodRegistry[settingMethod]["Settings"][settingName]["Value"] := settingValue
    }

    if methodRegistry[settingMethod]["Settings"][settingName]["Value"] > methodRegistry[settingMethod]["Settings"][settingName]["Ceiling"] {
        methodRegistry[settingMethod]["Settings"][settingName]["Value"] := methodRegistry[settingMethod]["Settings"][settingName]["Ceiling"]
    } else if methodRegistry[settingMethod]["Settings"][settingName]["Value"] < methodRegistry[settingMethod]["Settings"][settingName]["Floor"] {
        methodRegistry[settingMethod]["Settings"][settingName]["Value"] := methodRegistry[settingMethod]["Settings"][settingName]["Floor"]
    }
}

DecreaseMethodSetting(settingMethod, settingName) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("settingMethod As String, settingName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [settingMethod, settingName])

    global methodRegistry

    if methodRegistry[settingMethod]["Settings"][settingName].Has("Value") {
        newSettingValue := methodRegistry[settingMethod]["Settings"][settingName]["Value"] - methodRegistry[settingMethod]["Settings"][settingName]["Delta"]

        if newSettingValue >= methodRegistry[settingMethod]["Settings"][settingName]["Floor"] {
            methodRegistry[settingMethod]["Settings"][settingName]["Value"] := newSettingValue
        }
    }
}

IncreaseMethodSetting(settingMethod, settingName) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("settingMethod As String, settingName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [settingMethod, settingName])

    global methodRegistry

    if methodRegistry[settingMethod]["Settings"][settingName].Has("Ceiling") {
        newSettingValue := methodRegistry[settingMethod]["Settings"][settingName]["Value"] + methodRegistry[settingMethod]["Settings"][settingName]["Delta"]

        if newSettingValue <= methodRegistry[settingMethod]["Settings"][settingName]["Ceiling"] {
            methodRegistry[settingMethod]["Settings"][settingName]["Value"] := newSettingValue
        }
    }
}

SetMethodSetting(settingMethod, settingName, settingValue) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("settingMethod As String, settingName As String, settingValue As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [settingMethod, settingName, settingValue])

    global methodRegistry

    if !methodRegistry.Has(settingMethod) {
        methodRegistry[settingMethod] := Map()
        methodRegistry[settingMethod]["Symbol"] := ""
    }

    if !methodRegistry[settingMethod].Has("Settings") {
        methodRegistry[settingMethod]["Settings"] := Map()
    }

    if methodRegistry[settingMethod]["Settings"][settingName].Has("Ceiling") {
        if methodRegistry[settingMethod]["Settings"][settingName]["Ceiling"] != 0 {
            if settingValue > methodRegistry[settingMethod]["Settings"][settingName]["Ceiling"] {
                methodRegistry[settingMethod]["Settings"][settingName]["Value"] := methodRegistry[settingMethod]["Settings"][settingName]["Ceiling"]
            }
        }
    } else {
        methodRegistry[settingMethod]["Settings"][settingName]["Value"] := settingValue
    }
}

; **************************** ;
; System Methods: Environment  ;
; **************************** ;

GetActiveDisplayGpu() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    DISPLAY_DEVICE_ACTIVE_FLAG              := 0x00000001
    DISPLAY_DEVICE_PRIMARY_DEVICE_FLAG      := 0x00000004
    DISPLAY_DEVICEW_STRUCTURE_SIZE_IN_BYTES := 840

    primaryAdapterFriendlyName     := ""
    firstActiveAdapterFriendlyName := ""
    modelName                      := "Unknown GPU"

    displayDeviceBuffer := Buffer(DISPLAY_DEVICEW_STRUCTURE_SIZE_IN_BYTES, 0)
    displayDeviceIndex  := 0
    Loop {
        DllCall("Msvcrt\memset", "Ptr", displayDeviceBuffer.Ptr, "Int", 0, "UPtr", DISPLAY_DEVICEW_STRUCTURE_SIZE_IN_BYTES, "Int")
        NumPut("UInt", DISPLAY_DEVICEW_STRUCTURE_SIZE_IN_BYTES, displayDeviceBuffer, 0)

        enumerationSuccessful := DllCall("User32\EnumDisplayDevicesW", "Ptr", 0, "UInt", displayDeviceIndex, "Ptr", displayDeviceBuffer.Ptr, "UInt", 0, "Int")
        if enumerationSuccessful = 0 {
            break
        }

        displayDeviceStateFlags   := NumGet(displayDeviceBuffer, 68 + 256, "UInt")
        displayDeviceFriendlyName := StrGet(displayDeviceBuffer.Ptr + 68, "UTF-16")

        if InStr(displayDeviceFriendlyName, "Microsoft Basic Display") || InStr(displayDeviceFriendlyName, "Remote Display") || InStr(displayDeviceFriendlyName, "RDP") {
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

    activeModelNameFromWmi := ""
    firstModelNameFromWmi  := ""

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

GetActiveKeyboardLayout() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

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

GetActiveMonitor() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    DISPLAY_DEVICEW_SIZE  := 840
    OFFSET_DeviceString   := 68
    OFFSET_StateFlags     := 324
    OFFSET_DeviceID       := 328
    DISPLAY_DEVICE_ACTIVE := 0x00000001

    monitorNameResult        := "Unknown Monitor"
    primaryDisplayDeviceName := ""

    unifiedExtensibleFirmwareInterfacePlugaAndPlayIdOfficialRegistryHash      := GetFileHash(system["Directories"]["Mappings"] . "Unified Extensible Firmware Interface Plug and Play ID Official Registry.csv", "SHA-256")
    unifiedExtensibleFirmwareInterfacePlugaAndPlayIdOfficialRegistryContent   := ReadFileOnHashMatch(system["Directories"]["Mappings"] . "Unified Extensible Firmware Interface Plug and Play ID Official Registry.csv", unifiedExtensibleFirmwareInterfacePlugaAndPlayIdOfficialRegistryHash)
    unifiedExtensibleFirmwareInterfacePlugaAndPlayIdOfficialRegistryArray     := ParseDelimitedRowsToArrayOfMaps(unifiedExtensibleFirmwareInterfacePlugaAndPlayIdOfficialRegistryContent)
    unifiedExtensibleFirmwareInterfacePlugaAndPlayIdUnofficialRegistryHash    := GetFileHash(system["Directories"]["Mappings"] . "Unified Extensible Firmware Interface Plug and Play ID Unofficial Registry.csv", "SHA-256")
    unifiedExtensibleFirmwareInterfacePlugaAndPlayIdUnofficialRegistryContent := ReadFileOnHashMatch(system["Directories"]["Mappings"] . "Unified Extensible Firmware Interface Plug and Play ID Unofficial Registry.csv", unifiedExtensibleFirmwareInterfacePlugaAndPlayIdUnofficialRegistryHash)
    unifiedExtensibleFirmwareInterfacePlugaAndPlayIdUnofficialRegistryArray   := ParseDelimitedRowsToArrayOfMaps(unifiedExtensibleFirmwareInterfacePlugaAndPlayIdUnofficialRegistryContent)

    plugAndPlayManufacturers := Map()
    for manufacturer in unifiedExtensibleFirmwareInterfacePlugaAndPlayIdOfficialRegistryArray {
        plugAndPlayManufacturers[manufacturer["Vendor ID"]] := manufacturer["Vendor Name"]
    }

    for manufacturer in unifiedExtensibleFirmwareInterfacePlugaAndPlayIdUnofficialRegistryArray {
        plugAndPlayManufacturers[manufacturer["Vendor ID"]] := manufacturer["Vendor Name"]
    }

    primaryMonitorIndex := MonitorGetPrimary()
    if primaryMonitorIndex > 0 {
        primaryDisplayDeviceName := MonitorGetName(primaryMonitorIndex)
    }

    monitorDeviceInstanceId     := ""
    monitorFriendlyDeviceString := ""
    if primaryDisplayDeviceName != "" {
        enumerationIndex := 0
        displayDeviceBuffer := Buffer(DISPLAY_DEVICEW_SIZE, 0)
        Loop {
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

    vendorCode  := ""
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
                    Active = True
            )"

            queryResults := windowsManagementInstrumentationService.ExecQuery(wmiMonitorIDQuery)
            for record in queryResults {
                instanceName       := record.InstanceName
                manufacturerArray  := record.ManufacturerName
                candidateBrandCode := ""
                for codePoint in manufacturerArray {
                    if codePoint = 0 {
                        break
                    }
                    candidateBrandCode .= Chr(codePoint)
                }
                candidateBrandCode := StrUpper(Trim(candidateBrandCode))
                candidateBrand     := plugAndPlayManufacturers.Has(candidateBrandCode) ? plugAndPlayManufacturers[candidateBrandCode] : candidateBrandCode

                userFriendlyArray := record.UserFriendlyName
                candidateModel    := ""
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
                Loop Reg, registryBasePath, "K" {
                    instanceKeyName       := A_LoopRegName
                    parametersPath        := registryBasePath . "\" . instanceKeyName . "\Device Parameters"
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
                Loop Reg, registryBasePath, "K" {
                    instanceKeyName := A_LoopRegName
                    parametersPath  := registryBasePath . "\" . instanceKeyName . "\Device Parameters"
                    edidBuffer      := ""
                    try {
                        edidBuffer := RegRead(parametersPath, "EDID")
                    }
                    if IsObject(edidBuffer) && edidBuffer.Size >= 128 {
                        descriptorStart  := 54
                        descriptorLength := 18
                        Loop 4 {
                            descriptorOffset := descriptorStart + (A_Index - 1) * descriptorLength
                            byte0 := NumGet(edidBuffer, descriptorOffset + 0, "UChar")
                            byte1 := NumGet(edidBuffer, descriptorOffset + 1, "UChar")
                            tag   := NumGet(edidBuffer, descriptorOffset + 3, "UChar")
                            if byte0 = 0x00 && byte1 = 0x00 && tag = 0xFC {
                                modelFromEdid := ""
                                Loop 13 {
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

GetActiveMonitorRefreshRateHz() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    ENUM_CURRENT_SETTINGS     := -1
    DEVMODEW_BYTES            := 220
    OFFSET_dmSize             := 68
    OFFSET_dmFields           := 76
    OFFSET_dmDisplayFrequency := 120
    DM_DISPLAYFREQUENCY       := 0x00400000

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

GetBios() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    biosVersion   := ""
    biosDateIso   := ""
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
                    for versionEntry in biosVersionField {
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

            biosCharacteristics := record.BiosCharacteristics
            if IsObject(biosCharacteristics) {
                for characteristicCode in biosCharacteristics {
                    if characteristicCode + 0 = 75 {
                        uefiIsEnabled := true
                        break
                    }
                }
            }

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

                        if structureType = 0 {
                            if structureLength >= 0x14 {
                                biosCharacteristicsExtensionByte2 := NumGet(currentStructurePointer, 0x13, "UChar")
                                if biosCharacteristicsExtensionByte2 & 0x08 {
                                    uefiIsEnabled := true
                                }
                            }
                        }

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
                        if structureType = 127 {
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

GetColorMode() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    registryPath := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"

    hasAppsUseLightTheme     := false
    hasSystemUsesLightTheme  := false
    appsUseLightThemeFlag    := 0
    systemUsesLightThemeFlag := 0

    try {
        registryValue := RegRead(registryPath, "AppsUseLightTheme")
        hasAppsUseLightTheme  := true
        appsUseLightThemeFlag := (registryValue + 0) ? 1 : 0
    }

    try {
        registryValue := RegRead(registryPath, "SystemUsesLightTheme")
        hasSystemUsesLightTheme  := true
        systemUsesLightThemeFlag := (registryValue + 0) ? 1 : 0
    }

    presentFlagCount := (hasAppsUseLightTheme ? 1 : 0) + (hasSystemUsesLightTheme ? 1 : 0)
    colorMode := ""

    switch presentFlagCount {
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
            colorMode := "Light"
    }

    return colorMode
}

GetComputerIdentifier() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

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

GetCpu() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    registryPath      := "HKEY_LOCAL_MACHINE\Hardware\Description\System\CentralProcessor\0"
    registryValueName := "ProcessorNameString"
    modelName         := ""
    defaultModelName  := "Unknown CPU"

    rawName     := ""
    cleanedName := ""
    try {
        rawName     := RegRead(registryPath, registryValueName)
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

GetDiskModel(driveLetter) {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("driveLetter As String [Constraint: Drive Letter]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [driveLetter])

    diskModel := "Unknown Disk"

    if StrLen(driveLetter) = 1 {
        driveLetter := driveLetter . ":"
    } else {
        driveLetter := SubStr(driveLetter, 1, StrLen(driveLetter) - 1)
    }

    try {
        windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
        windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\CIMV2")
        windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3

        partitionObjectQuery := "
        (
            ASSOCIATORS OF
                {Win32_LogicalDisk.DeviceID='
        )" . driveLetter . "'}" . "
        (
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
        )" . selectedPartitionDeviceId . "'}" . "
        (
            WHERE
                AssocClass=Win32_DiskDriveToDiskPartition
        )"

        if selectedPartitionDeviceId != "" {
            for diskDriveObject in windowsManagementInstrumentationService.ExecQuery(diskDriveObjectQuery) {
                if Trim(diskDriveObject.Model) != "" {
                    diskModel := Trim(diskDriveObject.Model)
                }

                break
            }
        }
    }
    
    return diskModel
}

GetDisplayLanguage() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    MUI_LANGUAGE_NAME      := 0x8
    languageCount          := 0
    requiredCharacterCount := 0

    displayLanguage := ""

    callToDetermineRequiredBufferSizeDoneSuccessfully := DllCall("Kernel32\GetUserPreferredUILanguages", "UInt", MUI_LANGUAGE_NAME, "UInt*", &languageCount, "Ptr", 0, "UInt*", &requiredCharacterCount, "Int")
    if callToDetermineRequiredBufferSizeDoneSuccessfully && requiredCharacterCount > 0 {
        displayLanguageUtf16Buffer := Buffer(requiredCharacterCount * 2, 0)

        retrievedLanguageListSuccessfully := DllCall("Kernel32\GetUserPreferredUILanguages", "UInt", MUI_LANGUAGE_NAME, "UInt*", &languageCount, "Ptr", displayLanguageUtf16Buffer.Ptr, "UInt*", &requiredCharacterCount, "Int")
        if retrievedLanguageListSuccessfully {
            resolvedDisplayLanguage := StrGet(displayLanguageUtf16Buffer.Ptr, "UTF-16")
            
            if resolvedDisplayLanguage != "" {
                displayLanguage := resolvedDisplayLanguage
            }
        }
    }

    if displayLanguage = "" {
        LOCALE_NAME_MAX_LENGTH := 85
        BYTES_PER_WIDE_CHAR    := 2

        fallbackLocaleNameUtf16Buffer := Buffer(LOCALE_NAME_MAX_LENGTH * BYTES_PER_WIDE_CHAR, 0)

        userDefaultUiLanguageId := DllCall("Kernel32\GetUserDefaultUILanguage", "UShort")
        if userDefaultUiLanguageId = 0 {
            userDefaultUiLanguageId := DllCall("Kernel32\GetSystemDefaultUILanguage", "UShort")
        }

        if userDefaultUiLanguageId = 0 {
            displayLanguage := "Unknown Display Language"
        } else {
            SORT_DEFAULT    := 0
            constructedLCID := (userDefaultUiLanguageId & 0xFFFF) | (SORT_DEFAULT << 16)

            lcidToLocaleNameConvertedSuccessfully := DllCall("Kernel32\LCIDToLocaleName", "UInt", constructedLCID, "Ptr", fallbackLocaleNameUtf16Buffer.Ptr, "Int", LOCALE_NAME_MAX_LENGTH, "UInt", 0, "Int")
            if lcidToLocaleNameConvertedSuccessfully {
                resolvedLocaleName := StrGet(fallbackLocaleNameUtf16Buffer.Ptr, "UTF-16")

                if resolvedLocaleName != "" {
                    displayLanguage := resolvedLocaleName
                } else {
                    displayLanguage := "Unknown Display Language"
                }
            } else {
                displayLanguage := "Unknown Display Language"
            }
        }
    }

    return displayLanguage
}

GetInputLanguage() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

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

GetInternationalSnapshot() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    internationalRegistryKeyPath := "HKEY_CURRENT_USER\Control Panel\International"
    internationalSnapshot        := Map()

    Loop Reg, internationalRegistryKeyPath, "V" {
        try {
            if A_LoopRegType = "REG_SZ" {
                registryValue := RegRead(internationalRegistryKeyPath, A_LoopRegName)
                internationalSnapshot[A_LoopRegName] := registryValue
            }
        }
    }

    LOCALE_NAME_MAX_LENGTH := 85
    BYTES_PER_WIDE_CHAR    := 2

    userDefaultLocaleNameBuffer := Buffer(LOCALE_NAME_MAX_LENGTH * BYTES_PER_WIDE_CHAR, 0)

    userDefaultLocaleNameRetrievedSuccessfully := DllCall("Kernel32\GetUserDefaultLocaleName", "Ptr", userDefaultLocaleNameBuffer.Ptr, "Int", LOCALE_NAME_MAX_LENGTH, "Int")
    if !userDefaultLocaleNameRetrievedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve user default locale name. [Kernel32\GetUserDefaultLocaleName" . ", System Error Code: " . A_LastError . "]")
    }

    internationalSnapshot["LocaleName"] := StrGet(userDefaultLocaleNameBuffer, "UTF-16")

    Loop Reg, internationalRegistryKeyPath, "K" {
        subkey := A_LoopRegName

        if subkey = "User Profile System Backup" {
            continue
        }

        internationalSnapshot[subkey] := Map()

        Loop Reg, internationalRegistryKeyPath . "\" . subkey, "V" {
            if A_LoopRegType = "REG_DWORD" || A_LoopRegType = "REG_MULTI_SZ" || A_LoopRegType = "REG_SZ" {
                try {
                    registryValue := RegRead(internationalRegistryKeyPath . "\" . subkey, A_LoopRegName)

                    if A_LoopRegType = "REG_MULTI_SZ" {
                        registryValue := StrSplit(registryValue, "`n")

                        if registryValue[-1] = "" {
                            registryValue.Pop()
                        }
                    }

                    internationalSnapshot[subkey][A_LoopRegName] := registryValue
                }
            }
        }

        if internationalSnapshot[subkey].Count = 0 {
            internationalSnapshot.Delete(subkey)
        }
    }

    if !internationalSnapshot["Geo"].Has("Nation") {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Unable to retrieve Geographical Location Identifier (GeoID).")
    }

    geographicalLocationIdentifier := internationalSnapshot["Geo"]["Nation"] + 0
   
    requiredBufferSizeForCurrencyCode := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x000F, "Ptr", 0, "Int", 0, "UInt", 0, "Int")
    if requiredBufferSizeForCurrencyCode = 0 {
        currencyCodeValue := "Error (" . A_LastError . ")"
    } else {
        currencyCodeBuffer                := Buffer(requiredBufferSizeForCurrencyCode * 2)
        currencyCodeRetrievedSuccessfully := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x000F, "Ptr", currencyCodeBuffer.Ptr, "Int", requiredBufferSizeForCurrencyCode, "UInt", 0, "Int")
        currencyCodeValue                 := (currencyCodeRetrievedSuccessfully != 0) ? StrGet(currencyCodeBuffer, "UTF-16") : "Error (" . A_LastError . ")"
    }

    requiredBufferSizeForFriendlyName := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x0008, "Ptr", 0, "Int", 0, "UInt", 0, "Int")
    if requiredBufferSizeForFriendlyName = 0 {
        friendlyNameValue := "Error (" . A_LastError . ")"
    } else {
        friendlyNameBuffer                := Buffer(requiredBufferSizeForFriendlyName * 2)
        friendlyNameRetrievedSuccessfully := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x0008, "Ptr", friendlyNameBuffer.Ptr, "Int", requiredBufferSizeForFriendlyName, "UInt", 0, "Int")
        friendlyNameValue                 := (friendlyNameRetrievedSuccessfully != 0) ? StrGet(friendlyNameBuffer, "UTF-16") : "Error (" . A_LastError . ")"
    }

    requiredBufferSizeForIso2Code := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x0004, "Ptr", 0, "Int", 0, "UInt", 0, "Int")
    if requiredBufferSizeForIso2Code = 0 {
        iso2CodeValue := "Error (" . A_LastError . ")"
    } else {
        iso2CodeBuffer                := Buffer(requiredBufferSizeForIso2Code * 2)
        iso2CodeRetrievedSuccessfully := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x0004, "Ptr", iso2CodeBuffer.Ptr, "Int", requiredBufferSizeForIso2Code, "UInt", 0, "Int")
        iso2CodeValue                 := (iso2CodeRetrievedSuccessfully != 0) ? StrGet(iso2CodeBuffer, "UTF-16") : "Error (" . A_LastError . ")"
    }

    requiredBufferSizeForIso3Code := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x0005, "Ptr", 0, "Int", 0, "UInt", 0, "Int")
    if requiredBufferSizeForIso3Code = 0 {
        iso3CodeValue := "Error (" . A_LastError . ")"
    } else {
        iso3CodeBuffer                := Buffer(requiredBufferSizeForIso3Code * 2)
        iso3CodeRetrievedSuccessfully := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x0005, "Ptr", iso3CodeBuffer.Ptr, "Int", requiredBufferSizeForIso3Code, "UInt", 0, "Int")
        iso3CodeValue                 := (iso3CodeRetrievedSuccessfully != 0) ? StrGet(iso3CodeBuffer, "UTF-16") : "Error (" . A_LastError . ")"
    }

    requiredBufferSizeForIsoUnNumber := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x000C, "Ptr", 0, "Int", 0, "UInt", 0, "Int")
    if requiredBufferSizeForIsoUnNumber = 0 {
        isoUnNumberValue := "Error (" . A_LastError . ")"
    } else {
        isoUnNumberBuffer                := Buffer(requiredBufferSizeForIsoUnNumber * 2)
        isoUnNumberRetrievedSuccessfully := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x000C, "Ptr", isoUnNumberBuffer.Ptr, "Int", requiredBufferSizeForIsoUnNumber, "UInt", 0, "Int")
        isoUnNumberValue                 := (isoUnNumberRetrievedSuccessfully != 0) ? StrGet(isoUnNumberBuffer, "UTF-16") : "Error (" . A_LastError . ")"
    }

    requiredBufferSizeForOfficialName := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x0009, "Ptr", 0, "Int", 0, "UInt", 0, "Int")
    if requiredBufferSizeForOfficialName = 0 {
        officialNameValue := "Error (" . A_LastError . ")"
    } else {
        officialNameBuffer                := Buffer(requiredBufferSizeForOfficialName * 2)
        officialNameRetrievedSuccessfully := DllCall("Kernel32\GetGeoInfoW", "UInt", geographicalLocationIdentifier, "UInt", 0x0009, "Ptr", officialNameBuffer.Ptr, "Int", requiredBufferSizeForOfficialName, "UInt", 0, "Int")
        officialNameValue                 := (officialNameRetrievedSuccessfully != 0) ? StrGet(officialNameBuffer, "UTF-16") : "Error (" . A_LastError . ")"
    }

    internationalSnapshot["Geo"]["Currency Code"]      := currencyCodeValue
    internationalSnapshot["Geo"]["Friendly Name"]      := friendlyNameValue
    internationalSnapshot["Geo"]["ISO 3166-1 alpha-2"] := iso2CodeValue
    internationalSnapshot["Geo"]["ISO 3166-1 alpha-3"] := iso3CodeValue
    internationalSnapshot["Geo"]["ISO 3166-1 numeric"] := isoUnNumberValue
    internationalSnapshot["Geo"]["Official Name"]      := officialNameValue

    return internationalSnapshot
}

GetSessionStartupTime() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))
    
    tokenQuery      := 0x0008
    tokenStatistics := 10
    
    sessionStartupTime := ""
    tokenHandle        := 0
    requiredSize       := 0
    
    currentProcess := DllCall("GetCurrentProcess", "Ptr")
    
    openProcessTokenAndRetrieveInformationSuccessfully := DllCall("advapi32\OpenProcessToken", "Ptr", currentProcess, "UInt", tokenQuery, "Ptr*", &tokenHandle)
    
    if !openProcessTokenAndRetrieveInformationSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to open process token and retrieve information. [advapi32\OpenProcessToken" . ", System Error Code: " . A_LastError . "]")
    }
    
    DllCall("advapi32\GetTokenInformation", "Ptr", tokenHandle, "UInt", tokenStatistics, "Ptr", 0, "UInt", 0, "UInt*", &requiredSize)
    if A_LastError != 122 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve information from token. [advapi32\GetTokenInformation" . ", System Error Code: " . A_LastError . "]")
    }
    
    statisticsBuffer := Buffer(requiredSize, 0)
    
    tokenStatisticsRetrievedSuccessfully := DllCall("advapi32\GetTokenInformation", "Ptr", tokenHandle, "UInt", tokenStatistics, "Ptr", statisticsBuffer, "UInt", requiredSize, "UInt*", &requiredSize)
    if !tokenStatisticsRetrievedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve statistics from token. [advapi32\GetTokenInformation" . ", System Error Code: " . A_LastError . "]")
    }
    
    authenticationId := NumGet(statisticsBuffer, 8, "UInt64")
    
    try {        
        windowsManagementInstrumentationLocator := ComObject("WbemScripting.SWbemLocator")
        windowsManagementInstrumentationService := windowsManagementInstrumentationLocator.ConnectServer(".", "ROOT\CIMV2")
        windowsManagementInstrumentationService.Security_.ImpersonationLevel := 3
               
        win32LogonSessionQuery := "
        (
            SELECT
                StartTime
            FROM
                Win32_LogonSession
            WHERE
                LogonId =
        )" . " '" . authenticationId . "'"
        
        for currentSession in windowsManagementInstrumentationService.ExecQuery(win32LogonSessionQuery) {
            if currentSession.StartTime {
                timeStamp        := currentSession.StartTime
                sessionLocalTime := SubStr(timeStamp, 1, 14)

                offset := unset
                if RegExMatch(timeStamp, "([+-])(\d{3})$", &offset) {
                    utcOffsetMinutes := offset[2] + 0

                    if offset[1] = "-" {
                        utcOffsetMinutes := -utcOffsetMinutes
                    }

                    sessionStartupTime := DateAdd(sessionLocalTime, -utcOffsetMinutes, "Minutes")
                } else {
                    LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to convert logon session date and time to UTC. Session Local Time: " . sessionLocalTime . ". Offset: " . (IsSet(offset) ? offset[0] : "N/A") . ".")
                }

                break
            }
        }
    }  catch as queryError {
        LogConclusion("Failed", logConclusionData, queryError.Line, queryError.Message)
    } finally {
        if tokenHandle {
            DllCall("CloseHandle", "Ptr", tokenHandle)
        }
    }
    
    return sessionStartupTime
}

GetMemorySizeAndType() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    systemManagementBiosType17MemoryDeviceTypesHash    := GetFileHash(system["Directories"]["Mappings"] . "System Management BIOS Type 17 Memory Device - Type.csv", "SHA-256")
    systemManagementBiosType17MemoryDeviceTypesContent := ReadFileOnHashMatch(system["Directories"]["Mappings"] . "System Management BIOS Type 17 Memory Device - Type.csv", systemManagementBiosType17MemoryDeviceTypesHash)
    systemManagementBiosType17MemoryDeviceTypesArray   := ParseDelimitedRowsToArrayOfMaps(systemManagementBiosType17MemoryDeviceTypesContent)

    ramValues := Map()
    for systemManagementBiosType17MemoryDeviceType in systemManagementBiosType17MemoryDeviceTypesArray {
        ramValues[systemManagementBiosType17MemoryDeviceType["Value"] + 0] := systemManagementBiosType17MemoryDeviceType["Meaning"]
    }

    memoryTypeDetailFlagCounts := Map()
    memoryTypeCodeCounts       := Map()
    partNumberStrings          := []
    installedMemoryTypeDisplay := ""
    resolvedLegacySubtype      := ""
    resolvedLegacyCount        := -1

    installedKilobytes := 0

    retrievedTheAmountOfRamPhysicallyInstalledSuccessfully := DllCall("Kernel32\GetPhysicallyInstalledSystemMemory", "UInt64*", &installedKilobytes, "Int")
    if !retrievedTheAmountOfRamPhysicallyInstalledSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve the amount of RAM that is physically installed on the computer. [Kernel32\GetPhysicallyInstalledSystemMemory" . ", System Error Code: " . A_LastError . "]")
    }

    installedMemorySizeInGigabytes := (installedKilobytes > 0) ? (installedKilobytes // 1048576) : 0
    installedMemorySizeDisplay     := installedMemorySizeInGigabytes ? (installedMemorySizeInGigabytes . " GB") : "Unknown Size"

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

                        if structureType = 17 {
                            if structureLength >= 0x13 {
                                systemManagementBiosMemoryTypeCodeFromRaw := NumGet(rawSmbiosBuffer, parseOffset + 0x12, "UChar")
                                if systemManagementBiosMemoryTypeCodeFromRaw >= 3 {
                                    if !memoryTypeCodeCounts.Has(systemManagementBiosMemoryTypeCodeFromRaw) {
                                        memoryTypeCodeCounts[systemManagementBiosMemoryTypeCodeFromRaw] := 0
                                    }
                                    memoryTypeCodeCounts[systemManagementBiosMemoryTypeCodeFromRaw] += 1
                                }
                            }

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

    for legacyLabel, legacyCount in memoryTypeDetailFlagCounts {
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

GetMotherboard() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

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

GetOperatingSystem() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    currentVersionRegistryKey := "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion"

    family                    := "Unknown Windows"
    edition                   := "Unknown Edition"
    architecture              := A_Is64bitOS ? "x64" : "x86"
    currentBuildNumber        := 0
    version                   := "Unknown Version"
    updateBuildRevisionNumber := ""
    displayVersion            := ""

    try {
        currentBuildNumber := RegRead(currentVersionRegistryKey, "CurrentBuildNumber") + 0
        version := currentBuildNumber
    }

    if currentBuildNumber = 0 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Unable to retrieve Build Number from the registry.")
    }

    switch {
        case currentBuildNumber >= 22000: family := "Windows 11"
        case currentBuildNumber >= 10240: family := "Windows 10"
        case currentBuildNumber >= 9600:  family := "Windows 8.1"
        case currentBuildNumber >= 9200:  family := "Windows 8"
        case currentBuildNumber >= 7600:  family := "Windows 7"
    }
   
    if family = "Windows 7" {
        servicePackVersion := ""
        try {
            servicePackVersion := RegRead(currentVersionRegistryKey, "CSDVersion")
        }

        displayVersion := servicePackVersion
    } else {
        try {
            displayVersion := RegRead(currentVersionRegistryKey, "DisplayVersion")
        }
        if displayVersion = "" {
            try {
                displayVersion := RegRead(currentVersionRegistryKey, "ReleaseId")
            }
        }
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

    installationDate := GetWindowsInstallationDateUtcTimestamp()

    operatingSystem := Map(
        "Family",            family,
        "Edition",           edition,
        "Architecture",      architecture,
        "Version",           version,
        "Display Version",   displayVersion,
        "Full Name",         "Microsoft " . family . " " . edition . " " . "(" . architecture . ")" . " Build " . version . (displayVersion != "" ? " (" . displayVersion . ")" : ""),
        "Build Number",      currentBuildNumber,
        "Installation Date", installationDate
    )

    return operatingSystem
}

GetQueryPerformanceCounterFrequency() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    queryPerformanceCounterFrequencyBuffer := Buffer(8, 0)
    queryPerformanceCounterFrequencyRetrievedSuccessfully := DllCall("QueryPerformanceFrequency", "Ptr", queryPerformanceCounterFrequencyBuffer.Ptr, "Int")
    if !queryPerformanceCounterFrequencyRetrievedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve the frequency of the performance counter. [QueryPerformanceFrequency" . ", System Error Code: " . A_LastError . "]")
    }

    queryPerformanceCounterFrequency := NumGet(queryPerformanceCounterFrequencyBuffer, 0, "Int64")

    return queryPerformanceCounterFrequency
}

GetTimeoutBeforeLockInSeconds() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    effectiveTimeout := 0

    inactivityTimeout := RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "InactivityTimeoutSecs", 0)

    if inactivityTimeout > 0 {
        effectiveTimeout := inactivityTimeout
    }

    for registryKey in [
        "HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\Control Panel\Desktop",
        "HKEY_CURRENT_USER\Control Panel\Desktop"
    ] {
        if RegRead(registryKey, "ScreenSaveActive", 0) = 1 && RegRead(registryKey, "ScreenSaverIsSecure", 0) = 1 {
            screenSaverTimeout := RegRead(registryKey, "ScreenSaveTimeOut", 0)

            if screenSaverTimeout > 0 {
                if effectiveTimeout = 0 || screenSaverTimeout < effectiveTimeout {
                    effectiveTimeout := screenSaverTimeout
                }
            }

            break
        }
    }

    return effectiveTimeout
}

GetTimeZone() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    static bufferSize            := 432
    static biasOffset            := 0
    static standardBiasOffset    := 84
    static daylightBiasOffset    := 168
    static timeZoneKeyNameOffset := 172
    static timeZoneKeyNameMaxLen := 128

    static timeZoneIdStandard := 1
    static timeZoneIdDaylight := 2
    static invalidCallResult  := 0xFFFFFFFF

    timeZoneKeyName := ""

    dynamicTimeZoneInformationBuffer := Buffer(bufferSize, 0)

    timeZoneStateRetrievedSuccessfully := DllCall("Kernel32\GetDynamicTimeZoneInformation", "Ptr", dynamicTimeZoneInformationBuffer, "UInt")
    if timeZoneStateRetrievedSuccessfully != invalidCallResult {
        extractedTimeZoneKey := StrGet(dynamicTimeZoneInformationBuffer.Ptr + timeZoneKeyNameOffset, timeZoneKeyNameMaxLen, "UTF-16")
        if extractedTimeZoneKey != "" {
            timeZoneKeyName := extractedTimeZoneKey
        }
    }

    if timeZoneKeyName = "" {
        try {
            registryValue := RegRead("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation", "TimeZoneKeyName")
            if registryValue != "" {
                timeZoneKeyName := registryValue
            }
        }
    }

    if timeZoneKeyName = "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve a valid Time Zone Key Name.")
    }

    timeZonesBaseKey        := "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones\"
    fullTimeZoneRegistryKey := timeZonesBaseKey . timeZoneKeyName

    displayName := "(Display name not available)"

    try {
        displayName := RegRead(fullTimeZoneRegistryKey, "Display")
    }

    effectiveBias := 0

    if timeZoneStateRetrievedSuccessfully != invalidCallResult {
        bias         := NumGet(dynamicTimeZoneInformationBuffer.Ptr, biasOffset, "Int")
        standardBias := NumGet(dynamicTimeZoneInformationBuffer.Ptr, standardBiasOffset, "Int")
        daylightBias := NumGet(dynamicTimeZoneInformationBuffer.Ptr, daylightBiasOffset, "Int")

        effectiveBias := bias

        if timeZoneStateRetrievedSuccessfully = timeZoneIdStandard {
            effectiveBias += standardBias
        } else if timeZoneStateRetrievedSuccessfully = timeZoneIdDaylight {
            effectiveBias += daylightBias
        }
    } else {
        try {
            tziBinary     := RegRead(fullTimeZoneRegistryKey, "TZI")
            effectiveBias := NumGet(tziBinary, 0, "Int")
        }
    }

    displayOffsetMinutes := -effectiveBias
    utcHours             := displayOffsetMinutes // 60
    utcMinutesPart       := Mod(Abs(displayOffsetMinutes), 60)
    utcSign              := (displayOffsetMinutes >= 0) ? "+" : "-"
    utcOffsetString      := "(UTC" . utcSign . Format("{:02}", Abs(utcHours)) . ":" . Format("{:02}", utcMinutesPart) . ")"

    timeZone := Map(
        "Key Name",     timeZoneKeyName,
        "Display Name", displayName,
        "UTC Offset",   utcOffsetString
    )

    return timeZone
}

GetWindowsInstallationDateUtcTimestamp() {
    qpcPreBuffer    := Buffer(8, 0)
    timestampBuffer := Buffer(8, 0)
    qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    registryKeySystemSetup    := "HKEY_LOCAL_MACHINE\System\Setup"
    registryKeyCurrentVersion := "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion"

    oldestSeconds      := unset
    installDateSeconds := unset

    Loop Reg, registryKeySystemSetup, "K" {
        if !RegExMatch(A_LoopRegName, "^Source OS")
            continue

        try {
            candidate := RegRead(registryKeySystemSetup . "\" . A_LoopRegName, "InstallDate")
        } catch {
            continue
        }

        if !IsInteger(candidate) {
            continue
        }

        candidate := candidate & 0xFFFFFFFF

        if (candidate > 0 && (!IsSet(oldestSeconds) || candidate < oldestSeconds)) {
            oldestSeconds := candidate
        }
    }

    if IsSet(oldestSeconds) {
        installDateSeconds := oldestSeconds
    }  else {
        installDateSeconds := unset
        try {
            registryValue := RegRead(registryKeyCurrentVersion, "InstallDate")
            if IsInteger(registryValue) {
                installDateSeconds := registryValue & 0xFFFFFFFF
            }
        }

        if (!IsSet(installDateSeconds) || installDateSeconds <= 0) {
            try {
                fileTimeString := FileGetTime(A_WinDir, "C")
                if fileTimeString {
                    installDateSeconds := DateDiff(fileTimeString, "19700101000000", "Seconds")
                }
            }
        }

        if (!IsSet(installDateSeconds) || installDateSeconds <= 0) {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "No valid InstallDate found in any source.")
        }
    }

    windowsInstallationDateUtc := ConvertUnixTimeToUtcTimestamp(installDateSeconds)

    return windowsInstallationDateUtc
}

; **************************** ;
; System Methods: Telemetry    ;
; **************************** ;

GetDriveSpaceSnapshot(driveLetter) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("driveLetter As String [Constraint: Drive Letter]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [driveLetter])

    driveLetter := StrUpper(driveLetter)
    if StrLen(driveLetter) = 1 {
        driveLetter := driveLetter . ":\"
    }

    freeBytesAvailableToCaller := 0
    totalNumberOfBytes         := 0
    totalNumberOfFreeBytes     := 0

    freeDiskSpaceRetrievedSuccessfully := DllCall("Kernel32\GetDiskFreeSpaceExW", "Str", driveLetter, "Int64*", &freeBytesAvailableToCaller, "Int64*", &totalNumberOfBytes, "Int64*", &totalNumberOfFreeBytes, "Int")
    if !freeDiskSpaceRetrievedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve disk volume information. [Kernel32\GetDiskFreeSpaceExW" . ", System Error Code: " . A_LastError . "]")
    }

    bytesPerGB  := 1000000000
    bytesPerGiB := 1 << 30
    bytesPerTB  := 1000000000000
    bytesPerTiB := 1 << 40

    freeGB  := Round(freeBytesAvailableToCaller / bytesPerGB, 4)
    freeGiB := Round(freeBytesAvailableToCaller / bytesPerGiB, 4)
    freeTB  := Round(freeBytesAvailableToCaller / bytesPerTB, 4)
    freeTiB := Round(freeBytesAvailableToCaller / bytesPerTiB, 4)

    totalGB  := Round(totalNumberOfBytes / bytesPerGB, 4)
    totalGiB := Round(totalNumberOfBytes / bytesPerGiB, 4)
    totalTB  := Round(totalNumberOfBytes / bytesPerTB, 4)
    totalTiB := Round(totalNumberOfBytes / bytesPerTiB, 4)

    usedBytes   := Round(totalNumberOfBytes - totalNumberOfFreeBytes, 4)
    usedGB      := Round(usedBytes / bytesPerGB, 4)
    usedGiB     := Round(usedBytes / bytesPerGiB, 4)
    usedTB      := Round(usedBytes / bytesPerTB, 4)
    usedTiB     := Round(usedBytes / bytesPerTiB, 4)

    usedPercent := (totalNumberOfBytes > 0) ? Round((usedBytes / totalNumberOfBytes) * 100, 2) : 0.0
    freePercent := Round(100 - usedPercent, 2)

    static formattedFreeSizeBuffer := Buffer(128, 0)

    freeSizeFormattedSuccessfully := DllCall("Shlwapi\StrFormatByteSizeW", "Int64", freeBytesAvailableToCaller, "Ptr", formattedFreeSizeBuffer.Ptr, "Int", 64, "Ptr")
    if !freeSizeFormattedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to format the amount of free space available on disk volume. [Shlwapi\StrFormatByteSizeW" . ", System Error Code: " . A_LastError . "]")
    }

    windowsFreeSizeFormatted := StrGet(formattedFreeSizeBuffer, "UTF-16")

    static formattedTotalSizeBuffer := Buffer(128, 0)

    totalSizeFormattedSuccessfully := DllCall("Shlwapi\StrFormatByteSizeW", "Int64", totalNumberOfBytes, "Ptr", formattedTotalSizeBuffer.Ptr, "Int", 64, "Ptr")
    if !totalSizeFormattedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to format the amount of total space available on disk volume. [Shlwapi\StrFormatByteSizeW" . ", System Error Code: " . A_LastError . "]")
    }

    windowsTotalSizeFormatted := StrGet(formattedTotalSizeBuffer, "UTF-16")

    driveSpaceSnapshot := Map(
        "Drive",              driveLetter,
        "Free Bytes",         freeBytesAvailableToCaller,
        "Free GB",            freeGB,
        "Free GiB",           freeGiB,
        "Free TB",            freeTB,
        "Free TiB",           freeTiB,
        "Free Percent",       freePercent,
        "Total Bytes",        totalNumberOfBytes,
        "Total GB",           totalGB,
        "Total GiB",          totalGiB,
        "Total TB",           totalTB,
        "Total TiB",          totalTiB,
        "Used Bytes",         usedBytes,
        "Used GB",            usedGB,
        "Used GiB",           usedGiB,
        "Used TB",            usedTB,
        "Used TiB",           usedTiB,
        "Used Percent",       usedPercent,
        "Windows Free Size",  windowsFreeSizeFormatted,
        "Windows Total Size", windowsTotalSizeFormatted
    )

    return driveSpaceSnapshot
}

GetSystemResourceSnapshot() {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"))

    pointerSizeInBytes   := A_PtrSize
    structureSizeInBytes := 4 + (pointerSizeInBytes = 8 ? 4 : 0) + (10 * pointerSizeInBytes) + (3 * 4)
    if pointerSizeInBytes = 8 {
        structureSizeInBytes := (structureSizeInBytes + 7) & ~7
    }

    static performanceInformationBuffer := Buffer(structureSizeInBytes, 0)

    NumPut("UInt", structureSizeInBytes, performanceInformationBuffer, 0)

    getPerformanceInfoSucceeded := DllCall("Psapi\GetPerformanceInfo", "Ptr", performanceInformationBuffer.Ptr, "UInt", structureSizeInBytes, "Int")
    if !getPerformanceInfoSucceeded {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve system performance information. [Psapi\GetPerformanceInfo" . ", System Error Code: " . A_LastError . "]")
    }

    offsetInBytes := 4
    if pointerSizeInBytes = 8 {
        offsetInBytes += 4
    }

    commitTotalPages       := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    commitLimitPages       := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    commitPeakPages        := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    physicalTotalPages     := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    physicalAvailablePages := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    systemCachePages       := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    kernelTotalPages       := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    kernelPagedPages       := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    kernelNonpagedPages    := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    pageSizeInBytes        := NumGet(performanceInformationBuffer, offsetInBytes, "UPtr"), offsetInBytes += pointerSizeInBytes
    systemHandleCount      := NumGet(performanceInformationBuffer, offsetInBytes, "UInt"), offsetInBytes += 4
    systemProcessCount     := NumGet(performanceInformationBuffer, offsetInBytes, "UInt"), offsetInBytes += 4
    systemThreadCount      := NumGet(performanceInformationBuffer, offsetInBytes, "UInt"), offsetInBytes += 4

    totalPhysicalBytes     := physicalTotalPages * pageSizeInBytes
    usedPhysicalBytes      := (physicalTotalPages - physicalAvailablePages) * pageSizeInBytes

    usedPhysicalPercent    := (totalPhysicalBytes > 0) ? Round((usedPhysicalBytes / totalPhysicalBytes) * 100, 2) : 0.0
    commitUsedPercent      := (commitLimitPages > 0)   ? Round((commitTotalPages / commitLimitPages) * 100, 2)    : 0.0

    systemResourceSnapshot := Map(
        "Commit Total Bytes",    commitTotalPages * pageSizeInBytes,
        "Commit Limit Bytes",    commitLimitPages * pageSizeInBytes,
        "Commit Peak Bytes",     commitPeakPages * pageSizeInBytes,
        "Commit Used Percent",   commitUsedPercent,
        "Kernel Total Bytes",    kernelTotalPages * pageSizeInBytes,
        "Kernel Paged Bytes",    kernelPagedPages * pageSizeInBytes,
        "Kernel Nonpaged Bytes", kernelNonpagedPages * pageSizeInBytes,
        "Physical Used Bytes",   usedPhysicalBytes,
        "Physical Total Bytes",  totalPhysicalBytes,
        "Physical Used Percent", usedPhysicalPercent,
        "System Cache Bytes",    systemCachePages * pageSizeInBytes,
        "System Handle Count",   systemHandleCount,
        "System Process Count",  systemProcessCount,
        "System Thread Count",   systemThreadCount
    )

    return systemResourceSnapshot
}