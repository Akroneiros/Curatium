#Requires AutoHotkey v2.0
#Include Application Library.ahk
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Image Library.ahk

; Press Escape to abort the script early when running or to close the script when it's completed.
$Esc:: {
    if system["Logging"]["Log Engine State"] != "Pending" {
        Critical "On"
        AbortExecution()
    } else {
        ExitApp()
    }
}

AbortExecution() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [], "Abort Execution")

    LogConclusion("Failed", logConclusionData, A_LineNumber, "Execution aborted early by pressing escape.")
}

OverlayChangeTransparency(transparencyValue) {
    static methodName := RegisterMethod("transparencyValue As Integer [Constraint: Byte]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [transparencyValue], "Overlay Change Transparency (" . transparencyValue . ")")

    WinSetTransparent(transparencyValue, "ahk_id " . overlay["GUI"].Hwnd)

    LogConclusion("Completed", logConclusionData)
}

OverlayChangeVisibility() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [], "Overlay Change Visibility")

    if DllCall("User32\IsWindowVisible", "Ptr", overlay["GUI"].Hwnd) {
        overlay["GUI"].Hide()
    } else {
        overlay["GUI"].Show("NoActivate")
    }

    LogConclusion("Completed", logConclusionData)
}

OverlayHideLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("methodNameInput As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [methodNameInput], "Overlay Hide Log for Method (" . methodNameInput . ")")

    global methodRegistry
    
    if !methodRegistry.Has(methodNameInput) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, 'Method "' . methodNameInput . '" not registered.')
    }

    methodRegistry[methodNameInput]["Overlay Log"] := false

    LogConclusion("Completed", logConclusionData)
}

OverlayShowLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("methodNameInput As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [methodNameInput], "Overlay Show Log for Method (" . methodNameInput . ")")

    global methodRegistry

    if methodRegistry.Has(methodNameInput) {
        methodRegistry[methodNameInput]["Overlay Log"] := true
    } else {
        methodRegistry[methodNameInput] := Map(
            "Overlay Log", true
        )
    }

    LogConclusion("Completed", logConclusionData)
}

OverlayStart() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [], "Overlay Start")

    global overlay

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Base Logical Width", 960, 640, 7680)
        ConfigureMethodSetting(methodName, "Base Logical Height", 920, 480, 4320)
        ConfigureMethodSetting(methodName, "Overlay Transparency", 172, 0, 255)
        ConfigureMethodSetting(methodName, "Font Size", 10, 6, 24)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    baseLogicalWidth    := settings["Base Logical Width"].Get("Value")
    baseLogicalHeight   := settings["Base Logical Height"].Get("Value")
    overlayTransparency := settings["Overlay Transparency"].Get("Value")
    fontSize            := settings["Font Size"].Get("Value")

    overlay["GUI"].BackColor := "0x000000"
    overlay["GUI"].SetFont("s" . fontSize . " cWhite", "Consolas")
    overlay["GUI"].MarginX := 0
    overlay["GUI"].MarginY := 0

    statusTextControl := overlay["GUI"].Add("Text", "vStatusText w" . baseLogicalWidth . " h" . baseLogicalHeight . " +0x1", "")

    measureVisualRectangle := () => (
        overlay["GUI"].Show("Hide AutoSize"),
        rectBuffer := Buffer(16, 0),
        DllCall("Dwmapi\DwmGetWindowAttribute", "Ptr", overlay["GUI"].Hwnd, "Int", 9, "Ptr", rectBuffer, "Int", 16),
        Map(
            "left",   NumGet(rectBuffer,  0, "Int"),
            "top",    NumGet(rectBuffer,  4, "Int"),
            "right",  NumGet(rectBuffer,  8, "Int"),
            "bottom", NumGet(rectBuffer, 12, "Int")
        )
    )

    visualRectangle := measureVisualRectangle()
    visualWidth     := visualRectangle["right"]  - visualRectangle["left"]
    visualHeight    := visualRectangle["bottom"] - visualRectangle["top"]

    ; Ensure the *visual* size is even on both axes. If an axis is odd, nudge the client by +1 logical pixel on that axis and re-measure.
    adjustAttemptsForWidth := 0
    while Mod(visualWidth, 2) && adjustAttemptsForWidth < 6 {
        baseLogicalWidth += 1
        statusTextControl.Move(, , baseLogicalWidth, baseLogicalHeight)
        visualRectangle := measureVisualRectangle()
        visualWidth := visualRectangle["right"] - visualRectangle["left"]
        adjustAttemptsForWidth += 1
    }

    adjustAttemptsForHeight := 0
    while Mod(visualHeight, 2) && adjustAttemptsForHeight < 6 {
        baseLogicalHeight += 1
        statusTextControl.Move(, , baseLogicalWidth, baseLogicalHeight)
        visualRectangle := measureVisualRectangle()
        visualHeight := visualRectangle["bottom"] - visualRectangle["top"]
        adjustAttemptsForHeight += 1
    }

    MonitorGetWorkArea(1, &workLeft, &workTop, &workRight, &workBottom)
    workAreaWidth  := workRight  - workLeft
    workAreaHeight := workBottom - workTop

    centeredX := Round(workLeft + (workAreaWidth  - visualWidth)  / 2)
    centeredY := Round(workTop  + (workAreaHeight - visualHeight) / 2)

    overlay["GUI"].Show("x" . centeredX . " y" . centeredY . " NoActivate")
    WinSetTransparent(overlayTransparency, overlay["GUI"].Hwnd)

    LogConclusion("Completed", logConclusionData)
}

OverlayInsertSpacer() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [], "Overlay Insert Spacer")
    
    ; Method has Custom Overlay Rules: Executed directly in LogBeginning.

    LogConclusion("Completed", logConclusionData)
}

OverlayUpdateCustomLine(overlayKey, overlayValue) {
    static methodName := RegisterMethod("overlayKey As Integer, value As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [overlayKey, overlayValue], "Overlay Update Custom Line")

    ; Method has Custom Overlay Rules: Executed directly in LogBeginning.

    LogConclusion("Completed", logConclusionData)
}

; **************************** ;
; Core Methods                 ;
; **************************** ;

AppendLineToLog(line, logType) {
    static newLine := "`r`n"

    if !system["Logging"]["Log to Array"] {
        callerWasCritical := A_IsCritical
        if !callerWasCritical {
            Critical "On"
        }

        try {
            FileAppend(line . newLine, system["Logging"]["Log File Path"][logType], "UTF-8-RAW")
        } finally {
            if !callerWasCritical {
                Critical "Off"
            }
        }
    } else {
        system["Logging"]["Log Entries"][logType].Push(line)
    }
}

LogBeginning(methodName, arguments := [], overlayValue := unset) {
    static lastRunTelemetryTick := unset
    static runTelemetryInterval := 12 * 60 * 1000

    timestamp := LogTimestamp()

    runTelemetryTick := A_TickCount
    if !IsSet(lastRunTelemetryTick) {
        lastRunTelemetryTick := runTelemetryTick
    }

    logConclusionData := Map(
        "Method Name", methodName
    )

    customOverlayMethod := (methodName = "OverlayInsertSpacer" || methodName = "OverlayUpdateCustomLine")

    operationSequenceNumber := IncrementCounter("Operation Sequence Number")
    if IsSet(overlayValue) {
        logConclusionData["Operation Sequence Number"] := EncodeIntegerToBase(operationSequenceNumber, 94)
        logConclusionData["Query Performance Counter"] := EncodeIntegerToBase(timestamp["QPC Midpoint Tick"], 94)
        logConclusionData["UTC Timestamp Integer"]     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)
    } else {
        logConclusionData["Operation Sequence Number"] := operationSequenceNumber
        logConclusionData["Query Performance Counter"] := timestamp["QPC Midpoint Tick"]
        logConclusionData["UTC Timestamp Integer"]     := timestamp["UTC Timestamp Integer"]
    }

    logBeginning := unset
    overlayKey   := -1
    if IsSet(overlayValue) {
        if !customOverlayMethod {
            if methodRegistry[methodName].Has("Overlay Log") {
                if methodRegistry[methodName]["Overlay Log"] {
                    overlayKey := IncrementCounter("Overlay")
                } else {
                    overlayKey := 0
                }
            }

            if overlayKey >= 1 {
                OverlayUpdateLine(overlayKey, overlayValue . overlay["Status"]["Beginning"])
            }
        } else {
            if methodName = "OverlayInsertSpacer" {
                overlayKey := IncrementCounter("Overlay")
                OverlayUpdateLine(overlayKey, overlayValue := "")
            } else if methodName = "OverlayUpdateCustomLine" {
                OverlayUpdateLine(overlayKey := arguments[1], overlayValue := arguments[2])
            }
        }

        logBeginning :=
            logConclusionData["Operation Sequence Number"] . "|" . ; Operation Sequence Number
            "B" .                                            "|" . ; Status
            logConclusionData["Query Performance Counter"] . "|" . ; Query Performance Counter
            logConclusionData["UTC Timestamp Integer"] .     "|" . ; UTC Timestamp Integer
            methodRegistry[methodName]["Symbol"]                   ; Method or Context
    }

    logConclusionData["Overlay Key"] := overlayKey

    if arguments.Length != 0 {
        validation := ""

        for index, argument in arguments {
            parameterName  := methodRegistry[methodName]["Parameter Contracts"][index]["Parameter Name"]
            dataType       := methodRegistry[methodName]["Parameter Contracts"][index]["Data Type"]
            dataConstraint := methodRegistry[methodName]["Parameter Contracts"][index]["Data Constraint"]
            optional       := methodRegistry[methodName]["Parameter Contracts"][index]["Optional"]
            whitelist      := methodRegistry[methodName]["Parameter Contracts"][index]["Whitelist"]

            validationOfArgument := ""
            if dataType = "String" && optional = "" && argument = "" {
                validationOfArgument := "No value passed as argument when required."
            } else if dataType = "String" && optional = "Optional" && argument = "" {
                ; Skip validation as it's optional.
            } else {
                validationOfArgument := ValidateDataUsingSpecification(argument, dataType, dataConstraint, whitelist)
            }

            if validationOfArgument != "" {
                if validation != "" {
                    validation := validation . " "
                }

                validation := validation . 'Parameter "' . parameterName . '" failed validation. ' . validationOfArgument
            }
        }

        if validation != "" {
            logConclusionData["Validation"] := validation
        }

        if logConclusionData["Overlay Key"] != -1 {
            if !customOverlayMethod {
                logConclusionData := LogProcessArguments(logConclusionData, arguments)

                logBeginning := logBeginning . "|" . 
                    logConclusionData["Arguments Log"]             ; Arguments or Error Message
            } else {
                logBeginning := logBeginning . "|" . 
                    ""                                             ; Arguments or Error Message
            }
        } else {
            logConclusionData["Arguments"] := arguments
        }
    }

    if logConclusionData["Overlay Key"] >= 1 {
        if !symbolLedger["Overlay"].Has(overlayValue) {
            logOverlaySymbolLedgerLine := RegisterSymbol(overlayValue, "Overlay", false)
            AppendLineToLog(logOverlaySymbolLedgerLine, "Symbol Ledger")
        }

        logBeginning := logBeginning . "|" . 
            EncodeIntegerToBase(overlayKey, 94) . "|" .            ; Overlay Key
            symbolLedger["Overlay"][overlayValue]                  ; Overlay Value
    }

    if IsSet(logBeginning) {
        AppendLineToLog(logBeginning, "Operation Log")
    }

    if logConclusionData.Has("Validation") {
        LogConclusion("Failed", logConclusionData, A_LineNumber, logConclusionData["Validation"])
    }

    if IsSet(overlayValue) {
        if runTelemetryTick - lastRunTelemetryTick >= runTelemetryInterval {
            lastRunTelemetryTick := runTelemetryTick
            system["Logging"]["Log Engine State"] := "Intermission"
            LogEngine()
        }
    }

    return logConclusionData
}

LogConclusion(conclusionStatus, logConclusionData, errorLineNumber := unset, errorMessage := unset) {
    timestamp := LogTimestamp()

    conclusionStatus := StrUpper(SubStr(conclusionStatus, 1, 1)) . StrLower(SubStr(conclusionStatus, 2))

    logConclusion := unset
    if conclusionStatus = "Failed" && logConclusionData["Overlay Key"] = -1 {
        logConclusionData["Operation Sequence Number"] := EncodeIntegerToBase(logConclusionData["Operation Sequence Number"], 94)
        logConclusionData["Query Performance Counter"] := EncodeIntegerToBase(logConclusionData["Query Performance Counter"], 94)
        logConclusionData["UTC Timestamp Integer"]     := EncodeIntegerToBase(logConclusionData["UTC Timestamp Integer"], 94)

        logBeginning :=
            logConclusionData["Operation Sequence Number"] . "|" .      ; Operation Sequence Number
            "B" .                                            "|" .      ; Status
            logConclusionData["Query Performance Counter"] . "|" .      ; Query Performance Counter
            logConclusionData["UTC Timestamp Integer"] .     "|" .      ; UTC Timestamp Integer
            methodRegistry[logConclusionData["Method Name"]]["Symbol"]  ; Method or Context

            if logConclusionData.Has("Arguments") {
                logConclusionData := LogProcessArguments(logConclusionData, logConclusionData["Arguments"])

                logBeginning := logBeginning . "|" . 
                    logConclusionData["Arguments Log"]                  ; Arguments or Error Message
            }

            AppendLineToLog(logBeginning, "Operation Log")
    }

    logConclusion := 
        logConclusionData["Operation Sequence Number"] .          "|" . ; Operation Sequence Number
        SubStr(conclusionStatus, 1, 1) .                          "|" . ; Status
        EncodeIntegerToBase(timestamp["QPC Midpoint Tick"], 94) . "|" . ; Query Performance Counter
        EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)     ; UTC Timestamp Integer

    if !logConclusionData.Has("Context") && IsSet(errorMessage) {
        logConclusionData["Context"] := ""
    }

    if logConclusionData.Has("Context") {
        if !symbolLedger["Context"].Has(logConclusionData["Context"]) {
            logSymbolLedgerLine := RegisterSymbol(logConclusionData["Context"], "Context", false)
            AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
        }

        logConclusionData["Context"] := symbolLedger["Context"][logConclusionData["Context"]]

        logConclusion := logConclusion . "|" . 
            logConclusionData["Context"]                                ; Method or Context
    }

    errorWindow             := unset
    constructedErrorMessage := unset
    if IsSet(errorMessage) {
        windowTitle          := "AutoHotkey v" . system["Runtime"]["AutoHotkey Version"] . ": " . A_ScriptName
        currentUtcDateTime   := ConvertIntegerToUtcTimestamp(system["Telemetry"]["UTC Timestamp Integer"] + timestamp["UTC Timestamp Integer"])

        if logConclusionData.Has("Validation") {
            errorLineNumber := methodRegistry[logConclusionData["Method Name"]]["Validation Line"]
        }
        
        declaration := RegExReplace(methodRegistry[logConclusionData["Method Name"]]["Declaration"], " <\d+>$", "")

        newLine := "`r`n"
        constructedErrorMessage := "Declaration: " .  declaration . " (" . system["Runtime"]["Library Release"] . ")" . newLine
        if methodRegistry[logConclusionData["Method Name"]]["Parameters"] != "" {
            constructedErrorMessage := constructedErrorMessage .
                "Parameters: " . methodRegistry[logConclusionData["Method Name"]]["Parameters"] . newLine . 
                "Arguments: " . logConclusionData["Arguments Full"] . newLine
        }

        constructedErrorMessage := constructedErrorMessage . 
            "Line Number: " . errorLineNumber . newLine

        logErrorMessage := StrReplace(constructedErrorMessage . "Error Output: " . errorMessage, newLine, "|")
        if !symbolLedger["Error"].Has(logErrorMessage) {
            logSymbolLedgerLine := RegisterSymbol(logErrorMessage, "Error", false)
            AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
        }

        logConclusion := logConclusion . "|" . 
            symbolLedger["Error"][logErrorMessage]                      ; Arguments or Error Message

        if system["Environment"].Has("Time Zone") {
            currentLocalDateTime := ConvertUtcTimestampToLocalTimestampWithTimeZoneKey(currentUtcDateTime, system["Environment"]["Time Zone"]["Key Name"])

            constructedErrorMessage := constructedErrorMessage . 
                "Date Runtime: " . currentLocalDateTime
        } else {
            constructedErrorMessage := constructedErrorMessage . 
                "Date Runtime: " . currentUtcDateTime . " (UTC)"
        }

        constructedErrorMessage := constructedErrorMessage . 
            newLine . "Error Output: " . errorMessage

        errorWindow := Gui("-Resize +AlwaysOnTop +OwnDialogs", windowTitle)
        errorWindow.SetFont("s10", "Segoe UI")
        errorWindow.AddEdit("ReadOnly r10 w1024 -VScroll vErrorTextField", constructedErrorMessage)

        exitButton := errorWindow.AddButton("w60 Default", "Exit")
        exitButton.OnEvent("Click", (*) => ExitApp())
        exitButton.Focus()
        errorWindow.OnEvent("Close", (*) => ExitApp())

        copyButton := errorWindow.AddButton("x+10 yp wp", "Copy")
        copyButton.OnEvent("Click", (*) => A_Clipboard := constructedErrorMessage)
    }

    AppendLineToLog(logConclusion, "Operation Log")

    if logConclusionData["Overlay Key"] >= 1 {
        OverlayUpdateStatus(logConclusionData, conclusionStatus)
    }

    if IsSet(errorMessage) { 
        runTelemetryLogFilePath := system["Logging"]["Log Shared Name"] . system["Logging"]["Log Date and Time"] . " - " . "Run Telemetry.csv"
        if system["Environment"].Has("Time Zone") && FileExist(runTelemetryLogFilePath) {
            system["Logging"]["Log Engine State"] := "Failed"
            LogEngine()
        }

        if logConclusionData["Method Name"] = "AbortExecution" {
            ExitApp()
        }

        errorWindow.Show("AutoSize Center")
        WinWaitClose("ahk_id " . errorWindow.Hwnd)
    }
}

LogEngine() {
    global symbolLedger
    global system

    static configuration := system["Configuration"]
    static directories   := system["Directories"]
    static environment   := system["Environment"]
    static logging       := system["Logging"]
    static runtime       := system["Runtime"]
    static telemetry     := system["Telemetry"]

    static newLine     := "`r`n"
    static systemDrive := SubStr(A_WinDir, 1, 3)

    switch logging["Log Engine State"] {
        case "Pending":
            kernel32ModuleHandle   := DllCall("GetModuleHandle", "Str", "Kernel32", "Ptr")
            preciseFunctionAddress := DllCall("GetProcAddress", "Ptr", kernel32ModuleHandle, "AStr", "GetSystemTimePreciseAsFileTime", "Ptr")

            if !preciseFunctionAddress {
                MsgBox("Windows 8.1 or higher required to use this framework.", "Unsupported Operating System", "IconX")
                ExitApp()
            }

            logging["Log Engine State"] := "Beginning"
        case "Running":
            logging["Log Engine State"] := "Completed"
    }

    settings := methodRegistry["LogEngine"]["Settings"]
    
    startMillisecondsTreshold                := settings["Start Milliseconds Treshold"].Get("Value")
    telemetryTimestampDurationInMilliseconds := settings["Telemetry Timestamp Duration in Milliseconds"].Get("Value")

    startMillisecondsTresholdCeiling := settings["Start Milliseconds Treshold"].Get("Ceiling")
    startMillisecondsTresholdFloor   := settings["Start Milliseconds Treshold"].Get("Floor")

    operationLogLineNumber := unset

    if logging["Log Engine State"] = "Beginning" && startMillisecondsTreshold >= startMillisecondsTresholdFloor && startMillisecondsTreshold < startMillisecondsTresholdCeiling + 1 {
        while A_MSec > startMillisecondsTreshold {
            Sleep(16)
        }
    }

    if logging["Log Engine State"] != "Beginning" && logging["Log Engine State"] != "Intermission" {
        if OverlayIsVisible() {
            OverlayChangeTransparency(255)
        }
    }

    runTelemetryOrder   := IncrementCounter("Run Telemetry Order")
    system["Telemetry"] := TelemetryTimestamp(telemetryTimestampDurationInMilliseconds)
    telemetry           := system["Telemetry"]

    if logging["Log Engine State"] = "Beginning" {
        SplitPath(A_ScriptFullPath, , , , &projectName)
        SplitPath(A_LineFile, , &librariesFolderPath)
        SplitPath(librariesFolderPath, , &sharedFolderPath, , &librariesVersion)
        SplitPath(sharedFolderPath, , &curatiumFolderPath)

        runtime["Project Name"]       := projectName
        runtime["Library Release"]    := SubStr(librariesVersion, InStr(librariesVersion, "(") + 1, InStr(librariesVersion, ")") - InStr(librariesVersion, "(") - 1)
        runtime["AutoHotkey Version"] := A_AhkVersion

        directories["Curatium"]  := curatiumFolderPath . "\"
        directories["Log"]       := directories["Curatium"] . "Log\"
        directories["Project"]   := directories["Curatium"] . "Projects\" . RTrim(SubStr(projectName, 1, InStr(projectName, "(") - 1)) . "\"
        directories["Projects"]  := directories["Curatium"] . "Projects\"
        directories["Shared"]    := sharedFolderPath . "\"
        directories["Constants"] := directories["Shared"] . "Constants\"
        directories["Images"]    := directories["Shared"] . "Images\"
        directories["Libraries"] := directories["Shared"] . "Libraries (" . runtime["Library Release"] . ")" . "\"
        directories["Mappings"]  := directories["Shared"] . "Mappings\"
        directories["Spreadsheet Operations Template"] := directories["Shared"] . "Spreadsheet Operations Template\"

        logging["Log Shared Name"]   := directories["Log"] . projectName . " - "
        logging["Log Date and Time"] := StrReplace(StrSplit(telemetry["UTC Timestamp Precise"], ".")[1], ":", ".")
        logging["Log File Path"]     := Map(
            "Execution Log", logging["Log Shared Name"] . logging["Log Date and Time"] . " - Execution Log.csv",
            "Operation Log", logging["Log Shared Name"] . logging["Log Date and Time"] . " - Operation Log.csv",
            "Run Telemetry", logging["Log Shared Name"] . logging["Log Date and Time"] . " - Run Telemetry.csv",
            "Symbol Ledger", logging["Log Shared Name"] . logging["Log Date and Time"] . " - Symbol Ledger.csv"
        )

        for baseCharacterSet in [
            94, 92, 86, 66, 62, 52
        ] {
            GetBaseCharacterSet(baseCharacterSet)
        }
    }

    static configurationPath    := directories["Project"] . "Configuration (" . runtime["Project Name"] . ", " . "Library Release" . " " . runtime["Library Release"] . ").json"
    static defaultConfiguration := StrReplace(
    '{' . newLine . 
        '    "Application Whitelist": [' . newLine . 
            '        ' .  newLine . 
        '    ],' . newLine . 
        '    "Application Executable Directory Candidates": [' . newLine .
            '        '  . newLine . 
        '    ],' . newLine . 
        '    "Candidate Base Directories": [' . newLine . 
            '        "' . systemDrive . 'Portable Files' . '",' . newLine . 
            '        "' . systemDrive . 'Program Files (Portable)' . '"' . newLine . 
            '    ],' . newLine . 
        '    "Settings": {' . newLine . 
            '        "Image Variant Preset": "' . directories["Constants"] . 'Heroes (2025-09-20).csv' . '",' . newLine . 
            '        "Application Image Override Directory": "' . "" . '",' . newline . 
            '        "Computer Alias": "' . "N/A" . '"' . newline . 
        '    }' . newLine . 
    '}', "\", "\\")

    if logging["Log Engine State"] != "Beginning" {
        operationLogLineNumber := GetTextFileLineCount(logging["Log File Path"]["Operation Log"])
    } else {
        operationLogLineNumber := 2
    }

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [], "Log Engine")

    if logging["Log Engine State"] = "Beginning" {
        RemoveDuplicatesFromArray([])
        BatchAppendSymbolLedger("", [])
        BatchAppendRunTelemetry("Beginning", [])
        BatchAppendOperationLog([])
        BatchAppendExecutionLog("Beginning", [])

        for reference in [
            "",
            "|",
            "<Constraint: Base64>",
            "<Data Type: Array>",
            "<Data Type: Integer>",
            "<Data Type: Map>",
            "<Data Type: String>",
            "<Text Block: Length: " . StrLen(defaultConfiguration) . ", Rows: " . StrSplit(defaultConfiguration, "`n").Length . ">"
        ] {
            RegisterReference(reference)
        }
    }

    system["Telemetry"]["System Drive Space Snapshot"] := GetDriveSpaceSnapshot(systemDrive)

    if logging["Log Engine State"] = "Beginning" {
        EnsureDirectoryExists(directories["Log"])
        EnsureDirectoryExists(directories["Project"])

        logging["Log to Array"] := false
        AppendLineToLog("Log|Type", "Execution Log")
        AppendLineToLog("Operation Sequence Number|Status|Query Performance Counter|UTC Timestamp Integer|Method or Context|Arguments or Error Message|Overlay Key|Overlay Value", "Operation Log")
        AppendLineToLog("Log|Type", "Run Telemetry")
        AppendLineToLog("Reference|Type|Symbol", "Symbol Ledger")
        logging["Log to Array"] := true

        environment["Time Zone"]            := GetTimeZone()
        environment["QPC Frequency"]        := GetQueryPerformanceCounterFrequency()
        environment["Session Startup Time"] := GetSessionStartupTime()

        for reference in [
            directories["Constants"] . "BIP-39 (2025-09-20).csv",
            directories["Constants"] . "EFF Dice-Generated Passphrases (2026-06-02).csv",
            directories["Constants"] . "Heroes (2025-09-20).csv",
            directories["Constants"] . "Middle-earth (2025-12-20).csv",
            directories["Constants"] . "NATO Phonetic Alphabet (2026-06-02).csv",
            directories["Constants"] . "Resolutions (2025-09-20).csv",
            directories["Constants"] . "Scales (2025-09-20).csv",
            directories["Constants"] . "XKCD Color Survey (2026-06-02).csv",
            "bdeca5734c5c8ca4a1adb2b5863c0cd46ac74837f24321235b5b7b1b32879229",
            "63d2175db6fb24702e49fbd72d339c4d8bd50c5a37804cbfc666e0ed04e843bf",
            "221c6504b42787aff09b43cb85a93511e3e4c06f52c084694119637c6794817d",
            "ffc72a6b738fdd75ea16964e6d43695c843ef2dea986d173196795e7d11d5dbd",
            "4222037720c26e12cffba2514436bc4b5029cdc3b3ccaa34f827415e8d46bbcf",
            "cc45d04bc98d76c9aa8ceb1e455c21082dfd8e6695c84b5382464bee2cd20364",
            "91eb6122786767eb83c7d87c43610fb87018d20ef2c25e43d3d38f31f49ec18d",
            "b4e194b06581c27bebaada8375a3dffa88e12cf815841574a614cd2249bcef87"
        ] {
            RegisterReference(reference)
        }

        heroes := directories["Constants"] . "Heroes (2025-09-20).csv"
        ExtractFilename(heroes)
        GetFileHash(heroes, "SHA-256")

        system["Constants"] := Map(
            "BIP-39",                         ConvertCsvToArrayOfMaps(directories["Constants"] . "BIP-39 (2025-09-20).csv"),
            "EFF Dice-Generated Passphrases", ConvertCsvToArrayOfMaps(directories["Constants"] . "EFF Dice-Generated Passphrases (2026-06-02).csv"),
            "Heroes",                         ConvertCsvToArrayOfMaps(directories["Constants"] . "Heroes (2025-09-20).csv"),
            "Middle-earth",                   ConvertCsvToArrayOfMaps(directories["Constants"] . "Middle-earth (2025-12-20).csv"),
            "NATO Phonetic Alphabet",         ConvertCsvToArrayOfMaps(directories["Constants"] . "NATO Phonetic Alphabet (2026-06-02).csv"),
            "Resolutions",                    ConvertCsvToArrayOfMaps(directories["Constants"] . "Resolutions (2025-09-20).csv"),
            "Scales",                         ConvertCsvToArrayOfMaps(directories["Constants"] . "Scales (2025-09-20).csv"),
            "XKCD Color Survey",              ConvertCsvToArrayOfMaps(directories["Constants"] . "XKCD Color Survey (2026-06-02).csv")
        )

        for index, rowMap in system["Constants"]["Resolutions"] {
            rowMap["Counter"] := index
        }

        for index, rowMap in system["Constants"]["Scales"] {
            rowMap["Counter"] := index
        }

        uefi := "Unified Extensible Firmware Interface "

        for reference in [
            directories["Mappings"] . "Application Executable Directory Candidates.csv",
            directories["Mappings"] . "Applications.csv",
            directories["Mappings"] . "Command Line Executables.csv",
            directories["Mappings"] . "File Signatures.csv",
            directories["Mappings"] . "System Management BIOS Type 17 Memory Device - Type.csv",
            directories["Mappings"] . uefi . "Advanced Configuration and Power Interface ID Official Registry.csv",
            directories["Mappings"] . uefi . "Plug and Play ID Official Registry.csv",
            directories["Mappings"] . uefi . "Plug and Play ID Unofficial Registry.csv"
        ] {
            RegisterReference(reference)
        }
        
        system["Mappings"] := Map(
            "Application Executable Directory Candidates",                            ConvertCsvToArrayOfMaps(directories["Mappings"] . "Application Executable Directory Candidates.csv"),
            "Applications",                                                           ConvertCsvToArrayOfMaps(directories["Mappings"] . "Applications.csv"),
            "Command Line Executables",                                               ConvertCsvToArrayOfMaps(directories["Mappings"] . "Command Line Executables.csv"),
            "File Signatures",                                                        ConvertCsvToArrayOfMaps(directories["Mappings"] . "File Signatures.csv"),
            "System Management BIOS Type 17 Memory Device - Type",                    ConvertCsvToArrayOfMaps(directories["Mappings"] . "System Management BIOS Type 17 Memory Device - Type.csv"),
            uefi . "Advanced Configuration and Power Interface ID Official Registry", ConvertCsvToArrayOfMaps(directories["Mappings"] . uefi . "Advanced Configuration and Power Interface ID Official Registry.csv"),
            uefi . "Plug and Play ID Official Registry",                              ConvertCsvToArrayOfMaps(directories["Mappings"] . uefi . "Plug and Play ID Official Registry.csv"),
            uefi . "Plug and Play ID Unofficial Registry",                            ConvertCsvToArrayOfMaps(directories["Mappings"] . uefi . "Plug and Play ID Unofficial Registry.csv")
        )

        for fileSignature in system["Mappings"]["File Signatures"] {
            fileSignature["Maximum Base64 Signature"] := ConvertHexStringToBase64(fileSignature["Maximum Hex Signature"])
            fileSignature["Minimal Base64 Signature"] := ConvertHexStringToBase64(fileSignature["Minimal Hex Signature"])
        }

        ExtractDirectory(heroes)
        ExtractParentDirectory(heroes)
        GetTextFileLineCount(heroes)
        ModifyScreenCoordinates(2, 2, "0x0")
        CombineCode("Intro", "Main")
        ComputeMouseMoveSpeed("0x0", "2x2")
        GetFoldersFromDirectory(directories["Log"])
        FileExistsInDirectory("Ultra High Definition", directories["Images"])
        OverlayIsVisible()
        ConvertIntegerToUtcTimestamp(telemetry["UTC Timestamp Integer"])
        ConvertUtcTimestampToInteger(telemetry["UTC Timestamp Precise"])
        ConvertUtcTimestampToLocalTimestampWithTimeZoneKey(telemetry["UTC Timestamp Precise"], environment["Time Zone"]["Key Name"])
        ConvertLocalTimestampToUtcTimestampWithTimeZoneKey(telemetry["UTC Timestamp Precise"], environment["Time Zone"]["Key Name"])
    }

    AppendLineToLog("Run Telemetry Order: " . runTelemetryOrder . "|" . "Operation Log Line Number: " . operationLogLineNumber . 
        "|" . "Duration in Milliseconds: " . telemetry["Duration in Milliseconds"] . "|" . "Number of Readings: " .  telemetry["Number of Readings"] . 
        "|" . "UTC Timestamp Precise: " . telemetry["UTC Timestamp Precise"] . "|" . "QPC Before Tick: " . telemetry["QPC Before Tick"] . 
        "|" . "QPC After Tick: " . telemetry["QPC After Tick"] . "|" . "QPC Midpoint Tick: " . telemetry["QPC Midpoint Tick"] . "|" . SubStr(logging["Log Engine State"], 1, 1), "Run Telemetry")

    telemetry["Computer Uptime in Seconds"] := Round(telemetry["QPC Midpoint Tick"] / environment["QPC Frequency"])
    telemetry["Session Uptime in Seconds"]  := DateDiff(SubStr(telemetry["UTC Timestamp Integer"], 1, 14) . "", environment["Session Startup Time"], "Seconds")
    AppendLineToLog("Computer Uptime in Seconds: " . telemetry["Computer Uptime in Seconds"] . "|" . "Session Uptime in Seconds: " . telemetry["Session Uptime in Seconds"], "Run Telemetry")

    system["Telemetry"]["System Resource Snapshot"] := GetSystemResourceSnapshot()
    AppendLineToLog("Physical: " . telemetry["System Resource Snapshot"]["Physical Used Percent"] . "%" . "|" . "Commit: " . telemetry["System Resource Snapshot"]["Commit Used Percent"] . "%" . 
        "|" . "Processes: " . telemetry["System Resource Snapshot"]["System Process Count"] . "|" . "Threads: " . telemetry["System Resource Snapshot"]["System Thread Count"], "Run Telemetry")

    AppendLineToLog("Disk Free Bytes: " . telemetry["System Drive Space Snapshot"]["Free Bytes"] . "|" . "Windows Free Size: " . telemetry["System Drive Space Snapshot"]["Windows Free Size"], "Run Telemetry")

    if logging["Log Engine State"] != "Beginning" {
        logging["Log to Array"] := false
        BatchAppendRunTelemetry(logging["Log Engine State"], logging["Log Entries"]["Run Telemetry"])
        logging["Log Entries"]["Run Telemetry"] := []
    }

    if logging["Log Engine State"] = "Beginning" {
        environment["BIOS"]                 := GetBios()
        environment["Color Mode"]           := GetColorMode()
        environment["Computer Name"]        := A_ComputerName
        environment["CPU"]                  := GetCpu()
        environment["Display GPU"]          := GetActiveDisplayGpu()
        environment["Display Language"]     := GetDisplayLanguage()
        environment["Display Resolution"]   := A_ScreenWidth . "x" . A_ScreenHeight
        environment["DPI Scale"]            := Round(A_ScreenDPI / 96 * 100) . "%"
        environment["Input Language"]       := GetInputLanguage()
        environment["International"]        := GetInternationalSnapshot()
        environment["Keyboard Layout"]      := GetActiveKeyboardLayout()
        environment["Memory Size and Type"] := GetMemorySizeAndType()
        environment["Monitor"]              := GetActiveMonitor()
        environment["Motherboard"]          := GetMotherboard()
        environment["Operating System"]     := GetOperatingSystem()
        environment["Refresh Rate"]         := GetActiveMonitorRefreshRateHz()
        environment["Regional Format"]      := environment["International"]["LocaleName"]
        environment["System Disk"]          := GetDiskModel(systemDrive)
        environment["Timeout Before Lock"]  := GetTimeoutBeforeLockInSeconds()
        environment["Username"]             := A_UserName

        runtime["Project Hash"]             := GetFileHash(directories["Projects"] . runtime["Project Name"] . ".ahk", "SHA-256")
        runtime["Application Library Hash"] := GetFileHash(directories["Libraries"] . "Application Library" . ".ahk", "SHA-256")
        runtime["Base Library Hash"]        := GetFileHash(directories["Libraries"] . "Base Library" .        ".ahk", "SHA-256")
        runtime["Chrono Library Hash"]      := GetFileHash(directories["Libraries"] . "Chrono Library" .      ".ahk", "SHA-256")
        runtime["File Library Hash"]        := GetFileHash(directories["Libraries"] . "File Library" .        ".ahk", "SHA-256")
        runtime["Image Library Hash"]       := GetFileHash(directories["Libraries"] . "Image Library" .       ".ahk", "SHA-256")
        runtime["Logging Library Hash"]     := GetFileHash(directories["Libraries"] . "Logging Library" .     ".ahk", "SHA-256")

        DefineApplicationRegistry()

        if !FileExist(configurationPath) {
            WriteTextToFile(defaultConfiguration, configurationPath, "UTF-8", "Create")
        }

        ValidateConfiguration(configurationPath)
        configuration := system["Configuration"]
        
        configuration["Path"]            := configurationPath
        configuration["Number of Lines"] := GetTextFileLineCount(configuration["Path"])

        for executionLogLine in [
            "Project Name: " .             runtime["Project Name"],
            "Library Release: " .          runtime["Library Release"],
            "AutoHotkey Version: " .       runtime["AutoHotkey Version"],
            "Project Hash: " .             runtime["Project Hash"],
            "Application Library Hash: " . runtime["Application Library Hash"],
            "Base Library Hash: " .        runtime["Base Library Hash"],
            "Chrono Library Hash: " .      runtime["Chrono Library Hash"],
            "File Library Hash: " .        runtime["File Library Hash"],
            "Image Library Hash: " .       runtime["Image Library Hash"],
            "Logging Library Hash: " .     runtime["Logging Library Hash"],
            "Operating System: " .         environment["Operating System"]["Full Name"] . "|" . 
                "Installation Date: " .    environment["Operating System"]["Installation Date"],
            "Computer Name: " .            environment["Computer Name"],
            "Computer Alias: " .           configuration["Settings"]["Computer Alias"],
            "Username: " .                 environment["Username"],
            "Time Zone Key Name: " .       environment["Time Zone"]["Key Name"],
            "Country or Region: " .        environment["International"]["Geo"]["Friendly Name"] . "|" . 
                "ISO 3166-1 alpha-2: " .   environment["International"]["Geo"]["ISO 3166-1 alpha-2"] . "|" . 
                "ISO 3166-1 alpha-3: " .   environment["International"]["Geo"]["ISO 3166-1 alpha-3"] . "|" . 
                "ISO 3166-1 numeric: " .   environment["International"]["Geo"]["ISO 3166-1 numeric"],
            "Display Language: " .         environment["Display Language"],
            "Regional Format: " .          environment["Regional Format"],
            "Input Language: " .           environment["Input Language"],
            "Keyboard Layout: " .          environment["Keyboard Layout"],
            "Timeout Before Lock: " .      environment["Timeout Before Lock"],
            "Motherboard: " .              environment["Motherboard"],
            "CPU: " .                      environment["CPU"],
            "Memory Size and Type: " .     environment["Memory Size and Type"],
            "System Disk: " .              environment["System Disk"] . "|" . 
                "Disk Total Bytes: " .     telemetry["System Drive Space Snapshot"]["Total Bytes"] . "|" . 
                "Windows Total Size: " .   telemetry["System Drive Space Snapshot"]["Windows Total Size"],
            "Display GPU: " .              environment["Display GPU"],
            "Monitor: " .                  environment["Monitor"],
            "BIOS: " .                     environment["BIOS"],
            "QPC Frequency: " .            environment["QPC Frequency"],
            "Display Resolution: " .       environment["Display Resolution"],
            "Refresh Rate: " .             environment["Refresh Rate"],
            "DPI Scale: " .                environment["DPI Scale"],
            "Color Mode: " .               environment["Color Mode"]
        ] {
            AppendLineToLog(executionLogLine, "Execution Log")
        }
    }

    logConclusionData["Context"] := "Log Engine State: " . logging["Log Engine State"]
    LogConclusion("Completed", logConclusionData)

    if logging["Log Engine State"] != "Beginning" && logging["Log Engine State"] != "Intermission" {
        timestampNow := A_Now
        for logType, filePath in logging["Log File Path"] {
            if !FileExist(filePath) {
                continue
            }

            file := FileOpen(filePath, "rw")

            fileSize := file.Length
            if fileSize = 0 {
                file.Close()
                continue
            }
            
            bytesToRead := (fileSize >= 2) ? 2 : 1
            file.Seek(-bytesToRead, 2)
            tail := Buffer(bytesToRead)
            file.RawRead(tail, bytesToRead)

            if bytesToRead = 2 && NumGet(tail, 0, "UChar") = 13 && NumGet(tail, 1, "UChar") = 10 {
                file.Length := fileSize - 2
            } else if NumGet(tail, bytesToRead - 1, "UChar") = 10 {
                file.Length := fileSize - 1
            }

            file.Close()

            FileSetTime(timestampNow, logging["Log File Path"][logType], "M")
        }

        logging.Delete("Log File Path")
        logging.Delete("Log Shared Name")
        logging["Counters"] := Map(
            "Context",                   0,
            "Error",                     0,
            "Method",                    0,
            "Overlay",                   0,
            "Reference",                 0,
            "Operation Sequence Number", 0,
            "Run Telemetry Order",       0
        )

        logging["Log Engine State"] := "Pending"
    }

    if logging["Log Engine State"] = "Intermission" {
        logging["Log Engine State"] := "Running"
    }

    if logging["Log Engine State"] = "Beginning" {
        logging["Log to Array"] := false
        BatchAppendExecutionLog(logging["Log Engine State"], logging["Log Entries"]["Execution Log"])
        BatchAppendOperationLog(logging["Log Entries"]["Operation Log"])
        BatchAppendRunTelemetry(logging["Log Engine State"], logging["Log Entries"]["Run Telemetry"])
        BatchAppendSymbolLedger("", logging["Log Entries"]["Symbol Ledger"])

        logging["Log Entries"] := Map(
            "Execution Log", [],
            "Operation Log", [],
            "Run Telemetry", [],
            "Symbol Ledger", []
        )

        logging["Log Engine State"] := "Running"
    }
}

LogProcessArguments(logConclusionData, arguments) {
    global symbolLedger

    argumentsFormatted := Map(
        "Arguments Full", "",
        "Arguments Log",  ""
    )

    methodName := logConclusionData["Method Name"]
    for index, argument in arguments {
        argumentsFormatted["Parameter"]       := methodRegistry[methodName]["Parameter Contracts"][index]["Parameter Name"]
        argumentsFormatted["Argument"]        := argument
        argumentsFormatted["Data Type"]       := methodRegistry[methodName]["Parameter Contracts"][index]["Data Type"]
        argumentsFormatted["Data Constraint"] := methodRegistry[methodName]["Parameter Contracts"][index]["Data Constraint"]
        argumentsFormatted["Whitelist"]       := methodRegistry[methodName]["Parameter Contracts"][index]["Whitelist"]

        argumentValueFull := argument
        argumentValueLog  := argument
        switch argumentsFormatted["Data Type"] {
            case "Array", "Map", "Object":
                argumentValueFull := "<Data Type: " . Type(argument) . ">"
                argumentValueLog  := RegisterReference(argumentValueFull)

                argumentValueFull := Format('"{1}"', argumentValueFull)
                argumentValueLog  := Format('"{1}"', argumentValueLog)
            case "Boolean":
                if Type(argument) != "Integer" {
                    argumentValueFull := "<Data Type: " . Type(argument) . ">"
                    argumentValueLog  := RegisterReference(argumentValueFull)

                    argumentValueFull := Format('"{1}"', argumentValueFull)
                    argumentValueLog  := Format('"{1}"', argumentValueLog)
                }
            case "Integer":
                if Type(argument) != argumentsFormatted["Data Type"] {
                    argumentValueFull := "<Data Type: " . Type(argument) . ">"
                    argumentValueLog  := RegisterReference(argumentValueFull)

                    argumentValueFull := Format('"{1}"', argumentValueFull)
                    argumentValueLog  := Format('"{1}"', argumentValueLog)
                }
            case "String":
                if Type(argument) != argumentsFormatted["Data Type"] {
                    argumentValueFull := "<Data Type: " . Type(argument) . ">"
                } else if argumentsFormatted["Data Constraint"] = "Base64" {
                    argumentValueFull := "<Constraint: Base64>"
                } else if InStr(argument, "`n") || InStr(argument, "`r") {
                    argumentValueFull := "<Text Block: Length: " . StrLen(argument) . ", Rows: " . StrSplit(argument, "`n").Length . ">"
                } else {
                    if StrLen(argument) >= 255 {
                        argumentValueFull := SubStr(argument, 1, 255) . "…"
                    }
                }

                if argumentsFormatted["Whitelist"].Length != 0 && !logConclusionData.Has("Validation") {
                    argumentValueLog  := symbolLedger["Whitelist"][argumentValueLog]
                    argumentValueLog  := Format('\{1}\', argumentValueLog)
                } else {
                    argumentValueLog  := RegisterReference(argumentValueFull)
                    argumentValueLog  := Format('"{1}"', argumentValueLog)
                }

                argumentValueFull := Format('"{1}"', argumentValueFull)
            case "Variant":
                if Type(argument) = "String" {
                    argumentValueLog  := RegisterReference(argumentValueFull)

                    argumentValueFull := Format('"{1}"', argumentValueFull)
                    argumentValueLog  := Format('"{1}"', argumentValueLog)
                }
        }

        argumentsFormatted["Arguments Full"] .= argumentValueFull
        argumentsFormatted["Arguments Log"]  .= argumentValueLog

        if index < arguments.Length {
            argumentsFormatted["Arguments Full"] .= ", "
            argumentsFormatted["Arguments Log"]  .= ", "
        }
    }

    logConclusionData["Arguments Full"] := argumentsFormatted["Arguments Full"]
    logConclusionData["Arguments Log"]  := argumentsFormatted["Arguments Log"]

    return logConclusionData
}

LogTimestamp() {
    queryPerformanceCounterBefore           := GetQueryPerformanceCounter()
    utcTimestampInteger                     := GetUtcTimestampInteger()
    queryPerformanceCounterAfter            := GetQueryPerformanceCounter()

    utcTimestampInteger                     := utcTimestampInteger - system["Telemetry"]["UTC Timestamp Integer"]
    queryPerformanceCounterMeasurementDelta := queryPerformanceCounterAfter - queryPerformanceCounterBefore
    queryPerformanceCounterMidpointTick     := (queryPerformanceCounterBefore + (queryPerformanceCounterMeasurementDelta // 2)) - system["Telemetry"]["QPC Midpoint Tick"]

    logTimestamp := Map(
        "QPC Midpoint Tick",     queryPerformanceCounterMidpointTick,
        "UTC Timestamp Integer", utcTimestampInteger
    )

    return logTimestamp
}

OverlayUpdateLine(overlayKey, overlayValue) {
    global overlay

    if !overlay["Lines"].Has(overlayKey) {
        overlay["Order"].Push(overlayKey)
    }
    overlay["Lines"][overlayKey] := overlayValue

    newText := ""
    for lineKey in overlay["Order"] {
        newText .= (newText != "" ? "`n" : "") . overlay["Lines"][lineKey]
    }
    overlay["GUI"]["StatusText"].Text := newText
}

OverlayUpdateStatus(logConclusionData, newStatus) {
    global overlay

    overlaykey := logConclusionData["Overlay Key"]

    currentText := overlay["Lines"][overlayKey]

    if logConclusionData["Method Name"] !== "OverlayInsertSpacer" && logConclusionData["Method Name"] !== "OverlayUpdateCustomLine" {
        switch newStatus {
            case "Skipped":
                OverlayUpdateLine(overlayKey, StrReplace(currentText, overlay["Status"]["Beginning"], overlay["Status"]["Skipped"]))
            case "Completed":
                OverlayUpdateLine(overlayKey, StrReplace(currentText, overlay["Status"]["Beginning"], overlay["Status"]["Completed"]))
            case "Failed":
                OverlayUpdateLine(overlayKey, StrReplace(currentText, overlay["Status"]["Beginning"], overlay["Status"]["Failed"]))
        }
    }
}

RegisterReference(referenceValue) {
    if !symbolLedger["Reference"].Has(referenceValue) {
        logSymbolLedgerLine := RegisterSymbol(referenceValue, "Reference", false)
        AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
    }

    return symbolLedger["Reference"][referenceValue]
}

RegisterSymbol(value, type, addNewLine := true) {
    global symbolLedger

    static newLine := "`r`n"

    switch StrLower(type) {
        case "context":
            type := "Context"
        case "error":
            type := "Error"
        case "method":
            type := "Method"
        case "overlay":
            type := "Overlay"
        case "reference":
            type := "Reference"
        case "whitelist":
            type := "Whitelist"
    }

    symbolLine := ""
    if !symbolLedger[type].Has(value) {
        counter := IncrementCounter(type)

        if type = "Reference" || type = "Whitelist" {
            symbolLedger[type][value] := EncodeIntegerToBase(counter, 92)
        } else {
            symbolLedger[type][value] := EncodeIntegerToBase(counter, 94)
        }

        typeCharacter := SubStr(type, 1, 1)

        symbolLine :=
            value . "|" . 
            typeCharacter . "|" . 
            symbolLedger[type][value]
    }

    if addNewLine {
        symbolLine := symbolLine . newLine
    }

    return symbolLine
}

; **************************** ;
; Encoding & Decoding Methods  ;
; **************************** ;

GetBaseCharacterSet(baseType) {
    ; https://web.archive.org/web/20260310005954/https://www.utf8-chartable.de/
    static cachedBase52Result := unset
    static cachedBase62Result := unset
    static cachedBase66Result := unset
    static cachedBase86Result := unset
    static cachedBase92Result := unset
    static cachedBase94Result := unset

    static excludedAsciiCodePoints := Map(
        0x7C, 94, ; | VERTICAL LINE
        0x22, 92, ; " QUOTATION MARK
        0x5C, 92, ; \ REVERSE SOLIDUS
        0x2A, 86, ; * ASTERISK
        0x2F, 86, ; / SOLIDUS
        0x3A, 86, ; : COLON
        0x3C, 86, ; < LESS-THAN SIGN
        0x3E, 86, ; > GREATER-THAN SIGN
        0x3F, 86, ; ? QUESTION MARK
        0x20, 66, ;   SPACE
        0x21, 66, ; ! EXCLAMATION MARK
        0x23, 66, ; # NUMBER SIGN
        0x24, 66, ; $ DOLLAR SIGN
        0x25, 66, ; % PERCENT SIGN
        0x26, 66, ; & AMPERSAND
        0x27, 66, ; ' APOSTROPHE
        0x28, 66, ; ( LEFT PARENTHESIS
        0x29, 66, ; ) RIGHT PARENTHESIS
        0x2B, 66, ; + PLUS SIGN
        0x2C, 66, ; , COMMA
        0x3B, 66, ; ; SEMICOLON
        0x3D, 66, ; = EQUALS SIGN
        0x40, 66, ; @ COMMERCIAL AT
        0x5B, 66, ; [ LEFT SQUARE BRACKET
        0x5D, 66, ; ] RIGHT SQUARE BRACKET
        0x5E, 66, ; ^ CIRCUMFLEX ACCENT
        0x60, 66, ; ` GRAVE ACCENT
        0x7B, 66, ; { LEFT CURLY BRACKET
        0x7D, 66, ; } RIGHT CURLY BRACKET
        0x2D, 62, ; - HYPHEN-MINUS
        0x2E, 62, ; . FULL STOP
        0x5F, 62, ; _ LOW LINE
        0x7E, 62, ; ~ TILDE
        0x30, 52, ; 0 DIGIT ZERO
        0x31, 52, ; 1 DIGIT ONE
        0x32, 52, ; 2 DIGIT TWO
        0x33, 52, ; 3 DIGIT THREE
        0x34, 52, ; 4 DIGIT FOUR
        0x35, 52, ; 5 DIGIT FIVE
        0x36, 52, ; 6 DIGIT SIX
        0x37, 52, ; 7 DIGIT SEVEN
        0x38, 52, ; 8 DIGIT EIGHT
        0x39, 52, ; 9 DIGIT NINE
    )

    cachedResult := unset

    switch baseType {
        case 52:
            if IsSet(cachedBase52Result) {
                cachedResult := cachedBase52Result
            }
        case 62:
            if IsSet(cachedBase62Result) {
                cachedResult := cachedBase62Result
            }
        case 66:
            if IsSet(cachedBase66Result) {
                cachedResult := cachedBase66Result
            }
        case 86:
            if IsSet(cachedBase86Result) {
                cachedResult := cachedBase86Result
            }
        case 92:
            if IsSet(cachedBase92Result) {
                cachedResult := cachedBase92Result
            }
        case 94:
            if IsSet(cachedBase94Result) {
                cachedResult := cachedBase94Result
            }
    }

    if !IsSet(cachedResult) {
        baseCharacters := ""
        loop 0x7E - 0x20 + 1 {
            codePoint := 0x20 + A_Index - 1

            maxBaseForExclusion := excludedAsciiCodePoints.Get(codePoint, 0)
            if baseType <= maxBaseForExclusion {
                continue
            }

            baseCharacters .= Chr(codePoint)
        }

        baseRadix := StrLen(baseCharacters)

        baseDigitByCharacterMap := Map()
        loop StrLen(baseCharacters) {
            baseCharacter := SubStr(baseCharacters, A_Index, 1)
            baseDigitByCharacterMap[baseCharacter] := A_Index - 1
        }

        baseCharacterArray := []
        currentIndex := 1
        while currentIndex <= baseRadix {
            baseCharacterArray.Push(SubStr(baseCharacters, currentIndex, 1))
            currentIndex += 1
        }

        digitBytesBuffer := Buffer(baseRadix, 0)
        digitValueIndex := 0
        while digitValueIndex < baseRadix {
            NumPut("UChar", Ord(baseCharacterArray[digitValueIndex + 1]), digitBytesBuffer, digitValueIndex)
            digitValueIndex += 1
        }

        digitMapBytesBuffer := Buffer(256, 0xFF)
        digitValueIndex := 0
        while digitValueIndex < baseRadix {
            byteValue := Ord(baseCharacterArray[digitValueIndex + 1])
            NumPut("UChar", digitValueIndex, digitMapBytesBuffer, byteValue)
            digitValueIndex += 1
        }

        cachedResult := Map(
            "Characters",      baseCharacters,
            "Base Radix",      baseRadix,
            "Digit Map",       baseDigitByCharacterMap,
            "Char Array",      baseCharacterArray,
            "Digit Bytes",     digitBytesBuffer,
            "Digit Map Bytes", digitMapBytesBuffer
        )

        switch baseType {
            case 52: cachedBase52Result := cachedResult
            case 62: cachedBase62Result := cachedResult
            case 66: cachedBase66Result := cachedResult
            case 86: cachedBase86Result := cachedResult
            case 92: cachedBase92Result := cachedResult
            case 94: cachedBase94Result := cachedResult
        }
    }

    return cachedResult
}

EncodeIntegerToBase(integerValue, baseType) {
    baseCharacterSet := GetBaseCharacterSet(baseType)
    baseRadix        := baseCharacterSet["Base Radix"]
    charArray        := baseCharacterSet["Char Array"]

    baseText := ""
    if integerValue = 0 {
        baseText := charArray[1]
    } else {
        while integerValue >= baseRadix {
            digitValue   := Mod(integerValue, baseRadix)
            baseText     := charArray[digitValue + 1] . baseText
            integerValue := integerValue // baseRadix
        }

        baseText := charArray[integerValue + 1] . baseText
    }

    return baseText
}

DecodeBaseToInteger(baseText, baseType) {
    baseCharacterSet := GetBaseCharacterSet(baseType)
    baseRadix        := baseCharacterSet["Base Radix"]
    digitMap         := baseCharacterSet["Digit Map"]

    integerValue := 0
    loop parse, baseText {
        integerValue := integerValue * baseRadix + digitMap[A_LoopField]
    }

    return integerValue
}

EncodeSha256HexToBase(hexSha256, baseType) {
    baseCharacterSet  := GetBaseCharacterSet(baseType)
    characters        := baseCharacterSet["Characters"]
    baseRadix         := baseCharacterSet["Base Radix"]
    charArray         := baseCharacterSet["Char Array"]

    static requiredLengthByBaseMap := Map(52, 45, 66, 43, 86, 40, 94, 40)

    hexSha256 := StrLower(hexSha256)

    sha256BytesBuffer := Buffer(32, 0)
    writeOffset := 0
    loop 32 {
        twoHexDigits := SubStr(hexSha256, (A_Index - 1) * 2 + 1, 2)
        byteValue    := ("0x" . twoHexDigits) + 0
        NumPut("UChar", byteValue, sha256BytesBuffer, writeOffset)
        writeOffset += 1
    }

    baseDigitsLeastSignificantFirst := []
    isAllZero := true
    byteIndex := 0
    while byteIndex < 32 {
        if NumGet(sha256BytesBuffer, byteIndex, "UChar") {
            isAllZero := false
            break
        }

        byteIndex += 1
    }

    if isAllZero {
        baseDigitsLeastSignificantFirst.Push(0)
    } else {
        loop {
            remainderValue := 0
            hasNonZeroQuotientByte := false
            byteIndex := 0
            while byteIndex < 32 {
                currentByte    := NumGet(sha256BytesBuffer, byteIndex, "UChar")
                accumulator    := remainderValue * 256 + currentByte
                quotientByte   := accumulator // baseRadix
                remainderValue := accumulator - quotientByte * baseRadix
                NumPut("UChar", quotientByte, sha256BytesBuffer, byteIndex)
                if quotientByte != 0 {
                    hasNonZeroQuotientByte := true
                }
                byteIndex += 1
            }
            baseDigitsLeastSignificantFirst.Push(remainderValue)
            if !hasNonZeroQuotientByte {
                break
            }
        }
    }

    baseText := ""
    digitIndex := baseDigitsLeastSignificantFirst.Length
    while digitIndex >= 1 {
        digitValue := baseDigitsLeastSignificantFirst[digitIndex]
        baseText   .= charArray[digitValue + 1]
        digitIndex -= 1
    }

    requiredLength := requiredLengthByBaseMap.Has(baseType) ? requiredLengthByBaseMap[baseType] : Ceil(256 * Log(2) / Log(baseRadix))

    zeroDigit := charArray[1]
    while StrLen(baseText) < requiredLength {
        baseText := zeroDigit . baseText
    }

    return baseText
}

DecodeBaseToSha256Hex(baseText, baseType) {
    baseCharacterSet := GetBaseCharacterSet(baseType)
    baseRadix        := baseCharacterSet["Base Radix"]
    digitMap         := baseCharacterSet["Digit Map"]

    static sha256BytesBuffer := Buffer(32, 0)
    resetIndex := 0
    while resetIndex < 32 {
        NumPut("UChar", 0, sha256BytesBuffer, resetIndex)
        resetIndex += 1
    }

    loop parse, baseText {
        baseCharacter := A_LoopField
        digitValue    := digitMap[baseCharacter]

        carryValue := digitValue
        byteIndex := 31
        while byteIndex >= 0 {
            currentByte  := NumGet(sha256BytesBuffer, byteIndex, "UChar")
            productValue := currentByte * baseRadix + carryValue
            NumPut("UChar", productValue & 0xFF, sha256BytesBuffer, byteIndex)
            carryValue := productValue // 256
            byteIndex -= 1
        }
    }

    static lowercaseHexStringByByteArray := ""
    if Type(lowercaseHexStringByByteArray) != "Array" {
        temporary := []
        temporary.Capacity := 256
        index := 0
        while index < 256 {
            temporary.Push(Format("{:02x}", index))
            index += 1
        }
        lowercaseHexStringByByteArray := temporary
    }

    hexOutput := ""
    byteIndex := 0
    while byteIndex < 32 {
        currentByte := NumGet(sha256BytesBuffer, byteIndex, "UChar")
        hexOutput   .= lowercaseHexStringByByteArray[currentByte + 1]
        byteIndex   += 1
    }

    return hexOutput
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

BatchAppendExecutionLog(executionType, array) {
    static executionTypeWhitelist := Format('"{1}", "{2}"', "Application", "Beginning")
    static methodName := RegisterMethod("executionType As String [Whitelist: " . executionTypeWhitelist . "], array as Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [executionType, array])

    static newLine := "`r`n"

    switch StrLower(executionType) {
        case "application":
            executionType := "A"
        case "beginning":
            executionType := "B"
    }

    if array.Length != 0 {
        consolidatedExecutionLog := ""
        for index, value in array {
            if array.Length != index {
                consolidatedExecutionLog := consolidatedExecutionLog . value . "|" . executionType . newLine
            } else {
                consolidatedExecutionLog := consolidatedExecutionLog . value . "|" . executionType
            }
        }

        AppendLineToLog(consolidatedExecutionLog, "Execution Log")
    }
}

BatchAppendOperationLog(array) {
    static methodName := RegisterMethod("array as Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [array])

    static newLine := "`r`n"

    if array.Length != 0 {
        consolidatedOperationLog := ""
        for index, value in array {
            if array.Length != index {
                consolidatedOperationLog := consolidatedOperationLog . value . newLine
            } else {
                consolidatedOperationLog := consolidatedOperationLog . value
            }
        }

        AppendLineToLog(consolidatedOperationLog, "Operation Log")
    }
}

BatchAppendRunTelemetry(appendType, array) {
    static appendTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}"', "Beginning", "Completed", "Failed", "Intermission")
    static methodName := RegisterMethod("appendType As String [Whitelist: " . appendTypeWhitelist . "], array as Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [appendType, array])

    static newLine := "`r`n"

    switch StrLower(appendType) {
        case "beginning":
            appendType := "B"
        case "completed":
            appendType := "C"
        case "failed":
            appendType := "F"
        case "intermission":
            appendType := "I"
    }

    if array.Length != 0 {
        consolidatedRunTelemetry := ""
        for index, value in array {
            if array.Length != index {
                consolidatedRunTelemetry := consolidatedRunTelemetry . value . "|" . appendType . newLine
            } else {
                consolidatedRunTelemetry := consolidatedRunTelemetry . value . "|" . appendType
            }
        }

        AppendLineToLog(consolidatedRunTelemetry, "Run Telemetry")
    }
}

BatchAppendSymbolLedger(symbolType, array) {
    static symbolTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Context", "Error", "Method", "Overlay", "Reference", "Whitelist")
    static methodName := RegisterMethod("symbolType As String [Optional] [Whitelist: " . symbolTypeWhitelist . "], array As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [symbolType, array])

    static newLine := "`r`n"

    switch StrLower(symbolType) {
        case "context":
            symbolType := "Context"
        case "error":
            symbolType := "Error"
        case "method":
            symbolType := "Method"
        case "overlay":
            symbolType := "Overlay"
        case "reference":
            symbolType := "Reference"
        case "whitelist":
            symbolType := "Whitelist"
    }

    symbolLedgerArray := []
    if symbolType = "" {
        for value in array {
            symbolLedgerArray.Push(value)
        }
    } else {
        for value in array {
            if !symbolLedger[symbolType].Has(value) {
                symbolLedgerArray.Push(value)
            }
        }
    }

  
    if symbolLedgerArray.Length != 0 {
        symbolLedgerArray := RemoveDuplicatesFromArray(symbolLedgerArray)

        consolidatedSymbolLedger := ""
        if symbolType = "" {
            for index, value in symbolLedgerArray {
                if symbolLedgerArray.Length != index {
                    consolidatedSymbolLedger := consolidatedSymbolLedger . value . newLine
                } else {
                    consolidatedSymbolLedger := consolidatedSymbolLedger . value
                }
            }
        } else {
            for index, value in symbolLedgerArray {
                if symbolLedgerArray.Length != index {
                    consolidatedSymbolLedger := consolidatedSymbolLedger . RegisterSymbol(value, symbolType)
                } else {
                    consolidatedSymbolLedger := consolidatedSymbolLedger . RegisterSymbol(value, symbolType, false)
                }
            }
        }

        AppendLineToLog(consolidatedSymbolLedger, "Symbol Ledger")
    }
}

OverlayIsVisible() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName)

    windowHandle  := overlay["GUI"].Hwnd
    windowVisible := unset

    if DllCall("User32\IsWindowVisible", "Ptr", windowHandle) {
        windowVisible := true
    } else {
        windowVisible := false
    }

    return windowVisible
}