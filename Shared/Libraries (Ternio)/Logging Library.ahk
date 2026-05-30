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
        currentLocalDateTime := ConvertUtcTimestampToLocalTimestampWithTimeZoneKey(currentUtcDateTime, system["Environment"]["Time Zone Key Name"])

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

        constructedErrorMessage := constructedErrorMessage . 
            "Date Runtime: " . currentLocalDateTime . newLine . 
            "Error Output: " . errorMessage

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
        system["Logging"]["Log Engine State"] := "Failed"
        LogEngine()

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

    static newLine     := "`r`n"
    static systemDrive := SubStr(A_WinDir, 1, 3)

    static configuration := "Configuration"
    static directories   := "Directories"
    static environment   := "Environment"
    static logging       := "Logging"
    static runtime       := "Runtime"
    static telemetry     := "Telemetry"

    operationLogLineNumber := unset

    switch system[logging]["Log Engine State"] {
        case "Pending": system[logging]["Log Engine State"] := "Beginning"
        case "Running": system[logging]["Log Engine State"] := "Completed"
    }

    while true {
        milliseconds := A_MSec + 0

        if milliseconds <= 584 {
            break
        } else {
            Sleep(1016 - milliseconds)
            continue
        }
    }

    if system[logging]["Log Engine State"] != "Beginning" && system[logging]["Log Engine State"] != "Intermission" {
        if OverlayIsVisible() {
            OverlayChangeTransparency(255)
        }
    }

    runTelemetryOrder := IncrementCounter("Run Telemetry Order")
    system[telemetry] := TelemetryTimestamp(200)

    if system[logging]["Log Engine State"] = "Beginning" {
        SplitPath(A_ScriptFullPath, , , , &projectName)
        SplitPath(A_LineFile, , &librariesFolderPath)
        SplitPath(librariesFolderPath, , &sharedFolderPath, , &librariesVersion)
        SplitPath(sharedFolderPath, , &curatiumFolderPath)

        system[runtime]["Project Name"]       := projectName
        system[runtime]["Library Release"]    := SubStr(librariesVersion, InStr(librariesVersion, "(") + 1, InStr(librariesVersion, ")") - InStr(librariesVersion, "(") - 1)
        system[runtime]["AutoHotkey Version"] := A_AhkVersion

        system[directories]["Curatium"]  := curatiumFolderPath . "\"
        system[directories]["Log"]       := system[directories]["Curatium"] . "Log\"
        system[directories]["Project"]   := system[directories]["Curatium"] . "Projects\" . RTrim(SubStr(projectName, 1, InStr(projectName, "(") - 1)) . "\"
        system[directories]["Projects"]  := system[directories]["Curatium"] . "Projects\"
        system[directories]["Shared"]    := sharedFolderPath . "\"
        system[directories]["Constants"] := system[directories]["Shared"] . "Constants\"
        system[directories]["Images"]    := system[directories]["Shared"] . "Images\"
        system[directories]["Libraries"] := system[directories]["Shared"] . "Libraries (" . system[runtime]["Library Release"] . ")" . "\"
        system[directories]["Mappings"]  := system[directories]["Shared"] . "Mappings\"
        system[directories]["Spreadsheet Operations Template"] := system[directories]["Shared"] . "Spreadsheet Operations Template\"

        system[logging]["Log Shared Name"]   := system[directories]["Log"] . projectName . " - "
        system[logging]["Log Date and Time"] := StrReplace(StrSplit(system[telemetry]["UTC Timestamp Precise"], ".")[1], ":", ".")
        system[logging]["Log File Path"]     := Map(
            "Execution Log", system[logging]["Log Shared Name"] . system[logging]["Log Date and Time"] . " - Execution Log.csv",
            "Operation Log", system[logging]["Log Shared Name"] . system[logging]["Log Date and Time"] . " - Operation Log.csv",
            "Run Telemetry", system[logging]["Log Shared Name"] . system[logging]["Log Date and Time"] . " - Run Telemetry.csv",
            "Symbol Ledger", system[logging]["Log Shared Name"] . system[logging]["Log Date and Time"] . " - Symbol Ledger.csv"
        )
    }

    if system[logging]["Log Engine State"] != "Beginning" {
        operationLogLineNumber := GetTextFileLineCount(system["Logging"]["Log File Path"]["Operation Log"])
    }

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, [], "Log Engine")

    static configurationPath    := system[directories]["Project"] . "Configuration (" . system[runtime]["Project Name"] . ", " . "Library Release" . " " . system[runtime]["Library Release"] . ").json"
    static defaultConfiguration := StrReplace(
    '{' . newLine . 
        '    "Application Whitelist": [' . newLine . 
            '        ' .  newLine . 
        '    ],' . newLine . 
        '    "Application Executable Directory Candidates": [' . newLine .
            '        '  . newLine . 
        '    ],' . newLine . 
        '    "Candidate Base Directories": [' . newLine . 
            '        "' . ExtractDirectory(A_WinDir) . 'Portable Files' . '",' . newLine . 
            '        "' . ExtractDirectory(A_WinDir) . 'Program Files (Portable)' . '"' . newLine . 
            '    ],' . newLine . 
        '    "Settings": {' . newLine . 
            '        "Image Variant Preset": "' . system[directories]["Constants"] . 'Heroes (2025-09-20).csv' . '",' . newLine . 
            '        "Application Image Override Directory": "' . "" . '",' . newline . 
            '        "Computer Alias": "' . "N/A" . '"' . newline . 
        '    }' . newLine . 
    '}', "\", "\\")

    if system[logging]["Log Engine State"] = "Beginning" {
        for reference in [
            configurationPath,
            "",
            "|",
            "<Constraint: Base64>",
            "<Data Type: Array>",
            "<Data Type: Integer>",
            "<Data Type: Map>",
            "<Data Type: String>",
            system[directories]["Log"],
            system[directories]["Project"],
            system[directories]["Constants"] . "Resolutions (2025-09-20).csv",
            system[directories]["Constants"] . "Scales (2025-09-20).csv",
            system[directories]["Images"] .    "Image Library Catalog (Full High Definition).csv",
            system[directories]["Images"] .    "Image Library Catalog (Quad High Definition).csv",
            system[directories]["Images"] .    "Image Library Catalog (Ultra High Definition).csv",
            system[directories]["Mappings"] .  "Application Executable Directory Candidates.csv",
            system[directories]["Mappings"] .  "Applications.csv",
            system[directories]["Mappings"] .  "Command Line Executables.csv",
            system[directories]["Mappings"] .  "File Signatures.csv",
            system[directories]["Mappings"] .  "System Management BIOS Type 17 Memory Device - Type.csv",
            system[directories]["Mappings"] .  "Unified Extensible Firmware Interface Advanced Configuration and Power Interface ID Official Registry.csv",
            system[directories]["Mappings"] .  "Unified Extensible Firmware Interface Plug and Play ID Official Registry.csv",
            system[directories]["Mappings"] .  "Unified Extensible Firmware Interface Plug and Play ID Unofficial Registry.csv",
            system[directories]["Spreadsheet Operations Template"] . "Version Manifest.ini",
            "Log|Type",
            "Operation Sequence Number|Status|Query Performance Counter|UTC Timestamp Integer|Method or Context|Arguments or Error Message|Overlay Key|Overlay Value",
            "Reference|Type|Symbol",
            system[logging]["Log Shared Name"] . "Template.csv",
            "<Text Block: Length: " . StrLen(defaultConfiguration) . ", Rows: " . StrSplit(defaultConfiguration, "`n").Length . ">",
            "Full High Definition",
            "Quad High Definition",
            "Ultra High Definition"
        ] {
            RegisterReference(reference)
        }

        RemoveDuplicatesFromArray([])
        system[environment]["Time Zone Key Name"] := GetTimeZoneKeyName()
        system[environment]["QPC Frequency"]      := GetQueryPerformanceCounterFrequency()
        BatchAppendSymbolLedger("", [])
        BatchAppendRunTelemetry("Beginning", [])
        BatchAppendOperationLog([])
        BatchAppendExecutionLog("Beginning", [])

        EnsureDirectoryExists(system[directories]["Log"])

        WriteTextToFile("Log|Type", system[logging]["Log Shared Name"] . "Template.csv", "UTF-8", "Overwrite")
        FileMove(system[logging]["Log Shared Name"] . "Template.csv", system[logging]["Log File Path"]["Execution Log"])
        WriteTextToFile("Operation Sequence Number|Status|Query Performance Counter|UTC Timestamp Integer|Method or Context|Arguments or Error Message|Overlay Key|Overlay Value", system[logging]["Log Shared Name"] . "Template.csv", "UTF-8", "Overwrite")
        FileMove(system[logging]["Log Shared Name"] . "Template.csv", system[logging]["Log File Path"]["Operation Log"])
        WriteTextToFile("Log|Type", system[logging]["Log Shared Name"] . "Template.csv", "UTF-8", "Overwrite")
        FileMove(system[logging]["Log Shared Name"] . "Template.csv", system[logging]["Log File Path"]["Run Telemetry"])
        WriteTextToFile("Reference|Type|Symbol", system[logging]["Log Shared Name"] . "Template.csv", "UTF-8", "Overwrite")
        FileMove(system[logging]["Log Shared Name"] . "Template.csv", system[logging]["Log File Path"]["Symbol Ledger"])

        operationLogLineNumber := GetTextFileLineCount(system["Logging"]["Log File Path"]["Operation Log"])

        EnsureDirectoryExists(system[directories]["Project"])

        system[logging]["Log to Array"] := false
        AppendLineToLog("", "Execution Log")
        AppendLineToLog("", "Operation Log")
        AppendLineToLog("", "Run Telemetry")
        AppendLineToLog("", "Symbol Ledger")
    }

    if system[logging]["Log Engine State"] != "Beginning" {
        system[logging]["Log to Array"] := false
    }

    AppendLineToLog("Run Telemetry Order: " . runTelemetryOrder . "|" . "Operation Log Line Number: " . operationLogLineNumber . 
        "|" . "Duration in Milliseconds: " . system[telemetry]["Duration in Milliseconds"] . "|" . "Number of Readings: " .  system[telemetry]["Number of Readings"] . 
        "|" . "UTC Timestamp Precise: " . system[telemetry]["UTC Timestamp Precise"] . "|" . "QPC Before Tick: " . system[telemetry]["QPC Before Tick"] . 
        "|" . "QPC After Tick: " . system[telemetry]["QPC After Tick"] . "|" . "QPC Midpoint Tick: " . system[telemetry]["QPC Midpoint Tick"] . "|" . SubStr(system[logging]["Log Engine State"], 1, 1), "Run Telemetry")

    system[logging]["Log to Array"] := true

    if system[logging]["Log Engine State"] = "Beginning" {
        system[environment]["Session Startup Time"] := GetSessionStartupTime()
        ModifyScreenCoordinates(2, 2, "0x0")
        CombineCode("Intro", "Main")
        ComputeMouseMoveSpeed("0x0", "2x2")
        GetFoldersFromDirectory(system[directories]["Log"])
        FileExistsInDirectory("Ultra High Definition", system[directories]["Images"])
        ExtractParentDirectory(system[directories]["Log"])
        OverlayIsVisible()
        ConvertIntegerToUtcTimestamp(system[telemetry]["UTC Timestamp Integer"])
        ConvertUtcTimestampToInteger(system[telemetry]["UTC Timestamp Precise"])
        ConvertUtcTimestampToLocalTimestampWithTimeZoneKey(system[telemetry]["UTC Timestamp Precise"], system[environment]["Time Zone Key Name"])
    }

    system[telemetry]["Computer Uptime in Seconds"]  := Round(system[telemetry]["QPC Midpoint Tick"] / system[environment]["QPC Frequency"])
    system[telemetry]["Session Uptime in Seconds"]   := DateDiff(SubStr(system[telemetry]["UTC Timestamp Integer"], 1, 14) . "", system[environment]["Session Startup Time"], "Seconds")
    system[telemetry]["System Drive Space Snapshot"] := GetDriveSpaceSnapshot(systemDrive)
    system[telemetry]["System Resource Snapshot"]    := GetSystemResourceSnapshot()

    AppendLineToLog("Computer Uptime in Seconds: " . system[telemetry]["Computer Uptime in Seconds"] . "|" . "Session Uptime in Seconds: " . system[telemetry]["Session Uptime in Seconds"], "Run Telemetry")
    AppendLineToLog("Physical: " . system[telemetry]["System Resource Snapshot"]["Physical Used Percent"] . "%" . "|" . "Commit: " . system[telemetry]["System Resource Snapshot"]["Commit Used Percent"] . "%" . 
        "|" . "Processes: " . system[telemetry]["System Resource Snapshot"]["System Process Count"] . "|" . "Threads: " . system[telemetry]["System Resource Snapshot"]["System Thread Count"], "Run Telemetry")
    AppendLineToLog("Disk Free Bytes: " . system[telemetry]["System Drive Space Snapshot"]["Free Bytes"] . "|" . "Windows Free Size: " . system[telemetry]["System Drive Space Snapshot"]["Windows Free Size"], "Run Telemetry")

    if system[logging]["Log Engine State"] != "Beginning" {
        system[logging]["Log to Array"] := false
        BatchAppendRunTelemetry(system[logging]["Log Engine State"], system[logging]["Log Entries"]["Run Telemetry"])
        system[logging]["Log Entries"]["Run Telemetry"] := []
    }

    if system[logging]["Log Engine State"] = "Beginning" {
        system[environment]["BIOS"]                 := GetBios()
        system[environment]["Color Mode"]           := GetWindowsColorMode()
        system[environment]["CPU"]                  := GetCpu()
        system[environment]["Display GPU"]          := GetActiveDisplayGpu()
        system[environment]["Display Resolution"]   := A_ScreenWidth . "x" . A_ScreenHeight
        system[environment]["DPI Scale"]            := Round(A_ScreenDPI / 96 * 100) . "%"
        system[environment]["Input Language"]       := GetInputLanguage()
        system[environment]["International"]        := GetInternationalFormatting()
        system[environment]["Keyboard Layout"]      := GetActiveKeyboardLayout()
        system[environment]["Memory Size and Type"] := GetMemorySizeAndType()
        system[environment]["Monitor"]              := GetActiveMonitor()
        system[environment]["Motherboard"]          := GetMotherboard()
        system[environment]["Operating System"]     := GetOperatingSystem()
        system[environment]["OS Installation Date"] := GetWindowsInstallationDateUtcTimestamp()
        system[environment]["Computer Name"]        := A_ComputerName
        system[environment]["Refresh Rate"]         := GetActiveMonitorRefreshRateHz()
        system[environment]["Region Format"]        := GetRegionFormat()
        system[environment]["System Disk"]          := GetDiskModel(systemDrive)
        system[environment]["Timeout Before Lock"]  := GetTimeoutBeforeLockInSeconds()
        system[environment]["Username"]             := A_UserName

        system[runtime]["Project Hash"]             := GetFileHash(system[directories]["Projects"] . system[runtime]["Project Name"] . ".ahk", "SHA-256")
        system[runtime]["Application Library Hash"] := GetFileHash(system[directories]["Libraries"] . "Application Library" . ".ahk", "SHA-256")
        system[runtime]["Base Library Hash"]        := GetFileHash(system[directories]["Libraries"] . "Base Library" .        ".ahk", "SHA-256")
        system[runtime]["Chrono Library Hash"]      := GetFileHash(system[directories]["Libraries"] . "Chrono Library" .      ".ahk", "SHA-256")
        system[runtime]["File Library Hash"]        := GetFileHash(system[directories]["Libraries"] . "File Library" .        ".ahk", "SHA-256")
        system[runtime]["Image Library Hash"]       := GetFileHash(system[directories]["Libraries"] . "Image Library" .       ".ahk", "SHA-256")
        system[runtime]["Logging Library Hash"]     := GetFileHash(system[directories]["Libraries"] . "Logging Library" .     ".ahk", "SHA-256")

        DefineApplicationRegistry()

        if !FileExist(configurationPath) {
            WriteTextToFile(defaultConfiguration, configurationPath, "UTF-8", "Create")
        }

        ValidateConfiguration(configurationPath)
        system[configuration]["Path"]            := configurationPath
        system[configuration]["Number of Lines"] := GetTextFileLineCount(system[configuration]["Path"])

        for executionLogLine in [
            "Project Name: " .             system[runtime]["Project Name"],
            "Library Release: " .          system[runtime]["Library Release"],
            "AutoHotkey Version: " .       system[runtime]["AutoHotkey Version"],
            "Project Hash: " .             system[runtime]["Project Hash"],
            "Application Library Hash: " . system[runtime]["Application Library Hash"],
            "Base Library Hash: " .        system[runtime]["Base Library Hash"],
            "Chrono Library Hash: " .      system[runtime]["Chrono Library Hash"],
            "File Library Hash: " .        system[runtime]["File Library Hash"],
            "Image Library Hash: " .       system[runtime]["Image Library Hash"],
            "Logging Library Hash: " .     system[runtime]["Logging Library Hash"],
            "Operating System: " .         system[environment]["Operating System"],
            "OS Installation Date: " .     system[environment]["OS Installation Date"],
            "Computer Name: " .            system[environment]["Computer Name"],
            "Computer Alias: " .           system[configuration]["Settings"]["Computer Alias"],
            "Username: " .                 system[environment]["Username"],
            "Time Zone Key Name: " .       system[environment]["Time Zone Key Name"],
            "Region Format: " .            system[environment]["Region Format"],
            "Input Language: " .           system[environment]["Input Language"],
            "Keyboard Layout: " .          system[environment]["Keyboard Layout"],
            "Timeout Before Lock: " .      system[environment]["Timeout Before Lock"],
            "Motherboard: " .              system[environment]["Motherboard"],
            "CPU: " .                      system[environment]["CPU"],
            "Memory Size and Type: " .     system[environment]["Memory Size and Type"],
            "System Disk: " .              system[environment]["System Disk"] . "|" . 
                "Disk Total Bytes: " .     system[telemetry]["System Drive Space Snapshot"]["Total Bytes"] . "|" . 
                "Windows Total Size: " .   system[telemetry]["System Drive Space Snapshot"]["Windows Total Size"],
            "Display GPU: " .              system[environment]["Display GPU"],
            "Monitor: " .                  system[environment]["Monitor"],
            "BIOS: " .                     system[environment]["BIOS"],
            "QPC Frequency: " .            system[environment]["QPC Frequency"],
            "Display Resolution: " .       system[environment]["Display Resolution"],
            "Refresh Rate: " .             system[environment]["Refresh Rate"],
            "DPI Scale: " .                system[environment]["DPI Scale"],
            "Color Mode: " .               system[environment]["Color Mode"]
        ] {
            AppendLineToLog(executionLogLine, "Execution Log")
        }
    }

    logConclusionData["Context"] := "Log Engine State: " . system[logging]["Log Engine State"]
    LogConclusion("Completed", logConclusionData)

    if system[logging]["Log Engine State"] != "Beginning" && system[logging]["Log Engine State"] != "Intermission" {
        timestampNow := A_Now
        for logType, filePath in system[logging]["Log File Path"] {
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

            FileSetTime(timestampNow, system[logging]["Log File Path"][logType], "M")
        }

        system[logging].Delete("Log File Path")
        system[logging].Delete("Log Shared Name")
        system[logging]["Counters"] := Map(
            "Context",                   0,
            "Error",                     0,
            "Method",                    0,
            "Overlay",                   0,
            "Reference",                 0,
            "Operation Sequence Number", 0,
            "Run Telemetry Order",       0
        )

        system[logging]["Log Engine State"] := "Pending"
    }

    if system[logging]["Log Engine State"] = "Intermission" {
        system[logging]["Log Engine State"] := "Running"
    }

    if system[logging]["Log Engine State"] = "Beginning" {
        system[logging]["Log to Array"] := false
        BatchAppendExecutionLog(system[logging]["Log Engine State"], system[logging]["Log Entries"]["Execution Log"])
        BatchAppendOperationLog(system[logging]["Log Entries"]["Operation Log"])
        BatchAppendRunTelemetry(system[logging]["Log Engine State"], system[logging]["Log Entries"]["Run Telemetry"])
        BatchAppendSymbolLedger("", system[logging]["Log Entries"]["Symbol Ledger"])

        system[logging]["Log Entries"] := Map(
            "Execution Log", [],
            "Operation Log", [],
            "Run Telemetry", [],
            "Symbol Ledger", []
        )

        system[logging]["Log Engine State"] := "Running"
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

                if argumentsFormatted["Whitelist"].Length != 0 {
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
        case "context", "c":
            type := "Context"
        case "error", "e":
            type := "Error"
        case "method", "m":
            type := "Method"
        case "overlay", "o":
            type := "Overlay"
        case "reference", "r":
            type := "Reference"
        case "whitelist", "r":
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
        excludedAsciiCodePoints := Map()

        if baseType <= 94 {
            excludedAsciiCodePoints[0x7C] := true ; | VERTICAL LINE
        }

        if baseType <= 92 {
            excludedAsciiCodePoints[0x22] := true ; " QUOTATION MARK
            excludedAsciiCodePoints[0x5C] := true ; \ REVERSE SOLIDUS
        }

        if baseType <= 86 {
            excludedAsciiCodePoints[0x2A] := true ; * ASTERISK
            excludedAsciiCodePoints[0x2F] := true ; / SOLIDUS
            excludedAsciiCodePoints[0x3A] := true ; : COLON
            excludedAsciiCodePoints[0x3C] := true ; < LESS-THAN SIGN
            excludedAsciiCodePoints[0x3E] := true ; > GREATER-THAN SIGN
            excludedAsciiCodePoints[0x3F] := true ; ? QUESTION MARK

        }

        if baseType <= 66 {
            excludedAsciiCodePoints[0x20] := true ;   SPACE
            excludedAsciiCodePoints[0x21] := true ; ! EXCLAMATION MARK
            excludedAsciiCodePoints[0x23] := true ; # NUMBER SIGN
            excludedAsciiCodePoints[0x24] := true ; $ DOLLAR SIGN
            excludedAsciiCodePoints[0x25] := true ; % PERCENT SIGN
            excludedAsciiCodePoints[0x26] := true ; & AMPERSAND
            excludedAsciiCodePoints[0x27] := true ; ' APOSTROPHE
            excludedAsciiCodePoints[0x28] := true ; ( LEFT PARENTHESIS
            excludedAsciiCodePoints[0x29] := true ; ) RIGHT PARENTHESIS
            excludedAsciiCodePoints[0x2B] := true ; + PLUS SIGN
            excludedAsciiCodePoints[0x2C] := true ; , COMMA
            excludedAsciiCodePoints[0x3B] := true ; ; SEMICOLON
            excludedAsciiCodePoints[0x3D] := true ; = EQUALS SIGN
            excludedAsciiCodePoints[0x40] := true ; @ COMMERCIAL AT
            excludedAsciiCodePoints[0x5B] := true ; [ LEFT SQUARE BRACKET
            excludedAsciiCodePoints[0x5D] := true ; ] RIGHT SQUARE BRACKET
            excludedAsciiCodePoints[0x5E] := true ; ^ CIRCUMFLEX ACCENT
            excludedAsciiCodePoints[0x60] := true ; ` GRAVE ACCENT
            excludedAsciiCodePoints[0x7B] := true ; { LEFT CURLY BRACKET
            excludedAsciiCodePoints[0x7D] := true ; } RIGHT CURLY BRACKET
        }    

        if baseType <= 62 {
            excludedAsciiCodePoints[0x2D] := true ; - HYPHEN-MINUS
            excludedAsciiCodePoints[0x2E] := true ; . FULL STOP
            excludedAsciiCodePoints[0x5F] := true ; _ LOW LINE
            excludedAsciiCodePoints[0x7E] := true ; ~ TILDE
        }

        if baseType <= 52 {
            excludedAsciiCodePoints[0x30] := true ; 0 DIGIT ZERO
            excludedAsciiCodePoints[0x31] := true ; 1 DIGIT ONE
            excludedAsciiCodePoints[0x32] := true ; 2 DIGIT TWO
            excludedAsciiCodePoints[0x33] := true ; 3 DIGIT THREE
            excludedAsciiCodePoints[0x34] := true ; 4 DIGIT FOUR
            excludedAsciiCodePoints[0x35] := true ; 5 DIGIT FIVE
            excludedAsciiCodePoints[0x36] := true ; 6 DIGIT SIX
            excludedAsciiCodePoints[0x37] := true ; 7 DIGIT SEVEN
            excludedAsciiCodePoints[0x38] := true ; 8 DIGIT EIGHT
            excludedAsciiCodePoints[0x39] := true ; 9 DIGIT NINE

        }

        baseCharacters := ""
        loop 0x7E - 0x20 + 1 {
            codePoint := 0x20 + A_Index - 1

            if !excludedAsciiCodePoints.Has(codePoint) {
                baseCharacters .= Chr(codePoint)
            }
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
        byteValue := ("0x" . twoHexDigits) + 0
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
                currentByte := NumGet(sha256BytesBuffer, byteIndex, "UChar")
                accumulator := remainderValue * 256 + currentByte
                quotientByte := accumulator // baseRadix
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
        case "beginning", "b":
            appendType := "B"
        case "completed", "c":
            appendType := "C"
        case "failed", "f":
            appendType := "F"
        case "intermission", "i":
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
        case "context", "c":
            symbolType := "Context"
        case "error", "e":
            symbolType := "Error"
        case "method", "m":
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