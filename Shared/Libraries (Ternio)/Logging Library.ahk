#Requires AutoHotkey v2.0
#Include Application Library.ahk
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Image Library.ahk

; Press Escape to abort the script early when running or to close the script when it's completed.
$Esc:: {
    if !system["Logging"].Has("Log to Array") {
        Critical "On"
        AbortExecution()
    } else {
        ExitApp()
    }
}

AbortExecution() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Abort Execution")

    LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Execution aborted early by pressing escape.")
}

OverlayChangeTransparency(transparencyValue) {
    static methodName := RegisterMethod("transparencyValue As Integer [Constraint: Byte]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [transparencyValue], "Overlay Change Transparency (" . transparencyValue . ")")

    WinSetTransparent(transparencyValue, "ahk_id " . overlay["GUI"].Hwnd)

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayChangeVisibility() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Overlay Change Visibility")

    if DllCall("User32\IsWindowVisible", "Ptr", overlay["GUI"].Hwnd) {
        overlay["GUI"].Hide()
    } else {
        overlay["GUI"].Show("NoActivate")
    }

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayHideLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("methodNameInput As String [Constraint: Locator]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [methodNameInput], "Overlay Hide Log for Method (" . methodNameInput . ")")

    global methodRegistry
    
    if !methodRegistry.Has(methodNameInput) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, 'Method "' . methodNameInput . '" not registered.')
    }

    methodRegistry[methodNameInput]["Overlay Log"] := false

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayShowLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("methodNameInput As String [Constraint: Locator]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [methodNameInput], "Overlay Show Log for Method (" . methodNameInput . ")")

    global methodRegistry

    if methodRegistry.Has(methodNameInput) {
        methodRegistry[methodNameInput]["Overlay Log"] := true
    } else {
        methodRegistry[methodNameInput] := Map(
            "Overlay Log", true
        )
    }

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayStart() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Overlay Start")

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

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayInsertSpacer() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Overlay Insert Spacer")
    
    ; Method has Custom Overlay Rules: Executed directly in LogBeginning.

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayUpdateCustomLine(overlayKey, overlayValue) {
    static methodName := RegisterMethod("overlayKey As Integer, value As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [overlayKey, overlayValue], "Overlay Update Custom Line")

    ; Method has Custom Overlay Rules: Executed directly in LogBeginning.

    LogConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Core Methods                 ;
; **************************** ;

AppendLineToLog(line, logType) {
    static newLine := "`r`n"

    if !system["Logging"].Has("Log to Array") {
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
        system["Logging"]["Log Entries"].Push([line, logType])
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

    logValuesForConclusion := Map(
        "Method Name", methodName
    )

    if IsSet(overlayValue) {
        logValuesForConclusion["Operation Sequence Number"] := EncodeIntegerToBase(NextOperationSequenceNumber(), 94)
        logValuesForConclusion["Query Performance Counter"] := EncodeIntegerToBase(timestamp["QPC Midpoint Tick"], 94)
        logValuesForConclusion["UTC Timestamp Integer"]     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)
    } else {
        logValuesForConclusion["Operation Sequence Number"] := NextOperationSequenceNumber()
        logValuesForConclusion["Query Performance Counter"] := timestamp["QPC Midpoint Tick"]
        logValuesForConclusion["UTC Timestamp Integer"]     := timestamp["UTC Timestamp Integer"]
    }

    logBeginning := unset
    overlayKey   := unset
    if IsSet(overlayValue) {
        if methodName != "OverlayInsertSpacer" && methodName != "OverlayUpdateCustomLine" {
            overlayKey := OverlayGenerateNextKey(methodName)

            if overlayKey != 0 {
                OverlayUpdateLine(overlayKey, overlayValue . overlay["Status"]["Beginning"])
            }
        } else {
            if methodName = "OverlayInsertSpacer" {
                OverlayUpdateLine(overlayKey := OverlayGenerateNextKey(), overlayValue := "")
            } else if methodName = "OverlayUpdateCustomLine" {
                OverlayUpdateLine(overlayKey := arguments[1], overlayValue := arguments[2])
            }
        }

        logBeginning :=
            logValuesForConclusion["Operation Sequence Number"] . "|" . ; Operation Sequence Number
            "B" .                                                 "|" . ; Status
            logValuesForConclusion["Query Performance Counter"] . "|" . ; Query Performance Counter
            logValuesForConclusion["UTC Timestamp Integer"] .     "|" . ; UTC Timestamp Integer
            methodRegistry[methodName]["Symbol"]                        ; Method or Context
    } else {
        overlayKey := -1
    }

    logValuesForConclusion["Overlay Key"] := overlayKey

    if arguments.Length != 0 {
        validation := LogValidateMethodArguments(methodName, arguments)
        if validation != "" {
            logValuesForConclusion["Validation"] := validation
        }

        if logValuesForConclusion["Overlay Key"] != -1 {
            logValuesForConclusion := LogFormatMethodArguments(logValuesForConclusion, arguments)

            logBeginning := logBeginning . "|" . 
                logValuesForConclusion["Arguments Log"] ; Arguments or Error Message
        } else {
            logValuesForConclusion["Arguments"] := arguments
        }
    }

    if logValuesForConclusion["Overlay Key"] >= 1 {
        if !symbolLedger.Has(overlayValue . "|O") {
            logOverlaySymbolLedgerLine := RegisterSymbol(overlayValue, "Overlay", false)
            AppendLineToLog(logOverlaySymbolLedgerLine, "Symbol Ledger")
        }

        logBeginning := logBeginning . "|" . 
            EncodeIntegerToBase(overlayKey, 94) . "|" . ; Overlay Key
            symbolLedger[overlayValue . "|O"]           ; Overlay Value
    }

    if IsSet(logBeginning) {
        AppendLineToLog(logBeginning, "Operation Log")
    }

    if logValuesForConclusion.Has("Validation") {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, logValuesForConclusion["Validation"])
    }

    if IsSet(overlayValue) {
        if runTelemetryTick - lastRunTelemetryTick >= runTelemetryInterval {
            lastRunTelemetryTick := runTelemetryTick
            LogEngine("Intermission")
        }
    }

    return logValuesForConclusion
}

LogConclusion(conclusionStatus, logValuesForConclusion, errorLineNumber := unset, errorMessage := unset) {
    timestamp := LogTimestamp()

    conclusionStatus := StrUpper(SubStr(conclusionStatus, 1, 1)) . StrLower(SubStr(conclusionStatus, 2))

    logConclusion := unset
    if conclusionStatus = "Failed" && logValuesForConclusion["Overlay Key"] = -1 {
        logValuesForConclusion["Operation Sequence Number"] := EncodeIntegerToBase(logValuesForConclusion["Operation Sequence Number"], 94)
        logValuesForConclusion["Query Performance Counter"] := EncodeIntegerToBase(logValuesForConclusion["Query Performance Counter"], 94)
        logValuesForConclusion["UTC Timestamp Integer"]     := EncodeIntegerToBase(logValuesForConclusion["UTC Timestamp Integer"], 94)

        logBeginning :=
            logValuesForConclusion["Operation Sequence Number"] . "|" .     ; Operation Sequence Number
            "B" .                                                 "|" .     ; Status
            logValuesForConclusion["Query Performance Counter"] . "|" .     ; Query Performance Counter
            logValuesForConclusion["UTC Timestamp Integer"] .     "|" .     ; UTC Timestamp Integer
            methodRegistry[logValuesForConclusion["Method Name"]]["Symbol"] ; Method or Context

            if logValuesForConclusion.Has("Arguments") {
                logValuesForConclusion := LogFormatMethodArguments(logValuesForConclusion, logValuesForConclusion["Arguments"])

                logBeginning := logBeginning . "|" . 
                    logValuesForConclusion["Arguments Log"] ; Arguments or Error Message
            }

            AppendLineToLog(logBeginning, "Operation Log")
    }

    logConclusion := 
        logValuesForConclusion["Operation Sequence Number"] .     "|" . ; Operation Sequence Number
        SubStr(conclusionStatus, 1, 1) .                          "|" . ; Status
        EncodeIntegerToBase(timestamp["QPC Midpoint Tick"], 94) . "|" . ; Query Performance Counter
        EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)     ; UTC Timestamp Integer

    if !logValuesForConclusion.Has("Context") && IsSet(errorMessage) {
        logValuesForConclusion["Context"] := ""
    }

    if logValuesForConclusion.Has("Context") {
        if !symbolLedger.Has(logValuesForConclusion["Context"] . "|C") {
            logSymbolLedgerLine := RegisterSymbol(logValuesForConclusion["Context"], "Context", false)
            AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
        }

        logValuesForConclusion["Context"] := symbolLedger[logValuesForConclusion["Context"] . "|C"]

        logConclusion := logConclusion . "|" . 
            logValuesForConclusion["Context"] ; Method or Context
    }

    errorWindow             := unset
    constructedErrorMessage := unset
    if IsSet(errorMessage) {
        windowTitle          := "AutoHotkey v" . system["Runtime"]["AutoHotkey Version"] . ": " . A_ScriptName
        currentUtcDateTime   := ConvertIntegerToUtcTimestamp(system["Snapshot"]["UTC Timestamp Integer"] + timestamp["UTC Timestamp Integer"])
        currentLocalDateTime := ConvertUtcTimestampToLocalTimestampWithTimeZoneKey(currentUtcDateTime, system["Environment"]["Time Zone Key Name"])

        if logValuesForConclusion.Has("Validation") {
            errorLineNumber := methodRegistry[logValuesForConclusion["Method Name"]]["Validation Line"]
        }
        
        declaration := RegExReplace(methodRegistry[logValuesForConclusion["Method Name"]]["Declaration"], " <\d+>$", "")

        newLine := "`r`n"
        constructedErrorMessage := "Declaration: " .  declaration . " (" . system["Runtime"]["Library Release"] . ")" . newLine
        if methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] != "" {
            constructedErrorMessage := constructedErrorMessage .
                "Parameters: " . methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] . newLine . 
                "Arguments: " . logValuesForConclusion["Arguments Full"] . newLine
        }

        constructedErrorMessage := constructedErrorMessage . 
            "Line Number: " . errorLineNumber . newLine

        logErrorMessage := StrReplace(constructedErrorMessage . "Error Output: " . errorMessage, newLine, "|")
        if !symbolLedger.Has(logErrorMessage . "|E") {
            logSymbolLedgerLine := RegisterSymbol(logErrorMessage, "Error", false)
            AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
        }

        logConclusion := logConclusion . "|" . 
            symbolLedger[logErrorMessage . "|E"] ; Arguments or Error Message

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

    if logValuesForConclusion["Overlay Key"] >= 1 {
        OverlayUpdateStatus(logValuesForConclusion, conclusionStatus)
    }

    if IsSet(errorMessage) {
        if OverlayIsVisible() {
            WinSetTransparent(255, "ahk_id " . overlay["GUI"].Hwnd)
        }

        LogEngine("Failed")

        if logValuesForConclusion["Method Name"] = "AbortExecution" {
            ExitApp()
        }

        errorWindow.Show("AutoSize Center")
        WinWaitClose("ahk_id " . errorWindow.Hwnd)
    }
}

LogEngine(status) {
    global system

    static directories := "Directories"
    static environment := "Environment"
    static logging     := "Logging"
    static runtime     := "Runtime"
    static snapshot    := "Snapshot"

    static newLine           := "`r`n"
    static runTelemetryOrder := 0

    runTelemetryOrder := runTelemetryOrder + 1

    runTelemetryLines := []
    if status = "Beginning" {
        system["Logging"] := Map(
            "Log Entries",  [],
            "Log to Array", true
        )

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
        system[directories]["Shared"]    := sharedFolderPath . "\"
        system[directories]["Constants"] := system[directories]["Shared"] . "Constants\"
        system[directories]["Images"]    := system[directories]["Shared"] . "Images\"
        system[directories]["Mappings"]  := system[directories]["Shared"] . "Mappings\"
        system[directories]["Spreadsheet Operations Template"] := system[directories]["Shared"] . "Spreadsheet Operations Template\"

        system[snapshot] := LogTelemetryTimestamp()
        system[snapshot]["Run Telemetry Order"] := runTelemetryOrder

        system[logging]["Log Shared Name"] := system["Directories"]["Log"] . projectName . " - " . FormatTime(StrReplace(StrReplace(StrReplace(StrSplit(system[snapshot]["UTC Timestamp Precise"], ".")[1], "-"), " "), ":"), "yyyy-MM-dd HH.mm.ss")
        system[logging]["Log File Path"]   := Map(
            "Execution Log", system[logging]["Log Shared Name"] . " - Execution Log.csv",
            "Operation Log", system[logging]["Log Shared Name"] . " - Operation Log.csv",
            "Run Telemetry", system[logging]["Log Shared Name"] . " - Run Telemetry.csv",
            "Symbol Ledger", system[logging]["Log Shared Name"] . " - Symbol Ledger.csv"
        )

        system[environment]["International"] := GetInternationalFormatting()
        EnsureDirectoryExists(system["Directories"]["Log"])
        EnsureDirectoryExists(system["Directories"]["Project"])
        
        system[environment]["Operating System"]     := GetOperatingSystem()
        system[environment]["OS Installation Date"] := GetWindowsInstallationDateUtcTimestamp()
        system[environment]["Computer Name"]        := A_ComputerName
        system[environment]["Computer Identifier"]  := GetTextHash(GetComputerIdentifier(), "SHA-256")
        system[environment]["Username"]             := A_UserName
        system[environment]["Time Zone Key Name"]   := GetTimeZoneKeyName()
        system[environment]["Region Format"]        := GetRegionFormat()
        system[environment]["Input Language"]       := GetInputLanguage()
        system[environment]["Keyboard Layout"]      := GetActiveKeyboardLayout()
        system[environment]["Motherboard"]          := GetMotherboard()
        system[environment]["CPU"]                  := GetCpu()
        system[environment]["Memory Size and Type"] := GetMemorySizeAndType()
        system[environment]["System Disk"]          := GetSystemDisk()
        system[environment]["Display GPU"]          := GetActiveDisplayGpu()
        system[environment]["Monitor"]              := GetActiveMonitor()
        system[environment]["BIOS"]                 := GetBios()
        system[environment]["QPC Frequency"]        := GetQueryPerformanceCounterFrequency()
        system[environment]["Display Resolution"]   := A_ScreenWidth . "x" . A_ScreenHeight
        system[environment]["Refresh Rate"]         := GetActiveMonitorRefreshRateHz()
        system[environment]["DPI Scale"]            := Round(A_ScreenDPI / 96 * 100) . "%"
        system[environment]["Color Mode"]           := GetWindowsColorMode()

        DefineApplicationRegistry()

        configurationPath := system["Directories"]["Project"] . "Configuration (" . system[runtime]["Project Name"] . ", " . "Library Release" . " " . system[runtime]["Library Release"] . ").json"
        if !FileExist(configurationPath) {
            configurationData := '{' . newLine . 
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
                    '        "Image Variant Preset": "' . system["Directories"]["Constants"] . 'Heroes (2025-09-20).csv' . '",' . newLine . 
                    '        "Application Image Override Directory": "' . '"' . newline . 
                '    }' . newLine . '}'
            configurationData := StrReplace(configurationData, "\", "\\")
            WriteTextToFile(configurationData, configurationPath, "UTF-8", "Create")
        }

        ValidateConfiguration(configurationPath)

        newLine := "`r`n"
        WriteTextToFile("Log" . newLine, system[logging]["Log File Path"]["Execution Log"], "UTF-8-BOM", "Create")
        WriteTextToFile("Operation Sequence Number|Status|Query Performance Counter|UTC Timestamp Integer|Method or Context|Arguments or Error Message|Overlay Key|Overlay Value" . newLine, system[logging]["Log File Path"]["Operation Log"], "UTF-8-BOM", "Create")
        WriteTextToFile("Log" . newLine, system[logging]["Log File Path"]["Run Telemetry"], "UTF-8-BOM", "Create")
        WriteTextToFile("Reference|Type|Symbol" . newLine, system[logging]["Log File Path"]["Symbol Ledger"], "UTF-8-BOM", "Create")

        consolidatedOperationLog := ""
        consolidatedSymbolLedger := ""
        for logEntry in system["Logging"]["Log Entries"] {
            if logEntry[2] = "Operation Log" {
                if consolidatedOperationLog = "" {
                    consolidatedOperationLog := logEntry[1]
                } else {
                    consolidatedOperationLog := consolidatedOperationLog . newLine . logEntry[1]
                }
            }

            if logEntry[2] = "Symbol Ledger" {
                if consolidatedSymbolLedger = "" {
                    consolidatedSymbolLedger := logEntry[1]
                } else {
                    consolidatedSymbolLedger := consolidatedSymbolLedger . newLine . logEntry[1]
                }
            }
        }

        AppendLineToLog(consolidatedOperationLog, "Operation Log")
        AppendLineToLog(consolidatedSymbolLedger, "Symbol Ledger")

        system[logging].Delete("Log to Array")
        system[logging].Delete("Log Entries")

        executionLogLines := [
            system[runtime]["Project Name"],
            system[runtime]["Library Release"],
            system[runtime]["AutoHotkey Version"],
            system[environment]["Operating System"],
            system[environment]["OS Installation Date"],
            system[environment]["Computer Name"],
            system[environment]["Computer Identifier"],
            system[environment]["Username"],
            system[environment]["Time Zone Key Name"],
            system[environment]["Region Format"],
            system[environment]["Input Language"],
            system[environment]["Keyboard Layout"],
            system[environment]["Motherboard"],
            system[environment]["CPU"],
            system[environment]["Memory Size and Type"],
            system[environment]["System Disk"],
            system[environment]["Display GPU"],
            system[environment]["Monitor"],
            system[environment]["BIOS"],
            system[environment]["QPC Frequency"],
            system[environment]["Display Resolution"],
            system[environment]["Refresh Rate"],
            system[environment]["DPI Scale"],
            system[environment]["Color Mode"]
        ]

        BatchAppendExecutionLog("Beginning", executionLogLines)

        runTelemetryLines.Push(system[snapshot]["Run Telemetry Order"] . "|" . system[snapshot]["Number of Readings"] . "|" . system[snapshot]["Operation Log Line Number"] . 
            "|" . system[snapshot]["UTC Timestamp Precise"] . "|" . system[snapshot]["UTC Timestamp Integer"] . "|" . system[snapshot]["QPC Midpoint Tick"])
        runTelemetryLines.Push(GetPhysicalMemoryStatus())
        runTelemetryLines.Push(GetRemainingFreeDiskSpace())

        BatchAppendRunTelemetry("Beginning", runTelemetryLines)
    } else {
        system[snapshot] := LogTelemetryTimestamp()
        system[snapshot]["Run Telemetry Order"] := runTelemetryOrder

        runTelemetryLines.Push(system[snapshot]["Run Telemetry Order"] . "|" . system[snapshot]["Number of Readings"] . "|" . system[snapshot]["Operation Log Line Number"] . 
            "|" . system[snapshot]["UTC Timestamp Precise"] . "|" . system[snapshot]["UTC Timestamp Integer"] . "|" . system[snapshot]["QPC Midpoint Tick"])
        runTelemetryLines.Push(GetPhysicalMemoryStatus())
        runTelemetryLines.Push(GetRemainingFreeDiskSpace())

        if status = "Completed" {
            if OverlayIsVisible() {
                OverlayChangeTransparency(255)
            }
        }

        BatchAppendRunTelemetry(status, runTelemetryLines)
    }

    if status = "Completed" || status = "Failed" {
        timestampNow := A_Now

        for logType, filePath in system["Logging"]["Log File Path"] {
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

            FileSetTime(timestampNow, system["Logging"]["Log File Path"][logType], "M")
        }

        system["Logging"]["Log Entries"]  := []
        system["Logging"]["Log to Array"] := true
        system["Logging"].Delete("Log File Path")
        system["Logging"].Delete("Log Shared Name")
    }
}

LogFormatMethodArguments(logValuesForConclusion, arguments) {
    global symbolLedger

    methodName := logValuesForConclusion["Method Name"]

    if methodName = "OverlayInsertSpacer" || methodName = "OverlayUpdateCustomLine" {
        logValuesForConclusion["Arguments Full"] := ""
        logValuesForConclusion["Arguments Log"]  := ""

        return logValuesForConclusion
    }

    argumentsFormatted := Map(
        "Arguments Full",  "",
        "Arguments Log",   "",
        "Parameter",       "",
        "Argument",        "",
        "Data Type",       "",
        "Data Constraint", ""
    )

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
                argumentValueFull := "<" . argumentsFormatted["Data Type"] . ">"
                argumentValueLog  := "<" . argumentsFormatted["Data Type"] . ">"
            case "Boolean":
            case "Integer":
            case "String":
                if argumentsFormatted["Whitelist"].Length != 0 {
                    if !symbolLedger.Has(argument . "|W") {
                        logSymbolLedgerLine := RegisterSymbol(argument, "Whitelist", false)
                        AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                    }

                    argumentValueLog := symbolLedger[argument . "|W"]
                } else {
                    switch argumentsFormatted["Data Constraint"] {
                        case "Path", "Valid Path":
                            SplitPath(argument, &filename, &directoryPath)

                            if !symbolLedger.Has(directoryPath . "|D") {
                                logSymbolLedgerLine := RegisterSymbol(directoryPath, "Directory", false)
                                AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                            }

                            if !symbolLedger.Has(filename . "|F") {
                                logSymbolLedgerLine := RegisterSymbol(filename, "Filename", false)
                                AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[directoryPath . "|D"] . "\" . symbolLedger[filename . "|F"]
                        case "Base64", "Summary":
                            summary := "<Length: " . StrLen(argument) . ", Rows: " . StrSplit(argument, "`n").Length . ">"

                            if !symbolLedger.Has(summary . "|S") {
                                logSymbolLedgerLine := RegisterSymbol(summary, "Summary", false)
                                AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueFull := summary
                            argumentValueLog := symbolLedger[summary . "|S"]
                        case "Directory", "Valid Directory":
                            if !symbolLedger.Has(RTrim(argument, "\") . "|D") {
                                logSymbolLedgerLine := RegisterSymbol(argument, "Directory", false)
                                AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[RTrim(argument, "\") . "|D"]
                        case "Filename":
                            if !symbolLedger.Has(argument . "|F") {
                                logSymbolLedgerLine := RegisterSymbol(argument, "Filename", false)
                                AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[argument . "|F"]
                        case "Key":
                            argumentValueLog  := "<Key>"
                            argumentValueFull := "<Key>"
                        case "Locator":
                            if !symbolLedger.Has(argument . "|R") {
                                logSymbolLedgerLine := RegisterSymbol(argument, "Reference", false)
                                AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[argument . "|R"]
                        case "SHA-256":
                            encodedHash := EncodeSha256HexToBase(argument, 86)
                            if !symbolLedger.Has(encodedHash . "|H") {
                                logSymbolLedgerLine := RegisterSymbol(encodedHash, "Hash", false)
                                AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[encodedHash . "|H"]
                        default:
                            if StrLen(argument) > 192 {
                                argumentValueFull := SubStr(argument, 1, 224) . "…"
                                argumentValueLog  := SubStr(argument, 1, 192) . "…"
                            }
                    }
                }
                
                argumentValueFull := Format('"{1}"', argumentValueFull)
                argumentValueLog  := Format('"{1}"', argumentValueLog)
            case "Variant":
                if Type(argument) = "String" {
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

    logValuesForConclusion["Arguments Full"] := argumentsFormatted["Arguments Full"]
    logValuesForConclusion["Arguments Log"]  := argumentsFormatted["Arguments Log"]

    return logValuesForConclusion
}

LogTelemetryTimestamp() {
    global system

    static maxDurationMilliseconds := 240

    snapshot                        := Map()
    queryPerformanceCounterReadings := []
    utcTimestampPreciseReadings     := []
    combinedReadings                := []

    originalAutoHotkeyThreadHandle   := DllCall("GetCurrentThread", "Ptr")
    originalAutoHotkeyThreadPriority := DllCall("GetThreadPriority", "Ptr", originalAutoHotkeyThreadHandle, "Int")

    if originalAutoHotkeyThreadPriority = 0 {
        DllCall("SetThreadPriority", "Ptr", originalAutoHotkeyThreadHandle, "Int", 2) ; Change to Highest.
    }

    if !IsSet(logFilePath) {
        while true {
            milliseconds := A_MSec + 0

            if milliseconds <= 584 {
                break
            } else {
                Sleep(1016 - milliseconds)
                continue
            }
        }
    }

    startTime := A_TickCount
    while A_TickCount - startTime < maxDurationMilliseconds {
        queryPerformanceCounterReadings.Push(GetQueryPerformanceCounter())
        utcTimestampPreciseReadings.Push(GetUtcTimestampPrecise())
    }

    if originalAutoHotkeyThreadPriority = 0 {
        DllCall("SetThreadPriority", "Ptr", originalAutoHotkeyThreadHandle, "Int", 0) ; Change to Normal.
    }

    utcTimestampPreciseReadingsLength := utcTimestampPreciseReadings.Length
    for index, utcTimestampPrecise in utcTimestampPreciseReadings {
        if index = utcTimestampPreciseReadingsLength {
            continue
        }

        qpcBefore := queryPerformanceCounterReadings[index]
        qpcAfter  := queryPerformanceCounterReadings[index + 1]
        combinedReadings.Push([qpcBefore, utcTimestampPrecise, qpcAfter, qpcAfter - qpcBefore])
    }

    snapshot["Number of Readings"] := combinedReadings.Length

    bestIndex    := 1
    bestDuration := combinedReadings[1][4]

    for index, combinedReading in combinedReadings {
        if index = 1 {
            continue
        }

        duration := combinedReading[4]
        if duration < bestDuration {
            bestDuration := duration
            bestIndex    := index
        }
    }

    if IsSet(logFilePath) {
        snapshot["Operation Log Line Number"] := GetLastLineNumberFromTextFile(logFilePath["Operation Log"])
    } else {
        snapshot["Operation Log Line Number"] := 2
    }

    chosenReading                     := combinedReadings[bestIndex]
    qpcBeforeTimestamp                := chosenReading[1]
    snapshot["UTC Timestamp Precise"] := chosenReading[2]

    utcTimestampIntegerConversion := StrReplace(StrReplace(StrReplace(StrReplace(snapshot["UTC Timestamp Precise"], "-"), " "), ":"), ".")
    if StrLen(utcTimestampIntegerConversion) >= 18 {
        utcTimestampIntegerConversion := SubStr(utcTimestampIntegerConversion, 1, 17)
    }
    utcTimestampIntegerConversion := utcTimestampIntegerConversion + 0

    snapshot["UTC Timestamp Integer"] := utcTimestampIntegerConversion

    qpcMeasurementDelta               := chosenReading[4]
    snapshot["QPC Midpoint Tick"]     := qpcBeforeTimestamp + (qpcMeasurementDelta // 2)

    return snapshot
}

LogTimestamp() {
    queryPerformanceCounterBefore           := GetQueryPerformanceCounter()
    utcTimestampInteger                     := GetUtcTimestampInteger()
    queryPerformanceCounterAfter            := GetQueryPerformanceCounter()

    utcTimestampInteger                     := utcTimestampInteger - system["Snapshot"]["UTC Timestamp Integer"]
    queryPerformanceCounterMeasurementDelta := queryPerformanceCounterAfter - queryPerformanceCounterBefore
    queryPerformanceCounterMidpointTick     := (queryPerformanceCounterBefore + (queryPerformanceCounterMeasurementDelta // 2)) - system["Snapshot"]["QPC Midpoint Tick"]

    logTimestamp := Map(
        "QPC Midpoint Tick",     queryPerformanceCounterMidpointTick,
        "UTC Timestamp Integer", utcTimestampInteger
    )

    return logTimestamp
}

LogValidateMethodArguments(methodName, arguments) {
    validation := ""

    if methodName = "OverlayInsertSpacer" {
        return validation
    }

    for index, argument in arguments {
        parameterName  := methodRegistry[methodName]["Parameter Contracts"][index]["Parameter Name"]
        dataType       := methodRegistry[methodName]["Parameter Contracts"][index]["Data Type"]
        dataConstraint := methodRegistry[methodName]["Parameter Contracts"][index]["Data Constraint"]
        optional       := methodRegistry[methodName]["Parameter Contracts"][index]["Optional"]
        whitelist      := methodRegistry[methodName]["Parameter Contracts"][index]["Whitelist"]

        parameterMissingValue := 'Parameter "' . parameterName . '" has no value passed into it.'

        if dataType = "String" && optional = "" && argument = "" {
            validation := parameterMissingValue
        } else if dataType = "String" && optional != "" && argument = "" {
            ; Skip validation as it's Optional.
        } else {
            validation := ValidateDataUsingSpecification(argument, dataType, dataConstraint, whitelist)
        }

        if validation = parameterMissingValue {
            break
        } else if validation != "" {
            validation := 'Parameter "' . parameterName . '" failed validation. ' . validation
            break
        }
    }

    return validation
}

NextOperationSequenceNumber() {
    static operationSequenceNumber := -1

    operationSequenceNumber++

    return operationSequenceNumber
}

NextSymbolLedgerAlias() {
    static symbolLedgerIdentifier := -1

    symbolLedgerIdentifier++
    symbolLedgerAlias := EncodeIntegerToBase(symbolLedgerIdentifier, 86)

    return symbolLedgerAlias
}

OverlayGenerateNextKey(methodName := unset) {
    static counter := 1

    if !IsSet(methodName) {
        return counter++
    }

    if Type(methodName) != "String" {
        return 0
    }

    if !methodRegistry.Has(methodName) {
        return 0
    }

    if methodRegistry[methodName]["Overlay Log"] {
        return counter++
    } else {
        return 0
    }
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

OverlayUpdateStatus(logValuesForConclusion, newStatus) {
    global overlay

    overlaykey := logValuesForConclusion["Overlay Key"]

    currentText := overlay["Lines"][overlayKey]

    if logValuesForConclusion["Method Name"] !== "OverlayInsertSpacer" && logValuesForConclusion["Method Name"] !== "OverlayUpdateCustomLine" {
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

RegisterSymbol(value, type, addNewLine := true) {
    global symbolLedger

    static newLine := "`r`n"
    symbolLine     := ""

    switch StrLower(type) {
        case "context", "c":
            type := "C"
        case "directory", "d":
            type := "D"
            value := RTrim(value, "\")
        case "error", "e":
            type := "E"
        case "filename", "f":
            type := "F"
        case "hash", "h":
            type := "H"
        case "method", "m":
            type := "M"
        case "overlay", "o":
            type := "O"
        case "reference", "r":
            type := "R"
        case "summary", "s":
            type := "S"
        case "whitelist", "w":
            type := "W"
    }

    if !symbolLedger.Has(value . "|" . type) {
        symbolLedger[value . "|" . type] := NextSymbolLedgerAlias()

        symbolLine :=
            value . "|" . 
            type . "|" . 
            symbolLedger[value . "|" . type]
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
    ; https://www.utf8-chartable.de
    static cachedBase52Result := unset
    static cachedBase66Result := unset
    static cachedBase86Result := unset
    static cachedBase94Result := unset

    cachedResult := unset

    switch baseType {
        case 52:
            if IsSet(cachedBase52Result) {
                cachedResult := cachedBase52Result
            }
        case 66:
            if IsSet(cachedBase66Result) {
                cachedResult := cachedBase66Result
            }
        case 86:
            if IsSet(cachedBase86Result) {
                cachedResult := cachedBase86Result
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

       if baseType <= 86 {
            excludedAsciiCodePoints[0x22] := true ; " QUOTATION MARK
            excludedAsciiCodePoints[0x2A] := true ; * ASTERISK
            excludedAsciiCodePoints[0x2F] := true ; / SOLIDUS
            excludedAsciiCodePoints[0x3A] := true ; : COLON
            excludedAsciiCodePoints[0x3C] := true ; < LESS-THAN SIGN
            excludedAsciiCodePoints[0x3E] := true ; > GREATER-THAN SIGN
            excludedAsciiCodePoints[0x3F] := true ; ? QUESTION MARK
            excludedAsciiCodePoints[0x5C] := true ; \ REVERSE SOLIDUS
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

       if baseType <= 52 {
            excludedAsciiCodePoints[0x2D] := true ; - HYPHEN-MINUS
            excludedAsciiCodePoints[0x2E] := true ; . FULL STOP
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
            excludedAsciiCodePoints[0x5F] := true ; _ LOW LINE
            excludedAsciiCodePoints[0x7E] := true ; ~ TILDE
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
            case 52:
                cachedBase52Result := cachedResult
            case 66:
                cachedBase66Result := cachedResult
            case 86:
                cachedBase86Result := cachedResult
            case 94:
                cachedBase94Result := cachedResult
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
    static executionTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}"', "Application", "A", "Beginning", "B")
    static methodName := RegisterMethod("executionType As String [Whitelist: " . executionTypeWhitelist . "], array as Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [executionType, array])

    static newLine := "`r`n"

    switch StrLower(executionType) {
        case "application", "a":
            executionType := "A"
        case "beginning", "b":
            executionType := "B"
    }

    consolidatedExecutionLog := ""

    arrayLength := array.Length
    for index, value in array {
        if value = "" {
            continue
        }

        if arrayLength !== index {
            consolidatedExecutionLog := consolidatedExecutionLog . value . "|" . executionType . newLine
        } else {
            consolidatedExecutionLog := consolidatedExecutionLog . value . "|" . executionType
        }
    }

    AppendLineToLog(consolidatedExecutionLog, "Execution Log")
}

BatchAppendRunTelemetry(appendType, array) {
    static appendTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}"', "Beginning", "B", "Completed", "C", "Failed", "F", "Intermission", "I")
    static methodName := RegisterMethod("appendType As String [Whitelist: " . appendTypeWhitelist . "], array as Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [appendType, array])

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

    consolidatedRunTelemetry := ""

    arrayLength := array.Length
    for index, value in array {
        if value = "" {
            continue
        }

        if arrayLength !== index {
            consolidatedRunTelemetry := consolidatedRunTelemetry . value . "|" . appendType . newLine
        } else {
            consolidatedRunTelemetry := consolidatedRunTelemetry . value . "|" . appendType
        }
    }

    AppendLineToLog(consolidatedRunTelemetry, "Run Telemetry")
}

BatchAppendSymbolLedger(symbolType, array) {
    static symbolTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}", "{9}", "{10}", "{11}", "{12}", "{13}", "{14}", "{15}", "{16}", "{17}", "{18}"',
        "Context", "C", "Directory", "D", "Error", "E", "File", "F", "Hash", "H", "Overlay", "O", "Reference", "R", "Summary", "S", "Whitelist", "W")
    static methodName := RegisterMethod("symbolType As String [Whitelist: " . symbolTypeWhitelist . "], array As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [symbolType, array])

    static newLine := "`r`n"

    switch StrLower(symbolType) {
        case "context", "c":
            symbolType := "C"
        case "directory", "d":
            symbolType := "D"
        case "error", "e":
            symbolType := "E"
        case "file", "f":
            symbolType := "F"
        case "hash", "h":
            symbolType := "H"
        case "overlay", "o":
            symbolType := "O"
        case "reference", "r":
            symbolType := "R"
        case "summary", "s":
            symbolType := "S"
        case "whitelist", "w":
            symbolType := "W"
    }

    consolidatedSymbolLedger := ""

    symbolLedgerArray := []
    for value in array {
        if value = "" {
            continue
        }

        if symbolType = "D" {
            value := RTrim(value, "\")
        } else if symbolType = "H" {
            value := EncodeSha256HexToBase(value, 86)
        } else if symbolType = "S" {
            value := "<Length: " . StrLen(value) . ", Rows: " . StrSplit(value, "`n").Length . ">"
        }

        if !symbolLedger.Has(value . "|" . symbolType) {
            symbolLedgerArray.Push(value)
        }
    }
  
    arrayLength := symbolLedgerArray.Length
    if arrayLength !== 0 {
        symbolLedgerArray := RemoveDuplicatesFromArray(symbolLedgerArray)
        arrayLength := symbolLedgerArray.Length

        for index, value in symbolLedgerArray {
            if arrayLength !== index {
                consolidatedSymbolLedger := consolidatedSymbolLedger . RegisterSymbol(value, symbolType)
            } else {
                consolidatedSymbolLedger := consolidatedSymbolLedger . RegisterSymbol(value, symbolType, false)
            }
        }

        AppendLineToLog(consolidatedSymbolLedger, "Symbol Ledger")
    }
}

GetPhysicalMemoryStatus() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    pointerSizeInBytes := A_PtrSize
    structureSizeInBytes := 4 + (pointerSizeInBytes = 8 ? 4 : 0) + (10 * pointerSizeInBytes) + (3 * 4)
    if pointerSizeInBytes = 8 {
        structureSizeInBytes := (structureSizeInBytes + 7) & ~7
    }

    static performanceInformationBuffer := Buffer(structureSizeInBytes, 0)

    NumPut("UInt", structureSizeInBytes, performanceInformationBuffer, 0)

    getPerformanceInfoSucceeded := DllCall("Psapi\GetPerformanceInfo", "Ptr", performanceInformationBuffer.Ptr, "UInt", structureSizeInBytes, "Int")

    if !getPerformanceInfoSucceeded {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to retrieve performance values for memory. [Psapi\GetPerformanceInfo" . ", System Error Code: " . A_LastError . "]")
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

    bytesPerGibiByte := 1 << 30
    bytesPerMebiByte := 1 << 20

    totalPhysicalBytes     := physicalTotalPages * pageSizeInBytes
    usedPhysicalBytes      := (physicalTotalPages - physicalAvailablePages) * pageSizeInBytes
    totalPhysicalGigabytes := totalPhysicalBytes / bytesPerGibiByte
    usedPhysicalGigabytes  := usedPhysicalBytes  / bytesPerGibiByte

    usedPhysicalPercentHundredths := (totalPhysicalBytes > 0) ? ((usedPhysicalBytes * 10000 + (totalPhysicalBytes // 2)) // totalPhysicalBytes) : 0
    usedPhysicalPercentPrecise := usedPhysicalPercentHundredths / 100.0

    commitUsedGigabytes  := (commitTotalPages * pageSizeInBytes) / bytesPerGibiByte
    commitLimitGigabytes := (commitLimitPages * pageSizeInBytes) / bytesPerGibiByte

    commitUsedPercentHundredths := (commitLimitPages > 0) ? ((commitTotalPages * 10000 + (commitLimitPages // 2)) // commitLimitPages) : 0
    commitUsedPercentPrecise    := commitUsedPercentHundredths / 100.0

    kernelPagedMebiBytes    := (kernelPagedPages    * pageSizeInBytes) / bytesPerMebiByte
    kernelNonpagedMebiBytes := (kernelNonpagedPages * pageSizeInBytes) / bytesPerMebiByte
    systemCacheMebiBytes    := (systemCachePages    * pageSizeInBytes) / bytesPerMebiByte

    separator := " - "
    resultText := "Physical Used " . Format("{:.2f}", usedPhysicalGigabytes) . "/" . Format("{:.2f}", totalPhysicalGigabytes) . " GB" . " (" . Format("{:.2f}", usedPhysicalPercentPrecise) . "%)" . separator . "Commit " . Format("{:.2f}", commitUsedGigabytes)
        . "/" . Format("{:.2f}", commitLimitGigabytes) . " GB" . " (" . Format("{:.2f}", commitUsedPercentPrecise) . "%)" . separator . "Handles " . systemHandleCount . separator . "Processes " . systemProcessCount . separator . "Threads " . systemThreadCount
        . separator . "Kernel Paged " . Format("{:.2f}", kernelPagedMebiBytes) . " MB" . separator . "Kernel Nonpaged " . Format("{:.2f}", kernelNonpagedMebiBytes) . " MB" . separator . "System Cache " . Format("{:.2f}", systemCacheMebiBytes) . " MB"

    return resultText
}

GetRemainingFreeDiskSpace() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    systemDrive := SubStr(A_WinDir, 1, 3)

    freeBytesAvailableToCaller := 0
    totalNumberOfBytes := 0
    totalNumberOfFreeBytes := 0

    getDiskFreeSpaceSucceeded := DllCall("Kernel32\GetDiskFreeSpaceExW", "Str", systemDrive, "Int64*", &freeBytesAvailableToCaller, "Int64*", &totalNumberOfBytes, "Int64*", &totalNumberOfFreeBytes, "Int")

    bytesPerGibiByte := 1 << 30
    bytesPerTebiByte := 1 << 40
    resultText := ""

    if !getDiskFreeSpaceSucceeded {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to retrieve information about the amount of space that is available on system disk volume. [Kernel32\GetDiskFreeSpaceExW" . ", System Error Code: " . A_LastError . "]")
    }

    freeGibiBytes  := freeBytesAvailableToCaller / bytesPerGibiByte
    freeTebiBytes  := freeBytesAvailableToCaller / bytesPerTebiByte
    totalGibiBytes := totalNumberOfBytes / bytesPerGibiByte

    windowsReportedFreeGibiBytes := Round(freeGibiBytes, 1)

    usedPercentHundredths := (totalNumberOfBytes > 0) ? (((totalNumberOfBytes - freeBytesAvailableToCaller) * 10000 + (totalNumberOfBytes // 2)) // totalNumberOfBytes) : 0
    usedPercentPrecise := usedPercentHundredths / 100.0

    static windowsFormattedFreeSizeBuffer := Buffer(64 * 2, 0)
    DllCall("Shlwapi\StrFormatByteSizeW", "Int64", freeBytesAvailableToCaller, "Ptr", windowsFormattedFreeSizeBuffer.Ptr, "Int", 64, "Ptr")
    windowsFormattedFreeSize := StrGet(windowsFormattedFreeSizeBuffer, "UTF-16")

    separator  := " - "
    resultText := "Free Windows " . Format("{:.1f}", windowsReportedFreeGibiBytes) . " GB (" . windowsFormattedFreeSize . ")" . separator . "Free Detailed " . Format("{:.2f}", freeGibiBytes)
        . " GiB / " . Format("{:.2f}", freeTebiBytes) . " TiB" . separator . "Used " . Format("{:.2f}", usedPercentPrecise) . "% of " . Format("{:.2f}", totalGibiBytes) . " GB"

    return resultText
}

OverlayIsVisible() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName)

    windowHandle  := overlay["GUI"].Hwnd
    windowVisible := unset

    if DllCall("User32\IsWindowVisible", "Ptr", windowHandle) {
        windowVisible := true
    } else {
        windowVisible := false
    }

    return windowVisible
}