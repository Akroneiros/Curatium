#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Application Library.ahk
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include Image Library.ahk

global logEntries    := []
global logFilePath   := unset
global overlayGui    := unset
global overlayLines  := Map()
global overlayOrder  := []
global overlayStatus := Map(
    "Beginning", "... Beginning " . "▶️",
    "Skipped",   "... Skipped " .   "➡️",
    "Completed", "... Completed " . "✔️",
    "Failed",    "... Failed " .    "✖️"
)
global symbolLedger  := Map()
global system        := Map()

; Press Escape to abort the script early when running or to close the script when it's completed.
$Esc:: {
    if IsSet(logFilePath) {
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

    WinSetTransparent(transparencyValue, "ahk_id " . overlayGui.Hwnd)

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayChangeVisibility() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Overlay Change Visibility")

    if DllCall("User32\IsWindowVisible", "Ptr", overlayGui.Hwnd) {
        overlayGui.Hide()
    } else {
        overlayGui.Show("NoActivate")
    }

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayHideLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("methodNameInput As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [methodNameInput], "Overlay Hide Log for Method (" . methodNameInput . ")")

    global methodRegistry
    
    if !methodRegistry.Has(methodNameInput) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, 'Method "' . methodNameInput . '" not registered.')
    }

    methodRegistry[methodNameInput]["Overlay Log"] := false

    LogConclusion("Completed", logValuesForConclusion)
}

OverlayShowLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("methodNameInput As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
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

    global overlayGui

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Base Logical Width", 960, false)
        SetMethodSetting(methodName, "Base Logical Height", 920, false)
        SetMethodSetting(methodName, "Overlay Transparency", 172, false)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]
    baseLogicalWidth    := settings.Get("Base Logical Width")
    baseLogicalHeight   := settings.Get("Base Logical Height")
    overlayTransparency := settings.Get("Overlay Transparency")

    overlayGui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x80000 +DPIScale")
    overlayGui.BackColor := "0x000000"
    overlayGui.SetFont("s11 cWhite", "Consolas")
    overlayGui.MarginX := 0
    overlayGui.MarginY := 0

    statusTextControl := overlayGui.Add("Text", "vStatusText w" baseLogicalWidth " h" baseLogicalHeight " +0x1", "")

    measureVisualRectangle := () => (
        overlayGui.Show("Hide AutoSize"),
        rectBuffer := Buffer(16, 0),
        DllCall("Dwmapi\DwmGetWindowAttribute", "Ptr", overlayGui.Hwnd, "Int", 9, "Ptr", rectBuffer, "Int", 16),
        Map(
            "left",   NumGet(rectBuffer,  0, "Int"),
            "top",    NumGet(rectBuffer,  4, "Int"),
            "right",  NumGet(rectBuffer,  8, "Int"),
            "bottom", NumGet(rectBuffer, 12, "Int")
        )
    )

    visualRectangle := measureVisualRectangle()
    visualWidth  := visualRectangle["right"]  - visualRectangle["left"]
    visualHeight := visualRectangle["bottom"] - visualRectangle["top"]

    ; Ensure the *visual* size is even on both axes. If an axis is odd, nudge the client by +1 logical pixel on that axis and re-measure.
    adjustAttemptsForWidth := 0
    while Mod(visualWidth, 2) && adjustAttemptsForWidth < 6 {
        baseLogicalWidth += 1
        statusTextControl.Move(, , baseLogicalWidth, baseLogicalHeight)
        visualRectangle := measureVisualRectangle()
        visualWidth  := visualRectangle["right"]  - visualRectangle["left"]
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

    overlayGui.Show("x" . centeredX . " y" . centeredY . " NoActivate")
    WinSetTransparent(overlayTransparency, overlayGui.Hwnd)

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

    if IsSet(logFilePath) {
        callerWasCritical := A_IsCritical
        if !callerWasCritical {
            Critical "On"
        }

        try {
            FileAppend(line . newLine, logFilePath[logType], "UTF-8-RAW")
        } finally {
            if !callerWasCritical {
                Critical "Off"
            }
        }
    } else {
        logEntries.Push([line, logType])
    }
}

LogBeginning(methodName, arguments := [], overlayValue := unset) {
    static lastRunTelemetryTick := 0

    logValuesForConclusion := Map(
        "Method Name",    methodName,
        "Arguments Full", "",
        "Arguments Log",  "",
        "Validation",     "",
        "Context",        ""
    )

    timestamp := LogTimestamp()

    runTelemetryInterval := 6 * 60 * 1000
    runTelemetryTick     := A_TickCount

    if lastRunTelemetryTick = 0 {
        lastRunTelemetryTick := runTelemetryTick
    }

    logBeginning := unset
    overlayKey   := unset
    if IsSet(overlayValue) {
        encodedOperationSequenceNumber := EncodeIntegerToBase(NextOperationSequenceNumber(), 94)
        encodedQueryPerformanceCounter := EncodeIntegerToBase(timestamp["QPC Midpoint Tick"], 94)
        encodedUtcTimestampInteger     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)

        if methodName != "OverlayInsertSpacer" && methodName != "OverlayUpdateCustomLine" {
            overlayKey := OverlayGenerateNextKey(methodName)

            if overlayKey != 0 {
                OverlayUpdateLine(overlayKey, overlayValue . overlayStatus["Beginning"])
            }
        } else {
            if methodName = "OverlayInsertSpacer" {
                OverlayUpdateLine(overlayKey := OverlayGenerateNextKey(), overlayValue := "")
            } else if methodName = "OverlayUpdateCustomLine" {
                OverlayUpdateLine(overlayKey := arguments[1], overlayValue := arguments[2])
            }
        }

        logValuesForConclusion["Operation Sequence Number"] := encodedOperationSequenceNumber

        logBeginning :=
            encodedOperationSequenceNumber .       "|" . ; Operation Sequence Number
            "B" .                                  "|" . ; Status
            encodedQueryPerformanceCounter .       "|" . ; Query Performance Counter
            encodedUtcTimestampInteger .           "|" . ; UTC Timestamp Integer
            methodRegistry[methodName]["Symbol"]         ; Method or Context
    } else {
        overlayKey := 0
    }

    logValuesForConclusion["Overlay Key"] := overlayKey

    if arguments.Length != 0 {
        logValuesForConclusion["Validation"] := LogValidateMethodArguments(methodName, arguments)
        logValuesForConclusion               := LogFormatMethodArguments(logValuesForConclusion, arguments)

        if IsSet(overlayValue) {
            logBeginning := logBeginning . "|" . 
                logValuesForConclusion["Arguments Log"]  ; Arguments
        }

        if !IsSet(overlayValue) && logValuesForConclusion["Validation"] != "" {
            encodedOperationSequenceNumber := EncodeIntegerToBase(NextOperationSequenceNumber(), 94)
            encodedQueryPerformanceCounter := EncodeIntegerToBase(timestamp["QPC Midpoint Tick"], 94)
            encodedUtcTimestampInteger     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)

            logValuesForConclusion["Operation Sequence Number"] := encodedOperationSequenceNumber

            logBeginning :=
                encodedOperationSequenceNumber .       "|" . ; Operation Sequence Number
                "B" .                                  "|" . ; Status
                encodedQueryPerformanceCounter .       "|" . ; Query Performance Counter
                encodedUtcTimestampInteger .           "|" . ; UTC Timestamp Integer
                methodRegistry[methodName]["Symbol"] . "|" . ; Method or Context
                logValuesForConclusion["Arguments Log"]      ; Arguments
        }
    }

    if logValuesForConclusion["Overlay Key"] != 0 {
        if !symbolLedger.Has(overlayValue . "|O") {
            logOverlaySymbolLedgerLine := RegisterSymbol(overlayValue, "Overlay", false)
            AppendLineToLog(logOverlaySymbolLedgerLine, "Symbol Ledger")
        }

        if IsSet(overlayValue) {
            encodedOverlayKey := EncodeIntegerToBase(overlayKey, 94)

            logBeginning := logBeginning . "|" . 
                encodedOverlayKey . "|" .                ; Overlay Key
                symbolLedger[overlayValue . "|O"]        ; Overlay Value
        }
    }

    if IsSet(overlayValue) || logValuesForConclusion["Validation"] != "" {
        AppendLineToLog(logBeginning, "Operation Log")

        if logValuesForConclusion["Validation"] != "" {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, logValuesForConclusion["Validation"])
        }
    } else {
        logValuesForConclusion["QPC Midpoint Tick"]     := timestamp["QPC Midpoint Tick"]
        logValuesForConclusion["UTC Timestamp Integer"] := timestamp["UTC Timestamp Integer"]
    }

    if runTelemetryTick - lastRunTelemetryTick >= runTelemetryInterval {
        lastRunTelemetryTick := runTelemetryTick
        LogEngine("Intermission")
    }

    return logValuesForConclusion
}

LogConclusion(conclusionStatus, logValuesForConclusion, errorLineNumber := unset, errorMessage := unset) {
    timestamp := LogTimestamp()

    if !logValuesForConclusion.Has("Operation Sequence Number") {
        logValuesForConclusion["Operation Sequence Number"] := EncodeIntegerToBase(NextOperationSequenceNumber(), 94)
    }

    encodedOperationSequenceNumber := logValuesForConclusion["Operation Sequence Number"]
    encodedQueryPerformanceCounter := EncodeIntegerToBase(timestamp["QPC Midpoint Tick"], 94)
    encodedUtcTimestampInteger     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)

    conclusionStatus := StrUpper(SubStr(conclusionStatus, 1, 1)) . StrLower(SubStr(conclusionStatus, 2))
    status := SubStr(conclusionStatus, 1, 1)

    logConclusion := 
        encodedOperationSequenceNumber . "|" . ; Operation Sequence Number
        status .                         "|" . ; Status
        encodedQueryPerformanceCounter . "|" . ; Query Performance Counter
        encodedUtcTimestampInteger             ; UTC Timestamp Integer

    if logValuesForConclusion["Context"] != "" {
        logConclusion := logConclusion . "|" . 
            logValuesForConclusion["Context"]  ; Method or Context
    }
    
    switch conclusionStatus {
        case "Skipped":
            AppendLineToLog(logConclusion, "Operation Log")

            if logValuesForConclusion["Overlay Key"] !== 0 {
                OverlayUpdateStatus(logValuesForConclusion, "Skipped")
            }
        case "Completed":
            AppendLineToLog(logConclusion, "Operation Log")

            if logValuesForConclusion["Overlay Key"] !== 0 {
                OverlayUpdateStatus(logValuesForConclusion, "Completed")
            }
        case "Failed":
            if logValuesForConclusion.Has("QPC Midpoint Tick") && logValuesForConclusion.Has("UTC Timestamp Integer") {
                logBeginning :=
                    encodedOperationSequenceNumber .       "|" .                                     ; Operation Sequence Number
                    "B" .                                  "|" .                                     ; Status
                    EncodeIntegerToBase(logValuesForConclusion["QPC Midpoint Tick"], 94) . "|" .     ; Query Performance Counter
                    EncodeIntegerToBase(logValuesForConclusion["UTC Timestamp Integer"], 94) . "|" . ; UTC Timestamp Integer
                    methodRegistry[logValuesForConclusion["Method Name"]]["Symbol"]                  ; Method or Context

                    if logValuesForConclusion["Arguments Log"] != "" {
                        logBeginning := logBeginning . "|" . 
                            logValuesForConclusion["Arguments Log"]                                  ; Arguments
                    }

                AppendLineToLog(logBeginning, "Operation Log")
            }

            AppendLineToLog(logConclusion, "Operation Log")

            if logValuesForConclusion["Overlay Key"] != 0 {
                OverlayUpdateStatus(logValuesForConclusion, "Failed")
            }

            if logValuesForConclusion["Method Name"] = "ValidateApplicationFact" || logValuesForConclusion["Method Name"] = "ValidateApplicationInstalled" {
                OverlayUpdateLine(overlayOrder.Length, StrReplace(overlayLines[overlayOrder.Length], overlayStatus["Beginning"], overlayStatus["Failed"]))
            }

            windowTitle := "AutoHotkey v" . system["AutoHotkey Version"] . ": " . A_ScriptName
            currentDateTime := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
            newLine := "`r`n"

            if logValuesForConclusion["Validation"] != "" {
                errorLineNumber := methodRegistry[logValuesForConclusion["Method Name"]]["Validation Line"]
            }
            
            declaration := RegExReplace(methodRegistry[logValuesForConclusion["Method Name"]]["Declaration"], " <\d+>$", "")

            constructedErrorMessage := unset
            if methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] != "" {
                constructedErrorMessage :=
                    "Declaration: " .  declaration . " (" . system["Library Release"] . ")" . newLine . 
                    "Parameters: " .   methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] . newLine . 
                    "Arguments: " .    logValuesForConclusion["Arguments Full"] . newLine . 
                    "Line Number: " .  errorLineNumber . newLine . 
                    "Date Runtime: " . currentDateTime . newLine . 
                    "Error Output: " . errorMessage
            } else {
                constructedErrorMessage :=
                    "Declaration: " .  declaration . " (" . system["Library Release"] . ")" . newLine . 
                    "Line Number: " .  errorLineNumber . newLine . 
                    "Date Runtime: " . currentDateTime . newLine . 
                    "Error Output: " . errorMessage
            }

            LogEngine("Failed", constructedErrorMessage)

            if OverlayIsVisible() {
                WinSetTransparent(255, "ahk_id " . overlayGui.Hwnd)
            }

            if logValuesForConclusion["Method Name"] = "AbortExecution" {
                ExitApp()
            }

            errorWindow := Gui("-Resize +AlwaysOnTop +OwnDialogs", windowTitle)
            errorWindow.SetFont("s10", "Segoe UI")
            errorWindow.AddEdit("ReadOnly r10 w1024 -VScroll vErrorTextField", constructedErrorMessage)

            exitButton := errorWindow.AddButton("w60 Default", "Exit")
            exitButton.OnEvent("Click", (*) => ExitApp())
            exitButton.Focus()
            errorWindow.OnEvent("Close", (*) => ExitApp())

            copyButton := errorWindow.AddButton("x+10 yp wp", "Copy")
            copyButton.OnEvent("Click", (*) => A_Clipboard := constructedErrorMessage)

            errorWindow.Show("AutoSize Center")
            WinWaitClose("ahk_id " . errorWindow.Hwnd)
    }
}

LogEngine(status, constructedErrorMessage := unset) {
    global logEntries
    global logFilePath
    global system

    static newLine := "`r`n"

    runTelemetryLines := []
    if status = "Beginning" {
        SplitPath(A_ScriptFullPath, , , , &projectName)
        SplitPath(A_LineFile, , &librariesFolderPath)
        SplitPath(librariesFolderPath, , &sharedFolderPath, , &librariesVersion)
        SplitPath(sharedFolderPath, , &curatiumFolderPath)

        system["Project Name"]        := projectName
        system["Curatium Directory"]  := curatiumFolderPath . "\"
        system["Log Directory"]       := system["Curatium Directory"] . "Log\"
        system["Project Directory"]   := system["Curatium Directory"] . "Projects\" . RTrim(SubStr(projectName, 1, InStr(projectName, "(") - 1)) . "\"
        system["Shared Directory"]    := sharedFolderPath . "\"
        system["Constants Directory"] := system["Shared Directory"] . "Constants\"
        system["Images Directory"]    := system["Shared Directory"] . "Images\"
        system["Mappings Directory"]  := system["Shared Directory"] . "Mappings\"
        system["Library Release"]     := SubStr(librariesVersion, InStr(librariesVersion, "(") + 1, InStr(librariesVersion, ")") - InStr(librariesVersion, "(") - 1)
        system["AutoHotkey Version"]  := A_AhkVersion

        LogTelemetryTimestamp()

        EnsureDirectoryExists(system["Log Directory"])
        EnsureDirectoryExists(system["Project Directory"])

        system["International"]         := GetInternationalFormatting()
        
        system["Operating System"]      := GetOperatingSystem()
        system["OS Installation Date"]  := GetWindowsInstallationDateUtcTimestamp()
        system["Computer Name"]         := A_ComputerName
        system["Computer Identifier"]   := Hash.String("SHA256", GetComputerIdentifier())
        system["Username"]              := A_UserName
        system["Time Zone Key Name"]    := GetTimeZoneKeyName()
        system["Region Format"]         := GetRegionFormat()
        system["Input Language"]        := GetInputLanguage()
        system["Keyboard Layout"]       := GetActiveKeyboardLayout()
        system["Motherboard"]           := GetMotherboard()
        system["CPU"]                   := GetCpu()
        system["Memory Size and Type"]  := GetMemorySizeAndType()
        system["System Disk"]           := GetSystemDisk()
        system["Display GPU"]           := GetActiveDisplayGpu()
        system["Monitor"]               := GetActiveMonitor()
        system["BIOS"]                  := GetBios()
        system["QPC Frequency"]         := GetQueryPerformanceCounterFrequency()
        system["Display Resolution"]    := A_ScreenWidth . "x" . A_ScreenHeight
        system["Refresh Rate"]          := GetActiveMonitorRefreshRateHz()
        system["DPI Scale"]             := Round(A_ScreenDPI / 96 * 100) . "%"
        system["Color Mode"]            := GetWindowsColorMode()

        DefineApplicationRegistry()

        configurationPath := system["Project Directory"] . "Configuration (" . system["Project Name"] . ", " . "Library Release" . " " . system["Library Release"] . ").json"
        if !FileExist(configurationPath) {
            configurationData := '{' . newLine . 
                '    "Application Whitelist": [' . newLine . 
                    '        ' .  newLine . 
                '    ],' . newLine . 
                '    "Application Executable Directory Candidates": [' . newLine .
                    '        '  . newLine . 
                '    ],' . newLine . 
                '    "Candidate Base Directories": [' . newLine . 
                    '        "' . ExtractDirectory(A_WinDir) . 'Portable Files' . '", ' . newLine . 
                    '        "' . ExtractDirectory(A_WinDir) . 'Program Files (Portable)' . '"' . newLine . 
                    '    ],' . newLine . 
                '    "Settings": {' . newLine . 
                    '        "Image Variant Preset": "' . system["Constants Directory"] . 'Heroes (2025-09-20).csv' . '"' . newLine . 
                '    }' . newLine . '}'
            configurationData := StrReplace(configurationData, "\", "\\")
            WriteTextIntoFile(configurationData, configurationPath)
        }

        ValidateConfiguration(configurationPath)

        dateTimeOfToday := FormatTime(StrReplace(StrReplace(StrReplace(StrSplit(system["UTC Timestamp Precise"], ".")[1], "-"), " "), ":"), "yyyy-MM-dd HH.mm.ss")
        sharedStartName := system["Log Directory"] . projectName . " - " . dateTimeOfToday . " - "

        executionLogFilePath := sharedStartName . "Execution Log.csv"
        operationLogFilePath := sharedStartName . "Operation Log.csv"
        runTelemetryFilePath := sharedStartName . "Run Telemetry.csv"
        symbolLedgerFilePath := sharedStartName . "Symbol Ledger.csv"

        newLine := "`r`n"
        WriteTextIntoFile("Log" . newLine, executionLogFilePath, "UTF-8-BOM", false)
        WriteTextIntoFile("Operation Sequence Number|Status|Query Performance Counter|UTC Timestamp Integer|Method or Context|Arguments|Overlay Key|Overlay Value" . newLine, operationLogFilePath, "UTF-8-BOM", false)
        WriteTextIntoFile("Log" . newLine, runTelemetryFilePath, "UTF-8-BOM", false)
        WriteTextIntoFile("Reference|Type|Symbol" . newLine, symbolLedgerFilePath, "UTF-8-BOM", false)

        logFilePath := Map(
            "Execution Log", executionLogFilePath,
            "Operation Log", operationLogFilePath,
            "Run Telemetry", runTelemetryFilePath,
            "Symbol Ledger", symbolLedgerFilePath
        )

        consolidatedOperationLog := ""
        consolidatedSymbolLedger := ""
        for logEntry in logEntries {
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

        logEntries := []

        executionLogLines := [
            system["Project Name"],
            system["Library Release"],
            system["AutoHotkey Version"],
            system["Operating System"],
            system["OS Installation Date"],
            system["Computer Name"],
            system["Computer Identifier"],
            system["Username"],
            system["Time Zone Key Name"],
            system["Region Format"],
            system["Input Language"],
            system["Keyboard Layout"],
            system["Motherboard"],
            system["CPU"],
            system["Memory Size and Type"],
            system["System Disk"],
            system["Display GPU"],
            system["Monitor"],
            system["BIOS"],
            system["QPC Frequency"],
            system["Display Resolution"],
            system["Refresh Rate"],
            system["DPI Scale"],
            system["Color Mode"]
        ]

        BatchAppendExecutionLog("Beginning", executionLogLines)

        runTelemetryLines.Push(system["Run Telemetry Order"] . "|" . system["Number of Readings"] . "|" . system["Operation Log Line Number"] . "|" . system["UTC Timestamp Precise"] . "|" . system["UTC Timestamp Integer"] . "|" . system["QPC Midpoint Tick"])
        runTelemetryLines.Push(GetPhysicalMemoryStatus())
        runTelemetryLines.Push(GetRemainingFreeDiskSpace())

        BatchAppendRunTelemetry("Beginning", runTelemetryLines)
    } else {
        LogTelemetryTimestamp()

        runTelemetryLines.Push(system["Run Telemetry Order"] . "|" . system["Number of Readings"] . "|" . system["Operation Log Line Number"] . "|" . system["UTC Timestamp Precise"] . "|" . system["UTC Timestamp Integer"] . "|" . system["QPC Midpoint Tick"])
        runTelemetryLines.Push(GetPhysicalMemoryStatus())
        runTelemetryLines.Push(GetRemainingFreeDiskSpace())
    }

    switch status {
        case "Completed":
            if OverlayIsVisible() {
                OverlayChangeTransparency(255)
            }

            BatchAppendRunTelemetry("Completed", runTelemetryLines)
        case "Failed":
            BatchAppendRunTelemetry("Failed", runTelemetryLines)

            AppendLineToLog(constructedErrorMessage, "Run Telemetry")
        case "Intermission":
            BatchAppendRunTelemetry("Intermission", runTelemetryLines)
    }

    switch status {
        case "Completed", "Failed":
            if IsSet(logFilePath) {
                for index, filePath in logFilePath {
                    file := FileOpen(filePath, "rw")
                    if !file {
                        continue
                    }

                    fileSize := file.Length
                    if fileSize = 0 {
                        file.Close()
                        continue
                    }

                    ; Trim last trailing newline.
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

                    filePath := ""
                }

                timestampNow := A_Now
                FileSetTime(timestampNow, logFilePath["Execution Log"], "M")
                FileSetTime(timestampNow, logFilePath["Operation Log"], "M")
                FileSetTime(timestampNow, logFilePath["Run Telemetry"], "M")
                FileSetTime(timestampNow, logFilePath["Symbol Ledger"], "M")

                logFilePath := unset
            }
    }
}

LogFormatMethodArguments(logValuesForConclusion, arguments) {
    global symbolLedger

    methodName := logValuesForConclusion["Method Name"]

    if methodName = "OverlayInsertSpacer" || methodName = "OverlayUpdateCustomLine" {
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
            case "Array", "Object":
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
                        case "Absolute Path", "Absolute Save Path":
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
                        case "Directory":
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
                            if !symbolLedger.Has(argument . "|L") {
                                logSymbolLedgerLine := RegisterSymbol(argument, "Locator", false)
                                AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[argument . "|L"]
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

    static runTelemetryOrder        := 0
    runTelemetryOrder               := runTelemetryOrder + 1
    system["Run Telemetry Order"]   := runTelemetryOrder

    static maxDurationMilliseconds  := 240
    queryPerformanceCounterReadings := []
    utcTimestampPreciseReadings     := []
    combinedReadings                := []

    originalAutoHotkeyThreadHandle   := DllCall("GetCurrentThread", "Ptr")
    originalAutoHotkeyThreadPriority := DllCall("GetThreadPriority", "Ptr", originalAutoHotkeyThreadHandle, "Int")

    if originalAutoHotkeyThreadPriority = 0 {
        DllCall("SetThreadPriority", "Ptr", originalAutoHotkeyThreadHandle, "Int", 2) ; Change to Highest.
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

    startTime := A_TickCount

    while (A_TickCount - startTime < maxDurationMilliseconds) {
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

    system["Number of Readings"] := combinedReadings.Length

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
        system["Operation Log Line Number"] := GetLastLineNumberFromTextFile(logFilePath["Operation Log"])
    } else {
        system["Operation Log Line Number"] := 2
    }

    chosenReading                   := combinedReadings[bestIndex]
    qpcBeforeTimestamp              := chosenReading[1]
    system["UTC Timestamp Precise"] := chosenReading[2]

    utcTimestampIntegerConversion := StrReplace(StrReplace(StrReplace(StrReplace(system["UTC Timestamp Precise"], "-"), " "), ":"), ".")
    if StrLen(utcTimestampIntegerConversion) >= 18 {
        utcTimestampIntegerConversion := SubStr(utcTimestampIntegerConversion, 1, 17)
    }
    utcTimestampIntegerConversion := utcTimestampIntegerConversion + 0

    system["UTC Timestamp Integer"] := utcTimestampIntegerConversion

    qpcMeasurementDelta             := chosenReading[4]
    system["QPC Midpoint Tick"]     := qpcBeforeTimestamp + (qpcMeasurementDelta // 2)
}

LogTimestamp() {
    queryPerformanceCounterBefore           := GetQueryPerformanceCounter()
    utcTimestampInteger                     := GetUtcTimestampInteger()
    queryPerformanceCounterAfter            := GetQueryPerformanceCounter()

    utcTimestampInteger                     := utcTimestampInteger - system["UTC Timestamp Integer"]
    queryPerformanceCounterMeasurementDelta := queryPerformanceCounterAfter - queryPerformanceCounterBefore
    queryPerformanceCounterMidpointTick     := (queryPerformanceCounterBefore + (queryPerformanceCounterMeasurementDelta // 2)) - system["QPC Midpoint Tick"]

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
    global overlayGui
    global overlayLines
    global overlayOrder

    if !overlayLines.Has(overlayKey) {
        overlayOrder.Push(overlayKey)
    }
    overlayLines[overlayKey] := overlayValue

    newText := ""
    for index, lineKey in overlayOrder {
        newText .= (newText != "" ? "`n" : "") . overlayLines[lineKey]
    }
    overlayGui["StatusText"].Text := newText
}

OverlayUpdateStatus(logValuesForConclusion, newStatus) {
    global overlayLines

    overlaykey := logValuesForConclusion["Overlay Key"]

    currentText := overlayLines[overlayKey]

    if logValuesForConclusion["Method Name"] !== "OverlayInsertSpacer" && logValuesForConclusion["Method Name"] !== "OverlayUpdateCustomLine" {
        switch newStatus {
            case "Skipped":
                OverlayUpdateLine(overlayKey, StrReplace(currentText, overlayStatus["Beginning"], overlayStatus["Skipped"]))
            case "Completed":
                OverlayUpdateLine(overlayKey, StrReplace(currentText, overlayStatus["Beginning"], overlayStatus["Completed"]))
            case "Failed":
                OverlayUpdateLine(overlayKey, StrReplace(currentText, overlayStatus["Beginning"], overlayStatus["Failed"]))
        }
    }
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

    parameterContracts := []
    parameterParts := []
    currentParameterText := ""

    if contract != "" {
        squareBracketDepth := 0
        inQuotedString := false
        removeLeadingSpaceAfterComma := false

        loop parse contract {
            currentCharacter := A_LoopField

            ; While inQuotedString = true, commas and brackets are considered literal characters.
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
                        for index, piece in StrSplit(conceptValue, '", "')
                        {
                            cleanedValue := Trim(piece, '" ')
                            if cleanedValue != "" {
                                whitelist.Push(cleanedValue)
                            }
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

    symbol := unset
    if !symbolLedger.Has(methodWithDeclaration . "|" . "M") {
        logSymbolLedgerLine := RegisterSymbol(methodWithDeclaration, "M", false)
        AppendLineToLog(logSymbolLedgerLine, "Symbol Ledger")

        logParts := StrSplit(logSymbolLedgerLine, "|")
        symbol   := logParts[logParts.Length]
    } else {
        symbol   := symbolLedger[methodWithDeclaration . "|" . "M"]
    }

    methodWithDeclarationParsed := ParseMethodWithDeclaration(methodWithDeclaration)
    if methodRegistry.Has(methodName) {
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
            methodRegistry[methodName]["Overlay Log"]     := false
        }

        methodRegistry[methodName]["Symbol"]              := symbol
        
        if !methodRegistry[methodName].Has("Settings") {
            methodRegistry[methodName]["Settings"]        := Map()
        }
    } else {      
        methodRegistry[methodName] := Map(
            "Declaration",         methodWithDeclarationParsed["Declaration"],
            "Signature",           methodWithDeclarationParsed["Signature"],
            "Library",             methodWithDeclarationParsed["Library"],
            "Contract",            methodWithDeclarationParsed["Contract"],
            "Parameters",          methodWithDeclarationParsed["Parameters"],
            "Data Types",          methodWithDeclarationParsed["Data Types"],
            "Metadata",            methodWithDeclarationParsed["Metadata"],
            "Validation Line",     methodWithDeclarationParsed["Validation Line"],
            "Parameter Contracts", methodWithDeclarationParsed["Parameter Contracts"],
            "Overlay Log",         false,
            "Symbol",              symbol,
            "Settings",            Map()
        )
    }

    return methodName
}

RegisterSymbol(value, type, addNewLine := true) {
    global symbolLedger

    static newLine := "`r`n"
    symbolLine     := ""

    switch StrLower(type) {
        case "directory", "d":
            type := "D"
            value := RTrim(value, "\")
        case "filename", "f":
            type := "F"
        case "hash", "h":
            type := "H"
        case "locator", "l":
            type := "L"
        case "method", "m":
            type := "M"
        case "overlay", "o":
            type := "O"
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
    static symbolTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}", "{9}", "{10}", "{11}", "{12}", "{13}", "{14}"',
        "Directory", "D", "File", "F", "Hash", "H", "Locator", "L", "Overlay", "O", "Summary", "S", "Whitelist", "W")
    static methodName := RegisterMethod("symbolType As String [Whitelist: " . symbolTypeWhitelist . "], array As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [symbolType, array])

    static newLine := "`r`n"

    switch StrLower(symbolType) {
        case "directory", "d":
            symbolType := "D"
        case "file", "f":
            symbolType := "F"
        case "hash", "h":
            symbolType := "H"
        case "locator", "l":
            symbolType := "L"
        case "overlay", "o":
            symbolType := "O"
        case "summary", "s":
            symbolType := "S"
        case "whitelist", "w":
            symbolType := "W"
    }

    consolidatedSymbolLedger := ""

    symbolLedgerArray := []
    for index, value in array {
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

    if !IsSet(overlayGui) || !(overlayGui is Gui) {
        return false
    }

    windowHandle := overlayGui.Hwnd

    if !DllCall("User32\IsWindow", "Ptr", windowHandle) {
        return false
    }

    if DllCall("User32\IsWindowVisible", "Ptr", windowHandle) {
        return true
    } else {
        return false
    }
}