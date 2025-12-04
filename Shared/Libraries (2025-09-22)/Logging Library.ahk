#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include Image Library.ahk

global logFilePath := Map(
    "Execution Log", "",
    "Operation Log", "",
    "Runtime Trace", "",
    "Symbol Ledger", ""
)
global methodRegistry := Map()
global overlayGui := ""
global overlayLines := Map()
global overlayOrder := []
global overlayStatus := Map(
    "Beginning", "... Beginning ▶️",
    "Skipped",   "... Skipped ➡️",
    "Completed", "... Completed ✔️",
    "Failed",    "... Failed ✖️"
)
global symbolLedger := Map()
global system := Map()

; Press Escape to abort the script early when running or to close the script when it's completed.
$Esc:: {
    if logFilePath["Execution Log"] != "" && logFilePath["Operation Log"] != "" && logFilePath["Runtime Trace"] != "" && logFilePath["Symbol Ledger"] != "" {
        Critical "On"
        AbortExecution()
    } else {
        ExitApp()
    }
}

AbortExecution() {
    static methodName := RegisterMethod("AbortExecution()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Abort Execution", methodName)

    try {
        throw Error("Execution aborted early by pressing escape.")
    } catch as executionAbortedError {
        LogInformationConclusion("Failed", logValuesForConclusion, executionAbortedError)
    }
}

DisplayErrorMessage(logValuesForConclusion, errorObject, customLineNumber := unset) {
    windowTitle := "AutoHotkey v" . system["AutoHotkey Version"] . ": " . A_ScriptName
    currentDateTime := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    newLine := "`r`n"

    errorMessage := (errorObject.HasOwnProp("Message") ? errorObject.Message : errorObject)

    lineNumber := unset
    if IsSet(customLineNumber) {
        lineNumber := customLineNumber
    } else if logValuesForConclusion["Validation"] != "" {
        lineNumber := methodRegistry[logValuesForConclusion["Method Name"]]["Validation Line"]
    } else {
        lineNumber := errorObject.Line
    }
    
    declaration := RegExReplace(methodRegistry[logValuesForConclusion["Method Name"]]["Declaration"], " <\d+>$", "")

    fullErrorText := unset
    if methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] != "" {
        fullErrorText :=
            "Declaration: " .  declaration . " (" . system["Library Release"] . ")" . newLine . 
            "Parameters: " .   methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] . newLine . 
            "Arguments: " .    logValuesForConclusion["Arguments Full"] . newLine . 
            "Line Number: " .  lineNumber . newLine . 
            "Date Runtime: " . currentDateTime . newLine . 
            "Error Output: " . errorMessage
    } else {
        fullErrorText :=
            "Declaration: " .  declaration . " (" . system["Library Release"] . ")" . newLine . 
            "Line Number: " .  lineNumber . newLine . 
            "Date Runtime: " . currentDateTime . newLine . 
            "Error Output: " . errorMessage
    }

    LogEngine("Failed", fullErrorText)

    if OverlayIsVisible() {
        WinSetTransparent(255, "ahk_id " . overlayGui.Hwnd)
    }

    if logValuesForConclusion["Method Name"] !== "AbortExecution" {
        errorWindow := Gui("-Resize +AlwaysOnTop +OwnDialogs", windowTitle)
        errorWindow.SetFont("s10", "Segoe UI")
        errorWindow.AddEdit("ReadOnly r10 w1024 -VScroll vErrorTextField", fullErrorText)

        exitButton := errorWindow.AddButton("w60 Default", "Exit")
        exitButton.OnEvent("Click", (*) => ExitApp())
        exitButton.Focus()
        errorWindow.OnEvent("Close", (*) => ExitApp())

        copyButton := errorWindow.AddButton("x+10 yp wp", "Copy")
        copyButton.OnEvent("Click", (*) => A_Clipboard := fullErrorText)

        errorWindow.Show("AutoSize Center")
        WinWaitClose("ahk_id " errorWindow.Hwnd)
    } else {
        ExitApp()
    }
}

LogEngine(status, fullErrorText := "") {
    global logFilePath
    global system

    runtimeTraceLines := []
    if status = "Beginning" {
        SplitPath(A_ScriptFullPath, , , , &projectName)
        SplitPath(A_LineFile, , &librariesFolderPath)
        SplitPath(librariesFolderPath, , &sharedFolderPath, , &librariesVersion)
        SplitPath(sharedFolderPath, , &curatiumFolderPath)

        system["Project Name"]          := projectName
        system["Curatium Directory"]    := curatiumFolderPath . "\"
        system["Log Directory"]         := system["Curatium Directory"] . "Log\"
        system["Project Directory"]     := system["Curatium Directory"] . "Projects\" . RTrim(SubStr(projectName, 1, InStr(projectName, "(") - 1)) . "\"
        system["Shared Directory"]      := sharedFolderPath . "\"
        system["Constants Directory"]   := system["Shared Directory"] . "Constants\"
        system["Images Directory"]      := system["Shared Directory"] . "Images\"
        system["Mappings Directory"]    := system["Shared Directory"] . "Mappings\"
        system["Library Release"]       := SubStr(librariesVersion, InStr(librariesVersion, "(") + 1, InStr(librariesVersion, ")") - InStr(librariesVersion, "(") - 1)
        system["AutoHotkey Version"]    := A_AhkVersion

        if !DirExist(system["Log Directory"]) {
            DirCreate(system["Log Directory"])
        }

        while true {
            milliseconds    := A_MSec + 0
            dateTimeOfToday := FormatTime(A_NowUTC, "yyyy-MM-dd HH.mm.ss")
            sharedStartName := system["Log Directory"] . projectName . " - " . dateTimeOfToday . " - "

            logFilePath["Execution Log"] := sharedStartName . "Execution Log.csv"
            logFilePath["Operation Log"] := sharedStartName . "Operation Log.csv"
            logFilePath["Runtime Trace"] := sharedStartName . "Runtime Trace.csv"
            logFilePath["Symbol Ledger"] := sharedStartName . "Symbol Ledger.csv"

            if !FileExist(logFilePath["Execution Log"]) && !FileExist(logFilePath["Operation Log"]) && !FileExist(logFilePath["Runtime Trace"]) && !FileExist(logFilePath["Symbol Ledger"]) {
                if milliseconds >= 400 {
                    Sleep(1016 - milliseconds)
                    continue
                }

                break
            } else {
                Sleep(1016 - milliseconds)
            }
        }

        executionLogFileHandle := FileOpen(logFilePath["Execution Log"], "w", "UTF-8")
        executionLogFileHandle.WriteLine("Log")
        executionLogFileHandle.Close()

        operationLogFileHandle := FileOpen(logFilePath["Operation Log"], "w", "UTF-8")
        operationLogFileHandle.WriteLine("Operation Sequence Number|Status|Query Performance Counter|UTC Timestamp Integer|Method or Context|Arguments|Overlay Key|Overlay Value")
        operationLogFileHandle.Close()

        runtimeTraceFileHandle := FileOpen(logFilePath["Runtime Trace"], "w", "UTF-8")
        runtimeTraceFileHandle.WriteLine("Log")
        runtimeTraceFileHandle.Close()

        symbolLedgerFileHandle := FileOpen(logFilePath["Symbol Ledger"], "w", "UTF-8")
        symbolLedgerFileHandle.WriteLine("Reference|Type|Symbol")
        symbolLedgerFileHandle.Close()

        warmupUtcTimestamp              := GetUtcTimestamp()
        warmupUtcTimestampInteger       := GetUtcTimestampInteger()

        LogTimestampPrecise()

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

        runtimeTraceLines.Push(system["Runtime Trace Order"])
        runtimeTraceLines.Push(system["QPC Before Timestamp"])
        runtimeTraceLines.Push(system["UTC Timestamp Precise"])
        runtimeTraceLines.Push(system["QPC After Timestamp"])
        runtimeTraceLines.Push(system["QPC Midpoint Tick"])
        runtimeTraceLines.Push(system["UTC Timestamp Integer"])
        runtimeTraceLines.Push(GetPhysicalMemoryStatus())
        runtimeTraceLines.Push(GetRemainingFreeDiskSpace())

        BatchAppendRuntimeTrace("Beginning", runtimeTraceLines)
    } else {
        LogTimestampPrecise()

        runtimeTraceLines.Push(system["Runtime Trace Order"])
        runtimeTraceLines.Push(system["QPC Before Timestamp"])
        runtimeTraceLines.Push(system["UTC Timestamp Precise"])
        runtimeTraceLines.Push(system["QPC After Timestamp"])
        runtimeTraceLines.Push(system["QPC Midpoint Tick"])
        runtimeTraceLines.Push(system["UTC Timestamp Integer"])
        runtimeTraceLines.Push(GetPhysicalMemoryStatus())
        runtimeTraceLines.Push(GetRemainingFreeDiskSpace())
    }

    switch status {
        case "Completed":
            if OverlayIsVisible() {
                OverlayChangeTransparency(255)
            }

            BatchAppendRuntimeTrace("Completed", runtimeTraceLines)

            FinalizeLogs()
        case "Failed":
            BatchAppendRuntimeTrace("Failed", runtimeTraceLines)
            AppendCsvLineToLog(fullErrorText, "Runtime Trace")
            
            FinalizeLogs()
        case "Intermission":
            BatchAppendRuntimeTrace("Intermission", runtimeTraceLines)
    }
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

LogFormatMethodArguments(methodName, arguments, validation := "") {
    global symbolLedger

    argumentValueFull := ""
    argumentValueLog  := ""

    argumentsAndValidation := Map(
        "Arguments Full",  "",
        "Arguments Log",   "",
        "Parameter",       "",
        "Argument",        "",
        "Data Type",       "",
        "Data Constraint", "",
        "Validation",      validation
    )

    argumentsAndValidation["Arguments Full"] := ""
    argumentsAndValidation["Arguments Log"]  := ""

    if methodName = "OverlayInsertSpacer" {
        return argumentsAndValidation
    }

    for index, argument in arguments {
        argumentsAndValidation["Parameter"]       := methodRegistry[methodName]["Parameter Contracts"][index]["Parameter Name"]
        argumentsAndValidation["Argument"]        := argument
        argumentsAndValidation["Data Type"]       := methodRegistry[methodName]["Parameter Contracts"][index]["Data Type"]
        argumentsAndValidation["Data Constraint"] := methodRegistry[methodName]["Parameter Contracts"][index]["Data Constraint"]
        argumentsAndValidation["Whitelist"]       := methodRegistry[methodName]["Parameter Contracts"][index]["Whitelist"]

        argumentValueFull := argument
        argumentValueLog  := argument
        switch argumentsAndValidation["Data Type"] {
            case "Boolean":
            case "Integer":
            case "Object":
                argumentsAndValidation["Arguments Full"] .= "<Object>"
                argumentsAndValidation["Arguments Log"]  .= "<Object>"

                if index < arguments.Length {
                    argumentsAndValidation["Arguments Full"] .= ", "
                    argumentsAndValidation["Arguments Log"]  .= ", "
                }

                continue
            case "String":
                if argumentsAndValidation["Whitelist"].Length != 0 {
                    if !symbolLedger.Has(argument . "|W") {
                        csvSymbolLedgerLine := RegisterSymbol(argument, "Whitelist", false)
                        AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                    }

                    argumentValueLog := symbolLedger[argument . "|W"]
                } else {
                    switch argumentsAndValidation["Data Constraint"] {
                        case "Absolute Path", "Absolute Save Path":
                            SplitPath(argument, &filename, &directoryPath)

                            if !symbolLedger.Has(directoryPath . "|D") {
                                csvSymbolLedgerLine := RegisterSymbol(directoryPath, "Directory", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                            }

                            if !symbolLedger.Has(filename . "|F") {
                                csvSymbolLedgerLine := RegisterSymbol(filename, "Filename", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[directoryPath . "|D"] . "\" . symbolLedger[filename . "|F"]
                        case "Base64":
                            base64Summary := "<Base64 (Length: " . StrLen(argument) . ")>"

                            if !symbolLedger.Has(base64Summary . "|B") {
                                csvSymbolLedgerLine := RegisterSymbol(base64Summary, "Base64", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueFull := base64Summary
                            argumentValueLog := symbolLedger[base64Summary . "|B"]
                        case "Code":
                            codeSummary := "<Code (Length: " . StrLen(argument) . ", Rows: " . StrSplit(argument, "`n").Length . ")>"

                            if !symbolLedger.Has(codeSummary . "|C") {
                                csvSymbolLedgerLine := RegisterSymbol(codeSummary, "Code", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueFull := codeSummary
                            argumentValueLog := symbolLedger[codeSummary . "|C"]
                        case "Directory":
                            if !symbolLedger.Has(RTrim(argument, "\") . "|D") {
                                csvSymbolLedgerLine := RegisterSymbol(argument, "Directory", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[RTrim(argument, "\") . "|D"]
                        case "Filename":
                            if !symbolLedger.Has(argument . "|F") {
                                csvSymbolLedgerLine := RegisterSymbol(argument, "Filename", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[argument . "|F"]
                        case "Locator":
                            if !symbolLedger.Has(argument . "|L") {
                                csvSymbolLedgerLine := RegisterSymbol(argument, "Locator", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[argument . "|L"]
                        case "Secret":
                            if !symbolLedger.Has("<Secret>" . "|S") {
                                csvSymbolLedgerLine := RegisterSymbol("<Secret>", "Secret", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog  := symbolLedger["<Secret>" . "|S"]
                            argumentValueFull := symbolLedger["<Secret>" . "|S"]
                        case "SHA-256":
                            encodedHash := EncodeSha256HexToBase(argument, 86)
                            if !symbolLedger.Has(encodedHash . "|H") {
                                csvSymbolLedgerLine := RegisterSymbol(encodedHash, "Hash", false)
                                AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
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
        }

        argumentsAndValidation["Arguments Full"] .= argumentValueFull
        argumentsAndValidation["Arguments Log"]  .= argumentValueLog

        if index < arguments.Length {
            argumentsAndValidation["Arguments Full"] .= ", "
            argumentsAndValidation["Arguments Log"]  .= ", "
        }
    }

    return argumentsAndValidation
}

LogHelperError(logValuesForConclusion, errorLineNumber, errorMessage) {
    timestamp := LogTimestamp()

    encodedOperationSequenceNumber := logValuesForConclusion["Operation Sequence Number"]
    encodedQueryPerformanceCounter := EncodeIntegerToBase(timestamp["Query Performance Counter Midpoint Tick"], 94)
    encodedUtcTimestampInteger     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)

    operationSequenceNumber        := NextOperationSequenceNumber()
    encodedOperationSequenceNumber := EncodeIntegerToBase(operationSequenceNumber, 86)
    logValuesForConclusion["Operation Sequence Number"] := encodedOperationSequenceNumber

    csvConclusion := 
        encodedOperationSequenceNumber . "|" . ; Operation Sequence Number
        "F" .                            "|" . ; Status
        encodedQueryPerformanceCounter . "|" . ; Query Performance Counter
        encodedUtcTimestampInteger             ; UTC Timestamp Integer

    if logValuesForConclusion["Context"] != "" {
        csvConclusion := csvConclusion . "|" . 
            logValuesForConclusion["Context"]  ; Context
    }

    AppendCsvLineToLog(csvConclusion, "Operation Log")

    try {
        throw Error(errorMessage)
    } catch as customError {
        DisplayErrorMessage(logValuesForConclusion, customError, errorLineNumber)
    }
}

LogHelperValidation(methodName, arguments := unset) {
    logValuesForConclusion := Map(
        "Operation Sequence Number", 0,
        "Method Name",               methodName,
        "Arguments Full",            "",
        "Arguments Log",             "",
        "Overlay Key",               0,
        "Validation",                "",
        "Context",                   ""
    )

    argumentsAndValidationStatus := unset
    if IsSet(arguments) {
        logValuesForConclusion["Validation"] := LogValidateMethodArguments(methodName, arguments)
        argumentsAndValidationStatus := LogFormatMethodArguments(methodName, arguments, logValuesForConclusion["Validation"])

        logValuesForConclusion["Arguments Full"] := argumentsAndValidationStatus["Arguments Full"]
        logValuesForConclusion["Arguments Log"]  := argumentsAndValidationStatus["Arguments Log"]

        if logValuesForConclusion["Validation"] != "" {
            timestamp := LogTimestamp()

            operationSequenceNumber        := NextOperationSequenceNumber()
            encodedOperationSequenceNumber := EncodeIntegerToBase(operationSequenceNumber, 94)
            encodedQueryPerformanceCounter := EncodeIntegerToBase(timestamp["Query Performance Counter Midpoint Tick"], 94)
            encodedUtcTimestampInteger     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)

            logValuesForConclusion["Operation Sequence Number"] := encodedOperationSequenceNumber
            
            csvShared :=
                encodedOperationSequenceNumber .       "|" . ; Operation Sequence Number
                "B" .                                  "|" . ; Status
                encodedQueryPerformanceCounter .       "|" . ; Query Performance Counter
                encodedUtcTimestampInteger .           "|" . ; UTC Timestamp Integer
                methodRegistry[methodName]["Symbol"]         ; Method

            if logValuesForConclusion["Arguments Full"] != "" {
                csvShared := csvShared . "|" . 
                    logValuesForConclusion["Arguments Log"]  ; Arguments
            }

            AppendCsvLineToLog(csvShared, "Operation Log")

            try {
                throw Error(logValuesForConclusion["Validation"])
            } catch as validationError {
                LogInformationConclusion("Failed", logValuesForConclusion, validationError)
            }
        }
    }

    return logValuesForConclusion
}

LogInformationBeginning(overlayValue, methodName, arguments := unset, overlayCustomKey := 0) {
    static lastRuntimeTraceTick := 0

    timestamp := LogTimestamp()

    runtimeTraceInterval := 6 * 60 * 1000
    runtimeTraceTick := A_TickCount

    if lastRuntimeTraceTick = 0 {
        lastRuntimeTraceTick := runtimeTraceTick
    }

    operationSequenceNumber        := NextOperationSequenceNumber()
    encodedOperationSequenceNumber := EncodeIntegerToBase(operationSequenceNumber, 94)
    encodedQueryPerformanceCounter := EncodeIntegerToBase(timestamp["Query Performance Counter Midpoint Tick"], 94)
    encodedUtcTimestampInteger     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)

    csvShared :=
        encodedOperationSequenceNumber .       "|" . ; Operation Sequence Number
        "B" .                                  "|" . ; Status
        encodedQueryPerformanceCounter .       "|" . ; Query Performance Counter
        encodedUtcTimestampInteger .           "|" . ; UTC Timestamp Integer
        methodRegistry[methodName]["Symbol"]         ; Method

    overlayKey := unset
    if overlayCustomKey = 0 {
        overlayKey := OverlayGenerateNextKey(methodName)
    } else {
        overlayKey := overlayCustomKey
        if methodName = "OverlayInsertSpacer" || methodName = "OverlayUpdateCustomLine" {
            arguments[1] := overlayKey
        }
    }

    logValuesForConclusion := Map(
        "Operation Sequence Number", encodedOperationSequenceNumber,
        "Method Name",               methodName,
        "Arguments Full",            "",
        "Arguments Log",             "",
        "Overlay Key",               overlayKey,
        "Validation",                "",
        "Context",                   ""
    )

    argumentsAndValidationStatus := unset
    if IsSet(arguments) {
        logValuesForConclusion["Validation"]     := LogValidateMethodArguments(methodName, arguments)
        argumentsAndValidationStatus             := LogFormatMethodArguments(methodName, arguments, logValuesForConclusion["Validation"])
        logValuesForConclusion["Arguments Full"] := argumentsAndValidationStatus["Arguments Full"]
        logValuesForConclusion["Arguments Log"]  := argumentsAndValidationStatus["Arguments Log"]

        csvShared := csvShared . "|" . 
            argumentsAndValidationStatus["Arguments Log"] ; Arguments
    }

    if overlayKey !== 0 {
        if !symbolLedger.Has(overlayValue . "|O") {
            csvOverlaySymbolLedgerLine := RegisterSymbol(overlayValue, "Overlay", false)
            AppendCsvLineToLog(csvOverlaySymbolLedgerLine, "Symbol Ledger")
        }

        encodedOverlayKey := EncodeIntegerToBase(overlayKey, 94)

        csvShared := csvShared . "|" . 
            encodedOverlayKey . "|" .                   ; Overlay Key
            symbolLedger[overlayValue . "|O"]           ; Overlay Value

        if overlayCustomKey = 0 {
            OverlayUpdateLine(overlayKey, overlayValue . overlayStatus["Beginning"])
        } else {
            OverlayUpdateLine(overlayKey, overlayValue)
        }
    }

    AppendCsvLineToLog(csvShared, "Operation Log")

    try {
        if IsSet(argumentsAndValidationStatus) {
            logValuesForConclusion["Validation"] := argumentsAndValidationStatus["Validation"]

            if argumentsAndValidationStatus["Validation"] != "" {
                throw Error(argumentsAndValidationStatus["Validation"])
            }
        }
    } catch as validationError {
        LogInformationConclusion("Failed", logValuesForConclusion, validationError)
    }

    if runtimeTraceTick - lastRuntimeTraceTick >= runtimeTraceInterval {
        lastRuntimeTraceTick := runtimeTraceTick
        LogEngine("Intermission")
    }

    return logValuesForConclusion
}

LogInformationConclusion(conclusionStatus, logValuesForConclusion, errorObject := unset) {
    timestamp := LogTimestamp()

    encodedOperationSequenceNumber := logValuesForConclusion["Operation Sequence Number"]
    encodedQueryPerformanceCounter := EncodeIntegerToBase(timestamp["Query Performance Counter Midpoint Tick"], 94)
    encodedUtcTimestampInteger     := EncodeIntegerToBase(timestamp["UTC Timestamp Integer"], 94)

    csvConclusion := 
        encodedOperationSequenceNumber . "|" . ; Operation Sequence Number
        "[[Status]]" .                   "|" . ; Status
        encodedQueryPerformanceCounter . "|" . ; Query Performance Counter
        encodedUtcTimestampInteger             ; UTC Timestamp Integer

    if logValuesForConclusion["Context"] != "" {
        csvConclusion := csvConclusion . "|" . 
            logValuesForConclusion["Context"]  ; Context
    }

    conclusionStatus := StrUpper(SubStr(conclusionStatus, 1, 1)) . StrLower(SubStr(conclusionStatus, 2))
        switch conclusionStatus {
            case "Skipped":
                AppendCsvLineToLog(StrReplace(csvConclusion, "[[Status]]", "S"), "Operation Log")

                if logValuesForConclusion["Overlay Key"] !== 0 {
                    OverlayUpdateStatus(logValuesForConclusion, "Skipped")
                }
            case "Completed":
                AppendCsvLineToLog(StrReplace(csvConclusion, "[[Status]]", "C"), "Operation Log")

                if logValuesForConclusion["Overlay Key"] !== 0 {
                    OverlayUpdateStatus(logValuesForConclusion, "Completed")
                }
            case "Failed":
                AppendCsvLineToLog(StrReplace(csvConclusion, "[[Status]]", "F"), "Operation Log")

                if logValuesForConclusion["Overlay Key"] !== 0 {
                    OverlayUpdateStatus(logValuesForConclusion, "Failed")
                }

                if logValuesForConclusion["Method Name"] = "ValidateApplicationFact" || logValuesForConclusion["Method Name"] = "ValidateApplicationInstalled" {
                    OverlayUpdateLine(overlayOrder.Length, StrReplace(overlayLines[overlayOrder.Length], overlayStatus["Beginning"], overlayStatus["Failed"]))
                }

                DisplayErrorMessage(logValuesForConclusion, errorObject)
            default:
                try {
                    throw Error("Unsupported overlay status: " . conclusionStatus)
                } catch as unsupportedOverlayStatusError {
                    DisplayErrorMessage(logValuesForConclusion, unsupportedOverlayStatusError)
                }
        }
}

LogTimestamp() {
    queryPerformanceCounterBefore           := GetQueryPerformanceCounter()
    utcTimestampInteger                     := GetUtcTimestampInteger()
    queryPerformanceCounterAfter            := GetQueryPerformanceCounter()

    utcTimestampInteger                     := utcTimestampInteger - system["UTC Timestamp Integer"]
    queryPerformanceCounterMeasurementDelta := queryPerformanceCounterAfter - queryPerformanceCounterBefore
    queryPerformanceCounterMidpointTick     := (queryPerformanceCounterBefore + (queryPerformanceCounterMeasurementDelta // 2)) - system["QPC Midpoint Tick"]

    logTimestamp := Map(
        "Query Performance Counter Midpoint Tick", queryPerformanceCounterMidpointTick,
        "UTC Timestamp Integer",                   utcTimestampInteger
    )

    return logTimestamp
}

LogTimestampPrecise() {
    global system

    static runtimeTraceOrder        := 0
    runtimeTraceOrder               := runtimeTraceOrder + 1
    system["Runtime Trace Order"]   := runtimeTraceOrder

    static maxDurationMilliseconds  := 200
    startTime                       := A_TickCount
    queryPerformanceCounterReadings := []
    utcTimestampPreciseReadings     := []
    combinedReadings                := []

    autoHotkeyThreadPriority := GetAutoHotkeyThreadPriority()

    if autoHotkeyThreadPriority = 0 {
        autoHotkeyThreadHandle := DllCall("GetCurrentThread", "Ptr")
        DllCall("SetThreadPriority", "Ptr", autoHotkeyThreadHandle, "Int", 2) ; Change to Highest.
    }

    while (A_TickCount - startTime < maxDurationMilliseconds) {
        queryPerformanceCounterReadings.Push(GetQueryPerformanceCounter())
        utcTimestampPreciseReadings.Push(GetUtcTimestampPrecise())
    }

    if autoHotkeyThreadPriority = 0 {
        autoHotkeyThreadHandle := DllCall("GetCurrentThread", "Ptr")
        DllCall("SetThreadPriority", "Ptr", autoHotkeyThreadHandle, "Int", 0) ; Change to Normal.
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

    chosen := combinedReadings[bestIndex]
    system["QPC Before Timestamp"]  := chosen[1]
    system["UTC Timestamp Precise"] := chosen[2]
    system["QPC After Timestamp"]   := chosen[3]
    system["QPC Measurement Delta"] := chosen[4]
    system["QPC Midpoint Tick"]     := system["QPC Before Timestamp"] + (system["QPC Measurement Delta"] // 2)
    system["UTC Timestamp Integer"] := ConvertUtcTimestampToInteger(system["UTC Timestamp Precise"])

    static methodName := RegisterMethod("LogTimestampPrecise()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Log Timestamp Precise", methodName)

    logValuesForConclusion["Context"] := "Runtime Trace Order: " . system["Runtime Trace Order"]

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayChangeTransparency(transparencyValue) {
    static methodName := RegisterMethod("OverlayChangeTransparency(transparencyValue As Integer [Constraint: Byte])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Change Transparency (" . transparencyValue . ")", methodName, [transparencyValue])

    WinSetTransparent(transparencyValue, "ahk_id " . overlayGui.Hwnd)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayChangeVisibility() {
    static methodName := RegisterMethod("OverlayChangeVisibility()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Change Visibility", methodName)

    if DllCall("User32\IsWindowVisible", "Ptr", overlayGui.Hwnd) {
        overlayGui.Hide()
    } else {
        overlayGui.Show("NoActivate")
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayHideLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("OverlayHideLogForMethod(methodNameInput As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Hide Log for Method (" . methodName . ")", methodName, [methodName])

    global methodRegistry
    
    try {
        if !methodRegistry.Has(methodNameInput) {
            throw Error('Method "' . methodNameInput . '" not registered.')
        }
    } catch as methodNotRegisteredError {
        LogInformationConclusion("Failed", logValuesForConclusion, methodNotRegisteredError)
    }

    methodRegistry[methodNameInput]["Overlay Log"] := false

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayShowLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("OverlayShowLogForMethod(methodNameInput As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Show Log for Method (" . methodNameInput . ")", methodName, [methodNameInput])

    global methodRegistry

    if methodRegistry.Has(methodNameInput) {
        methodRegistry[methodNameInput]["Overlay Log"] := true
    } else {
        methodRegistry[methodNameInput] := Map(
            "Overlay Log", true,
            "Symbol",      ""
        )
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayGenerateNextKey(methodName := "") {
    static counter := 1

    if methodName = "[[Custom]]" {
        return counter++
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

OverlayIsVisible() {
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

OverlayStart(baseLogicalWidth := 960, baseLogicalHeight := 920) {
    global overlayGui

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
    adjustAttemptsForWidth  := 0
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

    overlayGui.Show("x" centeredX " y" centeredY " NoActivate")
    WinSetTransparent(172, overlayGui.Hwnd)
}

OverlayInsertSpacer() {
    static methodName := RegisterMethod("OverlayInsertSpacer()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("", methodName, [""], overlayKey := OverlayGenerateNextKey("[[Custom]]"))
    
    OverlayUpdateLine(overlayKey, "")

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayUpdateCustomLine(overlayKey, value) {
    static methodName := RegisterMethod("OverlayUpdateCustomLine(overlayKey As Integer, value As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning(value, methodName, [""], overlayKey) ; No arguments as log will show in overlay.

    ;  OverlayUpdateCustomLine(overlaySummaryKey := OverlayGenerateNextKey("[[Custom]]"), "Overlay Summary: " . "Project A")
    ; Update to overlay will be made in LogInformationBeginning based on passing in a custom overlay key.

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayUpdateLine(overlayKey, value) {
    global overlayGui
    global overlayLines
    global overlayOrder

    if !overlayLines.Has(overlayKey) {
        overlayOrder.Push(overlayKey)
    }
    overlayLines[overlayKey] := value

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

FinalizeLogs() {
    global logFilePath

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
    FileSetTime(timestampNow, logFilePath["Runtime Trace"], "M")
    FileSetTime(timestampNow, logFilePath["Symbol Ledger"], "M")

    logFilePath["Execution Log"] := ""
    logFilePath["Operation Log"] := ""
    logFilePath["Runtime Trace"] := ""
    logFilePath["Symbol Ledger"] := ""
}

; **************************** ;
; Base Encoding & Decoding     ;
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

AppendCsvLineToLog(csvLine, logType) {
    static newLine := "`r`n"

    callerWasCritical := A_IsCritical
    if !callerWasCritical {
        Critical "On"
    }

    try {
        FileAppend(csvLine . newLine, logFilePath[logType], "UTF-8-RAW")
    } finally {
        if !callerWasCritical {
            Critical "Off"
        }
    }
}

BatchAppendExecutionLog(executionType, array) {
    static executionTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}"', "Application", "A", "Beginning", "B")
    static methodName := RegisterMethod("BatchAppendExecutionLog(executionType As String [Whitelist: " . executionTypeWhitelist . "], array as Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [executionType, array])

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

    AppendCsvLineToLog(consolidatedExecutionLog, "Execution Log")
}

BatchAppendRuntimeTrace(appendType, array) {
    static appendTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}"', "Beginning", "B", "Completed", "C", "Failed", "F", "Intermission", "I")
    static methodName := RegisterMethod("BatchAppendRuntimeTrace(appendType As String [Whitelist: " . appendTypeWhitelist . "], array as Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [appendType, array])

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

    consolidatedRuntimeTrace := ""

    arrayLength := array.Length
    for index, value in array {
        if value = "" {
            continue
        }

        if arrayLength !== index {
            consolidatedRuntimeTrace := consolidatedRuntimeTrace . value . "|" . appendType . newLine
        } else {
            consolidatedRuntimeTrace := consolidatedRuntimeTrace . value . "|" . appendType
        }
    }

    AppendCsvLineToLog(consolidatedRuntimeTrace, "Runtime Trace")
}

BatchAppendSymbolLedger(symbolType, array) {
    static symbolTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}", "{8}", "{9}", "{10}", "{11}", "{12}", "{13}", "{14}", "{15}", "{16}"',
        "Base64", "B", "Code", "C", "Directory", "D", "File", "F", "Hash", "H", "Locator", "L", "Overlay", "O", "Whitelist", "W")
    static methodName := RegisterMethod("BatchAppendSymbolLedger(symbolType As String [Whitelist: " . symbolTypeWhitelist . "], array As Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [symbolType, array])

    static newLine := "`r`n"

    switch StrLower(symbolType) {
        case "base64", "b":
            symbolType := "B"
        case "code", "c":
            symbolType := "C"
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
        case "whitelist", "w":
            symbolType := "W"
    }

    consolidatedSymbolLedger := ""

    symbolLedgerArray := []
    for index, value in array {
        if value = "" {
            continue
        }

        if symbolType = "B" {
            value := "<Base64 (Length: " . StrLen(value) . ")>"
        } else if symbolType = "C" {
            value := "<Code (Length: " . StrLen(value) . ", Rows: " . StrSplit(value, "`n").Length . ")>"
        } else if symbolType = "D" {
            value := RTrim(value, "\")
        } else if symbolType = "H" {
            value := EncodeSha256HexToBase(value, 86)
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

        AppendCsvLineToLog(consolidatedSymbolLedger, "Symbol Ledger")
    }
}

GetPhysicalMemoryStatus() {
    static methodName := RegisterMethod("GetPhysicalMemoryStatus()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    pointerSizeInBytes := A_PtrSize
    structureSizeInBytes := 4 + (pointerSizeInBytes = 8 ? 4 : 0) + (10 * pointerSizeInBytes) + (3 * 4)
    if pointerSizeInBytes = 8 {
        structureSizeInBytes := (structureSizeInBytes + 7) & ~7
    }

    static performanceInformationBuffer := Buffer(structureSizeInBytes, 0)

    NumPut("UInt", structureSizeInBytes, performanceInformationBuffer, 0)

    getPerformanceInfoSucceeded := DllCall("Psapi\GetPerformanceInfo", "Ptr", performanceInformationBuffer.Ptr, "UInt", structureSizeInBytes, "Int")

    if !getPerformanceInfoSucceeded {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve performance values for memory. [Psapi\GetPerformanceInfo" . ", System Error Code: " . A_LastError . "]")
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
    static methodName := RegisterMethod("GetRemainingFreeDiskSpace()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    systemDrive := SubStr(A_WinDir, 1, 3)

    freeBytesAvailableToCaller := 0
    totalNumberOfBytes := 0
    totalNumberOfFreeBytes := 0

    getDiskFreeSpaceSucceeded := DllCall("Kernel32\GetDiskFreeSpaceExW", "Str", systemDrive, "Int64*", &freeBytesAvailableToCaller, "Int64*", &totalNumberOfBytes, "Int64*", &totalNumberOfFreeBytes, "Int")

    bytesPerGibiByte := 1 << 30
    bytesPerTebiByte := 1 << 40
    resultText := ""

    if !getDiskFreeSpaceSucceeded {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve information about the amount of space that is available on system disk volume. [Kernel32\GetDiskFreeSpaceExW" . ", System Error Code: " . A_LastError . "]")
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

ParseMethodDeclaration(declaration) {
    atParts     := StrSplit(declaration, "@", , 2)
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

    parsedMethod := Map(
        "Declaration",         declaration,
        "Signature",           signature,
        "Library",             library,
        "Contract",            contract,
        "Parameters",          parameters,
        "Data Types",          dataTypes,
        "Metadata",            metadata,
        "Validation Line",     lineNumberForValidation,
        "Parameter Contracts", parameterContracts,
        "Overlay Log",         "",
        "Symbol",              ""
    )

    return parsedMethod
}

RegisterMethod(declaration, sourceFilePath := "", validationLineNumber := 0) {
    global methodRegistry

    if sourceFilePath != "" && validationLineNumber !== 0 {
        SplitPath(sourceFilePath, , , , &filenameWithoutExtension)
        libraryTag := " @ " . filenameWithoutExtension
        validationLineNumber := " " . "<" . validationLineNumber . ">"
        declaration := declaration . libraryTag . validationLineNumber
    }

    methodName := SubStr(declaration, 1, InStr(declaration, "(") - 1)

    symbol := unset
    if !symbolLedger.Has(declaration . "|" . "M") {
        csvSymbolLedgerLine := RegisterSymbol(declaration, "M", false)
        AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")

        csvParts := StrSplit(csvSymbolLedgerLine, "|")
        symbol   := csvParts[csvParts.Length]
    } else {
        symbol   := symbolLedger[declaration . "|" . "M"]
    }

    parsedMethod := ParseMethodDeclaration(declaration)
    if methodRegistry.Has(methodName) {
        if methodRegistry[methodName]["Symbol"] = "" {
            methodRegistry[methodName]["Declaration"]         := parsedMethod["Declaration"]
            methodRegistry[methodName]["Signature"]           := parsedMethod["Signature"]
            methodRegistry[methodName]["Library"]             := parsedMethod["Library"]
            methodRegistry[methodName]["Contract"]            := parsedMethod["Contract"]
            methodRegistry[methodName]["Parameters"]          := parsedMethod["Parameters"]
            methodRegistry[methodName]["Data Types"]          := parsedMethod["Data Types"]
            methodRegistry[methodName]["Metadata"]            := parsedMethod["Metadata"]
            methodRegistry[methodName]["Validation Line"]     := parsedMethod["Validation Line"]
            methodRegistry[methodName]["Parameter Contracts"] := parsedMethod["Parameter Contracts"]
            methodRegistry[methodName]["Symbol"]              := symbol
        }
    } else {      
        methodRegistry[methodName] := Map(
            "Declaration",         parsedMethod["Declaration"],
            "Signature",           parsedMethod["Signature"],
            "Library",             parsedMethod["Library"],
            "Contract",            parsedMethod["Contract"],
            "Parameters",          parsedMethod["Parameters"],
            "Data Types",          parsedMethod["Data Types"],
            "Metadata",            parsedMethod["Metadata"],
            "Validation Line",     parsedMethod["Validation Line"],
            "Parameter Contracts", parsedMethod["Parameter Contracts"],
            "Overlay Log",         false,
            "Symbol",              symbol
        )
    }

    return methodName
}

RegisterSymbol(value, type, addNewLine := true) {
    global symbolLedger

    static newLine := "`r`n"
    symbolLine     := ""

    switch StrLower(type) {
        case "base64", "b":
            type := "B"
        case "code", "c":
            type := "C"
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
        case "secret", "s":
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
; Helper Methods: System       ;
; **************************** ;

GetInternationalFormatting() {
    static methodName := RegisterMethod("GetInternationalFormatting()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetOperatingSystem()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetWindowsInstallationDateUtcTimestamp()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
        LogHelperError(logValuesForConclusion, A_LineNumber, "InstallDate (DWORD) not found in SYSTEM\Setup snapshots or CurrentVersion.")
    }

    installDateSeconds := installDateSeconds & 0xFFFFFFFF
    if installDateSeconds <= 0 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Invalid InstallDate seconds: " . installDateSeconds)
    }

    utcTimestamp := ConvertUnixTimeToUtcTimestamp(installDateSeconds)

    return utcTimestamp
}

GetComputerIdentifier() {
    static methodName := RegisterMethod("GetComputerIdentifier()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetTimeZoneKeyName()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve time zone information.")
    }

    return timeZoneKeyName
}

GetRegionFormat() {
    static methodName := RegisterMethod("GetRegionFormat()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetInputLanguage()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetActiveKeyboardLayout()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetMotherboard()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetCpu()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetMemorySizeAndType()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve the amount of RAM that is physically installed on the computer. [Kernel32\GetPhysicallyInstalledSystemMemory" . ", System Error Code: " . A_LastError . "]")
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
    static methodName := RegisterMethod("GetSystemDisk()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetActiveDisplayGpu()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetActiveMonitor()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetBios()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetQueryPerformanceCounterFrequency()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    queryPerformanceCounterFrequencyBuffer := Buffer(8, 0)
    queryPerformanceCounterFrequencyRetrievedSuccessfully := DllCall("QueryPerformanceFrequency", "Ptr", queryPerformanceCounterFrequencyBuffer.Ptr, "Int")
    if !queryPerformanceCounterFrequencyRetrievedSuccessfully {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve the frequency of the performance counter. [QueryPerformanceFrequency" . ", System Error Code: " . A_LastError . "]")
    }

    queryPerformanceCounterFrequency := NumGet(queryPerformanceCounterFrequencyBuffer, 0, "Int64")

    return queryPerformanceCounterFrequency
}

GetActiveMonitorRefreshRateHz() {
    static methodName := RegisterMethod("GetActiveMonitorRefreshRateHz()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetWindowsColorMode()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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