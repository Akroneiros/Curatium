#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include Image Library.ahk

global logFilePath := Map(
    "Error Message", "",
    "Execution Log", "",
    "Operation Log", "",
    "Symbol Ledger", ""
)
global methodRegistry := Map()
global overlayGui := ""
global overlayLines := Map()
global overlayOrder := []
global overlayStatus := Map(
    "Beginning", "... Beginning 🔰",
    "Skipped",   "... Skipped ⏭️",
    "Completed", "... Completed ✔️",
    "Failed",    "... Failed 😞"
)
global symbolLedger := Map()
global system := Map()

; Press Escape to abort the script early when running or to close the script when it's completed.
$Esc:: {
    if logFilePath["Error Message"] !== "" && logFilePath["Execution Log"] !== "" && logFilePath["Operation Log"] !== "" && logFilePath["Symbol Ledger"] !== "" {
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

    errorMessage := (errorObject.HasOwnProp("Message") ? errorObject.Message : errorObject)

    lineNumber := unset
    if IsSet(customLineNumber) {
        lineNumber := customLineNumber
    } else if logValuesForConclusion["Validation"] !== "" {
        lineNumber := methodRegistry[logValuesForConclusion["Method Name"]]["Validation Line"]
    } else {
        lineNumber := errorObject.Line
    }
    
    declaration := RegExReplace(methodRegistry[logValuesForConclusion["Method Name"]]["Declaration"], " <\d+>$", "")

    fullErrorText := unset
    if methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] !== "" {
        fullErrorText :=
            "Declaration: " .  declaration . " (" . system["Library Release"] . ")" . "`n" . 
            "Parameters: " .   methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] . "`n" . 
            "Arguments: " .    logValuesForConclusion["Arguments Full"] . "`n" . 
            "Line Number: " .  lineNumber . "`n" . 
            "Date Runtime: " . currentDateTime . "`n" . 
            "Error Output: " . errorMessage
    } else {
        fullErrorText :=
            "Declaration: " .  declaration . " (" . system["Library Release"] . ")" . "`n" . 
            "Line Number: " .  lineNumber . "`n" . 
            "Date Runtime: " . currentDateTime . "`n" . 
            "Error Output: " . errorMessage
    }

    LogEngine("Failed", fullErrorText)

    if OverlayIsVisible() = true {
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
    static lastIntermissionFlushTick := 0
    static intermissionBuffer        := []
    static intermissionFlushInterval := 7200000 ; 120 * 60 * 1000

    if status = "Beginning" {
        global logFilePath

        SplitPath(A_ScriptFullPath, , &projectFolderPath, , &projectName)

        dateOfToday := FormatTime(A_Now, "yyyy-MM-dd")
        tickCount   := A_TickCount
      
        logFilePath["Error Message"] := projectFolderPath . "\Log\" . projectName . " - " . "Error Message" . " - " . dateOfToday . "." . tickCount . ".csv"
        logFilePath["Execution Log"] := projectFolderPath . "\Log\" . projectName . " - " . "Execution Log" . " - " . dateOfToday . "." . tickCount . ".csv"
        logFilePath["Operation Log"] := projectFolderPath . "\Log\" . projectName . " - " . "Operation Log" . " - " . dateOfToday . "." . tickCount . ".csv"
        logFilePath["Symbol Ledger"] := projectFolderPath . "\Log\" . projectName . " - " . "Symbol Ledger" . " - " . dateOfToday . "." . tickCount . ".csv"

        if !DirExist(projectFolderPath . "\Log\") {
            DirCreate(projectFolderPath . "\Log\")
        }

        errorMessageFileHandle := FileOpen(logFilePath["Error Message"], "w", "UTF-8")
        errorMessageFileHandle.Close()

        executionLogFileHandle := FileOpen(logFilePath["Execution Log"], "w", "UTF-8")
        executionLogFileHandle.Close()

        operationLogFileHandle := FileOpen(logFilePath["Operation Log"], "w", "UTF-8")
        operationLogFileHandle.Close()

        symbolLedgerFileHandle := FileOpen(logFilePath["Symbol Ledger"], "w", "UTF-8")
        symbolLedgerFileHandle.Close()

        global system
    
        system["Project Name"]         := projectName
        system["Project Directory"]    := RegExReplace(A_ScriptFullPath, "^(.*)\\([^\\]+?) \(.+\)\.ahk$", "$1\Projects\$2\")
        system["Library Release"]      := (RegExMatch(ExtractDirectory(A_LineFile), "\(([^()]*)\)", &regularExpressionMatch), regularExpressionMatch[1])
        system["AutoHotkey Version"]   := A_AhkVersion
        system["QPC Frequency"]        := GetQueryPerformanceCounterFrequency()

        warmupUtcTimestamp        := GetUtcTimestamp()
        warmupUtcTimestampPrecise := GetUtcTimestampPrecise()
        warmupQpcCounter          := GetQueryPerformanceCounter()

        system["QPC Counter Before"]   := GetQueryPerformanceCounter()
        system["Script Run Timestamp"] := GetUtcTimestampPrecise()
        system["QPC Counter After"]    := GetQueryPerformanceCounter()
        system["QPC Counter Midpoint"] := system["QPC Counter Before"] + ((system["QPC Counter After"] - system["QPC Counter Before"] + 1) // 2)
        system["Script Run Integer"]   := ConvertUtcTimestampToInteger(system["Script Run Timestamp"])
        system["Script File Hash"]     := Hash.File("SHA256", A_ScriptFullPath)
        system["Computer Name"]        := A_ComputerName
        system["Username"]             := A_UserName
        system["Operating System"]     := GetOperatingSystem()
        system["Input Language"]       := GetInputLanguage()
        system["Keyboard Layout"]      := GetActiveKeyboardLayout()
        system["Region Format"]        := GetRegionFormat()
        system["Time Zone Key Name"]   := GetTimeZoneKeyName()
        
        system["Display Resolution"]   := A_ScreenWidth . "x" . A_ScreenHeight
        system["DPI Scale"]            := Round(A_ScreenDPI / 96 * 100) . "%"
        system["Color Mode"]           := GetWindowsColorMode()
        system["Memory Size and Type"] := GetMemorySizeAndType()
        system["Motherboard"]          := GetMotherboard()
        system["CPU"]                  := GetCpu()
        system["Display GPU"]          := GetActiveDisplayGpu()
        system["System Disk"]          := GetSystemDisk()      
    }

    executionLogLines := []

    if status = "Beginning" {
        executionLogLines := [
            system["Project Name"]         ,
            system["Project Directory"]    ,
            system["Script File Hash"]     ,
            system["Library Release"]      ,
            system["AutoHotkey Version"]   ,
            system["Computer Name"]        ,
            system["Username"]             ,
            system["Operating System"]     ,
            system["Input Language"]       ,
            system["Keyboard Layout"]      ,
            system["Region Format"]        ,
            system["Time Zone Key Name"]   ,
            system["QPC Frequency"]        ,
            system["Display Resolution"]   ,
            system["DPI Scale"]            ,
            system["Color Mode"]           ,
            system["Memory Size and Type"] ,
            system["Motherboard"]          ,
            system["CPU"]                  ,
            system["Display GPU"]          ,
            system["System Disk"]
        ]
    } else if status !== "Intermission" {
        executionLogLines := [
            "Remaining Free Disk Space: " .              GetRemainingFreeDiskSpace()
        ]
    } else {
        physicalRamSituation := GetPhysicalMemoryStatus()

        for line in [
            "Physical RAM Situation: " .                 physicalRamSituation
        ] {
            intermissionBuffer.Push(line)
        }
    }

    static operationLogLine := "Operation Sequence Number|Status|Tick|Symbol|Arguments|Overlay Key|Overlay Value"
    static symbolLedgerLine := "Reference|Type|Symbol"

    Switch status {
        Case "Beginning":
            ExecutionLogBatchAppend("Beginning", executionLogLines)
            AppendCsvLineToLog(operationLogLine, "Operation Log")
            AppendCsvLineToLog(symbolLedgerLine, "Symbol Ledger")
        Case "Completed":
            if OverlayIsVisible() = true {
                OverlayChangeTransparency(255)
            }

            if intermissionBuffer.Length !== 0 {
                ExecutionLogBatchAppend("Intermission", intermissionBuffer)
                intermissionBuffer.Length := 0
            }

            ExecutionLogBatchAppend("Completed", executionLogLines)

            FileSetTime(A_Now, logFilePath["Error Message"], "M")

            FinalizeLogs()
        Case "Failed":
            if intermissionBuffer.Length !== 0 {
                ExecutionLogBatchAppend("Intermission", intermissionBuffer)
                intermissionBuffer.Length := 0
            }

            ExecutionLogBatchAppend("Failed", executionLogLines)
            AppendCsvLineToLog(fullErrorText, "Error Message")
            
            FinalizeLogs()
        Case "Intermission":
            currentTick := A_TickCount
            if currentTick - lastIntermissionFlushTick >= intermissionFlushInterval {
                lastIntermissionFlushTick := currentTick
                ExecutionLogBatchAppend("Intermission", intermissionBuffer)
                intermissionBuffer.Length := 0
            }
    }
}

LogValidateMethodArguments(methodName, arguments) {
    if methodName = "OverlayInsertSpacer" {
        return ""
    }

    validation := ""

    for index, argument in arguments {
        parameter := methodRegistry[methodName]["Parameter Contracts"][index]["Parameter Name"]
        dataType  := methodRegistry[methodName]["Parameter Contracts"][index]["Data Type"]
        optional  := methodRegistry[methodName]["Parameter Contracts"][index]["Optional"]
        pattern   := methodRegistry[methodName]["Parameter Contracts"][index]["Pattern"]
        typeValue := methodRegistry[methodName]["Parameter Contracts"][index]["Type"]
        whitelist := methodRegistry[methodName]["Parameter Contracts"][index]["Whitelist"]

        parameterMissingValue := "Parameter " . Chr(34) . parameter . Chr(34) . " has no value passed into it."

        switch dataType {
            case "Boolean":
                if optional = "" && argument = "" {
                    validation := parameterMissingValue
                } else if !(IsInteger(argument) && (argument = 0 || argument = 1)) {
                    validation := "Parameter " . Chr(34) . parameter . Chr(34) . " must be Boolean (true/false) or Integer (0/1)."
                }
            case "Integer":
                if optional = "" && argument = "" {
                    validation := parameterMissingValue
                } else if !IsInteger(argument) {
                    validation := "Parameter " . Chr(34) . parameter . Chr(34) . " must be an Integer"
                } else {
                    switch typeValue {
                        case "Byte":
                            if argument < 0 || argument > 255 {
                                validation := "Value out of byte range (0–255): " . argument
                            }
                        case "Year":
                            if !RegExMatch(argument, "^\d+$") {
                                validation := "Year value must be integer digits only: " . argument
                            } else if argument < 1900 || argument > 2100 {
                                validation := "Year value must be between 1900 and 2100: " . argument
                            }
                    }
                }
            case "Object":
                ; Helper methods only, no handling.
            case "String":
                if optional = "" && argument = "" {
                    validation := parameterMissingValue
                } else if optional = "Optional" && argument = "" {
                    ; Skip validation regardless of type as no value exists.
                } else if pattern !== "" {
                    if !RegExMatch(argument, pattern) {
                        validation := "Argument pattern (" . pattern . ") does not validate against argument: " . argument
                    }
                } else if whitelist.Length != 0 {
                    valueIsWhitelisted := false

                    for index, whitelistEntry in whitelist {
                        if argument = whitelistEntry {
                            valueIsWhitelisted := true
                            break
                        }
                    }

                    if valueIsWhitelisted = false {
                        validation := "Failed as whitelist did not match argument: " . argument
                    }
                } else if Type(argument) != "String" {
                    validation := "Parameter " . Chr(34) . parameter . Chr(34) . " must be a String."
                } else {
                    switch typeValue {
                        case "Absolute Path", "Absolute Save Path":
                            isDrive := RegExMatch(argument, "^[A-Za-z]:\\")
                            isUNC   := RegExMatch(argument, "^\\\\{2}[^\\\/]+\\[^\\\/]+\\")

                            if !(isDrive || isUNC) {
                                validation := "Path must start with drive (C:\) or UNC (\\server\share\): " . argument
                            } else if !FileExist(argument) && typeValue = "Absolute Path" {
                                validation := "File doesn't exist: " . argument
                            }
                        case "Base64":
                            if !RegExMatch(argument, "^[A-Za-z0-9+/]*={0,2}$") {
                                validation := "Invalid Base64 content. Only A–Z, a–z, 0–9, +, /, and = allowed."
                            } else if Mod(StrLen(argument), 4) != 0 {
                                validation := "Invalid Base64 length. mMust be multiple of 4."
                            } else if RegExMatch(argument, "=[^=]") {
                                validation := "Invalid Base64 padding. The character = can only appear at the end."
                            }
                        case "Code":
                        case "Directory":
                            isDrive := RegExMatch(argument, "^[A-Za-z]:\\")
                            isUNC   := RegExMatch(argument, "^\\\\{2}[^\\\/]+\\[^\\\/]+\\")

                            if !(isDrive || isUNC) {
                                validation := "Path must start with drive (C:\) or UNC (\\server\share\): " . argument
                            } else if !DirExist(argument) && methodName !== "EnsureDirectoryExists" {
                                validation := "Directory doesn't exist: " . argument
                            } else if SubStr(argument, -1) != "\" {
                                validation := "Directory path must end with a backslash \: " . argument
                            }
                        case "Filename":
                            pattern := "[\\/:*?" . Chr(34) . "<>|]"

                            if RegExMatch(argument, pattern) {
                                forbiddenList := "\ / : * ? " Chr(34) " < > |"
                                validation := "Filename contains forbidden characters (" . forbiddenList . "): " . argument
                            } else if argument = "." || argument = ".." {
                                validation := "Filename reserved: " . argument
                            }
                        case "ISO Date Time":
                            if !RegExMatch(argument, "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$") {
                                validation := "Invalid ISO 8601 Date Time: " . argument . " (must be YYYY-MM-DD HH:MM:SS)."
                            } else {
                                dateTimeparts := StrSplit(argument, " ")
                                dateParts     := StrSplit(dateTimeparts[1], "-")
                                timeParts     := StrSplit(dateTimeparts[2], ":")

                                year   := dateParts[1] + 0
                                month  := dateParts[2] + 0
                                day    := dateParts[3] + 0
                                hour   := timeParts[1] + 0
                                minute := timeParts[2] + 0
                                second := timeParts[3] + 0

                                if validation = "" {
                                    validation := StrReplace(ValidateIsoDate(year, month, day, hour, minute, second), "ISO 8601 Date:", "ISO 8601 Date Time:")
                                }
                            }
                        case "ISO Date":
                            if !RegExMatch(argument, "^\d{4}-\d{2}-\d{2}$") {
                                validation := "Invalid ISO 8601 Date: " . argument . " (must be YYYY-MM-DD)."
                            } else {
                                dateParts := StrSplit(argument, "-")
                                year      := dateParts[1] + 0
                                month     := dateParts[2] + 0
                                day       := dateParts[3] + 0

                                if validation = "" {
                                    validation := ValidateIsoDate(year, month, day)
                                }
                            }
                        case "Percent Range":
                            if !RegExMatch(argument, "^\d{1,3}-\d{1,3}$") {
                                validation := "Must be two integers separated by the character -: " . argument
                            } else {
                                parts  := StrSplit(argument, "-")
                                first  := parts[1] + 0
                                second := parts[2] + 0

                                if first < 0 || first > 100 || second < 0 || second > 100 {
                                    validation := "Values must be between 0 and 100: " . argument
                                } else if first >= second {
                                    validation := "First value must be lower than second: " . argument
                                }
                            }
                        case "Raw Date Time":
                            if !RegExMatch(argument, "^\d{14}$") {
                                validation := "Must be in the format YYYYMMDDHHMMSS: " . argument
                            } else {
                                year   := SubStr(argument, 1, 4) + 0
                                month  := SubStr(argument, 5, 2) + 0
                                day    := SubStr(argument, 7, 2) + 0
                                hour   := SubStr(argument, 9, 2) + 0
                                minute := SubStr(argument, 11, 2) + 0
                                second := SubStr(argument, 13, 2) + 0

                                if validation = "" {
                                    validation := ValidateIsoDate(year, month, day, hour, minute, second, true)
                                }
                            }
                        case "Search", "Search Open":
                            pattern := "[\\/:*?" . Chr(34) . "<>|]"

                            if typeValue = "Search" && RegExMatch(argument, pattern) {
                                forbiddenList := "\ / : * ? " . Chr(34) . " < > |"
                                validation := "Contains forbidden characters (" . forbiddenList . "): " . argument
                            }
                        case "SHA-256":
                            if StrLen(argument) != 64 {
                                validation := "Expected length of 64 but instead got: " . StrLen(argument)
                            } else if !RegExMatch(argument, "^[0-9a-fA-F]{64}$") {
                                validation := "Must be hex digits only."
                            }
                    }
                }
        }
    }

    return validation
}

LogFormatMethodArguments(methodName, arguments, validation := "") {
    global symbolLedger

    argumentValueFull := ""
    argumentValueLog  := ""

    argumentsAndValidation := Map(
        "Arguments Full", "",
        "Arguments Log",  "",
        "Parameter",      "",
        "Argument",       "",
        "Data Type",      "",
        "Type",           "",
        "Validation",     validation
    )

    argumentsAndValidation["Arguments Full"] := ""
    argumentsAndValidation["Arguments Log"]  := ""

    if methodName = "OverlayInsertSpacer" {
        return argumentsAndValidation
    }

    for index, argument in arguments {
        argumentsAndValidation["Parameter"] := methodRegistry[methodName]["Parameter Contracts"][index]["Parameter Name"]
        argumentsAndValidation["Argument"]  := argument
        argumentsAndValidation["Data Type"] := methodRegistry[methodName]["Parameter Contracts"][index]["Data Type"]
        argumentsAndValidation["Type"]      := methodRegistry[methodName]["Parameter Contracts"][index]["Type"]

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
                switch argumentsAndValidation["Type"] {
                    case "Absolute Path", "Absolute Save Path":
                        SplitPath(argument, &filename, &directoryPath)

                        if !symbolLedger.Has(directoryPath . "|D") {
                            symbolLedger[directoryPath . "|D"] := Map(
                                "Symbol", NextSymbolLedgerAlias()
                            )

                            csvSymbolLedger :=
                                directoryPath . "|" . 
                                "D" . "|" . 
                                symbolLedger[directoryPath . "|D"]["Symbol"]

                            AppendCsvLineToLog(csvSymbolLedger, "Symbol Ledger")
                        }

                        if !symbolLedger.Has(filename . "|F") {
                            symbolLedger[filename . "|F"] := Map(
                                "Symbol", NextSymbolLedgerAlias()
                            )

                            csvSymbolLedger :=
                                filename . "|" . 
                                "F" . "|" . 
                                symbolLedger[filename . "|F"]["Symbol"]

                            AppendCsvLineToLog(csvSymbolLedger, "Symbol Ledger")
                        }

                        argumentValueLog := symbolLedger[directoryPath . "|D"]["Symbol"] . "\" . symbolLedger[filename . "|F"]["Symbol"]
                    case "Base64":
                        base64Summary := "<Base64 (Length: " . StrLen(argument) . ")>"

                        if !symbolLedger.Has(base64Summary . "|B") {
                            csvSymbolLedgerLine := RegisterSymbol(base64Summary, "Base64", false)
                            AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                        }

                        argumentValueFull := base64Summary
                        argumentValueLog := symbolLedger[base64Summary . "|B"]["Symbol"]
                    case "Code":
                        codeSummary := "<Code (Length: " . StrLen(argument) . ", Rows: " . StrSplit(argument, "`n").Length . ")>"

                        if !symbolLedger.Has(codeSummary . "|C") {
                            csvSymbolLedgerLine := RegisterSymbol(codeSummary, "Code", false)
                            AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                        }

                        argumentValueFull := codeSummary
                        argumentValueLog := symbolLedger[codeSummary . "|C"]["Symbol"]
                    case "Directory":
                        if !symbolLedger.Has(RTrim(argument, "\") . "|D") {
                            csvSymbolLedgerLine := RegisterSymbol(argument, "Directory", false)
                            AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                        }

                        argumentValueLog := symbolLedger[RTrim(argument, "\") . "|D"]["Symbol"]
                    case "Filename":
                        if !symbolLedger.Has(argument . "|F") {
                            csvSymbolLedgerLine := RegisterSymbol(argument, "Filename", false)
                            AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                        }

                        argumentValueLog := symbolLedger[argument . "|F"]["Symbol"]
                    case "ISO Date Time":
                        argumentValueLog := LocalIsoWithUtcTag(argument)
                    case "Search", "Search Open":
                        if !symbolLedger.Has(argument . "|S") {
                            csvSymbolLedgerLine := RegisterSymbol(argument, "Search", false)
                            AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                        }

                        argumentValueLog := symbolLedger[argument . "|S"]["Symbol"]
                    case "SHA-256":
                        encodedHash := EncodeSha256HexToBase(argument, 86)
                        if !symbolLedger.Has(encodedHash . "|H") {
                            csvSymbolLedgerLine := RegisterSymbol(encodedHash, "Hash", false)
                            AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
                        }

                        argumentValueLog := symbolLedger[encodedHash . "|H"]["Symbol"]
                    default:
                        if StrLen(argument) > 192 {
                            argumentValueFull := SubStr(argument, 1, 224) . "…"
                            argumentValueLog  := SubStr(argument, 1, 192) . "…"
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
    operationSequenceNumber        := NextOperationSequenceNumber()
    encodedOperationSequenceNumber := EncodeIntegerToBase(operationSequenceNumber, 86)
    encodedTickCount               := EncodeIntegerToBase(A_TickCount, 86)

    logValuesForConclusion["Operation Sequence Number"] := encodedOperationSequenceNumber

    csvConclusion := 
        logValuesForConclusion["Operation Sequence Number"] . "|" . ; Operation Sequence Number
        "F" .                                                 "|" . ; Status
        encodedTickCount                                            ; Tick

    if logValuesForConclusion["Context"] !== "" {
        csvConclusion := csvConclusion . "|" . logValuesForConclusion["Context"]
    }

    AppendCsvLineToLog(csvConclusion, "Operation Log")

    try {
        throw Error(errorMessage)
    } catch as customError {
        DisplayErrorMessage(logValuesForConclusion, customError, errorLineNumber)
    }
}

LogHelperValidation(methodName, arguments := unset) {
    argumentsAndValidationStatus := unset
    if IsSet(arguments) {
        validation := LogValidateMethodArguments(methodName, arguments)
        argumentsAndValidationStatus := LogFormatMethodArguments(methodName, arguments, validation)
    }

    argumentsFull := ""
    argumentsLog  := ""
    validation    := ""
    if IsSet(argumentsAndValidationStatus) {
        argumentsFull := argumentsAndValidationStatus["Arguments Full"]
        argumentsLog  := argumentsAndValidationStatus["Arguments Log"]
    }

    logValuesForConclusion := Map(
        "Operation Sequence Number", 0,
        "Method Name",               methodName,
        "Arguments Full",            argumentsFull,
        "Arguments Log",             argumentsLog,
        "Overlay Key",               0,
        "Validation",                validation,
        "Context",                   ""
    )

    if validation !== "" {
        operationSequenceNumber        := NextOperationSequenceNumber()
        encodedOperationSequenceNumber := EncodeIntegerToBase(operationSequenceNumber, 86)
        encodedTickCount               := EncodeIntegerToBase(A_TickCount, 86)

        logValuesForConclusion["Operation Sequence Number"] := encodedOperationSequenceNumber
        
        csvShared :=
            encodedOperationSequenceNumber .     "|" . ; Operation Sequence Number
            "B" .                                "|" . ; Status
            encodedTickCount .                   "|" . ; Tick
            methodRegistry[methodName]["Symbol"]       ; Symbol

        if logValuesForConclusion["Arguments Full"] !== "" {
            csvShared := csvShared . "|" . logValuesForConclusion["Arguments Log"] ; Arguments
        }

        AppendCsvLineToLog(csvShared, "Operation Log")

        try {
            throw Error(argumentsAndValidationStatus["Validation"])
        } catch as validationError {
            LogInformationConclusion("Failed", logValuesForConclusion, validationError)
        }
    }

    return logValuesForConclusion
}

LogInformationBeginning(overlayValue, methodName, arguments := unset, overlayCustomKey := 0) {
    static lastIntermissionTick := 0
    intermissionInterval := 6 * 60 * 1000
    intermissionTick := A_TickCount

    if lastIntermissionTick = 0 {
        lastIntermissionTick := intermissionTick
    }

    operationSequenceNumber        := NextOperationSequenceNumber()
    encodedOperationSequenceNumber := EncodeIntegerToBase(operationSequenceNumber, 86)
    encodedTickCount               := EncodeIntegerToBase(A_TickCount, 86)

    if overlayCustomKey = 0 {
        overlayKey                 := OverlayGenerateNextKey(methodName)
    } else {
        overlayKey := overlayCustomKey
        if methodName = "OverlayInsertSpacer" || methodName = "OverlayUpdateCustomLine" {
            arguments[1] := overlayKey
        }
    }

    if intermissionTick - lastIntermissionTick >= intermissionInterval {
        lastIntermissionTick := intermissionTick
        LogEngine("Intermission")
    }

    argumentsAndValidationStatus := unset
    if IsSet(arguments) {
        validation := LogValidateMethodArguments(methodName, arguments)
        argumentsAndValidationStatus := LogFormatMethodArguments(methodName, arguments, validation)
    }

    csvShared :=
        encodedOperationSequenceNumber .     "|" . ; Operation Sequence Number
        "B" .                                "|" . ; Status
        encodedTickCount .                   "|" . ; Tick
        methodRegistry[methodName]["Symbol"]       ; Symbol

    argumentsFull := ""
    argumentsLog  := ""
    if IsSet(argumentsAndValidationStatus) {
        csvShared := csvShared . "|" . argumentsAndValidationStatus["Arguments Log"] ; Arguments
        argumentsFull := argumentsAndValidationStatus["Arguments Full"]
        argumentsLog := argumentsAndValidationStatus["Arguments Log"]
    }

    if overlayKey !== 0 {
        csvShared := csvShared . "|" . ; Shared
        overlayKey . "|" .             ; Overlay Key
        overlayValue                   ; Overlay Value

        if overlayCustomKey = 0 {
            OverlayUpdateLine(overlayKey, overlayValue . overlayStatus["Beginning"])
        } else {
            OverlayUpdateLine(overlayKey, overlayValue)
        }
    }

    AppendCsvLineToLog(csvShared, "Operation Log")

    logValuesForConclusion := Map(
        "Operation Sequence Number", encodedOperationSequenceNumber,
        "Method Name",               methodName,
        "Arguments Full",            argumentsFull,
        "Arguments Log",             argumentsLog,
        "Overlay Key",               overlayKey,
        "Validation",                "",
        "Context",                   ""
    )

    try {
        if IsSet(argumentsAndValidationStatus) {
            logValuesForConclusion["Validation"] := argumentsAndValidationStatus["Validation"]

            if argumentsAndValidationStatus["Validation"] !== "" {
                throw Error(argumentsAndValidationStatus["Validation"])
            }
        }
    } catch as validationError {
        LogInformationConclusion("Failed", logValuesForConclusion, validationError)
    }

    return logValuesForConclusion
}

LogInformationConclusion(conclusionStatus, logValuesForConclusion, errorObject := unset) {
    encodedTickCount := EncodeIntegerToBase(A_TickCount, 86)

    csvConclusion := 
        logValuesForConclusion["Operation Sequence Number"] . "|" . ; Operation Sequence Number
        "[[Status]]" .                                        "|" . ; Status
        encodedTickCount                                            ; Tick

    if logValuesForConclusion["Context"] !== "" {
        csvConclusion := csvConclusion . "|" . logValuesForConclusion["Context"]
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

OverlayChangeTransparency(transparencyValue) {
    static methodName := RegisterMethod("OverlayChangeTransparency(transparencyValue As Integer [Type: Byte])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Change Transparency (" . transparencyValue . ")", methodName, [transparencyValue])

    WinSetTransparent(transparencyValue, "ahk_id " . overlayGui.Hwnd)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayChangeVisibility() {
    static methodName := RegisterMethod("OverlayChangeVisibility()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Change Visibility", methodName)

    if DllCall("user32\IsWindowVisible", "Ptr", overlayGui.Hwnd) {
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
            throw Error("Method " . Chr(34) . methodNameInput . Chr(34) . " not registered.")
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
            "Overlay Log",     true,
            "Symbol",          "",
            "Declaration",     "",
            "Signature",       "",
            "Library",         "",
            "Contract",        "",
            "Parameters",      "",
            "Data Types",      "",
            "Metadata",        "",
            "Validation Line", ""
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

    if methodRegistry[methodName]["Overlay Log"] = true {
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

    if !DllCall("user32\IsWindow", "Ptr", windowHandle) {
        return false
    }

    if DllCall("user32\IsWindowVisible", "Ptr", windowHandle) {
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
        DllCall("dwmapi\DwmGetWindowAttribute", "Ptr", overlayGui.Hwnd, "Int", 9, "Ptr", rectBuffer, "Int", 16),
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
    global overlayGui, overlayLines, overlayOrder

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
            default:
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

    logFilePath["Error Message"] := ""
    logFilePath["Execution Log"] := ""
    logFilePath["Operation Log"] := ""
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
        Loop 0x7E - 0x20 + 1 {
            codePoint := 0x20 + A_Index - 1

            if !excludedAsciiCodePoints.Has(codePoint) {
                baseCharacters .= Chr(codePoint)
            }
        }

        baseDigitByCharacterMap := Map()
        Loop StrLen(baseCharacters) {
            baseCharacter := SubStr(baseCharacters, A_Index, 1)
            baseDigitByCharacterMap[baseCharacter] := A_Index - 1
        }

        cachedResult := Map(
            "Characters", baseCharacters,
            "Base",       StrLen(baseCharacters),
            "Digit Map",  baseDigitByCharacterMap
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
    ; Maximum possible value: 9223372036854775807
    characterSetInfo := GetBaseCharacterSet(baseType)
    baseCharacters   := characterSetInfo["Characters"]
    baseRadix        := characterSetInfo["Base"]

    baseText  := ""
    if integerValue = 0 {
        baseText := SubStr(baseCharacters, 1, 1)
    } else {
        while integerValue > 0 {
            digitValue   := Mod(integerValue, baseRadix)
            baseText     := SubStr(baseCharacters, digitValue + 1, 1) . baseText
            integerValue := integerValue // baseRadix
        }
    }

    return baseText
}

DecodeBaseToInteger(baseText, baseType) {
    characterSetInfo        := GetBaseCharacterSet(baseType)
    baseRadix               := characterSetInfo["Base"]
    baseDigitByCharacterMap := characterSetInfo["Digit Map"]

    integerValue := 0
    Loop StrLen(baseText) {
        baseCharacter := SubStr(baseText, A_Index, 1)
        digitValue := baseDigitByCharacterMap[baseCharacter]
        integerValue := integerValue * baseRadix + digitValue
    }

    return integerValue
}

EncodeSha256HexToBase(hexSha256, baseType) {
    characterSetInfo := GetBaseCharacterSet(baseType)
    baseCharacters   := characterSetInfo["Characters"]
    baseRadix        := characterSetInfo["Base"]

    hexSha256 := StrLower(hexSha256)

    sha256Bytes := Buffer(32, 0)
    writeOffset := 0
    Loop 32 {
        twoHexDigits := SubStr(hexSha256, (A_Index - 1) * 2 + 1, 2)
        byteValue := ("0x" . twoHexDigits) + 0
        NumPut("UChar", byteValue, sha256Bytes, writeOffset)
        writeOffset += 1
    }

    isAllZero := true
    byteIndex := 0
    while byteIndex < 32 {
        if NumGet(sha256Bytes, byteIndex, "UChar") {
            isAllZero := false
            break
        }

        byteIndex += 1
    }

    baseDigitsLeastSignificantFirst := []
    if isAllZero {
        baseDigitsLeastSignificantFirst.Push(0)
    } else {
        loop {
            remainderValue := 0
            hasNonZeroQuotientByte := false
            byteIndex := 0
            while byteIndex < 32 {
                currentByte := NumGet(sha256Bytes, byteIndex, "UChar")
                accumulator := remainderValue * 256 + currentByte
                quotientByte := Floor(accumulator / baseRadix)
                remainderValue := Mod(accumulator, baseRadix)
                NumPut("UChar", quotientByte, sha256Bytes, byteIndex)
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
        baseText .= SubStr(baseCharacters, digitValue + 1, 1)
        digitIndex -= 1
    }

    requiredLength := Ceil(256 * Log(2) / Log(baseRadix))

    zeroDigit := SubStr(baseCharacters, 1, 1)
    while StrLen(baseText) < requiredLength {
        baseText := zeroDigit . baseText
    }

    return baseText
}

DecodeBaseToSha256Hex(baseText, baseType) {
    characterSetInfo          := GetBaseCharacterSet(baseType)
    baseRadix                 := characterSetInfo["Base"]
    baseDigitByCharacterMap   := characterSetInfo["Digit Map"]

    static sha256BytesBuffer := Buffer(32, 0)
    baseLength := StrLen(baseText)
    basePosition := 1
    while basePosition <= baseLength {
        baseCharacter := SubStr(baseText, basePosition, 1)
        digitValue    := baseDigitByCharacterMap[baseCharacter]

        carryValue := digitValue
        byteIndex := 31
        while byteIndex >= 0 {
            currentByte := NumGet(sha256BytesBuffer, byteIndex, "UChar")
            productValue := currentByte * baseRadix + carryValue
            NumPut("UChar", Mod(productValue, 256), sha256BytesBuffer, byteIndex)
            carryValue := Floor(productValue / 256)

            byteIndex -= 1
        }

        basePosition += 1
    }

    hexOutput := ""
    byteIndex := 0
    while byteIndex < 32 {
        hexOutput .= Format("{:02x}", NumGet(sha256BytesBuffer, byteIndex, "UChar"))

        byteIndex += 1
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

ExecutionLogBatchAppend(executionType, array) {
    static methodName := RegisterMethod("ExecutionLogBatchAppend(executionType As String, array as Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [executionType, array])

    static newLine := "`r`n"

    switch StrLower(executionType) {
        case "beginning":
            executionType  := "Beginning"
        case "completed":
            executionType  := "Completed"
        case "failed":
            executionType  := "Failed"
        case "intermission":
            executionType  := "Intermission"
        default:
            return
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

GetActiveKeyboardLayout() {
    static methodName := RegisterMethod("GetActiveKeyboardLayout()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    ; Returns: "en-US - US-International (00020409)"
    keyboardLayoutKlid := ""
    try {
        buf := Buffer(9*2, 0)
        if DllCall("user32\GetKeyboardLayoutNameW", "Ptr", buf) {
            keyboardLayoutKlid := StrGet(buf)
        }
    }

    if keyboardLayoutKlid = "" {
        ; Fallback: build KLID from HKL if needed
        try {
            hkl := DllCall("user32\GetKeyboardLayout", "UInt", 0, "Ptr")
            if hkl {
                keyboardLayoutKlid := Format("{:08X}", hkl & 0xFFFFFFFF)
            }
        }
    }

    ; Language (locale) from LANGID (last 4 hex digits)
    keyboardLayoutLanguageId := (keyboardLayoutKlid != "") ? SubStr(keyboardLayoutKlid, -3) : ""
    keyboardLayoutLocaleName := ""
    if keyboardLayoutLanguageId != "" {
        try {
            buf2 := Buffer(85*2, 0)
            if DllCall("kernel32\LCIDToLocaleName", "UInt", ("0x" . keyboardLayoutLanguageId) + 0, "Ptr", buf2, "Int", 85, "UInt", 0) {
                keyboardLayoutLocaleName := StrGet(buf2)
            }
        }
    }
    if keyboardLayoutLocaleName = "" {
        keyboardLayoutLocaleName := (keyboardLayoutLanguageId != "") ? keyboardLayoutLanguageId : "Unknown Language"
    }

    ; Layout text from registry
    keyboardLayoutLayoutText := ""
    if keyboardLayoutKlid != "" {
        try {
            keyboardLayoutLayoutText := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Keyboard Layouts\" . keyboardLayoutKlid, "Layout Text")
        }

        if keyboardLayoutLayoutText = "" {
            try {
                keyboardLayoutLayoutText := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Keyboard Layouts\" . keyboardLayoutKlid, "Layout Display Name")
            }

            ; Resolve @-style resource names when possible
            if keyboardLayoutLayoutText != "" && SubStr(keyboardLayoutLayoutText, 1, 1) = "@" {
                try {
                    buf3 := Buffer(260*2, 0)
                    if DllCall("shlwapi\SHLoadIndirectString", "WStr", keyboardLayoutLayoutText, "Ptr", buf3, "Int", 260, "Ptr", 0) = 0 {
                        keyboardLayoutLayoutText := StrGet(buf3)
                    }
                }
            }
        }
    }

    if keyboardLayoutLayoutText = "" {
        keyboardLayoutLayoutText := "Unknown Layout"
    }

    if keyboardLayoutKlid = "" {
        keyboardLayoutKlid := "????????"
    }

    keyboardLayout := keyboardLayoutLocaleName . " - " . keyboardLayoutLayoutText . " (" . keyboardLayoutKlid . ")"

    return keyboardLayout
}

GetActiveDisplayGpu() {
    static methodName := RegisterMethod("GetActiveDisplayGpu()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    windowsManagementInstrumentation := ComObjGet("winmgmts:\\.\root\CIMV2")
    activeModelName := ""
    firstModelName := ""
    for controller in windowsManagementInstrumentation.ExecQuery(
        "SELECT Name, CurrentHorizontalResolution, CurrentVerticalResolution FROM Win32_VideoController"
    ) {
        if firstModelName = "" {
            firstModelName := controller.Name
        }

        if controller.CurrentHorizontalResolution > 0 && controller.CurrentVerticalResolution > 0 {
            activeModelName := controller.Name
            break
        }
    }

    resultModelName := "Unknown GPU"
    if activeModelName != "" {
        resultModelName := activeModelName
    } else if firstModelName != "" {
        resultModelName := firstModelName
    }

    return resultModelName
}

GetCpu() {
    static methodName := RegisterMethod("GetCpu()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    resultModelName := ""
    defaultModelName := "Unknown CPU"
    registryPath := "HKLM\HARDWARE\DESCRIPTION\System\CentralProcessor\0"
    registryValueName := "ProcessorNameString"

    rawName := ""
    cleanedName := ""
    try {
        rawName := RegRead(registryPath, registryValueName)
        cleanedName := Trim(RegExReplace(rawName, "\s+", " "))
    }

    if cleanedName != "" {
        resultModelName := cleanedName
    }

    if resultModelName = "" {
        resultModelName := defaultModelName
    }

    return resultModelName
}

GetInputLanguage() {
    static methodName := RegisterMethod("GetInputLanguage()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    inputLanguageLocaleName := ""
    try {
        hkl := DllCall("user32\GetKeyboardLayout", "UInt", 0, "Ptr")
        langId := hkl & 0xFFFF
        buf := Buffer(85*2, 0)
        if DllCall("kernel32\LCIDToLocaleName", "UInt", langId, "Ptr", buf, "Int", 85, "UInt", 0) {
            inputLanguageLocaleName := StrGet(buf)
        }
    }
    if inputLanguageLocaleName = "" {
        inputLanguageLocaleName := "Unknown Language"
    }

    return inputLanguageLocaleName
}

GetMemorySizeAndType() {
    static methodName := RegisterMethod("GetMemorySizeAndType()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    windowsManagementInstrumentationService := ComObjGet("winmgmts:root\cimv2")
    memoryModuleRecords := windowsManagementInstrumentationService.ExecQuery(
        "SELECT Capacity, SMBIOSMemoryType, PartNumber FROM Win32_PhysicalMemory"
    )

    totalInstalledMemoryBytes := 0
    memoryTypeCodeCounts := Map()
    partNumberStrings := []

    for memoryModuleRecord in memoryModuleRecords {
        try {
            totalInstalledMemoryBytes += memoryModuleRecord.Capacity + 0
        }

        systemManagementBiosMemoryTypeCode := ""
        try {
            systemManagementBiosMemoryTypeCode := memoryModuleRecord.SMBIOSMemoryType + 0
        }

        if systemManagementBiosMemoryTypeCode != "" {
            if !memoryTypeCodeCounts.Has(systemManagementBiosMemoryTypeCode) {
                memoryTypeCodeCounts[systemManagementBiosMemoryTypeCode] := 0
            }

            memoryTypeCodeCounts[systemManagementBiosMemoryTypeCode] += 1
        }

        partNumberValue := ""
        try {
            partNumberValue := Trim(memoryModuleRecord.PartNumber . "")
        }

        if partNumberValue != "" {
            partNumberStrings.Push(partNumberValue)
        }
    }

    installedMemorySizeInGigabytes := (totalInstalledMemoryBytes > 0) ? Round(totalInstalledMemoryBytes / 1024 / 1024 / 1024) : 0
    installedMemorySizeDisplay := (installedMemorySizeInGigabytes > 0) ? (installedMemorySizeInGigabytes . " GB") : "Unknown Size"

    memoryTypeDisplayByCode := Map(
        14, "SDRAM",
        17, "SGRAM",
        20, "DDR SDRAM",
        21, "DDR2 SDRAM",
        22, "DDR2 FB-DIMM",
        24, "DDR3 SDRAM",
        26, "DDR4 SDRAM",
        34, "DDR5 SDRAM"
    )

    installedMemoryTypeDisplay := ""
    if memoryTypeCodeCounts.Count > 0 {
        mostCommonMemoryTypeCode := ""
        mostCommonMemoryTypeCount := -1
        for memoryTypeCode, memoryTypeCount in memoryTypeCodeCounts {
            if memoryTypeCount > mostCommonMemoryTypeCount {
                mostCommonMemoryTypeCount := memoryTypeCount
                mostCommonMemoryTypeCode := memoryTypeCode
            }
        }

        if memoryTypeDisplayByCode.Has(mostCommonMemoryTypeCode) {
            installedMemoryTypeDisplay := memoryTypeDisplayByCode[mostCommonMemoryTypeCode]
        } else {
            installedMemoryTypeDisplay := "Unknown Type (code " . mostCommonMemoryTypeCode . ")"
        }
    }

    if installedMemoryTypeDisplay = "" || InStr(installedMemoryTypeDisplay, "Unknown") {
        combinedPartNumbers := ""
        for partNumberItem in partNumberStrings {
            combinedPartNumbers .= partNumberItem . " "
        }
        combinedPartNumbersLower := StrLower(combinedPartNumbers)

        switch true {
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
    static methodName := RegisterMethod("GetMotherboard()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    windowsManagementInstrumentationService := ComObjGet("winmgmts:root\cimv2")
    baseboardRecords := windowsManagementInstrumentationService.ExecQuery(
        "SELECT Manufacturer, Product FROM Win32_BaseBoard"
    )

    rawManufacturer := ""
    rawProduct := ""
    for baseboardRecord in baseboardRecords {
        try {
            rawManufacturer := Trim(baseboardRecord.Manufacturer . "")
        }

        try {
            rawProduct := Trim(baseboardRecord.Product . "")
        }
        
        break
    }

    if rawManufacturer = "" {
        rawManufacturer := "Unknown Manufacturer"
    }

    if rawProduct = "" {
        rawProduct := "Unknown Product"
    }

    normalizedManufacturer := rawManufacturer
    switch StrLower(rawManufacturer) {
        case "asustek computer inc.", "asustek computer inc", "asustek computer incorporated":
            normalizedManufacturer := "ASUS"
        case "micro-star international co., ltd.", "micro-star international co.,ltd.":
            normalizedManufacturer := "MSI"
        case "gigabyte technology co., ltd.", "giga-byte technology co., ltd.":
            normalizedManufacturer := "GIGABYTE"
        case "hewlett-packard", "hp", "hp inc.", "hewlett packard":
            normalizedManufacturer := "HP"
    }

    motherboard := Trim(normalizedManufacturer . " " . rawProduct)
    
    return motherboard
}

GetOperatingSystem() {
    static methodName := RegisterMethod("GetOperatingSystem()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    currentBuildNumber := ""
    try {
        currentBuildNumber := RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuildNumber")
    }

    currentBuildNumberNumeric := currentBuildNumber + 0

    family := "Unknown Windows"
    switch true {
        case currentBuildNumberNumeric >= 22000:
            family := "Windows 11"
        case currentBuildNumberNumeric >= 10240:
            family := "Windows 10"
        case currentBuildNumberNumeric >= 9600:
            family := "Windows 8.1"
        case currentBuildNumberNumeric >= 9200:
            family := "Windows 8"
        case currentBuildNumberNumeric >= 7600:
            family := "Windows 7"
    }

    edition := "Unknown Edition"
    try {
        edition := RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "EditionID")
    }

    architectureTag := A_Is64bitOS ? "x64" : "x86"

    version := "Build "
    switch true {
        case currentBuildNumberNumeric >= 10240:
            version := version . SubStr(A_OSVersion, 6) ; Remove "10.0." for Windows 10 and upward.
        case currentBuildNumberNumeric >= 7600:
            version := version . A_OSVersion
        default:
            version := "Unknown Version"
    }

    updateBuildRevisionNumber := ""
    try {
        updateBuildRevisionNumber := RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "UBR")
    }

    if updateBuildRevisionNumber !== "" {
        updateBuildRevisionNumber := "." . updateBuildRevisionNumber
    }

    releaseDisplay := ""
    try {
        releaseDisplay := RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "DisplayVersion")
    }

    if releaseDisplay = "" {
        try {
            releaseDisplay := RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId")
        }
    }

    if releaseDisplay !== "" {
        releaseDisplay := " (" . releaseDisplay . ")"
    }

    operatingSystem := "Microsoft " . family . " " . edition . " " . "(" . architectureTag . ")" . " " . version . updateBuildRevisionNumber . releaseDisplay

    return operatingSystem
}

GetPhysicalMemoryStatus() {
    static methodName := RegisterMethod("GetPhysicalMemoryStatus()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    memoryStatusExSize := 64
    memoryBuffer := Buffer(memoryStatusExSize, 0)
    NumPut("UInt", memoryStatusExSize, memoryBuffer, 0)

    if !DllCall("Kernel32.dll\GlobalMemoryStatusEx", "Ptr", memoryBuffer.Ptr, "Int") {
        physicalRamSituation := "GlobalMemoryStatusEx failed."

        return physicalRamSituation
    } else {
        memoryLoadPercent              := NumGet(memoryBuffer, 4,  "UInt")
        totalPhysicalMemoryBytes       := NumGet(memoryBuffer, 8,  "UInt64")
        availablePhysicalMemoryBytes   := NumGet(memoryBuffer, 16, "UInt64")

        ; Divide by 1024^3 so results align with Task Manager (labeled GB)
        bytesPerGB := 1024**3
        totalGB     := totalPhysicalMemoryBytes / bytesPerGB
        availableGB := availablePhysicalMemoryBytes / bytesPerGB
        usedGB      := totalGB - availableGB

        physicalRamSituation := "Total " . Format("{:.1f}", totalGB) . " GB, " . "Available " . Format("{:.1f}", availableGB) . " GB " . "(" . memoryLoadPercent . "%" . ")"

        return physicalRamSituation
    }
}

GetQueryPerformanceCounterFrequency() {
    static methodName := RegisterMethod("GetQueryPerformanceCounterFrequency()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    static queryPerformanceCounterFrequencyBuffer := Buffer(8, 0)
    queryPerformanceCounterFrequencyRetrievedSuccessfully := DllCall("QueryPerformanceFrequency", "Ptr", queryPerformanceCounterFrequencyBuffer.Ptr, "Int")
    if queryPerformanceCounterFrequencyRetrievedSuccessfully = false {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve the frequency of the performance counter. [QueryPerformanceFrequency" . ", System Error Code: " . A_LastError . "]")
    }

    queryPerformanceCounterFrequency := NumGet(queryPerformanceCounterFrequencyBuffer, 0, "Int64")

    return queryPerformanceCounterFrequency
}

GetRegionFormat() {
    static methodName := RegisterMethod("GetRegionFormat()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    regionFormat := ""
    try {
        regionFormat  := RegRead("HKEY_CURRENT_USER\Control Panel\International", "LocaleName")
    }
    if regionFormat  = "" {
        regionFormat  := "Unknown"
    }

    return regionFormat
}

GetRemainingFreeDiskSpace() {
    static methodName := RegisterMethod("GetRemainingFreeDiskSpace()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    freeMB := DriveGetSpaceFree(A_MyDocuments)
    result := freeMB . " MB (" . Format("{:.2f}", Floor((freeMB/1024/1024)*100) / 100) . " TB)"

    return result
}

GetSystemDisk() {
    static methodName := RegisterMethod("GetSystemDisk()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    diskModelText := "Unknown Disk"
    diskCapacityText := "Unknown"
    systemPartitionCapacityText := "Unknown"

    try {
        windowsManagementLocator := ComObject("WbemScripting.SWbemLocator")
        windowsManagementService := windowsManagementLocator.ConnectServer(".", "root\CIMV2")
        windowsManagementService.Security_.ImpersonationLevel := 3

        logicalDriveLetter := SubStr(A_WinDir, 1, 2)

        systemPartitionByteCount := ""
        for logicalDisk in windowsManagementService.ExecQuery(
            "SELECT Size FROM Win32_LogicalDisk WHERE DeviceID='" . logicalDriveLetter . "'"
        ) {
            systemPartitionByteCount := logicalDisk.Size + 0
            break
        }

        selectedPartitionDeviceId := ""
        for partitionObject in windowsManagementService.ExecQuery(
            "ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='" . logicalDriveLetter . "'} " .
            "WHERE AssocClass=Win32_LogicalDiskToPartition"
        ) {
            selectedPartitionDeviceId := partitionObject.DeviceID
            break
        }

        physicalDiskByteCount := ""
        if selectedPartitionDeviceId != "" {
            for diskDriveObject in windowsManagementService.ExecQuery(
                "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" . selectedPartitionDeviceId . "'} " .
                "WHERE AssocClass=Win32_DiskDriveToDiskPartition"
            ) {
                if Trim(diskDriveObject.Model) != "" {
                    diskModelText := Trim(diskDriveObject.Model)
                }

                if diskDriveObject.Size != "" && diskDriveObject.Size >= 0 {
                    physicalDiskByteCount := diskDriveObject.Size + 0
                }

                break
            }
        }

        ; Use StrFormatByteSizeW to format exactly like Windows Explorer.
        if physicalDiskByteCount != "" && physicalDiskByteCount >= 0 {
            bufferDisk := Buffer(64, 0)
            DllCall("shlwapi\StrFormatByteSizeW", "Int64", physicalDiskByteCount, "Ptr", bufferDisk, "UInt", 64)
            diskCapacityText := StrGet(bufferDisk, "UTF-16")
        }

        if systemPartitionByteCount != "" && systemPartitionByteCount >= 0 {
            bufferPartition := Buffer(64, 0)
            DllCall("shlwapi\StrFormatByteSizeW", "Int64", systemPartitionByteCount, "Ptr", bufferPartition, "UInt", 64)
            systemPartitionCapacityText := StrGet(bufferPartition, "UTF-16")
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

GetTimeZoneKeyName() {
    static methodName := RegisterMethod("GetTimeZoneKeyName()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    timeZoneKeyName := "Unknown"

    dynamicTimeZoneInformationBuffer := Buffer(432, 0)
    callResult := DllCall("kernel32\GetDynamicTimeZoneInformation", "Ptr", dynamicTimeZoneInformationBuffer, "UInt")
    if callResult != 0xFFFFFFFF {
        extractedKey := StrGet(dynamicTimeZoneInformationBuffer.Ptr + 172, 128, "UTF-16")

        if extractedKey != "" {
            timeZoneKeyName := extractedKey
        }
    }

    return timeZoneKeyName
}

GetWindowsColorMode() {
    static methodName := RegisterMethod("GetWindowsColorMode()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    registryPath := "HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"

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
    resultColorMode := ""

    switch presentFlagCount
    {
        Case 2:
            if appsUseLightThemeFlag = systemUsesLightThemeFlag {
                if (appsUseLightThemeFlag = 1) {
                    resultColorMode := "Light"
                } else {
                    resultColorMode := "Dark"
                }
            } else {
                resultColorMode := "Custom"
            }
        Case 1:
            onlyFlag := hasAppsUseLightTheme ? appsUseLightThemeFlag : systemUsesLightThemeFlag
            if onlyFlag = 1 {
                resultColorMode := "Light"
            } else {
                resultColorMode := "Dark"
            }
        Default:
            ; Keys do not exist, treat as Light (closest equivalent).
            resultColorMode := "Light"
    }

    return resultColorMode
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

    if contract !== "" {
        squareBracketDepth := 0               ; Tracks how many [ … ] levels we are inside
        inQuotedString := false               ; True while inside " … "
        removeLeadingSpaceAfterComma := false ; True right after a delimiter comma has been processed

        Loop Parse contract
        {
            currentCharacter := A_LoopField

            ; Toggle quoted-string mode on a double quote (").
            ; While inQuotedString = true, commas and brackets are considered literal characters.
            if currentCharacter = Chr(34) { ; Chr(34) = "
                inQuotedString := !inQuotedString
                currentParameterText .= currentCharacter
                continue
            }

            ; If not inside quotes, structural characters may affect parsing.
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

                ; Top-level comma → this marks the end of one parameter.
                if currentCharacter = "," && squareBracketDepth = 0 {
                    parameterParts.Push(Trim(currentParameterText))
                    currentParameterText := ""
                    removeLeadingSpaceAfterComma := true
                    continue
                }
            }

            ; Immediately after a delimiter comma, drop exactly one space if it exists.
            if removeLeadingSpaceAfterComma && currentCharacter = " " {
                removeLeadingSpaceAfterComma := false
                continue
            }
            removeLeadingSpaceAfterComma := false

            currentParameterText .= currentCharacter
        }

        ; Push the final parameter (even if empty → reflects missing/empty parameter).
        parameterParts.Push(Trim(currentParameterText))

        for index, parameterClause in parameterParts {
            RegExMatch(parameterClause, "^[A-Za-z_][A-Za-z0-9_]*", &matchObject)
            parameterName := matchObject[0]
            metadataValue := Trim(SubStr(parameterClause, StrLen(parameterName) + 1))
            metadataValue := RegExReplace(metadataValue, "^(?i)As\s+", "")

            dataTypesValue := StrSplit(metadataValue, " ")[1]
            metadataValue := SubStr(metadataValue, StrLen(dataTypesValue) + 2)

            optionalValue  := ""
            patternValue   := ""
            typeValue      := ""
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
                            optionalValue  := conceptValue
                        }
                    case "pattern":
                        patternValue   := conceptValue
                    case "type":
                        typeValue      := conceptValue
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
                "Parameter Name", parameterName,
                "Data Type",      dataTypesValue,
                "Optional",       optionalValue,
                "Pattern",        patternValue,
                "Type",           typeValue,
                "Whitelist",      whitelist
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
        "Overlay Log",         "",
        "Symbol",              "",
        "Declaration",         declaration,
        "Signature",           signature,
        "Library",             library,
        "Contract",            contract,
        "Parameters",          parameters,
        "Data Types",          dataTypes,
        "Metadata",            metadata,
        "Validation Line",     lineNumberForValidation,
        "Parameter Contracts", parameterContracts
    )

    return parsedMethod
}

RegisterMethod(declaration, sourceFilePath := "", validationLineNumber := 0) {
    if sourceFilePath !== "" && validationLineNumber !== 0 {
        SplitPath(sourceFilePath, , , , &filenameWithoutExtension)
        libraryTag := " @ " . filenameWithoutExtension
        validationLineNumber := " " . "<" . validationLineNumber . ">"
        declaration := declaration . libraryTag . validationLineNumber
    }

    parsedMethod := ParseMethodDeclaration(declaration)

    methodName := RTrim(SubStr(parsedMethod["Signature"], 1, InStr(parsedMethod["Signature"], "(") - 1))

    if !symbolLedger.Has(declaration . "|" . "M") {
        csvSymbolLedgerLine := RegisterSymbol(declaration, "M", false)
        AppendCsvLineToLog(csvSymbolLedgerLine, "Symbol Ledger")
        csvParts := StrSplit(csvSymbolLedgerLine, "|")
        symbol   := csvParts[csvParts.Length]
    } else {
        ; Later logic for dealing with re-use of Symbol Ledger.
    }

    global methodRegistry

    if methodRegistry.Has(methodName) {
        methodRegistry[methodName]["Declaration"]         := parsedMethod["Declaration"]
        methodRegistry[methodName]["Symbol"]              := symbol
        methodRegistry[methodName]["Signature"]           := parsedMethod["Signature"]
        methodRegistry[methodName]["Library"]             := parsedMethod["Library"]
        methodRegistry[methodName]["Contract"]            := parsedMethod["Contract"]
        methodRegistry[methodName]["Parameters"]          := parsedMethod["Parameters"]
        methodRegistry[methodName]["Data Types"]          := parsedMethod["Data Types"]
        methodRegistry[methodName]["Metadata"]            := parsedMethod["Metadata"]
        methodRegistry[methodName]["Validation Line"]     := parsedMethod["Validation Line"]
        methodRegistry[methodName]["Parameter Contracts"] := parsedMethod["Parameter Contracts"]
    } else {
        methodRegistry[methodName] := Map(
            "Overlay Log",         false,
            "Symbol",              symbol,
            "Declaration",         parsedMethod["Declaration"],
            "Signature",           parsedMethod["Signature"],
            "Library",             parsedMethod["Library"],
            "Contract",            parsedMethod["Contract"],
            "Parameters",          parsedMethod["Parameters"],
            "Data Types",          parsedMethod["Data Types"],
            "Metadata",            parsedMethod["Metadata"],
            "Validation Line",     parsedMethod["Validation Line"],
            "Parameter Contracts", parsedMethod["Parameter Contracts"]
        )
    }

    return methodName
}

RegisterSymbol(value, type, addNewLine := true) {
    global symbolLedger
    static newLine := "`r`n"
    symbolLine     := ""

    switch StrLower(type) {
        case "application", "a":
            type := "A"
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
        case "method", "m":
            type := "M"
        case "search", "s":
            type := "S"
        default:
            type := "[[Invalid]]"
    }

    if !symbolLedger.Has(value . "|" . type) {
        symbolLedger[value . "|" . type] := Map(
            "Symbol", NextSymbolLedgerAlias()
        )

        symbolLine :=
            value . "|" . 
            type . "|" . 
            symbolLedger[value . "|" . type]["Symbol"]
    }

    if type !== "[[Invalid]]" && addNewLine = true {
        symbolLine := symbolLine . newLine
    }

    return symbolLine
}

SymbolLedgerBatchAppend(symbolType, array) {
    static methodName := RegisterMethod("SymbolLedgerBatchAppend(symbolType As String, array As Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [symbolType, array])

    static newLine := "`r`n"

    switch StrLower(symbolType) {
        case "application", "a":
            symbolType := "A"
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
        case "search", "s":
            symbolType := "S"
        default:
            return
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
        } else if symbolType = "H" {
            value := EncodeSha256HexToBase(value, 86)
        }

        if !symbolLedger.Has(value . "|" . symbolType) {
            symbolLedgerArray.Push(value)
        }
    }
  
    arrayLength := symbolLedgerArray.Length
    if arrayLength !== 0 {
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