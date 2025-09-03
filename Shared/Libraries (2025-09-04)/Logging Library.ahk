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
    static methodName := RegisterMethod("AbortExecution()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Abort Execution", methodName)

    try {
        throw Error("Execution aborted early by pressing escape.")
    } catch as executionAbortedError {
        LogInformationConclusion("Failed", logValuesForConclusion, executionAbortedError)
    }    
}

DisplayErrorMessage(logValuesForConclusion, errorObject) {
    windowTitle := "AutoHotkey v" . A_AhkVersion . ": " . A_ScriptName
    currentDateTime := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")

    errorMessage := (errorObject.HasOwnProp("Message") ? errorObject.Message : errorObject)

    lineNumber := errorObject.Line
    if logValuesForConclusion["Validation"] !== "" {
        lineNumber := methodRegistry[logValuesForConclusion["Method Name"]]["Validation Line"]
    }

    fullErrorText := unset
    if methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] !== "" {
        fullErrorText :=
            "Declaration: " .  methodRegistry[logValuesForConclusion["Method Name"]]["Declaration"] . "`n" . 
            "Parameters: " .   methodRegistry[logValuesForConclusion["Method Name"]]["Parameters"] . "`n" . 
            "Arguments: " .    logValuesForConclusion["Arguments"] . "`n" . 
            "Line Number: " .  lineNumber . "`n" . 
            "Date Runtime: " . currentDateTime . "`n" . 
            "Error Output: " . errorMessage
    } else {
        fullErrorText :=
            "Declaration: " .  methodRegistry[logValuesForConclusion["Method Name"]]["Declaration"] . "`n" . 
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
        errorWindow.AddEdit("ReadOnly r10 w960 -VScroll vErrorTextField", fullErrorText)

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
    static intermissionFlushInterval := 120 * 60 * 1000
    static newLine := "`r`n"
    executionLogLines := []
    SplitPath(A_ScriptFullPath, , &projectFolderPath, , &filenameWithoutExtension)
    timeAnchor := CaptureTimeAnchor()

    if status = "Beginning" && logFilePath["Error Message"] = "" && logFilePath["Execution Log"] = "" && logFilePath["Operation Log"] = "" && logFilePath["Symbol Ledger"] = "" {
        global logFilePath

        if !DirExist(projectFolderPath . "\History") {
            DirCreate(projectFolderPath . "\History")
        }

        if !DirExist(projectFolderPath . "\Log") {
            DirCreate(projectFolderPath . "\Log")
        }

         if !DirExist(projectFolderPath . "\Projects") {
            DirCreate(projectFolderPath . "\Projects")
        }
        
        logFileDateOfToday := StrReplace(timeAnchor["Local Date Time ISO"], ":", ".") . "." . timeAnchor["Milliseconds Part"]
        
        logFilePath["Error Message"] := projectFolderPath . "\Log\" . filenameWithoutExtension . " - " . "Error Message" . " - " . logFileDateOfToday . ".csv"
        logFilePath["Execution Log"] := projectFolderPath . "\Log\" . filenameWithoutExtension . " - " . "Execution Log" . " - " . logFileDateOfToday . ".csv"
        logFilePath["Operation Log"] := projectFolderPath . "\Log\" . filenameWithoutExtension . " - " . "Operation Log" . " - " . logFileDateOfToday . ".csv"
        logFilePath["Symbol Ledger"] := projectFolderPath . "\Log\" . filenameWithoutExtension . " - " . "Symbol Ledger" . " - " . logFileDateOfToday . ".csv"
    }

    if status !== "Intermission" {
        inputLanguage            := GetInputLanguage()
        keyboardLayout           := GetActiveKeyboardLayout()
        regionFormat             := GetRegionFormat()
        primaryDisplayResolution := A_ScreenWidth . "x" . A_ScreenHeight
        remainingDiskSpace       := GetRemainingFreeDiskSpace()

        executionLogLines := [
            "Project Name: " .                           filenameWithoutExtension,
            "AutoHotkey Runtime Version: " .             A_AhkVersion,
            "Script File Hash: " .                       Hash.File("SHA256", A_ScriptFullPath),
            "Tick Before Change: " .                     timeAnchor["Tick Before Change"],
            "Tick After Change: " .                      timeAnchor["Tick After Change"],
            "Precise UTC FileTime Midpoint: " .          timeAnchor["Precise UTC FileTime Midpoint"],
            "UTC Date Time ISO: " .                      timeAnchor["UTC Date Time ISO"],
            "Local Date Time ISO: " .                    timeAnchor["Local Date Time ISO"],
            "Milliseconds Part: " .                      timeAnchor["Milliseconds Part"],
            "QueryPerformanceCounter Ticks Midpoint: " . timeAnchor["QueryPerformanceCounter Ticks Midpoint"],
            "Computer Name: " .                          A_ComputerName,
            "Username: " .                               A_UserName,
            "Operating System Family: " .                GetOperatingSystemFamilyAndEdition(),
            "Operating System Version: " .               A_OSVersion,
            "Operating System Architecture: " .          (A_Is64bitOS ? "64-bit" : "32-bit"),
            "Input Language: " .                         inputLanguage,
            "Keyboard Layout: " .                        keyboardLayout,
            "Region Format: " .                          regionFormat,
            "Primary Display Resolution: " .             primaryDisplayResolution,
            "Remaining Free Disk Space: " .              remainingDiskSpace
        ]
    } else {
        for line in [
            "Tick Before Change: " .                     timeAnchor["Tick Before Change"],
            "Tick After Change: " .                      timeAnchor["Tick After Change"],
            "Precise UTC FileTime Midpoint: " .          timeAnchor["Precise UTC FileTime Midpoint"],
            "Local Date Time ISO: " .                    timeAnchor["Local Date Time ISO"],
            "Milliseconds Part: " .                      timeAnchor["Milliseconds Part"],
            "QueryPerformanceCounter Ticks Midpoint: " . timeAnchor["QueryPerformanceCounter Ticks Midpoint"]
        ] {
            intermissionBuffer.Push(line)
        }
    }

    static operationLogLine := "Operation Sequence Number|Status|Tick|Symbol|Arguments|Overlay Key|Overlay Value"
    static symbolLedgerLine := "Reference|Type|Symbol"

    Switch status {
        Case "Beginning":
            errorMessageFileHandle := FileOpen(logFilePath["Error Message"], "w", "UTF-8")
            errorMessageFileHandle.Close()

            executionLogFileHandle := FileOpen(logFilePath["Execution Log"], "w", "UTF-8")
            executionLogFileHandle.Close()

            ExecutionLogBatchAppend("Beginning", executionLogLines)

            operationLogFileHandle := FileOpen(logFilePath["Operation Log"], "w", "UTF-8")
            operationLogFileHandle.Close()

            AppendCsvLineToLog(operationLogLine, "Operation Log")

            symbolLedgerFileHandle := FileOpen(logFilePath["Symbol Ledger"], "w", "UTF-8")
            symbolLedgerFileHandle.Close()

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
            if (currentTick - lastIntermissionFlushTick >= intermissionFlushInterval) {
                lastIntermissionFlushTick := currentTick
                ExecutionLogBatchAppend("Intermission", intermissionBuffer)
                intermissionBuffer.Length := 0
            }
        default:
    }
}

LogFormatArgumentsAndValidate(methodName, arguments) {
    global symbolLedger

    argumentValueFull := ""
    argumentValueLog  := ""

    argumentsAndValidation := Map(
        "Arguments Full", "",
        "Arguments Log",  "",
        "Parameter",      "",
        "Argument",       "",
        "Data Type",      "",
        "Optional",       "",
        "Pattern",        "",
        "Type",           "",
        "Whitelist",      [],
        "Validation",     ""
    )

    argumentsAndValidation["Arguments Full"] := ""
    argumentsAndValidation["Arguments Log"]  := ""
    for index, argumentValue in arguments {
        argumentsAndValidation["Parameter"] := methodRegistry[methodName]["Parameter Contracts"][index]["Parameter Name"]
        argumentsAndValidation["Argument"]  := argumentValue
        argumentsAndValidation["Data Type"] := methodRegistry[methodName]["Parameter Contracts"][index]["Data Type"]
        argumentsAndValidation["Optional"]  := methodRegistry[methodName]["Parameter Contracts"][index]["Optional"]
        argumentsAndValidation["Pattern"]   := methodRegistry[methodName]["Parameter Contracts"][index]["Pattern"]
        argumentsAndValidation["Type"]      := methodRegistry[methodName]["Parameter Contracts"][index]["Type"]
        argumentsAndValidation["Whitelist"] := methodRegistry[methodName]["Parameter Contracts"][index]["Whitelist"]

        argumentValueFull := argumentValue
        argumentValueLog  := argumentValue
        switch argumentsAndValidation["Data Type"] {
            case "Boolean":
                if argumentsAndValidation["Optional"] = "" && argumentValue = "" {
                    argumentsAndValidation["Validation"] := "Parameter " . argumentsAndValidation["Parameter"] . " has no value passed into it."
                } else if (Type(argumentValue) = "Boolean") {
                    argumentValueFull := (argumentValue ? "true" : "false")
                    argumentValueLog  := argumentValueFull
                } else if (Type(argumentValue) = "Integer" && (argumentValue = 0 || argumentValue = 1)) {
                    argumentValueFull := (argumentValue ? "true" : "false")
                    argumentValueLog  := argumentValueFull
                } else {
                    argumentsAndValidation["Validation"] := "Parameter " . argumentsAndValidation["Parameter"] . " must be Boolean (true/false) or Integer (0/1)."
                }
            case "Integer":
                if !(Type(argumentValue) = "Integer") {
                    argumentsAndValidation["Validation"] := "Parameter " . argumentsAndValidation["Parameter"] . " must be an Integer"
                } else {
                    switch argumentsAndValidation["Type"] {
                        case "Byte":
                            if (argumentValue < 0 || argumentValue > 255) {
                                argumentsAndValidation["Validation"] := "Value out of byte range (0–255): " . argumentValue
                            }
                        case "Year":
                            if !RegExMatch(argumentValue, "^\d+$") {
                                argumentsAndValidation["Validation"] := "Invalid Year value: " . argumentValue . " (must be integer digits only)."
                            } else if (argumentValue < 1900 || argumentValue > 2100) {
                                argumentsAndValidation["Validation"] := "Invalid Year value: " . argumentValue . " (must be between 1900 and 2100)."
                            }
                        default:
                    }
                }
            case "String":
                if argumentsAndValidation["Optional"] = "" && argumentValue = "" {
                    argumentsAndValidation["Validation"] := "Parameter " . argumentsAndValidation["Parameter"] . " has no value passed into it."
                } else if argumentsAndValidation["Optional"] = "Optional" && argumentValue = "" {
                    ; Skip validation regardless of type as no value exists.
                }
                else if argumentsAndValidation["Pattern"] !== "" {
                    if !RegExMatch(argumentValue, argumentsAndValidation["Pattern"]) {
                        argumentsAndValidation["Validation"] := "Argument does not qualify: " . argumentValue . " (Pattern: " . argumentsAndValidation["Pattern"] ")."
                    }
                } else if argumentsAndValidation["Whitelist"].Length != 0 {
                    valueIsWhitelisted := false

                    for index, whitelistEntry in argumentsAndValidation["Whitelist"] {
                        if (StrLower(Trim(argumentValue)) = StrLower(Trim(whitelistEntry))) {
                            valueIsWhitelisted := true
                            break
                        }
                    }

                    if (valueIsWhitelisted = false) {
                        argumentsAndValidation["Validation"] := "Failed as argument not in whitelist: " . argumentValue
                    }
                } else {
                    switch argumentsAndValidation["Type"] {
                        case "Absolute Path", "Absolute Save Path":
                            isDrive := RegExMatch(argumentValue, "^[A-Za-z]:\\")
                            isUNC   := RegExMatch(argumentValue, "^\\\\{2}[^\\\/]+\\[^\\\/]+\\")

                            if !(isDrive || isUNC) {
                                argumentsAndValidation["Validation"] := "Invalid Absolute Path: " . argumentValue . " (must start with drive (C:\) or UNC (\\server\share\)."
                            } else if !FileExist(argumentValue) && argumentsAndValidation["Type"] = "Absolute Path" {
                                argumentsAndValidation["Validation"] := "Invalid Absolute Path: " . argumentValue . " (file does not exist)."
                            } else if InStr(FileExist(argumentValue), "D") {
                                argumentsAndValidation["Validation"] := "Invalid Absolute Path: " . argumentValue . " (path is a directory, expected file)."
                            }

                            SplitPath(argumentValue, &filename, &directoryPath)

                            if !symbolLedger.Has(directoryPath . "|D") {
                                symbolLedger[directoryPath . "|D"] := Map(
                                    "Symbol", SymbolLedgerAlias()
                                )

                                csvSymbolLedger :=
                                    directoryPath . "|" . 
                                    "D" . "|" . 
                                    symbolLedger[directoryPath . "|D"]["Symbol"]

                                AppendCsvLineToLog(csvSymbolLedger, "Symbol Ledger")
                            }

                            if !symbolLedger.Has(filename . "|F") {
                                symbolLedger[filename . "|F"] := Map(
                                    "Symbol", SymbolLedgerAlias()
                                )

                                csvSymbolLedger :=
                                    filename . "|" . 
                                    "F" . "|" . 
                                    symbolLedger[filename . "|F"]["Symbol"]

                                AppendCsvLineToLog(csvSymbolLedger, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[directoryPath . "|D"]["Symbol"] . "\" . symbolLedger[filename . "|F"]["Symbol"]
                        case "Base64":
                            argumentValueClean := RegExReplace(argumentValue, "\s+")

                            if !RegExMatch(argumentValueClean, "^[A-Za-z0-9+/]*={0,2}$") {
                                argumentsAndValidation["Validation"] := "Invalid Base64 content: only A–Z, a–z, 0–9, +, /, and = allowed."
                            } else if Mod(StrLen(argumentValueClean), 4) != 0 {
                                argumentsAndValidation["Validation"] := "Invalid Base64 length: must be multiple of 4."
                            } else if RegExMatch(argumentValueClean, "=[^=]") {
                                argumentsAndValidation["Validation"] := "Invalid Base64 padding: '=' can only appear at the end."
                            } else {
                                argumentValueFull := "<Base64 (Length: " . StrLen(argumentValueClean) . ")>"
                                argumentValueLog  := argumentValueFull
                            }
                        case "Code":
                            argumentValueFull := "<Code (Length: " . StrLen(argumentValue) . ", Rows: " . StrSplit(argumentValue, "`n").Length . ")>"
                            argumentValueLog  := argumentValueFull
                        case "Directory":
                            isDrive := RegExMatch(argumentValue, "^[A-Za-z]:\\")
                            isUNC   := RegExMatch(argumentValue, "^\\\\{2}[^\\\/]+\\[^\\\/]+\\")
                            if !(isDrive || isUNC) {
                                argumentsAndValidation["Validation"] := "Invalid Directory: " . argumentValue . " (must start with drive (C:\) or UNC (\\server\share\))."
                            } else if !FileExist(argumentValue) && methodName !== "EnsureDirectoryExists" {
                                argumentsAndValidation["Validation"] := "Invalid Directory: " . argumentValue . " (path does not exist)."
                            } else if !InStr(FileExist(argumentValue), "D") && methodName !== "EnsureDirectoryExists" {
                                argumentsAndValidation["Validation"] := "Invalid Directory: " . argumentValue . " (path is a file, expected directory)."
                            } else if SubStr(argumentValue, -1) != "\" {
                                argumentsAndValidation["Validation"] := "Invalid Directory: " . argumentValue . " (must end with backslash \)."
                            }

                            if !symbolLedger.Has(RTrim(argumentValue, "\") . "|D") {
                                csvsymbolLedgerLine := RegisterSymbol(argumentValue, "Directory", false)
                                AppendCsvLineToLog(csvsymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[RTrim(argumentValue, "\") . "|D"]["Symbol"]
                        case "Filename":
                            pattern := "[\\/:*?" Chr(34) "<>|]"

                            if RegExMatch(argumentValue, pattern) {
                                forbiddenList := "\ / : * ? " Chr(34) " < > |"
                                argumentsAndValidation["Validation"] := "Invalid Filename: " . argumentValue . " (contains forbidden characters " . forbiddenList . ")."
                            } else if (argumentValue = "." || argumentValue = "..") {
                                argumentsAndValidation["Validation"] := "Invalid Filename: " . argumentValue . " (reserved)."
                            }

                            if !symbolLedger.Has(argumentValue . "|F") {
                                csvsymbolLedgerLine := RegisterSymbol(argumentValue, "Filename", false)
                                AppendCsvLineToLog(csvsymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[argumentValue . "|F"]["Symbol"]
                        case "ISO Date Time":
                            if !RegExMatch(argumentValue, "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$") {
                                argumentsAndValidation["Validation"] := "Invalid ISO 8601 Date Time: " . argumentValue . " (must be YYYY-MM-DD HH:MM:SS)."
                            } else {
                                dateTimeparts := StrSplit(argumentValue, " ")
                                dateParts     := StrSplit(dateTimeparts[1], "-")
                                timeParts     := StrSplit(dateTimeparts[2], ":")

                                year   := dateParts[1] + 0
                                month  := dateParts[2] + 0
                                day    := dateParts[3] + 0
                                hour   := timeParts[1] + 0
                                minute := timeParts[2] + 0
                                second := timeParts[3] + 0

                                if argumentsAndValidation["Validation"] = "" {
                                    argumentsAndValidation["Validation"] := StrReplace(ValidateIsoDate(year, month, day, hour, minute, second), "ISO 8601 Date:", "ISO 8601 Date Time:")
                                }

                                argumentValueLog  := LocalIsoWithUtcTag(argumentValue)
                            }
                        case "ISO Date":
                            if !RegExMatch(argumentValue, "^\d{4}-\d{2}-\d{2}$") {
                                argumentsAndValidation["Validation"] := "Invalid ISO 8601 Date: " . argumentValue . " (must be YYYY-MM-DD)."
                            } else {
                                dateParts := StrSplit(argumentValue, "-")
                                year      := dateParts[1] + 0
                                month     := dateParts[2] + 0
                                day       := dateParts[3] + 0

                                if argumentsAndValidation["Validation"] = "" {
                                    argumentsAndValidation["Validation"] := ValidateIsoDate(year, month, day)
                                }
                            }
                        case "Percent Range":
                            if !RegExMatch(argumentValue, "^\d{1,3}-\d{1,3}$") {
                                argumentsAndValidation["Validation"] := "Invalid Percent Range: " . argumentValue . " (must be two integers separated by '-')."
                            } else {
                                parts  := StrSplit(argumentValue, "-")
                                first  := parts[1] + 0
                                second := parts[2] + 0

                                if (first < 0 || first > 100 || second < 0 || second > 100) {
                                    argumentsAndValidation["Validation"] := "Invalid Percent Range: " . argumentValue . " (values must be between 0 and 100)."
                                } else if (first >= second) {
                                    argumentsAndValidation["Validation"] := "Invalid Percent Range: " . argumentValue . " (first value must be lower than second)."
                                }
                            }
                        case "Raw Date Time":
                            if !RegExMatch(argumentValue, "^\d{14}$") {
                                argumentsAndValidation["Validation"] := "Invalid Raw Date Time: " . argumentValue . " (must be YYYYMMDDHHMMSS)."
                            } else {
                                year   := SubStr(argumentValue, 1, 4) + 0
                                month  := SubStr(argumentValue, 5, 2) + 0
                                day    := SubStr(argumentValue, 7, 2) + 0
                                hour   := SubStr(argumentValue, 9, 2) + 0
                                minute := SubStr(argumentValue, 11, 2) + 0
                                second := SubStr(argumentValue, 13, 2) + 0

                                if argumentsAndValidation["Validation"] = "" {
                                    argumentsAndValidation["Validation"] := ValidateIsoDate(year, month, day, hour, minute, second)
                                }
                            }
                        case "Screen Delta":
                            if !RegExMatch(argumentValue, "^(0|[1-9]\d*|-0|-[1-9]\d*)$") {
                                argumentsAndValidation["Validation"] := "Invalid Screen Delta: " . argumentValue . " (must be 0 or integer without leading zeros, optional leading '-')."
                            }
                        case "Search":
                            pattern := "[\\/:*?" Chr(34) "<>|]"

                            if RegExMatch(argumentValue, pattern) {
                                forbiddenList := "\ / : * ? " Chr(34) " < > |"
                                argumentsAndValidation["Validation"] := "Invalid Search Filename: " . argumentValue . " (contains forbidden characters " . forbiddenList . ")."
                            } else if (argumentValue = "." || argumentValue = "..") {
                                argumentsAndValidation["Validation"] := "Invalid Search Filename: " . argumentValue . " (reserved)."
                            }

                            if !symbolLedger.Has(argumentValue . "|S") {
                                csvsymbolLedgerLine := RegisterSymbol(argumentValue, "Search", false)
                                AppendCsvLineToLog(csvsymbolLedgerLine, "Symbol Ledger")
                            }

                            argumentValueLog := symbolLedger[argumentValue . "|S"]["Symbol"]
                        case "SHA-256":
                            if StrLen(argumentValue) != 64 {
                                argumentsAndValidation["Validation"] := "Invalid SHA-256 hash length: " . StrLen(argumentValue) . "."
                            } else if !RegExMatch(argumentValue, "^[0-9a-fA-F]{64}$") {
                                argumentsAndValidation["Validation"] := "Invalid SHA-256 hash content: must be hex digits only."
                            } else {
                                argumentValueLog  := EncodeSha256HexToBase80(argumentValue)
                            }
                        default:
                            if (StrLen(argumentValue) > 192) {
                                argumentValueFull := SubStr(argumentValue, 1, 224) . "…"
                                argumentValueLog  := SubStr(argumentValue, 1, 192) . "…"
                            }
                    }
                }
                
                argumentValueFull := Format('"{1}"', argumentValueFull)
                argumentValueLog  := Format('"{1}"', argumentValueLog)
            default:
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

LogInformationBeginning(overlayValue, methodName, arguments := unset, overlayCustomKey := 0) {
    static operationSequenceNumber := 0
    static lastIntermissionTick := 0
    intermissionInterval := 6 * 60 * 1000
    intermissionTick := A_TickCount

    if (lastIntermissionTick = 0) {
        lastIntermissionTick := intermissionTick
    }

    encodedOperationSequenceNumber := EncodeIntegerToBase80(operationSequenceNumber)
    encodedTickCount               := EncodeIntegerToBase80(A_TickCount)
    if overlayCustomKey = 0 {
        overlayKey                 := OverlayGenerateNextKey(methodName)
    } else {
        overlayKey := overlayCustomKey
        if methodName = "OverlayUpdateCustomLine" {
            arguments[1] := overlayKey
        }
    }

    if (intermissionTick - lastIntermissionTick >= intermissionInterval) {
        lastIntermissionTick := intermissionTick
        LogEngine("Intermission")
    }

    argumentsLineAndValidationStatus := unset
    if IsSet(arguments) {
        argumentsLineAndValidationStatus := LogFormatArgumentsAndValidate(methodName, arguments)
    }

    csvShared :=
        encodedOperationSequenceNumber .     "|" . ; Operation Sequence Number
        "B" .                                "|" . ; Status
        encodedTickCount .                   "|" . ; Tick
        methodRegistry[methodName]["Symbol"]       ; Symbol

    argumentsLine := ""
    if IsSet(argumentsLineAndValidationStatus) {
        csvShared := csvShared . "|" . argumentsLineAndValidationStatus["Arguments Log"] ; Arguments
        argumentsLine := argumentsLineAndValidationStatus["Arguments Full"]
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
        "Arguments",                 argumentsLine,
        "Overlay Key",               overlayKey,
        "Validation",                "",
        "Context",                   ""
    )

    try {
        if IsSet(argumentsLineAndValidationStatus) {
            logValuesForConclusion["Validation"] := argumentsLineAndValidationStatus["Validation"]

            if argumentsLineAndValidationStatus["Validation"] !== "" {
                throw Error(argumentsLineAndValidationStatus["Validation"])
            }
        }
    } catch as validationError {
        LogInformationConclusion("Failed", logValuesForConclusion, validationError)
    }

    operationSequenceNumber++
    return logValuesForConclusion
}

LogInformationConclusion(conclusionStatus, logValuesForConclusion, errorObject := unset) {
    encodedTickCount := EncodeIntegerToBase80(A_TickCount)

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
    static methodName := RegisterMethod("OverlayChangeTransparency(transparencyValue As Integer [Type: Byte])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Change Transparency (" . transparencyValue . ")", methodName, [transparencyValue])

    try {
        if !IsInteger(transparencyValue) || transparencyValue < 0 || transparencyValue > 255 {
            throw ValueError("Transparency must be an integer between 0 and 255.", -1, transparencyValue)
        }
    } catch as transparencyValueError {
        LogInformationConclusion("Failed", logValuesForConclusion, transparencyValueError)
    }

    WinSetTransparent(transparencyValue, "ahk_id " . overlayGui.Hwnd)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayChangeVisibility() {
    static methodName := RegisterMethod("OverlayChangeVisibility()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Change Visibility", methodName)

    if DllCall("user32\IsWindowVisible", "ptr", overlayGui.Hwnd) {
        overlayGui.Hide()
    } else {
        overlayGui.Show("NoActivate")
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayHideLogForMethod(methodNameInput) {
    static methodName := RegisterMethod("OverlayHideLogForMethod(methodNameInput As String)" . LibraryTag(A_LineFile), A_LineNumber + 1)
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
    static methodName := RegisterMethod("OverlayShowLogForMethod(methodNameInput As String)" . LibraryTag(A_LineFile), A_LineNumber + 1)
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

    if !DllCall("user32\IsWindow", "ptr", windowHandle) {
        return false
    }

    if DllCall("user32\IsWindowVisible", "ptr", windowHandle) {
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
        DllCall("dwmapi\DwmGetWindowAttribute"
            , "ptr", overlayGui.Hwnd
            , "int", 9
            , "ptr", rectBuffer
            , "int", 16),
        Map(
            "left",   NumGet(rectBuffer,  0, "int"),
            "top",    NumGet(rectBuffer,  4, "int"),
            "right",  NumGet(rectBuffer,  8, "int"),
            "bottom", NumGet(rectBuffer, 12, "int")
        )
    )

    visualRectangle := measureVisualRectangle()
    visualWidth  := visualRectangle["right"]  - visualRectangle["left"]
    visualHeight := visualRectangle["bottom"] - visualRectangle["top"]

    ; Ensure the *visual* size is even on both axes. If an axis is odd, nudge the client by +1 logical pixel on that axis and re-measure.
    adjustAttemptsForWidth  := 0
    while (Mod(visualWidth, 2) && adjustAttemptsForWidth < 6) {
        baseLogicalWidth += 1
        statusTextControl.Move(, , baseLogicalWidth, baseLogicalHeight)
        visualRectangle := measureVisualRectangle()
        visualWidth  := visualRectangle["right"]  - visualRectangle["left"]
        adjustAttemptsForWidth += 1
    }

    adjustAttemptsForHeight := 0
    while (Mod(visualHeight, 2) && adjustAttemptsForHeight < 6) {
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
    static methodName := RegisterMethod("OverlayInsertSpacer()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Overlay Insert Spacer", methodName)
    
    OverlayUpdateLine(overlayKey := OverlayGenerateNextKey(), "")
    logValuesForConclusion["Overlay Key"] := overlayKey

    LogInformationConclusion("Completed", logValuesForConclusion)
}

OverlayUpdateCustomLine(overlayKey, value) {
    static methodName := RegisterMethod("OverlayUpdateCustomLine(overlayKey As Integer, value As String)" . LibraryTag(A_LineFile), A_LineNumber + 1)
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

    if logValuesForConclusion["Method Name"] !== "OverlayUpdateCustomLine" {
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

; ******************** ;
; Helper Methods       ;
; ******************** ;

AssignBase80CharacterSet() {
    static cachedResult := Unset

    if !IsSet(cachedResult) {
        excludedAsciiCodePoints := Map(
            0x20, true, ;   SPACE
            0x22, true, ; " QUOTATION MARK
            0x27, true, ; ' APOSTROPHE
            0x2C, true, ; , COMMA
            0x2F, true, ; / SOLIDUS
            0x3A, true, ; : COLON
            0x3B, true, ; ; SEMICOLON
            0x3C, true, ; < LESS-THAN SIGN
            0x3E, true, ; > GREATER-THAN SIGN
            0x3F, true, ; ? QUESTION MARK
            0x5B, true, ; [ LEFT SQUARE BRACKET
            0x5C, true, ; \ REVERSE SOLIDUS
            0x5D, true, ; ] RIGHT SQUARE BRACKET
            0x60, true, ; ` GRAVE ACCENT
            0x7C, true  ; | VERTICAL LINE
        )

        base80Characters := ""
        ; Build ASCII printable range U+0020..U+007E, skipping exclusions.
        Loop (0x7E - 0x20 + 1) {
            codePoint := 0x20 + A_Index - 1
            if !excludedAsciiCodePoints.Has(codePoint) {
                base80Characters .= Chr(codePoint)
            }
        }

        ; Build character -> digit map (0..79).
        base80DigitByCharacterMap := Map()
        Loop StrLen(base80Characters) {
            base80Character := SubStr(base80Characters, A_Index, 1)
            base80DigitByCharacterMap[base80Character] := A_Index - 1
        }

        cachedResult := Map(
            "Characters", base80Characters,
            "Base",       StrLen(base80Characters),
            "DigitMap",   base80DigitByCharacterMap
        )
    }

    return cachedResult
}

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

DecodeBase80ToInteger(base80Text) {
    static base80Radix := 0, base80DigitByCharacterMap := ""

    if base80Radix = 0 {
        alphabetInfo              := AssignBase80CharacterSet()
        base80Radix               := alphabetInfo["Base"]
        base80DigitByCharacterMap := alphabetInfo["DigitMap"]
    }

    integerValue := 0
    Loop StrLen(base80Text) {
        base80Character := SubStr(base80Text, A_Index, 1)
        digitValue := base80DigitByCharacterMap[base80Character]
        integerValue := integerValue * base80Radix + digitValue
    }
    
    return integerValue
}

DecodeBase80ToSha256Hex(base80Text) {
    characterSetInfo := AssignBase80CharacterSet()
    base80Characters := characterSetInfo["Characters"]
    base80Radix := characterSetInfo["Base"]
    base80DigitByCharacterMap := characterSetInfo["DigitMap"]

    base80Text := Trim(base80Text)
    if StrLen(base80Text) < 1 {
        throw Error("DecodeAliasBase80ToSha256Hex: alias text must be non-empty.")
    }

    ; Initialize 32-byte big-endian integer to zero
    sha256Bytes := Buffer(32, 0)

    ; value = value * base80Radix + digit  (in place on the 32-byte buffer)
    base80Length := StrLen(base80Text)
    base80Position := 1
    while (base80Position <= base80Length) {
        base80Character := SubStr(base80Text, base80Position, 1)
        if !base80DigitByCharacterMap.Has(base80Character) {
            throw Error("DecodeAliasBase80ToSha256Hex: character not in Base80 set: " . base80Character)
        }
        digitValue := base80DigitByCharacterMap[base80Character]

        carryValue := digitValue
        byteIndex := 31
        while (byteIndex >= 0) {
            currentByte := NumGet(sha256Bytes, byteIndex, "UChar")
            productValue := currentByte * base80Radix + carryValue
            NumPut("UChar", Mod(productValue, 256), sha256Bytes, byteIndex)
            carryValue := Floor(productValue / 256)
            byteIndex -= 1
        }
        if carryValue != 0 {
            throw Error("DecodeAliasBase80ToSha256Hex: overflow beyond 32 bytes.")
        }
        base80Position += 1
    }

    ; 32 bytes → canonical 64-char lowercase hex string
    hexOutput := ""
    byteIndex := 0
    while (byteIndex < 32) {
        hexOutput .= Format("{:02x}", NumGet(sha256Bytes, byteIndex, "UChar"))
        byteIndex += 1
    }
    
    return hexOutput
}

EncodeIntegerToBase80(identifier) {
    static base80Characters := "", base80Radix := 0

    if base80Radix = 0 {
        alphabetInfo     := AssignBase80CharacterSet()
        base80Characters := alphabetInfo["Characters"]
        base80Radix        := alphabetInfo["Base"]
    }

    if identifier = 0 {
        return SubStr(base80Characters, 1, 1)
    }

    base80Text  := ""
    integerValue := identifier
    while (integerValue > 0) {
        digitValue  := Mod(integerValue, base80Radix)
        base80Text := SubStr(base80Characters, digitValue + 1, 1) . base80Text
        integerValue := Floor(integerValue / base80Radix)
    }

    return base80Text
}

EncodeSha256HexToBase80(hexSha256) {
    characterSetInfo := AssignBase80CharacterSet()
    base80Characters := characterSetInfo["Characters"]
    base80Radix := characterSetInfo["Base"]

    ; Validate and parse hex → 32-byte big-endian buffer
    hexSha256 := Trim(hexSha256)
    if !RegExMatch(hexSha256, "^[0-9A-Fa-f]{64}$") {
        return hexSha256
    }
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
    while (byteIndex < 32) {
        if NumGet(sha256Bytes, byteIndex, "UChar") {
            isAllZero := false
            break
        }
        byteIndex += 1
    }

    base80DigitsLeastSignificantFirst := []
    if isAllZero {
        base80DigitsLeastSignificantFirst.Push(0)
    } else {
        while true {
            remainderValue := 0
            hasNonZeroQuotientByte := false
            byteIndex := 0
            while (byteIndex < 32) {
                currentByte := NumGet(sha256Bytes, byteIndex, "UChar")
                accumulator := remainderValue * 256 + currentByte
                quotientByte := Floor(accumulator / base80Radix)
                remainderValue := Mod(accumulator, base80Radix)
                NumPut("UChar", quotientByte, sha256Bytes, byteIndex)
                if quotientByte != 0 {
                    hasNonZeroQuotientByte := true
                }
                byteIndex += 1
            }
            base80DigitsLeastSignificantFirst.Push(remainderValue)
            if !hasNonZeroQuotientByte {
                break
            }
        }
    }

    base80Text := ""
    digitIndex := base80DigitsLeastSignificantFirst.Length
    while (digitIndex >= 1) {
        digitValue := base80DigitsLeastSignificantFirst[digitIndex]
        base80Text .= SubStr(base80Characters, digitValue + 1, 1)
        digitIndex -= 1
    }

    ; Left-pad with the zero digit to fixed length 41
    if StrLen(base80Text) > 41 {
        return hexSha256
    }
    base80ZeroDigit := SubStr(base80Characters, 1, 1)
    while (StrLen(base80Text) < 41) {
        base80Text := base80ZeroDigit . base80Text
    }

    return base80Text
}

ExecutionLogBatchAppend(executionType, array) {
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
    ; Returns: "en-US - US-International (00020409)"
    keyboardLayoutKlid := ""
    try {
        buf := Buffer(9*2, 0)
        if DllCall("user32\GetKeyboardLayoutNameW", "ptr", buf) {
            keyboardLayoutKlid := StrGet(buf)
        }
    }
    if keyboardLayoutKlid = "" {
        ; Fallback: build KLID from HKL if needed
        try {
            hkl := DllCall("user32\GetKeyboardLayout", "uint", 0, "ptr")
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
            if DllCall("kernel32\LCIDToLocaleName", "uint", ("0x" . keyboardLayoutLanguageId) + 0, "ptr", buf2, "int", 85, "uint", 0) {
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
        try keyboardLayoutLayoutText := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Keyboard Layouts\" . keyboardLayoutKlid, "Layout Text")
        if keyboardLayoutLayoutText = "" {
            try keyboardLayoutLayoutText := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Keyboard Layouts\" . keyboardLayoutKlid, "Layout Display Name")
            ; Resolve @-style resource names when possible
            if keyboardLayoutLayoutText != "" && SubStr(keyboardLayoutLayoutText, 1, 1) = "@" {
                try {
                    buf3 := Buffer(260*2, 0)
                    if DllCall("shlwapi\SHLoadIndirectString", "wstr", keyboardLayoutLayoutText, "ptr", buf3, "int", 260, "ptr", 0) = 0 {
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

    return keyboardLayoutLocaleName . " - " . keyboardLayoutLayoutText . " (" . keyboardLayoutKlid . ")"
}

GetInputLanguage() {
    inputLanguageLocaleName := ""
    try {
        hkl := DllCall("user32\GetKeyboardLayout", "uint", 0, "ptr")
        langId := hkl & 0xFFFF
        buf := Buffer(85*2, 0)
        if DllCall("kernel32\LCIDToLocaleName", "uint", langId, "ptr", buf, "int", 85, "uint", 0) {
            inputLanguageLocaleName := StrGet(buf)
        }
    }
    if inputLanguageLocaleName = "" {
        inputLanguageLocaleName := "Unknown Language"
    }

    return inputLanguageLocaleName
}

GetOperatingSystemFamilyAndEdition() {
    currentBuildNumber := ""
    try {
        currentBuildNumber := RegRead(
        "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", 
        "CurrentBuildNumber"
    )
    }
    currentBuildNumberNumeric := currentBuildNumber + 0

    switch true {
        case (currentBuildNumberNumeric >= 22000):
            operatingSystemFamily := "Windows 11"
        case (currentBuildNumberNumeric >= 10240):
            operatingSystemFamily := "Windows 10"
        case (currentBuildNumberNumeric >= 9600):
            operatingSystemFamily := "Windows 8.1"
        case (currentBuildNumberNumeric >= 9200):
            operatingSystemFamily := "Windows 8"
        case (currentBuildNumberNumeric >= 7600):
            operatingSystemFamily := "Windows 7"
        default:
            operatingSystemFamily := "Unknown Windows"
    }

    operatingSystemEdition := ""
    try {
        operatingSystemEdition := RegRead(
        "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", 
        "EditionID"
    )
    }
    if operatingSystemEdition = "" {
        operatingSystemEdition := "Unknown Edition"
    }

    return operatingSystemFamily . " (" . operatingSystemEdition . ")"
}

GetRegionFormat() {
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
    freeMB := DriveGetSpaceFree(A_MyDocuments)
    result := freeMB . " MB (" . Format("{:.2f}", Floor((freeMB/1024/1024)*100) / 100) . " TB)"

    return result
}

LibraryTag(sourceFilePath) {
    static cacheByFilePath := Map()

    if cacheByFilePath.Has(sourceFilePath) {
        return cacheByFilePath[sourceFilePath]
    }

    SplitPath(sourceFilePath, &filenameWithExtension, &parentFolderPath, &fileExtension, &filenameWithoutExtension)

    if RegExMatch(parentFolderPath, "\((\d{4}-\d{2}-\d{2})\)", &parentFolderMatch) {
        libraryTag := " @ " . filenameWithoutExtension . " (" . parentFolderMatch[1] . ")"
    } else {
        libraryTag := " @ " . filenameWithoutExtension
    }

    cacheByFilePath[sourceFilePath] := libraryTag
    return libraryTag
}

ParseMethodDeclaration(declaration) {
    atParts     := StrSplit(declaration, "@", , 2)
    signature   := RTrim(atParts[1])
    library     := LTrim(atParts[2])

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
            if (currentCharacter = Chr(34)) {    ; Chr(34) = "
                inQuotedString := !inQuotedString
                currentParameterText .= currentCharacter
                continue
            }

            ; If not inside quotes, structural characters may affect parsing.
            if !inQuotedString {
                if (currentCharacter = "[") {
                    squareBracketDepth += 1
                    currentParameterText .= currentCharacter
                    continue
                }
                if (currentCharacter = "]" && squareBracketDepth > 0) {
                    squareBracketDepth -= 1
                    currentParameterText .= currentCharacter
                    continue
                }

                ; Top-level comma → this marks the end of one parameter.
                if (currentCharacter = "," && squareBracketDepth = 0) {
                    parameterParts.Push(Trim(currentParameterText))
                    currentParameterText := ""
                    removeLeadingSpaceAfterComma := true
                    continue
                }
            }

            ; Immediately after a delimiter comma, drop exactly one space if it exists.
            if (removeLeadingSpaceAfterComma && currentCharacter = " ") {
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
                if (metadataBlock = "") {
                    continue
                }
                blockContent := Trim(metadataBlock, "[ `t")
                if (blockContent = "") {
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

    parsedMethod := Map()
    parsedMethod[methodName] := Map(
        "Overlay Log",         "",
        "Symbol",              "",
        "Declaration",         declaration,
        "Signature",           signature,
        "Library",             library,
        "Contract",            contract,
        "Parameters",          parameters,
        "Data Types",          dataTypes,
        "Metadata",            metadata,
        "Validation Line",     "",
        "Parameter Contracts", parameterContracts
    )

    return parsedMethod
}

RegisterMethod(declaration, lineNumber) {
    parsedMethod := ParseMethodDeclaration(declaration)

    methodName := ""
    for outerKey, innerMap in parsedMethod {
        methodName := outerKey
        break
    }

    csvsymbolLedgerLine := RegisterSymbol(declaration, "M", false)
    AppendCsvLineToLog(csvsymbolLedgerLine, "Symbol Ledger")
    csvParts := StrSplit(csvsymbolLedgerLine, "|")
    symbol   := csvParts[csvParts.Length]

    global methodRegistry

    if methodRegistry.Has(methodName) {
        methodRegistry[methodName]["Declaration"]         := parsedMethod[methodName]["Declaration"]
        methodRegistry[methodName]["Symbol"]              := symbol
        methodRegistry[methodName]["Signature"]           := parsedMethod[methodName]["Signature"]
        methodRegistry[methodName]["Library"]             := parsedMethod[methodName]["Library"]
        methodRegistry[methodName]["Contract"]            := parsedMethod[methodName]["Contract"]
        methodRegistry[methodName]["Parameters"]          := parsedMethod[methodName]["Parameters"]
        methodRegistry[methodName]["Data Types"]          := parsedMethod[methodName]["Data Types"]
        methodRegistry[methodName]["Metadata"]            := parsedMethod[methodName]["Metadata"]
        methodRegistry[methodName]["Validation Line"]     := lineNumber
        methodRegistry[methodName]["Parameter Contracts"] := parsedMethod[methodName]["Parameter Contracts"]
    } else {
        methodRegistry[methodName] := Map(
            "Overlay Log",         false,
            "Symbol",              symbol,
            "Declaration",         parsedMethod[methodName]["Declaration"],
            "Signature",           parsedMethod[methodName]["Signature"],
            "Library",             parsedMethod[methodName]["Library"],
            "Contract",            parsedMethod[methodName]["Contract"],
            "Parameters",          parsedMethod[methodName]["Parameters"],
            "Data Types",          parsedMethod[methodName]["Data Types"],
            "Metadata",            parsedMethod[methodName]["Metadata"],
            "Validation Line",     lineNumber,
            "Parameter Contracts", parsedMethod[methodName]["Parameter Contracts"]
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
            directoryPath := RTrim(value, "\")
            if !symbolLedger.Has(directoryPath . "|D") {
                symbolLedger[directoryPath . "|D"] := Map(
                    "Symbol", SymbolLedgerAlias()
                )

                symbolLine :=
                    directoryPath . "|" . 
                    "D" . "|" . 
                    symbolLedger[directoryPath . "|D"]["Symbol"]
            }
        case "filename", "f":
            if !symbolLedger.Has(value . "|F") {
                symbolLedger[value . "|F"] := Map(
                    "Symbol", SymbolLedgerAlias()
                )

                symbolLine :=
                    value . "|" . 
                    "F" . "|" . 
                    symbolLedger[value . "|F"]["Symbol"]
            }
        case "method", "m":
            if !symbolLedger.Has(value . "|M") {
                symbolLedger[value . "|M"] := Map(
                    "Symbol", SymbolLedgerAlias()
                )

                symbolLine :=
                    value . "|" . 
                    "M" . "|" . 
                    symbolLedger[value . "|M"]["Symbol"]
            }
        case "search", "s":
            if !symbolLedger.Has(value . "|S") {
                symbolLedger[value . "|S"] := Map(
                    "Symbol", SymbolLedgerAlias()
                )

                symbolLine :=
                    value . "|" . 
                    "S" . "|" . 
                    symbolLedger[value . "|S"]["Symbol"]
            }
        default:
            type := "[[Invalid]]"
    }

    if type !== "[[Invalid]]" && addNewLine = true {
        symbolLine := symbolLine . newLine
    }

    return symbolLine
}

SymbolLedgerAlias() {
    static symbolLedgerIdentifier := -1
    symbolLedgerIdentifier++

    symbolLedgerAlias := EncodeIntegerToBase80(symbolLedgerIdentifier)

    return symbolLedgerAlias
}

SymbolLedgerBatchAppend(symbolType, array) {
    static newLine := "`r`n"

    switch StrLower(symbolType) {
        case "directory", "d":
            symbolType := "D"
        case "file", "f":
            symbolType := "F"
        case "search", "s":
            symbolType := "S"
        default:
            return
    }

    consolidatedSymbolLedger := ""

    arrayLength := array.Length
    for index, value in array {
        if value = "" {
            continue
        }

        if arrayLength !== index {
            consolidatedSymbolLedger := consolidatedSymbolLedger . RegisterSymbol(value, symbolType)
        } else {
            consolidatedSymbolLedger := consolidatedSymbolLedger . RegisterSymbol(value, symbolType, false)
        }
    }

    AppendCsvLineToLog(consolidatedSymbolLedger, "Symbol Ledger")
}