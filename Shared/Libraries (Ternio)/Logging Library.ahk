#Requires AutoHotkey v2.0
#Include ..\jsongo_AHKv2 (2025-02-26)\jsongo.v2.ahk
#Include Application Library.ahk
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Image Library.ahk

; Press Escape to abort the script early when running or to close the script when it's completed.
$Esc:: {
    if system["Logging"]["Cycle"] != "Stopped" {
        Critical "On"
        AbortExecution()
    } else {
        ExitApp()
    }
}

AbortExecution() {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [], "Abort Execution")

    LogConclusion("Failed", logConclusionData, A_LineNumber, "Execution aborted early by pressing escape.")
}

LogEngine(runtimeOverride := Map()) {
    global methodRegistry
    global system
    global logToFile

    static configuration := system["Configuration"]
    static constants     := system["Constants"]
    static directories   := system["Directories"]
    static environment   := system["Environment"]
    static hardware      := system["Hardware"]
    static logging       := system["Logging"]
    static mappings      := system["Mappings"]
    static paths         := system["Paths"]
    static runtime       := system["Runtime"]
    static telemetry     := system["Telemetry"]

    settings := methodRegistry["LogEngine"]["Settings"]

    fileLoggingEnabled              := settings["File Logging Enabled"]["Value"]
    telemetryDurationInMilliseconds := settings["Telemetry Duration in Milliseconds"]["Value"]

    system["Telemetry"] := TelemetryTimestamp(telemetryDurationInMilliseconds)
    telemetry           := system["Telemetry"]
    runTelemetryOrder   := IncrementCounter("Run Telemetry Order")

    operationLogLineNumber := unset

    switch logging["Cycle"] {
        case "Pending":
            constants["New Line"]     := "`r`n"
            constants["System Drive"] := SubStr(A_WinDir, 1, 3)

            operationLogLineNumber := 2

            logging["Cycle"] := "Beginning"
        case "Running":
            logging["Cycle"] := "Completed"
    }

    newLine     := constants["New Line"]
    systemDrive := constants["System Drive"]

    if logging["Cycle"] != "Beginning" && logging["Cycle"] != "Intermission" {
        if OverlayIsVisible() {
            OverlayChangeTransparency(255)
        }
    }

    if logging["Cycle"] = "Beginning" {
        ; https://learn.microsoft.com/en-us/windows/win32/sysinfo/acquiring-high-resolution-time-stamps
        queryPerformanceCounterFrequencyBuffer := Buffer(8, 0)
        DllCall("QueryPerformanceFrequency", "Ptr", queryPerformanceCounterFrequencyBuffer.Ptr, "Int")
        environment["QPC Frequency"] := NumGet(queryPerformanceCounterFrequencyBuffer, 0, "Int64")

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

        uefi := "Unified Extensible Firmware Interface "

        for baseCharacterSet in [
            94, 92, 86, 66, 62, 52
        ] {
            GetBaseCharacterSet(baseCharacterSet)
        }

        SetLogFilenames(StrSplit(telemetry["UTC Timestamp Precise"], ".")[1])

        logging["Execution Log"].Push("Log|Value")
        logging["Operation Log"].Push("Operation Sequence Number|Status|Query Performance Counter|UTC Timestamp Integer|Method or Context|Arguments or Error Message|Overlay Key|Overlay Value")
        logging["Run Telemetry"].Push("Log|Value|Type")
        logging["Symbol Ledger"].Push("Reference|Type|Symbol")

        for context in [
            "Cycle: Beginning.", "Cycle: Completed.", "Cycle: Failed.", "Cycle: Intermission."
        ] {
            RegisterSymbol(context, "Context")
        }

        for error in [
            "Execution aborted early by pressing escape."
        ] {
            RegisterSymbol(error, "Error")
        }

        for overlayValue in [
            "",
            "Initializing Variables" . overlay["Status"]["Beginning"], "Verifying Requirements" . overlay["Status"]["Beginning"], "Loading Code to Memory" . overlay["Status"]["Beginning"], "Selecting Configuration" . overlay["Status"]["Beginning"],
            "Initializing Variables" . overlay["Status"]["Completed"], "Verifying Requirements" . overlay["Status"]["Completed"], "Loading Code to Memory" . overlay["Status"]["Completed"], "Selecting Configuration" . overlay["Status"]["Completed"]
        ] {
            RegisterSymbol(overlayValue, "Overlay")
        }

        for reference in [
            "", "|",
            "<Constraint: Base64>",
            "<Data Type: Array>", "<Data Type: Map>", "<Data Type: Object>"
        ] {
            RegisterSymbol(reference, "Reference")
        }

        for whitelist in [
            "",
            "Context", "Error", "Method", "Overlay", "Reference", "Whitelist",
            "Beginning", "Completed", "Failed", "Intermission",
            "ALT", "CTRL", "CONTROL", "SHIFT", "WIN", "WINDOWS",
            "Day-Month-Year", "Month-Day-Year", "Year-Month-Day",
            "Accessed", "Created", "Modified",
            "MD2", "MD4", "MD5", "SHA-1", "SHA-256", "SHA-384", "SHA-512",
            "'", "--", "#", "%", "//", ";",
            "Left", "Middle", "Move", "Move Smooth", "Right", "Wheel Down", "Wheel Up",
            "UTF-8", "UTF-8-BOM", "UTF-16 LE BOM",
            "Append", "Append Break", "Create", "Overwrite"
        ] {
            RegisterSymbol(whitelist, "Whitelist")
        }

        paths["Project"]               := A_ScriptFullPath
        paths["Project Configuration"] := directories["Project"] . "Configuration (" . runtime["Project Name"] . ", " . "Library Release" . " " . runtime["Library Release"] . ").json"
        paths["Project Symbol Ledger"] := directories["Project"] . "Symbol Ledger (" . runtime["Project Name"] . ", " . "Library Release" . " " . runtime["Library Release"] . ").csv"
        paths["Application Library"] := directories["Libraries"] . "Application Library.ahk"
        paths["Base Library"]        := directories["Libraries"] . "Base Library.ahk"
        paths["Chrono Library"]      := directories["Libraries"] . "Chrono Library.ahk"
        paths["File Library"]        := directories["Libraries"] . "File Library.ahk"
        paths["Image Library"]       := directories["Libraries"] . "Image Library.ahk"
        paths["Logging Library"]     := directories["Libraries"] . "Logging Library.ahk"
        paths["BIP-39"]                         := directories["Constants"] . "BIP-39 (2025-09-20).csv"
        paths["EFF Dice-Generated Passphrases"] := directories["Constants"] . "EFF Dice-Generated Passphrases (2026-06-02).csv"
        paths["Excel International"]            := directories["Constants"] . "Excel International (2025-09-26).csv"
        paths["Heroes"]                         := directories["Constants"] . "Heroes (2025-09-20).csv"
        paths["Middle-earth"]                   := directories["Constants"] . "Middle-earth (2025-12-20).csv"
        paths["NATO Phonetic Alphabet"]         := directories["Constants"] . "NATO Phonetic Alphabet (2026-06-02).csv"
        paths["Resolutions"]                    := directories["Constants"] . "Resolutions (2025-09-20).csv"
        paths["Scales"]                         := directories["Constants"] . "Scales (2025-09-20).csv"
        paths["Word International"]             := directories["Constants"] . "Word International (2025-09-26).csv"
        paths["XKCD Color Survey"]              := directories["Constants"] . "XKCD Color Survey (2026-06-02).csv"
        paths["Application Executable Directory Candidates"]                            := directories["Mappings"] . "Application Executable Directory Candidates.csv"
        paths["Application Executable Versions"]                                        := directories["Mappings"] . "Application Executable Versions.csv"
        paths["Applications"]                                                           := directories["Mappings"] . "Applications.csv"
        paths["Command Line Executables"]                                               := directories["Mappings"] . "Command Line Executables.csv"
        paths["Excel Default Cell Styles"]                                              := directories["Mappings"] . "Excel Default Cell Styles.csv"
        paths["File Signatures"]                                                        := directories["Mappings"] . "File Signatures.csv"
        paths["System Management BIOS Type 17 Memory Device - Type"]                    := directories["Mappings"] . "System Management BIOS Type 17 Memory Device - Type.csv"
        paths[uefi . "Advanced Configuration and Power Interface ID Official Registry"] := directories["Mappings"] . uefi . "Advanced Configuration and Power Interface ID Official Registry.csv"
        paths[uefi . "Plug and Play ID Official Registry"]                              := directories["Mappings"] . uefi . "Plug and Play ID Official Registry.csv"
        paths[uefi . "Plug and Play ID Unofficial Registry"]                            := directories["Mappings"] . uefi . "Plug and Play ID Unofficial Registry.csv"

        projectSymbolLedgerContent := unset
        if FileExist(paths["Project Symbol Ledger"]) {
            try {
                projectSymbolLedgerContent := FileRead(paths["Project Symbol Ledger"], "UTF-8")
            } catch {
                logConclusionData["Context"] := "Cycle: " . logging["Cycle"] . ". Project Symbol Ledger found but failed to read it."
            }

            if IsSet(projectSymbolLedgerContent) {
                projectSymbolLedgerLines  := StrSplit(projectSymbolLedgerContent, newLine)
                parsedProjectSymbolLedger := []

                for line in projectSymbolLedgerLines {
                    if line = "Reference|Type|Symbol" || line = "" {
                        continue
                    }

                    positionOfReferenceDivider := InStr(line, "|", , , -2)
                    positionOfTypeDivider      := InStr(line, "|", , , -1)
                    referenceSection           := SubStr(line, 1, positionOfReferenceDivider - 1)
                    typeSection                := SubStr(line, positionOfReferenceDivider + 1, positionOfTypeDivider - positionOfReferenceDivider - 1)

                    switch typeSection {
                        case "C": typeSection := "Context"
                        case "E": typeSection := "Error"
                        case "M": typeSection := "Method"
                        case "O": typeSection := "Overlay"
                        case "R": typeSection := "Reference"
                        case "W": typeSection := "Whitelist"
                    }

                    if typeSection != "Whitelist" {
                        parsedProjectSymbolLedger.Push([referenceSection, typeSection])
                    }
                }

                for parsedEntry in parsedProjectSymbolLedger {
                    if parsedEntry[2] != "Method" {
                        RegisterSymbol(parsedEntry[1], parsedEntry[2])
                    } else {
                        positionOfLastOpeningAngle := InStr(parsedEntry[1], "<", , -1)
                        positionOfLastClosingAngle := InStr(parsedEntry[1], ">", , -1)
                        validationLineNumber       := SubStr(parsedEntry[1], positionOfLastOpeningAngle + 1, positionOfLastClosingAngle - positionOfLastOpeningAngle - 1) + 0
                        everythingExceptValidation := SubStr(parsedEntry[1], 1, positionOfLastOpeningAngle - 2)
                        positionAfterAt            := InStr(everythingExceptValidation, ") @ ", , -1) + 3
                        sourceLibrary              := SubStr(everythingExceptValidation, positionAfterAt + 1)
                        methodNameAndDeclaration   := SubStr(everythingExceptValidation, 1, positionAfterAt - 3)
                        positionOfFirstParenthesis := InStr(methodNameAndDeclaration, "(")
                        declaration                := RTrim(SubStr(methodNameAndDeclaration, positionOfFirstParenthesis + 1), ")")
                        methodName                 := SubStr(methodNameAndDeclaration, 1, positionOfFirstParenthesis - 1)

                        sourceFilePath := unset
                        switch sourceLibrary {
                            case "Application Library": sourceFilePath := paths["Application Library"]
                            case "Base Library":        sourceFilePath := paths["Base Library"]
                            case "Chrono Library":      sourceFilePath := paths["Chrono Library"]
                            case "File Library":        sourceFilePath := paths["File Library"]
                            case "Image Library":       sourceFilePath := paths["Image Library"]
                            case "Logging Library":     sourceFilePath := paths["Logging Library"]
                        }
                        
                        RegisterMethod(declaration, methodName, sourceFilePath, validationLineNumber)
                    }
                }
            }
        }
    } else {
        if logToFile {
            operationLogLineNumber := GetTextFileLineCount(paths["Operation Log"])
        } else {
            operationLogLineNumber := logging["Operation Log"].Length
        }
    }

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("runtimeOverride As Map [Optional]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, telemetry["QPC Pre Timestamp"], telemetry["UTC Timestamp File Time"], telemetry["QPC Post Timestamp"], [runtimeOverride], "Log Engine")

    if logging["Cycle"] = "Beginning" {
        RemoveDuplicatesFromArray([])

        BatchAppendExecutionLog([])
        BatchAppendOperationLog([])
        BatchAppendSymbolLedger("", [])
        BatchAppendRunTelemetry("Beginning", [])

        if runtimeOverride.Count != 0 {
            if runtimeOverride.Has("Disable File Logging") && runtimeOverride["Disable File Logging"] = true {
                methodRegistry["LogEngine"]["Settings"]["File Logging Enabled"]["Default"] := false
                methodRegistry["LogEngine"]["Settings"]["File Logging Enabled"]["Value"]   := false
                fileLoggingEnabled := false
                logToFile          := false
            }
        } else {
            if fileLoggingEnabled {
                logToFile := true
            }
        }

        if fileLoggingEnabled && logToFile {
            BatchAppendExecutionLog(logging["Execution Log"])
            BatchAppendOperationLog(logging["Operation Log"])
            BatchAppendRunTelemetry(logging["Cycle"], logging["Run Telemetry"], true)
            BatchAppendSymbolLedger("", logging["Symbol Ledger"])

            logging["Execution Log"] := []
            logging["Operation Log"] := []
            logging["Run Telemetry"] := []
            logging["Symbol Ledger"] := []
        }

        environment["Computer Name"]      := A_ComputerName
        environment["Display Resolution"] := A_ScreenWidth . "x" . A_ScreenHeight
        environment["DPI Scale"]          := Round(A_ScreenDPI / 96 * 100) . "%"
        environment["Username"]           := A_UserName

        environment["Time Zone"]            := GetTimeZone()
        environment["Session Startup Time"] := GetSessionStartupTime()
        environment["Operating System"]     := GetOperatingSystem()

        for directory in [
            directories["Log"],
            directories["Project"]
        ] {
            if !DirExist(directory) {
                try {
                    DirCreate(directory)
                } catch as directoryCreationError {
                    LogConclusion("Failure", logConclusionData, directoryCreationError.Line, directoryCreationError.Message)
                }
            }
        }
    }

    telemetry["System Disk Space Snapshot"] := GetDriveSpaceSnapshot(systemDrive)
    telemetry["System Resource Snapshot"]   := GetSystemResourceSnapshot()
    telemetry["Computer Uptime in Seconds"] := Round(telemetry["QPC Midpoint Timestamp"] / environment["QPC Frequency"])
    telemetry["Session Uptime in Seconds"]  := DateDiff(SubStr(telemetry["UTC Timestamp Integer"], 1, 14) . "", environment["Session Startup Time"], "Seconds")

    for runTelemetryLine in [
        "Run Telemetry Order|" .           runTelemetryOrder,
        "Operation Log Line Number|" .     operationLogLineNumber,
        "Duration in Milliseconds|" .      telemetry["Duration in Milliseconds"],
        "Number of Readings|" .            telemetry["Number of Readings"],
        "QPC Pre Timestamp|" .             telemetry["QPC Pre Timestamp"] - telemetry["QPC Midpoint Timestamp"],
        "QPC Post Timestamp|" .            telemetry["QPC Post Timestamp"] - telemetry["QPC Midpoint Timestamp"],
        "QPC Midpoint Timestamp|" .        telemetry["QPC Midpoint Timestamp"],
        "Tick Count|" .                    telemetry["Tick Count"],
        "Tick Count Pre|" .                telemetry["Tick Count Pre"] - telemetry["Tick Count"],
        "UTC Timestamp Precise|" .         telemetry["UTC Timestamp Precise"],
        "Computer Uptime in Seconds|" .    telemetry["Computer Uptime in Seconds"],
        "Session Uptime in Seconds|" .     telemetry["Session Uptime in Seconds"],
        "Commit Total Bytes|" .            telemetry["System Resource Snapshot"]["Commit Total Bytes"],
        "Commit Limit Bytes" .             telemetry["System Resource Snapshot"]["Commit Limit Bytes"],
        "Commit Peak Bytes|" .             telemetry["System Resource Snapshot"]["Commit Peak Bytes"],
        "Commit Used Percent|" .           telemetry["System Resource Snapshot"]["Commit Used Percent"],
        "Kernel Total Bytes|" .            telemetry["System Resource Snapshot"]["Kernel Total Bytes"],
        "Kernel Paged Bytes|" .            telemetry["System Resource Snapshot"]["Kernel Paged Bytes"],
        "Kernel Nonpaged Bytes|" .         telemetry["System Resource Snapshot"]["Kernel Nonpaged Bytes"],
        "Physical Used Bytes|" .           telemetry["System Resource Snapshot"]["Physical Used Bytes"],
        "Physical Total Bytes|" .          telemetry["System Resource Snapshot"]["Physical Total Bytes"],
        "Physical Used Percent|" .         telemetry["System Resource Snapshot"]["Physical Used Percent"],
        "System Cache Bytes|" .            telemetry["System Resource Snapshot"]["System Cache Bytes"],
        "System Handle Count|" .           telemetry["System Resource Snapshot"]["System Handle Count"],
        "System Process Count|" .          telemetry["System Resource Snapshot"]["System Process Count"],
        "System Thread Count|" .           telemetry["System Resource Snapshot"]["System Thread Count"],
        "System Disk Free Bytes|" .        telemetry["System Disk Space Snapshot"]["Free Bytes"],
        "System Disk Windows Free Size|" . telemetry["System Disk Space Snapshot"]["Windows Free Size"]
    ] {
        logging["Run Telemetry"].Push(runTelemetryLine)
    }

    if logging["Cycle"] = "Beginning" {
        CombineCode("Intro", "Main")
        ComputeMouseMoveSpeed("0x0", "2x2")
        ConvertArrayToLineSeparatedString(["1st Line", "2nd Line"])
        ConvertHexStringToBase64("48656c6c6f20576f726c6421")
        GetBase64FromFile(paths["Scales"])
        KeyboardShortcut("CTRL", "F16")
        ModifyScreenCoordinates(2, 2, "0x0")

        ConvertIntegerToUtcTimestamp(telemetry["UTC Timestamp Integer"])
        ConvertUtcTimestampToInteger(telemetry["UTC Timestamp Precise"])
        ExtractTrailingDateAsIso("(01.01.2000)", "Month-Day-Year")
        GetDirectoryTimeAsUtc(directories["Constants"], "Created")
        GetFileTimeAsUtc(paths["Scales"], "Created")

        DetermineWindowsBinaryType("C:\Windows\System32\find.exe")
        GetFileHash(paths["Scales"], "SHA-256")
        GetFilesFromDirectory(directories["Constants"])
        GetFoldersFromDirectory(directories["Constants"])
        GetPathComponents(paths["Scales"])
        GetTextFileLineCount(paths["Scales"])
        ReadFile(paths["Scales"])
        SearchForUniqueFileInDirectory("Scales (2025-09-20)", directories["Constants"], "csv")

        OverlayIsVisible()

        ValidateDataUsingSpecification("v0.39, 2024-02-16", "String", "Spreadsheet Operations Template")

        constantValues := Map(
            "BIP-39",                         "bdeca5734c5c8ca4a1adb2b5863c0cd46ac74837f24321235b5b7b1b32879229",
            "EFF Dice-Generated Passphrases", "63d2175db6fb24702e49fbd72d339c4d8bd50c5a37804cbfc666e0ed04e843bf",
            "Excel International",            "f22a6b4c3a81f479bb7844429d5effff494023ae29fdd414bed848d54143f0f0",
            "Heroes",                         "221c6504b42787aff09b43cb85a93511e3e4c06f52c084694119637c6794817d",
            "Middle-earth",                   "ffc72a6b738fdd75ea16964e6d43695c843ef2dea986d173196795e7d11d5dbd",
            "NATO Phonetic Alphabet",         "4222037720c26e12cffba2514436bc4b5029cdc3b3ccaa34f827415e8d46bbcf",
            "Resolutions",                    "cc45d04bc98d76c9aa8ceb1e455c21082dfd8e6695c84b5382464bee2cd20364",
            "Scales",                         "91eb6122786767eb83c7d87c43610fb87018d20ef2c25e43d3d38f31f49ec18d",
            "Word International",             "d586eccccd709b85ebabbcd09a339a828fc46945df05e680c6ca52403dae8755",
            "XKCD Color Survey",              "b4e194b06581c27bebaada8375a3dffa88e12cf815841574a614cd2249bcef87"
        )

        for constant in constantValues {
            RegisterSymbol(paths[constant], "Reference")
        }

        for constant, hashValue in constantValues {
            RegisterSymbol(hashValue, "Reference")
        }

        for constant, hashValue in constantValues {
            if constant = "Excel International" || constant = "Word International" {
                continue
            }

            content := ReadFileOnHashMatch(paths[constant], hashValue)
            constants[constant] := ParseDelimitedRowsToArrayOfMaps(content)
        }

        for index, rowMap in system["Constants"]["BIP-39"] {
            rowMap["Counter"] := index
        }

        for rowMap in system["Constants"]["EFF Dice-Generated Passphrases"] {
            rowMap["Dice Sequence"] := rowMap["Dice Sequence"] + 0
        }

        for index, rowMap in system["Constants"]["Resolutions"] {
            rowMap["Counter"] := index

            parts  := StrSplit(rowMap["Resolution"], "x")
            width  := parts[1]
            height := parts[2]

            rowMap["Total Pixel Count"] := width * height
        }

        for index, rowMap in system["Constants"]["Scales"] {
            rowMap["Counter"] := index
        }

        constants["Resolution Counters"] := Map()
        for resolution in constants["Resolutions"] {
            constants["Resolution Counters"][resolution["Resolution"]] := resolution["Counter"]
        }

        constants["Scale Counters"] := Map()
        for scale in constants["Scales"] {
            constants["Scale Counters"][scale["Scale"]] := scale["Counter"]
        }

        if !constants["Scale Counters"].Has(environment["DPI Scale"]) {
            dpiScalePercent := RTrim(environment["DPI Scale"], "%") + 0
            scalePercent    := constants["Scale Counters"].Clone()
            for outerKey, innerValue in scalePercent {
                scalePercent[outerKey] := RTrim(outerKey, "%") + 0
            }

            closestScaleValue                 := 0
            minimumDifferenceBetweenDpiValues := 400

            for outerKey, innerValue in scalePercent {
                currentDpiScaleValue := innerValue

                differenceBetweenCurrentAndTarget := Abs(currentDpiScaleValue - dpiScalePercent)

                if closestScaleValue = 0 || differenceBetweenCurrentAndTarget < minimumDifferenceBetweenDpiValues || (differenceBetweenCurrentAndTarget = minimumDifferenceBetweenDpiValues && currentDpiScaleValue > closestScaleValue) {
                    minimumDifferenceBetweenDpiValues := differenceBetweenCurrentAndTarget
                    closestScaleValue                 := currentDpiScaleValue
                }
            }

            LogConclusion("Error", logConclusionData, A_LineNumber, "DPI Scale currently set to " . environment["DPI Scale"] . " which is invalid. Closest supported value is " . closestScaleValue . "%.")
        }

        if !constants["Resolution Counters"].Has(environment["Display Resolution"]) {
            displayResolutionParts      := StrSplit(environment["Display Resolution"], "x")
            displayResolutionPixelCount := displayResolutionParts[1] * displayResolutionParts[2]
            resolutionPixelCount        := constants["Resolution Counters"].Clone()
            for outerKey, innerValue in resolutionPixelCount {
                resolutionParts := StrSplit(outerKey, "x")
                resolutionPixelCount[outerKey] := resolutionParts[1] * resolutionParts[2]
            }

            closestResolutionKey := ""
            smallestDifference   := -1
            
            for resolutionKey, knownPixelCount in resolutionPixelCount {
                currentDifference := Abs(displayResolutionPixelCount - knownPixelCount)
                
                if currentDifference < smallestDifference {
                    smallestDifference   := currentDifference
                    closestResolutionKey := resolutionKey
                }
            }

            LogConclusion("Error", logConclusionData, A_LineNumber, "Display Resolution currently set to " . environment["Display Resolution"] . " which is invalid. Closest supported value is " . closestResolutionKey . ".")
        }
        
        environment["Display Resolution Counter"] := constants["Resolution Counters"][environment["Display Resolution"]]
        environment["DPI Scale Counter"]          := constants["Scale Counters"][environment["DPI Scale"]]

        mappingValues := [
            "Application Executable Directory Candidates",
            "Applications",
            "Command Line Executables",
            "File Signatures"
        ]

        for mapping in mappingValues {
            RegisterSymbol(paths[mapping], "Reference")
        }

        for mapping in mappingValues {
            fileHash := GetFileHash(paths[mapping], "SHA-256")
            content  := ReadFileOnHashMatch(paths[mapping], fileHash)
            mappings[mapping] := ParseDelimitedRowsToArrayOfMaps(content)
        }

        applicationsWithSharedImageLibraryData := GetFilesFromDirectory(directories["Images"], "Image Library Data (")
        for imageLibraryDataFile in applicationsWithSharedImageLibraryData {
            applicationName := GetPathComponents(imageLibraryDataFile)["Filename No Extension"]
            applicationName := SubStr(applicationName, StrLen("Image Library Data (") + 1)
            applicationName := SubStr(applicationName, 1, -1)

            for application in mappings["Applications"] {
                if application["Name"] = applicationName {
                    application["Shared Images"] := true
                    break
                }
            }
        }

        for application in mappings["Applications"] {
            application["Counter"] := application["Counter"] + 0
        }

        ValidateDataUsingSpecification("Excel", "String", "Application Name")

        for applicationExecutableDirectoryCandidate in mappings["Application Executable Directory Candidates"] {
            applicationExecutableDirectoryCandidate["Source"] := "Shared"
        }

        for commandLineExecutable in mappings["Command Line Executables"] {
            for application in mappings["Applications"] {
                if application["Name"] = commandLineExecutable["Name"] {
                    application["Command Line Executable"] := commandLineExecutable["Command Line Executable"]
                }
            }
        }

        for fileSignature in mappings["File Signatures"] {
            fileSignature["Maximum Base64 Signature"] := ConvertHexStringToBase64(fileSignature["Maximum Hex Signature"])
            fileSignature["Minimal Base64 Signature"] := ConvertHexStringToBase64(fileSignature["Minimal Hex Signature"])
        }

        if !FileExist(paths["Project Configuration"]) {
            defaultConfiguration := StrReplace(
            '{' . newLine . 
                '    "Application Executable Directory Candidates": [' . newLine .
                    '        '  . newLine . 
                '    ],' . newLine . 
                '    "Application Whitelist": [' . newLine . 
                    '        ' .  newLine . 
                '    ],' . newLine . 
                '    "Candidate Base Directories": [' . newLine . 
                    '        "' . systemDrive . 'Portable Files\' . '",' . newLine . 
                    '        "' . systemDrive . 'Program Files (Portable)\' . '"' . newLine . 
                    '    ],' . newLine . 
                '    "Settings": {' . newLine . 
                    '        "Advanced Mode": ' . 'false' . ',' . newLine . 
                    '        "Application Image Override Directory": "' . "" . '",' . newline . 
                    '        "Computer Alias": "' . "" . '",' . newline . 
                    '        "Image Variant Preset": "' . 'NATO Phonetic Alphabet' . '"' . newLine . 
                '    }' . newLine . 
            '}', "\", "\\")
            WriteTextToFile(defaultConfiguration, paths["Project Configuration"], "UTF-8", "Create")
        }

        jsongo.silent_error := false
        try {
            system["Configuration"] := jsongo.Parse(FileRead(paths["Project Configuration"]))
            configuration           := system["Configuration"]
        } catch as invalidJsonError {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to load Configuration File. " . StrReplace(invalidJsonError.Message, "`n", " "))
        }

        ValidateConfiguration(configuration)

        SetLogFilenames(StrSplit(telemetry["UTC Timestamp Precise"], ".")[1])

        if configuration["Application Whitelist"].Length != 0 {
            for application in mappings["Applications"] {
                for applicationWhitelist in configuration["Application Whitelist"] {
                    if application["Name"] = applicationWhitelist {
                        application["Whitelisted"] := true
                        break
                    }

                    application["Whitelisted"] := false
                }
            }
        } else {
            for application in mappings["Applications"] {
                application["Whitelisted"] := true
            }
        }

        for applicationExecutableDirectoryCandidate in configuration["Application Executable Directory Candidates"] {
            for application in mappings["Applications"] {
                if application["Name"] = applicationExecutableDirectoryCandidate[1] {
                    if application["Whitelisted"] {
                        mappings["Application Executable Directory Candidates"].Push(Map(
                            "Directory",  applicationExecutableDirectoryCandidate[3],
                            "Executable", applicationExecutableDirectoryCandidate[2],
                            "Name",       applicationExecutableDirectoryCandidate[1],
                            "Source",    "Project"
                        ))
                    }
                }
            }
        }

        mappings["Candidate Base Directories"] := []
        for candidateBaseDirectory in [
            EnvGet("LOCALAPPDATA"), EnvGet("LOCALAPPDATA") . "\Programs", EnvGet("ProgramData"), EnvGet("ProgramFiles"), EnvGet("ProgramFiles(x86)"), EnvGet("ProgramW6432"), EnvGet("SystemDrive"), EnvGet("USERPROFILE")
        ] {
            if DirExist(candidateBaseDirectory) {
                if SubStr(candidateBaseDirectory, -1) = "\" {
                    mappings["Candidate Base Directories"].Push(candidateBaseDirectory)
                } else {
                    mappings["Candidate Base Directories"].Push(candidateBaseDirectory . "\")
                }
            }
        }

        for configurationCandidateBaseDirectory in system["Configuration"]["Candidate Base Directories"] {
            if DirExist(configurationCandidateBaseDirectory) {
                mappings["Candidate Base Directories"].Push(configurationCandidateBaseDirectory)
            }
        }

        mappings["Candidate Base Directories"] := RemoveDuplicatesFromArray(mappings["Candidate Base Directories"])

        configuration["Image Variant Preset"] := Map()
        for index, name in system["Constants"][system["Configuration"]["Settings"]["Image Variant Preset"]] {
            variantName := unset
            if system["Configuration"]["Settings"]["Image Variant Preset"] = "NATO Phonetic Alphabet" {
                variantName := name["Code Word"]
            } else {
                variantName := name["Name"]
            }

            firstCharacter := StrLower(SubStr(variantName, 1, 1))

            configuration["Image Variant Preset"][firstCharacter] := variantName

            if index = 16 {
                break
            }
        }

        if configuration["Settings"]["Advanced Mode"] {
            advancedMappingValues := [
                "System Management BIOS Type 17 Memory Device - Type",
                "Unified Extensible Firmware Interface Advanced Configuration and Power Interface ID Official Registry",
                "Unified Extensible Firmware Interface Plug and Play ID Official Registry",
                "Unified Extensible Firmware Interface Plug and Play ID Unofficial Registry"
            ]

            for advancedMapping in advancedMappingValues {
                RegisterSymbol(paths[advancedMapping], "Reference")
            }

            for advancedMapping in advancedMappingValues {
                fileHash := GetFileHash(paths[advancedMapping], "SHA-256")
                content  := ReadFileOnHashMatch(paths[advancedMapping], fileHash)
                mappings[advancedMapping] := ParseDelimitedRowsToArrayOfMaps(content)
            }

            for type17MemoryDeviceType in mappings["System Management BIOS Type 17 Memory Device - Type"] {
                type17MemoryDeviceType["Value"] := type17MemoryDeviceType["Value"] + 0
            }

            mappings["Unified Extensible Firmware Interface Plug and Play ID Curated Registry"] := Map()
            for manufacturer in mappings["Unified Extensible Firmware Interface Plug and Play ID Official Registry"] {
                mappings["Unified Extensible Firmware Interface Plug and Play ID Curated Registry"][manufacturer["Vendor ID"]] := manufacturer["Vendor Name"]
            }

            for manufacturer in mappings["Unified Extensible Firmware Interface Plug and Play ID Unofficial Registry"] {
                mappings["Unified Extensible Firmware Interface Plug and Play ID Curated Registry"][manufacturer["Vendor ID"]] := manufacturer["Vendor Name"]
            }
        }

        environment["Color Mode"]          := GetColorMode()
        environment["Display Language"]    := GetDisplayLanguage()
        environment["Input Language"]      := GetInputLanguage()
        environment["International"]       := GetInternationalSnapshot()
        environment["Keyboard Layout"]     := GetActiveKeyboardLayout()
        environment["Refresh Rate"]        := GetActiveMonitorRefreshRateHz()
        environment["Regional Format"]     := environment["International"]["LocaleName"]
        environment["Timeout Before Lock"] := GetTimeoutBeforeLockInSeconds()

        hardware["Monitor Count"] := MonitorGetCount()

        if configuration["Settings"]["Advanced Mode"] {
            environment["Operating System"]["Installation Date"] := GetWindowsInstallationDateUtcTimestamp()

            hardware["CPU"]                  := GetCpu()
            hardware["Display GPU"]          := GetActiveDisplayGpu()
            hardware["Memory Size and Type"] := GetMemorySizeAndType()
            hardware["Monitor"]              := GetActiveMonitor()
            hardware["Motherboard"]          := GetMotherboard()
            hardware["Motherboard"]["BIOS"]  := GetBios()
            hardware["System Disk"]          := GetDiskModel(systemDrive)
        }

        runtime["Project Hash"]             := GetFileHash(paths["Project"],             "SHA-256")
        runtime["Application Library Hash"] := GetFileHash(paths["Application Library"], "SHA-256")
        runtime["Base Library Hash"]        := GetFileHash(paths["Base Library"],        "SHA-256")
        runtime["Chrono Library Hash"]      := GetFileHash(paths["Chrono Library"],      "SHA-256")
        runtime["File Library Hash"]        := GetFileHash(paths["File Library"],        "SHA-256")
        runtime["Image Library Hash"]       := GetFileHash(paths["Image Library"],       "SHA-256")
        runtime["Logging Library Hash"]     := GetFileHash(paths["Logging Library"],     "SHA-256")

        for executionLogLine in [
            "Project Name|" .             runtime["Project Name"],
            "Library Release|" .          runtime["Library Release"],
            "AutoHotkey Version|" .       runtime["AutoHotkey Version"],
            "Project Hash|" .             EncodeSha256HexToBase(runtime["Project Hash"], 86),
            "Application Library Hash|" . EncodeSha256HexToBase(runtime["Application Library Hash"], 86),
            "Base Library Hash|" .        EncodeSha256HexToBase(runtime["Base Library Hash"], 86),
            "Chrono Library Hash|" .      EncodeSha256HexToBase(runtime["Chrono Library Hash"], 86),
            "File Library Hash|" .        EncodeSha256HexToBase(runtime["File Library Hash"], 86),
            "Image Library Hash|" .       EncodeSha256HexToBase(runtime["Image Library Hash"], 86),
            "Logging Library Hash|" .     EncodeSha256HexToBase(runtime["Logging Library Hash"], 86),
            "Operating System|" .         environment["Operating System"]["Full Name"]
        ] {
            logging["Execution Log"].Push(executionLogLine)
        }

        if configuration["Settings"]["Advanced Mode"] {
            logging["Execution Log"].Push("Windows Installation Date|" . environment["Operating System"]["Installation Date"])
        }

        for executionLogLine in [
            "Computer Name|" .          environment["Computer Name"],
            "Computer Alias|" .         configuration["Settings"]["Computer Alias"],
            "Username|" .               environment["Username"],
            "Time Zone Key Name|" .     environment["Time Zone"]["Key Name"],
            "Time Zone UTC Offset|" .   environment["Time Zone"]["UTC Offset"],
            "QPC Frequency|" .          environment["QPC Frequency"],
            "Country or Region|" .      environment["International"]["Geo"]["Friendly Name"],
            "Display Language|" .       environment["Display Language"],
            "Regional Format|" .        environment["Regional Format"],
            "Input Language|" .         environment["Input Language"],
            "Keyboard Layout|" .        environment["Keyboard Layout"]
        ] {
            logging["Execution Log"].Push(executionLogLine)
        }

        if configuration["Settings"]["Advanced Mode"] {
            for executionLogLine in [
                "Motherboard|" .                    hardware["Motherboard"]["Full Name"],
                "BIOS|" .                           hardware["Motherboard"]["BIOS"],
                "CPU|" .                            hardware["CPU"],
                "Memory Size and Type|" .           hardware["Memory Size and Type"],
                "System Disk|" .                    hardware["System Disk"],
                "System Disk Total Bytes|" .        telemetry["System Disk Space Snapshot"]["Total Bytes"],
                "System Disk Windows Total Size|" . telemetry["System Disk Space Snapshot"]["Windows Total Size"],
                "Display GPU|" .                    hardware["Display GPU"],
                "Monitor|" .                        hardware["Monitor"]
            ] {
                logging["Execution Log"].Push(executionLogLine)
            }
        }

        for executionLogLine in [
            "Monitor Count|" .       hardware["Monitor Count"],
            "Display Resolution|" .  environment["Display Resolution"],
            "Refresh Rate|" .        environment["Refresh Rate"],
            "DPI Scale|" .           environment["DPI Scale"],
            "Color Mode|" .          environment["Color Mode"],
            "Timeout Before Lock|" . environment["Timeout Before Lock"]
        ] {
            logging["Execution Log"].Push(executionLogLine)
        }

        if fileLoggingEnabled && logToFile {
            BatchAppendExecutionLog(logging["Execution Log"])

            logging["Execution Log"] := []
        }
    } else {
        if fileLoggingEnabled && logToFile {
            BatchAppendRunTelemetry(logging["Cycle"], logging["Run Telemetry"])

            logging["Run Telemetry"] := []
        }
    }

    if !logConclusionData.Has("Context") {
        logConclusionData["Context"] := "Cycle: " . logging["Cycle"] . "."
    }

    LogConclusion("Completed", logConclusionData)

    if logging["Cycle"] != "Beginning" && logging["Cycle"] != "Intermission" {
        timestampNow := A_Now

        if fileLoggingEnabled = false || logToFile = false {
            if runtimeOverride.Count != 0 {
                if runtimeOverride.Has("Enable File Logging") && runtimeOverride["Enable File Logging"] = true {
                    methodRegistry["LogEngine"]["Settings"]["File Logging Enabled"]["Default"] := true
                    methodRegistry["LogEngine"]["Settings"]["File Logging Enabled"]["Value"]   := true
                    fileLoggingEnabled := true
                    logToFile          := true
                }

                BatchAppendExecutionLog(logging["Execution Log"])
                BatchAppendOperationLog(logging["Operation Log"])
                BatchAppendRunTelemetry(logging["Cycle"], logging["Run Telemetry"], true)
                BatchAppendSymbolLedger("", logging["Symbol Ledger"])

                logging["Execution Log"] := []
                logging["Operation Log"] := []
                logging["Run Telemetry"] := []
                logging["Symbol Ledger"] := []
            }
        }

        for logFilePath in [
            paths["Execution Log"],
            paths["Operation Log"],
            paths["Run Telemetry"],
            paths["Symbol Ledger"],
        ] {
            if !FileExist(logFilePath) {
                continue
            }

            logFile := FileOpen(logFilePath, "rw")

            logFileSize := logFile.Length
            if logFileSize = 0 {
                logFile.Close()
                continue
            }
            
            bytesToInspect := (logFileSize >= 2) ? 2 : 1
            logFile.Seek(-bytesToInspect, 2)
            trailingBytesBuffer := Buffer(bytesToInspect)
            logFile.RawRead(trailingBytesBuffer, bytesToInspect)

            if bytesToInspect = 2 && NumGet(trailingBytesBuffer, 0, "UChar") = 13 && NumGet(trailingBytesBuffer, 1, "UChar") = 10 {
                logFile.Length := logFileSize - 2
            } else if NumGet(trailingBytesBuffer, bytesToInspect - 1, "UChar") = 10 {
                logFile.Length := logFileSize - 1
            }

            logFile.Close()

            FileSetTime(timestampNow, logFilePath, "M")
        }

        logging["Cycle"] := "Stopped"
    }

    if logging["Cycle"] = "Beginning" || logging["Cycle"] = "Intermission" {
        logging["Cycle"] := "Running"
    }
}

OverlayChangeTransparency(transparencyValue) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("transparencyValue As Integer [Constraint: Byte]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [transparencyValue], "Overlay Change Transparency (" . transparencyValue . ")")

    WinSetTransparent(transparencyValue, "ahk_id " . overlay["GUI"].Hwnd)

    LogConclusion("Completed", logConclusionData)
}

OverlayChangeVisibility() {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [], "Overlay Change Visibility")

    if DllCall("User32\IsWindowVisible", "Ptr", overlay["GUI"].Hwnd) {
        overlay["GUI"].Hide()
    } else {
        overlay["GUI"].Show("NoActivate")
    }

    LogConclusion("Completed", logConclusionData)
}

OverlayHideLogForMethod(methodNameInput) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("methodNameInput As String", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [methodNameInput], "Overlay Hide Log for Method (" . methodNameInput . ")")

    global methodRegistry
    
    if !methodRegistry.Has(methodNameInput) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, 'Method "' . methodNameInput . '" not registered.')
    }

    methodRegistry[methodNameInput]["Overlay Log"] := false

    LogConclusion("Completed", logConclusionData)
}

OverlayShowLogForMethod(methodNameInput) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("methodNameInput As String", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [methodNameInput], "Overlay Show Log for Method (" . methodNameInput . ")")

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
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("", methodName, A_LineFile, A_LineNumber + 6, Map(
            "Base Logical Width", Map("Default", 960, "Floor", 640, "Ceiling", 7680),
            "Base Logical Height", Map("Default", 920, "Floor", 480, "Ceiling", 4320),
            "Overlay Transparency", Map("Default", 172, "Floor", 0, "Ceiling", 255),
            "Font Size", Map("Default", 10, "Floor", 6, "Ceiling", 24)))
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [], "Overlay Start")

    global overlay

    settings := methodRegistry[methodName]["Settings"]

    baseLogicalWidth    := settings["Base Logical Width"]["Value"]
    baseLogicalHeight   := settings["Base Logical Height"]["Value"]
    overlayTransparency := settings["Overlay Transparency"]["Value"]
    fontSize            := settings["Font Size"]["Value"]

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
            "Left",   NumGet(rectBuffer,  0, "Int"),
            "Top",    NumGet(rectBuffer,  4, "Int"),
            "Right",  NumGet(rectBuffer,  8, "Int"),
            "Bottom", NumGet(rectBuffer, 12, "Int")
        )
    )

    visualRectangle := measureVisualRectangle()
    visualWidth     := visualRectangle["Right"]  - visualRectangle["Left"]
    visualHeight    := visualRectangle["Bottom"] - visualRectangle["Top"]

    ; Ensure the *visual* size is even on both axes. If an axis is odd, nudge the client by +1 logical pixel on that axis and re-measure.
    adjustAttemptsForWidth := 0
    while Mod(visualWidth, 2) && adjustAttemptsForWidth < 6 {
        baseLogicalWidth += 1
        statusTextControl.Move(, , baseLogicalWidth, baseLogicalHeight)
        visualRectangle := measureVisualRectangle()
        visualWidth     := visualRectangle["Right"] - visualRectangle["Left"]
        adjustAttemptsForWidth += 1
    }

    adjustAttemptsForHeight := 0
    while Mod(visualHeight, 2) && adjustAttemptsForHeight < 6 {
        baseLogicalHeight += 1
        statusTextControl.Move(, , baseLogicalWidth, baseLogicalHeight)
        visualRectangle := measureVisualRectangle()
        visualHeight    := visualRectangle["Bottom"] - visualRectangle["Top"]
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
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [], "Overlay Insert Spacer", true)
    
    ; Method has Custom Overlay Rules: Executed directly in LogBeginning.

    LogConclusion("Completed", logConclusionData)
}

OverlayUpdateCustomLine(overlayKey, overlayValue) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("overlayKey As Integer, value As String", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [overlayKey, overlayValue], "Overlay Update Custom Line", true)

    ; Method has Custom Overlay Rules: Executed directly in LogBeginning.

    LogConclusion("Completed", logConclusionData)
}

; **************************** ;
; Core Methods                 ;
; **************************** ;

LogBeginning(methodName, qpcPre, timestamp, qpcPost, arguments := [], overlayValue := unset, customOverlayMethod := unset) {
    static lastRunTelemetryTickCount := unset
    static runTelemetryInterval      := 12 * 60 * 1000
    static fileTimeBuffer   := Buffer(8, 0)
    static systemTimeBuffer := Buffer(16, 0)

    newLine := system["Constants"]["New Line"]

    runTelemetryTickCount := DllCall("Kernel32\GetTickCount64", "UInt64")
    if !IsSet(lastRunTelemetryTickCount) {
        lastRunTelemetryTickCount := runTelemetryTickCount
    }

    logConclusionData := Map(
        "Method Name", methodName
    )

    operationSequenceNumber := IncrementCounter("Operation Sequence Number")
    if IsSet(overlayValue) {
        NumPut("UInt64", timestamp, fileTimeBuffer, 0)
        DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")

        year        := NumGet(systemTimeBuffer,  0, "UShort")
        month       := NumGet(systemTimeBuffer,  2, "UShort")
        day         := NumGet(systemTimeBuffer,  6, "UShort")
        hour        := NumGet(systemTimeBuffer,  8, "UShort")
        minute      := NumGet(systemTimeBuffer, 10, "UShort")
        second      := NumGet(systemTimeBuffer, 12, "UShort")
        millisecond := NumGet(systemTimeBuffer, 14, "UShort")

        qpcMidpointTimestampDelta := (qpcPre + (qpcPost - qpcPre) // 2) - system["Telemetry"]["QPC Midpoint Timestamp"]
        utcTimestampIntegerDelta  := Format("{:04}{:02}{:02}{:02}{:02}{:02}{:03}", year, month, day, hour, minute, second, millisecond) + 0 - system["Telemetry"]["UTC Timestamp Integer"]

        logConclusionData["Operation Sequence Number"] := EncodeIntegerToBase(operationSequenceNumber, 94)
        logConclusionData["Query Performance Counter"] := EncodeIntegerToBase(qpcMidpointTimestampDelta, 94)
        logConclusionData["UTC Timestamp Integer"]     := utcTimestampIntegerDelta
    } else {
        logConclusionData["Operation Sequence Number"] := operationSequenceNumber
        logConclusionData["QPC Pre"]                   := qpcPre
        logConclusionData["Timestamp"]                 := timestamp
        logConclusionData["QPC Post"]                  := qpcPost
    }

    logBeginning := unset
    overlayKey   := -1
    if IsSet(overlayValue) {
        if !IsSet(customOverlayMethod) {
            if methodRegistry[methodName].Has("Overlay Log") {
                if methodRegistry[methodName]["Overlay Log"] {
                    overlayKey := AssignNewOverlayKey()
                } else {
                    overlayKey := 0
                }
            }

            if overlayKey >= 1 {
                OverlayUpdateLine(overlayKey, overlayValue . overlay["Status"]["Beginning"])
            }
        } else {
            if methodName = "OverlayInsertSpacer" {
                overlayKey := AssignNewOverlayKey()
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
            if !IsSet(customOverlayMethod) {
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
        RegisterSymbol(overlayValue, "Overlay")

        logBeginning := logBeginning . "|" . 
            overlayKey .                          "|" .            ; Overlay Key
            symbolLedger["Overlay"][overlayValue]                  ; Overlay Value
    }

    if IsSet(logBeginning) {
        if logToFile {
            FileAppend(logBeginning . newLine, system["Paths"]["Operation Log"], "UTF-8-RAW")
        } else {
            system["Logging"]["Operation Log"].Push(logBeginning)
        }
    }

    if logConclusionData.Has("Validation") {
        LogConclusion("Failed", logConclusionData, A_LineNumber, logConclusionData["Validation"])
    }

    if IsSet(overlayValue) {
        if runTelemetryTickCount - lastRunTelemetryTickCount >= runTelemetryInterval {
            lastRunTelemetryTickCount := runTelemetryTickCount
            system["Logging"]["Cycle"] := "Intermission"
            LogEngine()
        }
    }

    return logConclusionData
}

LogConclusion(conclusionStatus, logConclusionData, errorLineNumber := unset, errorMessage := unset) {
    static fileTimeBuffer   := Buffer(8, 0)
    static systemTimeBuffer := Buffer(16, 0)

    newLine := system["Constants"]["New Line"]

    conclusionStatus := StrUpper(SubStr(conclusionStatus, 1, 1)) . StrLower(SubStr(conclusionStatus, 2))

    if conclusionStatus = "Failed" && logConclusionData["Overlay Key"] = -1 {
        NumPut("UInt64", logConclusionData["Timestamp"], fileTimeBuffer, 0)
        DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")

        year        := NumGet(systemTimeBuffer,  0, "UShort")
        month       := NumGet(systemTimeBuffer,  2, "UShort")
        day         := NumGet(systemTimeBuffer,  6, "UShort")
        hour        := NumGet(systemTimeBuffer,  8, "UShort")
        minute      := NumGet(systemTimeBuffer, 10, "UShort")
        second      := NumGet(systemTimeBuffer, 12, "UShort")
        millisecond := NumGet(systemTimeBuffer, 14, "UShort")

        qpcMidpointTimestampDelta := (logConclusionData["QPC Pre"] + (logConclusionData["QPC Post"] - logConclusionData["QPC Pre"]) // 2) - system["Telemetry"]["QPC Midpoint Timestamp"]
        utcTimestampIntegerDelta  := Format("{:04}{:02}{:02}{:02}{:02}{:02}{:03}", year, month, day, hour, minute, second, millisecond) + 0 - system["Telemetry"]["UTC Timestamp Integer"]

        logConclusionData["Operation Sequence Number"] := EncodeIntegerToBase(logConclusionData["Operation Sequence Number"], 94)
        logConclusionData["Query Performance Counter"] := EncodeIntegerToBase(qpcMidpointTimestampDelta, 94)
        logConclusionData["UTC Timestamp Integer"]     := utcTimestampIntegerDelta

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

            if logToFile {
                FileAppend(logBeginning . newLine, system["Paths"]["Operation Log"], "UTF-8-RAW")
            } else {
                system["Logging"]["Operation Log"].Push(logBeginning)
            }
    }

    logConclusion := 
        logConclusionData["Operation Sequence Number"] .          "|" . ; Operation Sequence Number
        SubStr(conclusionStatus, 1, 1)                                  ; Status

    if !logConclusionData.Has("Context") && IsSet(errorMessage) {
        logConclusionData["Context"] := ""
    }

    if logConclusionData.Has("Context") {
        contextSymbol := RegisterSymbol(logConclusionData["Context"], "Context")
        logConclusionData["Context"] := contextSymbol
    }

    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    NumPut("UInt64", NumGet(timestampPointer, "Int64"), fileTimeBuffer, 0)
    DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")

    year        := NumGet(systemTimeBuffer,  0, "UShort")
    month       := NumGet(systemTimeBuffer,  2, "UShort")
    day         := NumGet(systemTimeBuffer,  6, "UShort")
    hour        := NumGet(systemTimeBuffer,  8, "UShort")
    minute      := NumGet(systemTimeBuffer, 10, "UShort")
    second      := NumGet(systemTimeBuffer, 12, "UShort")
    millisecond := NumGet(systemTimeBuffer, 14, "UShort")

    qpcMidpointTimestampDelta := (NumGet(qpcPrePointer, "Int64") + (NumGet(qpcPostPointer, "Int64") - NumGet(qpcPrePointer, "Int64")) // 2) - system["Telemetry"]["QPC Midpoint Timestamp"]
    utcTimestampIntegerDelta  := Format("{:04}{:02}{:02}{:02}{:02}{:02}{:03}", year, month, day, hour, minute, second, millisecond) + 0 - system["Telemetry"]["UTC Timestamp Integer"]

    logConclusion := logConclusion . "|" . 
        EncodeIntegerToBase(qpcMidpointTimestampDelta, 94) .      "|" . ; Query Performance Counter
        utcTimestampIntegerDelta                                        ; UTC Timestamp Integer

    if logConclusionData.Has("Context") {
        logConclusion := logConclusion . "|" . 
            logConclusionData["Context"]                                ; Method or Context
    }

    errorWindow             := unset
    constructedErrorMessage := unset
    if IsSet(errorMessage) {
        windowTitle          := "AutoHotkey v" . system["Runtime"]["AutoHotkey Version"] . ": " . A_ScriptName
        currentUtcDateTime   := ConvertIntegerToUtcTimestamp(system["Telemetry"]["UTC Timestamp Integer"] + utcTimestampIntegerDelta)

        if logConclusionData.Has("Validation") {
            errorLineNumber := methodRegistry[logConclusionData["Method Name"]]["Validation Line"]
        }
        
        declaration := RegExReplace(methodRegistry[logConclusionData["Method Name"]]["Declaration"], " <\d+>$", "")

        newLine := system["Constants"]["New Line"]
        constructedErrorMessage := "Declaration: " .  declaration . " (" . system["Runtime"]["Library Release"] . ")" . newLine
        if methodRegistry[logConclusionData["Method Name"]]["Parameters"] != "" {
            constructedErrorMessage := constructedErrorMessage .
                "Parameters: " . methodRegistry[logConclusionData["Method Name"]]["Parameters"] . newLine . 
                "Arguments: " . logConclusionData["Arguments Full"] . newLine
        }

        constructedErrorMessage := constructedErrorMessage . 
            "Line Number: " . errorLineNumber . newLine

        logErrorMessage := StrReplace(constructedErrorMessage . "Error Output: " . errorMessage, newLine, "|")
        RegisterSymbol(logErrorMessage, "Error")

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

    if logToFile {
        FileAppend(logConclusion . newLine, system["Paths"]["Operation Log"], "UTF-8-RAW")
    } else {
        system["Logging"]["Operation Log"].Push(logConclusion)
    }

    if logConclusionData["Overlay Key"] >= 1 {
        OverlayUpdateStatus(logConclusionData, conclusionStatus)
    }

    if IsSet(errorMessage) { 
        if system["Environment"].Has("Time Zone") && FileExist(system["Paths"]["Run Telemetry"]) {
            system["Logging"]["Cycle"] := "Failed"
            LogEngine()
        }

        if logConclusionData["Method Name"] = "AbortExecution" {
            ExitApp()
        }

        errorWindow.Show("AutoSize Center")
        WinWaitClose("ahk_id " . errorWindow.Hwnd)
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
            case "Array", "Map":
                argumentValueFull := "<Data Type: " . Type(argument) . ">"
                argumentValueLog  := RegisterSymbol(argumentValueFull, "Reference")

                argumentValueFull := Format('"{1}"', argumentValueFull)
                argumentValueLog  := Format('"{1}"', argumentValueLog)
            case "Integer":
                if Type(argument) != argumentsFormatted["Data Type"] {
                    argumentValueFull := "<Data Type: " . Type(argument) . ">"
                    argumentValueLog  := RegisterSymbol(argumentValueFull, "Reference")

                    argumentValueFull := Format('"{1}"', argumentValueFull)
                    argumentValueLog  := Format('"{1}"', argumentValueLog)
                }
            case "Object":
                argumentValueFull := "<Data Type: Object>"
                argumentValueLog  := RegisterSymbol(argumentValueFull, "Reference")

                argumentValueFull := Format('"{1}"', argumentValueFull)
                argumentValueLog  := Format('"{1}"', argumentValueLog)
            case "String":
                if Type(argument) != argumentsFormatted["Data Type"] {
                    argumentValueFull := "<Data Type: " . Type(argument) . ">"
                } else if argumentsFormatted["Data Constraint"] = "Base64" {
                    argumentValueFull := "<Constraint: Base64>"
                } else if InStr(argument, "`n") || InStr(argument, "`r") {
                    argumentValueFull := "<Text Block Rows: " . StrSplit(argument, "`n").Length . ">"
                } else {
                    if StrLen(argument) >= 255 {
                        argumentValueFull := SubStr(argument, 1, 255) . "…"
                    }
                }

                if argumentsFormatted["Whitelist"].Length != 0 && !logConclusionData.Has("Validation") {
                    argumentValueLog  := symbolLedger["Whitelist"][argumentValueLog]
                    argumentValueLog  := Format('\{1}\', argumentValueLog)
                } else {
                    argumentValueLog  := RegisterSymbol(argumentValueFull, "Reference")
                    argumentValueLog  := Format('"{1}"', argumentValueLog)
                }

                argumentValueFull := Format('"{1}"', argumentValueFull)
            case "Variant":
                if Type(argument) = "String" {
                    argumentValueLog  := RegisterSymbol(argumentValueFull, "Reference")

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

RegisterSymbol(value, type, writeToSymbolLedger := true) {
    global symbolLedger

    newLine := system["Constants"]["New Line"]

    typeCharacter := SubStr(type, 1, 1)
    entryExists   := true

    symbolLine := unset
    if !symbolLedger[type].Has(value) {
        entryExists := false
        counter     := IncrementCounter(type)

        if type = "Reference" || type = "Whitelist" {
            symbolLedger[type][value] := EncodeIntegerToBase(counter, 92)
        } else {
            symbolLedger[type][value] := EncodeIntegerToBase(counter, 94)
        }
    }

    if !entryExists && writeToSymbolLedger {
        symbolLine :=
            value . "|" . 
            typeCharacter . "|" . 
            symbolLedger[type][value]

        if logToFile {
            FileAppend(symbolLine . newLine, system["Paths"]["Symbol Ledger"], "UTF-8-RAW")
        } else {
            system["Logging"]["Symbol Ledger"].Push(symbolLine)
        }
    }

    symbol := symbolLedger[type][value]

    return symbol
}

SetLogFilenames(utcTimestamp) {
    global system

    validation := ValidateDataUsingSpecification(utcTimestamp, "String", "ISO Date Time")
    if validation != "" {
        return
    }

    oldLogFilenames  := Map()
    logSharedName    := system["Directories"]["Log"] . system["Runtime"]["Project Name"] . " - "
    logDateAndTime   := StrReplace(utcTimestamp, ":", ".")
    logComputerAlias := ""

    if system["Paths"].Has("Execution Log") && system["Paths"].Has("Operation Log") && system["Paths"].Has("Run Telemetry") && system["Paths"].Has("Symbol Ledger") {
        oldLogFilenames["Execution Log"] := system["Paths"]["Execution Log"]
        oldLogFilenames["Operation Log"] := system["Paths"]["Operation Log"]
        oldLogFilenames["Run Telemetry"] := system["Paths"]["Run Telemetry"]
        oldLogFilenames["Symbol Ledger"] := system["Paths"]["Symbol Ledger"]
    }

    if system["Configuration"].Has("Settings") {
        if system["Configuration"]["Settings"].Has("Computer Alias") {
            if system["Configuration"]["Settings"]["Computer Alias"] != "" {
                logComputerAlias := system["Configuration"]["Settings"]["Computer Alias"] . " - "
            }
        }
    }

    system["Paths"]["Execution Log"] := logSharedName . logComputerAlias . logDateAndTime . " - Execution Log.csv"
    system["Paths"]["Operation Log"] := logSharedName . logComputerAlias . logDateAndTime . " - Operation Log.csv"
    system["Paths"]["Run Telemetry"] := logSharedName . logComputerAlias . logDateAndTime . " - Run Telemetry.csv"
    system["Paths"]["Symbol Ledger"] := logSharedName . logComputerAlias . logDateAndTime . " - Symbol Ledger.csv"

    if oldLogFilenames.Has("Execution Log") {
        if FileExist(system["Paths"]["Execution Log"]) || FileExist(system["Paths"]["Operation Log"]) || FileExist(system["Paths"]["Run Telemetry"]) || FileExist(system["Paths"]["Symbol Ledger"]) {
            system["Paths"]["Execution Log"] := oldLogFilenames["Execution Log"]
            system["Paths"]["Operation Log"] := oldLogFilenames["Operation Log"]
            system["Paths"]["Run Telemetry"] := oldLogFilenames["Run Telemetry"]
            system["Paths"]["Symbol Ledger"] := oldLogFilenames["Symbol Ledger"]

            return
        } else {
            try {
                FileMove(oldLogFilenames["Execution Log"], system["Paths"]["Execution Log"])
                FileMove(oldLogFilenames["Operation Log"], system["Paths"]["Operation Log"])
                FileMove(oldLogFilenames["Run Telemetry"], system["Paths"]["Run Telemetry"])
                FileMove(oldLogFilenames["Symbol Ledger"], system["Paths"]["Symbol Ledger"])
            }
        }
    }

    system["Runtime"]["Run Identifier"] := system["Runtime"]["Project Name"] . " - " . logComputerAlias . logDateAndTime
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
        Loop 0x7E - 0x20 + 1 {
            codePoint := 0x20 + A_Index - 1

            maxBaseForExclusion := excludedAsciiCodePoints.Get(codePoint, 0)
            if baseType <= maxBaseForExclusion {
                continue
            }

            baseCharacters .= Chr(codePoint)
        }

        baseRadix := StrLen(baseCharacters)

        baseDigitByCharacterMap := Map()
        Loop StrLen(baseCharacters) {
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
    Loop Parse, baseText {
        integerValue := integerValue * baseRadix + digitMap[A_LoopField]
    }

    return integerValue
}

EncodeSha256HexToBase(hexSha256, baseType) {
    baseCharacterSet  := GetBaseCharacterSet(baseType)
    baseRadix         := baseCharacterSet["Base Radix"]
    charArray         := baseCharacterSet["Char Array"]

    static requiredLengthByBaseMap := Map(52, 45, 66, 43, 86, 40, 94, 40)

    hexSha256 := StrLower(hexSha256)

    sha256BytesBuffer := Buffer(32, 0)
    writeOffset := 0
    Loop 32 {
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
        Loop {
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

    Loop Parse, baseText {
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

BatchAppendExecutionLog(array) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("array as Array", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [array])

    newLine := system["Constants"]["New Line"]

    if array.Length != 0 {
        consolidatedExecutionLog := ""
        for index, value in array {
            if array.Length != index {
                consolidatedExecutionLog := consolidatedExecutionLog . value . newLine
            } else {
                consolidatedExecutionLog := consolidatedExecutionLog . value
            }
        }

        if logToFile {
            FileAppend(consolidatedExecutionLog . newLine, system["Paths"]["Execution Log"], "UTF-8-RAW")
        } else {
            system["Logging"]["Execution Log"].Push(consolidatedExecutionLog)
        }
    }
}

BatchAppendOperationLog(array) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("array as Array", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [array])

    newLine := system["Constants"]["New Line"]

    if array.Length != 0 {
        consolidatedOperationLog := ""
        for index, value in array {
            if array.Length != index {
                consolidatedOperationLog := consolidatedOperationLog . value . newLine
            } else {
                consolidatedOperationLog := consolidatedOperationLog . value
            }
        }

        if logToFile {
            FileAppend(consolidatedOperationLog . newLine, system["Paths"]["Operation Log"], "UTF-8-RAW")
        } else {
            system["Logging"]["Operation Log"].Push(consolidatedOperationLog)
        }
    }
}

BatchAppendRunTelemetry(appendType, array, ignoreAppendTypeOnFirstLine := false) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static appendTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}"', "Beginning", "Completed", "Failed", "Intermission")
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("appendType As String [Whitelist: " . appendTypeWhitelist . "], array as Array", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [appendType, array])

    newLine := system["Constants"]["New Line"]

    switch appendType {
        case "Beginning":
            appendType := "B"
        case "Completed":
            appendType := "C"
        case "Failed":
            appendType := "F"
        case "Intermission":
            appendType := "I"
    }

    if array.Length != 0 {
        consolidatedRunTelemetry := ""
        for index, value in array {
            if ignoreAppendTypeOnFirstLine && index = 1 {
                if array.Length != index {
                    consolidatedRunTelemetry := consolidatedRunTelemetry . value . newLine
                    continue
                } else {
                    consolidatedRunTelemetry := consolidatedRunTelemetry . value
                    continue
                }
            }

            if array.Length != index {
                consolidatedRunTelemetry := consolidatedRunTelemetry . value . "|" . appendType . newLine
            } else {
                consolidatedRunTelemetry := consolidatedRunTelemetry . value . "|" . appendType
            }
        }

        if logToFile {
            FileAppend(consolidatedRunTelemetry . newLine, system["Paths"]["Run Telemetry"], "UTF-8-RAW")
        } else {
            system["Logging"]["Run Telemetry"].Push(consolidatedRunTelemetry)
        }
    }
}

BatchAppendSymbolLedger(symbolType, array) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static symbolTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Context", "Error", "Method", "Overlay", "Reference", "Whitelist")
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("symbolType As String [Optional] [Whitelist: " . symbolTypeWhitelist . "], array As Array", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [symbolType, array])

    newLine := system["Constants"]["New Line"]

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
            typeCharacter := SubStr(symbolType, 1, 1)
            for index, value in symbolLedgerArray {
                symbol := RegisterSymbol(value, symbolType, false)

                if symbolLedgerArray.Length != index {
                    consolidatedSymbolLedger := consolidatedSymbolLedger . value . "|" . typeCharacter . "|" . symbol . newLine
                } else {
                    consolidatedSymbolLedger := consolidatedSymbolLedger . value . "|" . typeCharacter . "|" . symbol
                }
            }
        }

        if logToFile {
            FileAppend(consolidatedSymbolLedger . newLine, system["Paths"]["Symbol Ledger"], "UTF-8-RAW")
        } else {
            system["Logging"]["Symbol Ledger"].Push(consolidatedSymbolLedger)
        }
    }
}

OverlayIsVisible() {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"))

    windowHandle  := overlay["GUI"].Hwnd
    windowVisible := unset

    if DllCall("User32\IsWindowVisible", "Ptr", windowHandle) {
        windowVisible := true
    } else {
        windowVisible := false
    }

    return windowVisible
}