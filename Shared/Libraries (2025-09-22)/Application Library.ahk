#Requires AutoHotkey v2.0
#Include Base Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

global applicationRegistry := Map()

; ******************** ;
; Application Registry ;
; ******************** ;

RegisterApplications() {
    global applicationRegistry

    for index, applicationName in [
        "Excel",
        "Notepad++",
        "SQL Server Management Studio",
        "Toad for Oracle"
    ] {
        applicationRegistry[applicationName] := Map()

        applicationRegistry[applicationName]["Executable Path"] := ExecutablePathResolve(applicationName)
        applicationRegistry[applicationName]["Facts"]           := FactsResolve(applicationName)
    }
}

ExecutablePathResolve(applicationName) {
    executablePath := ""

    switch applicationName
    {
        case "Excel":
            try {
                executablePath := RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", "")

                if executablePath = "" {
                    executablePath := RegRead("HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe", "")
                }
            }
        case "Notepad++":
            try {
                executablePath := RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\notepad++.exe", "")
            }

            if executablePath = "" {
                try {
                    executablePath := RegRead("HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\notepad++.exe", "")
                }
            }

            ; Fallback: vendor key stores install directory in "Path"
            if (executablePath = "" || !FileExist(executablePath)) {
                try {
                    installDirectory := RegRead("HKLM\SOFTWARE\Notepad++", "Path")
                    if (installDirectory != "" && FileExist(installDirectory . "\notepad++.exe")) {
                        executablePath := installDirectory . "\notepad++.exe"
                    }
                }
            }

            ; Fallbacks: typical install locations
            if (executablePath = "" || !FileExist(executablePath)) {
                if FileExist(A_ProgramFiles . "\Notepad++\notepad++.exe") {
                    executablePath := A_ProgramFiles . "\Notepad++\notepad++.exe"
                }
            }
            if (executablePath = "" || !FileExist(executablePath)) {
                if FileExist(A_ProgramFiles . " (x86)\Notepad++\notepad++.exe") {
                    executablePath := A_ProgramFiles . " (x86)\Notepad++\notepad++.exe"
                }
            }
        case "SQL Server Management Studio":
            try {
                ; Attempt 1: SSMSInstallRoot (SSMS 18–20)
                baseKeys := [
                    "HKLM\\SOFTWARE\\Microsoft\\Microsoft SQL Server Management Studio",
                    "HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Microsoft SQL Server Management Studio"
                ]
                for baseKey in baseKeys {
                    Loop Reg, baseKey, "K" {
                        try {
                            installRoot := RegRead(A_LoopRegKey, "SSMSInstallRoot")
                        } catch {
                            installRoot := ""
                        }
                        if installRoot != "" {
                            candidate := installRoot "\Common7\IDE\SSMS.exe"
                            if FileExist(candidate) {
                                executablePath := candidate
                                break 2
                            }
                        }
                    }
                }

                ; Attempt 2: pre‑18 releases (sqlwb.exe)
                if executablePath = "" {
                    regKeys := [
                        "HKLM\\SOFTWARE\\Classes\\Applications\\sqlwb.exe\\shell\\open\\command",
                        "HKLM\\SOFTWARE\\WOW6432Node\\Classes\\Applications\\sqlwb.exe\\shell\\open\\command"
                    ]
                    for key in regKeys {
                        try {
                            executableCommandLine := RegRead(key, "")
                        } catch {
                            executableCommandLine := ""
                        }
                        if executableCommandLine != "" {
                            executablePath := RegExReplace(executableCommandLine, '^\s*"([^"]+\.exe).*', "$1")
                            break
                        }
                    }
                }

                ; Attempt 3: SSMS 21 (file‑system probe)
                if executablePath = "" {
                    versionRoots := [
                        "C:\Program Files\Microsoft SQL Server Management Studio*",
                        "C:\Program Files (x86)\Microsoft SQL Server Management Studio*"
                    ]
                    for rootPattern in versionRoots {
                        Loop Files, rootPattern, "D" {
                            candidate := A_LoopFilePath "\Release\Common7\IDE\SSMS.exe"
                            if FileExist(candidate) {
                                executablePath := candidate
                                break 2
                            }
                            candidate := A_LoopFilePath "\Common7\IDE\SSMS.exe"
                            if FileExist(candidate) {
                                executablePath := candidate
                                break 2
                            }
                        }
                    }
                }
            }
        case "Toad for Oracle":
            try {
                ; Attempt 1: Uninstall keys (DisplayIcon or InstallLocation)
                if executablePath = "" {
                    for uninstallBaseKeyPath in [
                        "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                        "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                    ] {
                        Loop Reg, uninstallBaseKeyPath, "K" {
                            uninstallSubKeyPath := A_LoopRegKey "\" A_LoopRegName

                            displayName := ""
                            try {
                                displayName := RegRead(uninstallSubKeyPath, "DisplayName")
                            }
                            if !(displayName && InStr(StrLower(displayName), "toad") && InStr(StrLower(displayName), "oracle")) {
                                continue
                            }

                            displayIcon := ""
                            try {
                                displayIcon := RegRead(uninstallSubKeyPath, "DisplayIcon")
                            }
                            if displayIcon {
                                pathToExecutable := Trim(StrSplit(displayIcon, ",")[1], ' "')
                                if FileExist(pathToExecutable) && (SubStr(StrLower(pathToExecutable), -9) = "\toad.exe") {
                                    executablePath := pathToExecutable
                                }
                            }

                            installLocation := ""
                            try {
                                installLocation := RegRead(uninstallSubKeyPath, "InstallLocation")
                            }
                            if installLocation {
                                pathToExecutable := RTrim(installLocation, "\/") "\Toad.exe"
                                if FileExist(pathToExecutable) {
                                    executablePath := pathToExecutable
                                }
                            }
                        }
                    }
                }

                ; Attempt 2: App Paths (usually exact exe)
                if executablePath = "" {
                    for registryKeyPath in [
                        "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Toad.exe",
                        "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\Toad.exe"
                    ] {
                        try {
                            registryDefaultValue := RegRead(registryKeyPath, "")
                            if registryDefaultValue {
                                pathToExecutable := Trim(StrSplit(registryDefaultValue, ",")[1], ' "')
                                if FileExist(pathToExecutable) {
                                    executablePath := pathToExecutable
                                }
                            }
                        }
                    }
                }

                ; Attempt 3: Vendor keys (Quest/Dell rebrand)
                if executablePath = "" {
                    for vendorRegistryKeyPath in [
                        "HKLM\SOFTWARE\Quest Software\Toad for Oracle",
                        "HKLM\SOFTWARE\WOW6432Node\Quest Software\Toad for Oracle",
                        "HKLM\SOFTWARE\Dell\Toad for Oracle",
                        "HKLM\SOFTWARE\WOW6432Node\Dell\Toad for Oracle"
                    ] {
                        for vendorValueName in ["InstallPath","InstallLocation","Path"] {
                            try {
                                registryValue := RegRead(vendorRegistryKeyPath, vendorValueName)
                                if registryValue {
                                    pathToExecutable := (SubStr(StrLower(registryValue), -9) = "\toad.exe") ? Trim(StrSplit(registryValue, ",")[1], ' "') : RTrim(registryValue, "\/") "\Toad.exe"
                                    pathToExecutable := Trim(pathToExecutable, ' "')
                                    if FileExist(pathToExecutable) {
                                        executablePath := pathToExecutable
                                    }
                                }
                            }
                        }
                    }
                }

                ; Attempt 4: Narrow Program Files scan (no full-disk search)
                if executablePath = "" {
                    for programFilesDirectory in [EnvGet("ProgramFiles"), EnvGet("ProgramFiles(x86)")] {
                        if !programFilesDirectory {
                            continue
                        }
                        for vendorName in ["Quest Software","Dell"] {
                            vendorRootDirectory := programFilesDirectory "\" vendorName
                            if !DirExist(vendorRootDirectory) {
                                continue
                            }
                            Loop Files, vendorRootDirectory "\Toad for Oracle*\Toad.exe", "R" {
                                executablePath := A_LoopFileFullPath
                            }
                        }
                    }
                }
            }
        default:
    }

    return executablePath
}

FactsResolve(applicationName) {
    facts := ""

    switch applicationName
    {
        case "Excel":
            if applicationRegistry[applicationName]["Executable Path"] !== "" {
                facts := ApplicationBaseFacts(applicationName, facts)

                personalMacroWorkbookPath := EnvGet("AppData") "\Microsoft\Excel\XLSTART\PERSONAL.XLSB"
                if FileExist(personalMacroWorkbookPath) {
                    facts := facts . "Personal Macro Workbook: Enabled" . "|"
                } else {
                    facts := facts . "Personal Macro Workbook: Disabled" . "|"
                }

                Run('"' . applicationRegistry["Excel"]["Executable Path"] . '"')
                WaitForExcelToLoad()
                excelApplication := ComObjActive("Excel.Application")
                activeWorkbook := excelApplication.ActiveWorkbook

                excelMacroTestCode := 'Sub CellPing(): Range("A1").Value = "Cell": End Sub'

                ExcelScriptExecution(excelMacroTestCode, true)
                Sleep(160)
                SendEvent("^{F4}") ; CTRL+F4 (Close Window: Module)
                Sleep(160)

                if excelApplication.ActiveSheet.Range("A1").Value = "Cell" {
                    facts := facts . "Code Execution: Basic"

                    excelApplication.ActiveSheet.Range("A1").Value := ""

                    ExcelScriptExecution(excelMacroTestCode)
                    Sleep(160)

                    if excelApplication.ActiveSheet.Range("A1").Value = "Cell" {
                        facts := StrReplace(facts, "Code Execution: Basic", "Code Execution: Full")
                    }
                } else {
                    facts := facts . "Code Execution: Failed"
                }

                activeWorkbook.Close(false)
                excelApplication.DisplayAlerts := false
                excelApplication.Quit()

                activeWorkbook := 0
                excelApplication := 0
            } else {
                facts := "Installed: No"
            }
        case "Notepad++":
            if applicationRegistry[applicationName]["Executable Path"] !== "" {
                facts := ApplicationBaseFacts(applicationName, facts, true)
            } else {
                facts := "Installed: No"
            }
        case "SQL Server Management Studio":
            if applicationRegistry[applicationName]["Executable Path"] !== "" {
                facts := ApplicationBaseFacts(applicationName, facts, true)
            } else {
                facts := "Installed: No"
            }
        case "Toad for Oracle":
            if applicationRegistry[applicationName]["Executable Path"] !== "" {
                facts := ApplicationBaseFacts(applicationName, facts, true)
            } else {
                facts := "Installed: No"
            }
        default:
    }

    return facts
}

ApplicationBaseFacts(applicationName, facts, trimLastDelimiter := false) {
    facts := "Installed: Yes" . "|"
    facts := facts . "Executable Path: " . applicationRegistry[applicationName]["Executable Path"] . "|"
    facts := facts . "Executable Hash: " . Hash.File("SHA256", applicationRegistry[applicationName]["Executable Path"]) . "|"

    if trimLastDelimiter = true {
        facts := RTrim(facts, "|")
    }

    return facts
}

ValidateApplicationFact(applicationName, factName, factValue) {
    static methodName := RegisterMethod("ValidateApplicationFact(applicationName As String, factName As String, factValue As String)" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Validate Application Fact (" . applicationName . ", " . factName . ", " . factValue . ")", methodName, [applicationName, factName, factValue])

    try {
        if !applicationRegistry.Has(applicationName) {
            throw Error("Application " . Chr(34) . applicationName . Chr(34) . " not registered.")
        }
    } catch as applicationNotRegisteredError {
        LogInformationConclusion("Failed", logValuesForConclusion, applicationNotRegisteredError)
    }

    factMap := Map()
    for index, factPair in StrSplit(applicationRegistry[applicationName]["Facts"], "|") {
        positionOfDelimiter := InStr(factPair, ":")
        if positionOfDelimiter = 0 {
            continue
        }
        factNamePart  := Trim(SubStr(factPair, 1, positionOfDelimiter - 1))
        factValuePart := Trim(SubStr(factPair, positionOfDelimiter + 1))
        factMap[factNamePart] := factValuePart
    }

    try {
        if !factMap.Has(factName) {
            throw Error(Chr(34) . applicationName . Chr(34) . " does not have a valid fact name " . Chr(34) . factName . Chr(34) . ".")
        }
    } catch as factNameMissingError {
        LogInformationConclusion("Failed", logValuesForConclusion, factNameMissingError)
    }

    try {
        if factMap[factName] !== factValue {
            throw Error(Chr(34) . applicationName . Chr(34) . " with fact name of " . Chr(34) . factName . Chr(34) . " does not match fact value " . Chr(34) . factValue . Chr(34) . ".")
        }
    } catch as factValueMissingError {
        LogInformationConclusion("Failed", logValuesForConclusion, factValueMissingError)
    }

    if factName = "Installed" && factMap["Installed"] = "Yes" {
        factsRevised := "|" . applicationRegistry[applicationName]["Facts"] "|"
        factsRevised := StrReplace(factsRevised, "|Installed: Yes|", "|")
        AppendCsvLineToLog(applicationName . factsRevised . "Application", "Execution Log")
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; ******************** ;
; Excel                ;
; ******************** ;

ExcelExtensionRun(documentName, saveDirectory, code, displayName := "", aboutRange := "", aboutCondition := "") {
    static methodName := RegisterMethod("ExcelExtensionRun(documentName As String [Type: Search], saveDirectory As String [Type: Directory], code As String [Type: Code], displayName As String [Optional], aboutRange As String [Optional], aboutCondition As String [Optional])" . LibraryTag(A_LineFile), A_LineNumber + 7)
    overlayValue := ""
    if displayName = "" {
        overlayValue := documentName . " Excel Extension Run"
    } else {
        overlayValue := displayName . " Excel Extension Run"
    }
    logValuesForConclusion := LogInformationBeginning(overlayValue, methodName, [documentName, saveDirectory, code, displayName, aboutRange, aboutCondition])

    excelFilePath := FileExistsInDirectory(documentName, saveDirectory, "xlsx")
    try {
        if excelFilePath = "" {
            throw Error("documentName not found: " . documentName)
        }
    } catch as documentNameNotFoundError {
        LogInformationConclusion("Failed", logValuesForConclusion, documentNameNotFoundError)
    }

    Run('"' . applicationRegistry["Excel"]["Executable Path"] . '"  "' . excelFilePath . '"')
    WaitForExcelToLoad()
    excelWorkbook := ComObjGet(excelFilePath)
    excelApplication := excelWorkbook.Application

    aboutWorksheet := ""
    aboutWorksheetFound := false

    for sheet in excelWorkbook.Worksheets {
        if sheet.Name = "About" {
            aboutWorksheet := sheet
            aboutWorksheetFound := true
            break
        }
    }

    if (aboutRange !== "" || aboutCondition !== "") && aboutWorksheetFound = false {
        try {
            throw Error("Worksheet About not found with arguments passed in.")
        } catch as worksheetAboutMissingError {
            LogInformationConclusion("Failed", logValuesForConclusion, worksheetAboutMissingError)
        }
    }

    if aboutRange = "Progression Status" {
        aboutRange := "ProgressionStatus"
    }

    if aboutRange = "Augmentation Modules" {
        aboutRange := "AugmentationModules"
    }

    if aboutRange = "Dependencies List" {
        aboutRange := "DependenciesList"
    }

    if aboutRange = "Retrieved Date" {
        aboutRange := "RetrievedDate"
    } 

    if aboutWorksheetFound && (aboutRange !== "" || aboutCondition !== "") {
        aboutValues := Map(
            "ProgressionStatus",   "A3",
            "AugmentationModules", "A4",
            "RetrievedDate",       "C1",
            "EditionName",         "C2"
        )

        for fieldName, cellAddress in aboutValues {
            aboutValues[fieldName] := aboutWorksheet.Range(cellAddress).Value
        }

        aboutValues["ProgressionStatus"] := StrReplace(aboutValues["ProgressionStatus"], "Progression Status: ", "")
        aboutValues["AugmentationModules"] := StrReplace(aboutValues["AugmentationModules"], "Augmentation Modules: ", "")
        aboutValues["RetrievedDate"] := str := SubStr(StrReplace(aboutValues["RetrievedDate"], "Retrieved Date: ", ""), 1, -1)
        aboutValues["EditionName"] := StrReplace(aboutValues["EditionName"], "Edition Name: ", "")

        if aboutValues[aboutRange] = aboutCondition {
            ExcelScriptExecution(code)
            WaitForExcelToClose()

            excelWorkbook    := 0
            excelApplication := 0

            LogInformationConclusion("Completed", logValuesForConclusion)
        } else if aboutRange = "ProgressionStatus" {
            conditionParts := StrSplit(aboutCondition, ", ")
            builtPrefix := ""
            matchedIndex := 0
            conditionPartsCount := conditionParts.Length

            for index, currentPart in conditionParts {
                if index = 1 {
                    builtPrefix := currentPart
                } else {
                    builtPrefix := builtPrefix ", " currentPart
                }
                candidateValue := builtPrefix "."
                if aboutValues[aboutRange] = candidateValue {
                    matchedIndex := index
                    break
                }
            }

            if matchedIndex > 0 {
                ExcelScriptExecution(code)
                WaitForExcelToClose()

                excelWorkbook    := 0
                excelApplication := 0

                LogInformationConclusion("Completed", logValuesForConclusion)
            } else {
                activeWorkbook := excelApplication.ActiveWorkbook
                activeWorkbook.Close(false)
                excelApplication.DisplayAlerts := false
                excelApplication.Quit()

                activeWorkbook   := 0
                excelWorkbook    := 0
                excelApplication := 0

                LogInformationConclusion("Skipped", logValuesForConclusion)
            }
        } else {
            activeWorkbook := excelApplication.ActiveWorkbook
            activeWorkbook.Close(false)
            excelApplication.DisplayAlerts := false
            excelApplication.Quit()

            activeWorkbook   := 0
            excelWorkbook    := 0
            excelApplication := 0

            LogInformationConclusion("Skipped", logValuesForConclusion)
        }
    } else {
        ExcelScriptExecution(code)
        WaitForExcelToClose()

        excelWorkbook    := 0
        excelApplication := 0

        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

ExcelScriptExecution(code, insertModule := false) {
    static methodName := RegisterMethod("ExcelScriptExecution(code As String [Type: Code], insertModule As Boolean [Optional]" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Excel Script Execution (Length: " . StrLen(code) . ")", methodName, [code, insertModule])
   
    SendEvent("!{F11}") ; F11 (Microsoft Visual Basic for Applications)
    WinWait("ahk_class wndclass_desked_gsk", , 10)
    WinActivate("ahk_class wndclass_desked_gsk")
    WinWaitActive("ahk_class wndclass_desked_gsk", , 2)

    if insertModule = true {
        SendEvent("!i") ; ALT+I (Insert)
        Sleep(280)
        SendEvent("m") ; M (Module)
    }

    PasteCode(code, "'")

    SendEvent("{F5}") ; F5 (Run Sub/UserForm)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

ExcelStartingRun(documentName, saveDirectory, code, displayName := "") {
    static methodName := RegisterMethod("ExcelStartingRun(documentName As String [Type: Search], saveDirectory As String [Type: Directory], code As String [Type: Code], displayName As String [Optional])" . LibraryTag(A_LineFile), A_LineNumber + 7)
    overlayValue := ""
    if displayName = "" {
        overlayValue := documentName . " Excel Starting Run"
    } else {
        overlayValue := displayName . " Excel Starting Run"
    }
    logValuesForConclusion := LogInformationBeginning(overlayValue, methodName, [documentName, saveDirectory, code, displayName])

    xlsxPath := FileExistsInDirectory(documentName, saveDirectory, "xlsx")
    txtPath  := FileExistsInDirectory(documentName, saveDirectory, "txt")

    if txtPath !== "" && xlsxPath !== "" {
        DeleteFile(txtPath)
        DeleteFile(xlsxPath)
        xlsxPath := ""
    } else {
        if txtPath !== "" {
            DeleteFile(txtPath)
        }
    }

    if xlsxPath = "" {
        sidecarPath := saveDirectory . documentName . ".txt"
        FileAppend("", sidecarPath, "UTF-8-RAW")

        Run('"' . applicationRegistry["Excel"]["Executable Path"] . '"')
        WaitForExcelToLoad()
        ExcelScriptExecution(code)
        WaitForExcelToClose()

        DeleteFile(sidecarPath) ; Remove sidecar after a successful run.
        LogInformationConclusion("Completed", logValuesForConclusion)
    } else {
        LogInformationConclusion("Skipped", logValuesForConclusion)
    }
}

WaitForExcelToClose(maxWaitMinutes := 240, mouseMoveIntervalSec := 120) {
    static methodName := RegisterMethod("WaitForExcelToClose(maxWaitMinutes As Integer [Optional: 240], mouseMoveIntervalSec As Integer [Optional: 120])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Wait for Excel to Close", methodName, [maxWaitMinutes, mouseMoveIntervalSec])

    totalSecondsToWait := maxWaitMinutes * 60
    secondsSinceLastMouseMove := 0

    Loop totalSecondsToWait {
        Sleep(1000)
        secondsSinceLastMouseMove += 1

        if secondsSinceLastMouseMove >= mouseMoveIntervalSec {
            MouseMove 0, 0, 0, "R" ; For preventing screen saver from activating.
            secondsSinceLastMouseMove := 0
        }

        if !WinExist("ahk_class XLMAIN ahk_exe EXCEL.EXE") {
            break
        }
    }

    try {
        if WinExist("ahk_class XLMAIN ahk_exe EXCEL.EXE") {
            throw Error("Excel did not close within " . maxWaitMinutes . " minutes.")
        }
    } catch as excelCloseError {
        LogInformationConclusion("Failed", logValuesForConclusion, excelCloseError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

WaitForExcelToLoad() {
    static methodName := RegisterMethod("WaitForExcelToLoad()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Wait for Excel to Load", methodName)

    try {
        if !WinWait("ahk_class XLMAIN", , 480) {
            throw Error("Excel document did not appear within 480 seconds.")
        }
    } catch as documentNotLoadedError {
        LogInformationConclusion("Failed", logValuesForConclusion, documentNotLoadedError)
    }

    WinActivate("ahk_class XLMAIN")
    WinActivate("ahk_class XLMAIN")
    WinWaitActive("ahk_class XLMAIN", , 10)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; ******************** ;
; SQL Server MS        ;
; ******************** ;

StartMicrosoftSqlServerManagementStudioAndConnect() {
    static methodName := RegisterMethod("StartMicrosoftSqlServerManagementStudioAndConnect()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Start Microsoft SQL Server Management Studio and Connect", methodName)

    Run('"' . applicationRegistry["SQL Server Management Studio"]["Executable Path"] . '"')

    WinWait("Connect to Server",, 20)
    Sleep(2000)

    SendEvent("{Enter}")

    try {
        if WinWaitClose("Connect to Server",, 40) {
        } else {
            throw Error("Connection failed.")
        }
    } catch as connectError {
        LogInformationConclusion("Failed", logValuesForConclusion, connectError)
    }

    windowTitle := "Microsoft SQL Server Management Studio"
    WinWait(windowTitle,, 20)
    WinActivate(windowTitle)
    WinWaitActive(windowTitle,, 10)
    WinMaximize(windowTitle)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

ExecuteSqlQueryAndSaveAsCsv(code, saveDirectory, filename) {
    static methodName := RegisterMethod("ExecuteSqlQueryAndSaveAsCsv(code As String [Type: Code], saveDirectory As String [Type: Directory], filename As String [Type: Search])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Execute SQL Query and Save (" . filename . ")", methodName, [code, saveDirectory, filename])

    savePath       := saveDirectory . filename . ".csv"

    SendInput("^n") ; CTRL+N (Query with Current Connection)
    Sleep(2000)

    PasteCode(code, "--")

    SendInput("!x") ; ALT+X (Execute)
    sqlQuerySuccessfulCoordinates := RetrieveImageCoordinatesFromSegment("SMMS Query Successful", "6-26", "88-96", 360)
    sqlQueryResultsWindow := ModifyScreenCoordinates(80, -80, sqlQuerySuccessfulCoordinates)
    PerformMouseActionAtCoordinates("Left", sqlQueryResultsWindow)
    Sleep(480)
    PerformMouseActionAtCoordinates("Right", sqlQueryResultsWindow)
    Sleep(480)
    SendEvent("v") ; V (Save Results As...)
    Sleep(2000)
    SendEvent("!n") ; ALT+N (File name)

    PastePath(savePath)

    SendEvent("{Enter}") ; ENTER (Save)
    Sleep(480)
    SendEvent("y") ; Y (Yes)
    Sleep(240)
}

CloseMicrosoftSqlServerManagementStudio() {
    static methodName := RegisterMethod("CloseMicrosoftSqlServerManagementStudio()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Close Microsoft SQL Server Management Studio", methodName)

    fullPath := applicationRegistry["SQL Server Management Studio"]["Executable Path"]
    SplitPath(fullPath, &executableName)
    ProcessClose(executableName)
    ProcessWaitClose(executableName, 4)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; ******************** ;
; Toad for Oracle      ;
; ******************** ;

ExecuteAutomationApp(appName, runtimeDate := "") {
    static methodName := RegisterMethod("ExecuteAutomationApp(appName As String [Type: Search], runtimeDate As String [Optional] [Type: Raw Date Time])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Verify Toad for Oracle Works", methodName, [appName, runtimeDate])

    toadForOracleExecutablePath := applicationRegistry["Toad for Oracle"]["Executable Path"]
    SplitPath(toadForOracleExecutablePath, &toadExecutableFilename)

    windowCriteria := "ahk_exe " . toadExecutableFilename . " ahk_class TfrmMain"

    try {
        if !ProcessExist(toadExecutableFilename) {
            throw Error("Toad for Oracle process is not running.")
        }
    } catch as processNotRunningError {
        LogInformationConclusion("Failed", logValuesForConclusion, processNotRunningError)
    }

    try {
        if WinExist("ahk_exe " . toadExecutableFilename . " ahk_class TfrmLogin") {
            throw Error("No server connection is active in Toad for Oracle (login dialog is open).")
        }
    } catch as noActiveConnectionError {
        LogInformationConclusion("Failed", logValuesForConclusion, noActiveConnectionError)
    }

    WinActivate(windowCriteria)
    WinMaximize(windowCriteria)

    SendEvent("!s")      ; ALT+S (Session)
    Sleep(400)
    SendEvent("s")       ; S (Test Connection)
    Sleep(400)
    SendEvent("{Enter}") ; ENTER (Apply)
    
    overallStartTickCount := A_TickCount
    firstSeenTickCount := 0
    dialogHasAppeared := false

    while true {
        dialogExists := WinExist("ahk_exe " . toadExecutableFilename . " ahk_class TReconnectForm")

        if (dialogHasAppeared = false) {
            ; Phase 1: waiting for the dialog to appear
            if (dialogExists != false) {
                dialogHasAppeared := true
                firstSeenTickCount := A_TickCount
            } else if (A_TickCount - overallStartTickCount >= 2000) {
                ; No dialog ever appeared -> likely instant/local reconnect
                break
            }
        } else {
            ; Phase 2: dialog has appeared; wait until it closes (disappears)
            if (dialogExists = false) {
                ; Closed = reconnect finished
                break
            }
            try {
                if (A_TickCount - firstSeenTickCount >= 30000) {
                    throw Error("Reconnect dialog did not close within " . Round(30000 / 1000) . " seconds.")
                }
            } catch as reconnectFailedError {
                LogInformationConclusion("Failed", logValuesForConclusion, reconnectFailedError)
            }
        }

        Sleep(32)
    }

    Sleep(800)

    try {
        SendEvent("!u") ; ALT+U (Utilities)
        submenuWindowHandle := WinWait("ahk_exe " toadExecutableFilename " ahk_class TdxBarSubMenuControl", , 1000)
        if !submenuWindowHandle {
            throw Error("Failed to open the Utilities menu (submenu was not detected).")
        }
    } catch as openUtilitiesError {
        LogInformationConclusion("Failed", logValuesForConclusion, openUtilitiesError)
    }

    try {
        SendEvent("{Enter}") ; ENTER (Automation Designer)

        if !WinWaitClose("ahk_id " submenuWindowHandle, , 1000) {
            throw Error("Failed to launch Automation Designer from the Utilities menu.")
        }
    } catch as selectAutomationDesignerError {
        LogInformationConclusion("Failed", logValuesForConclusion, selectAutomationDesignerError)
    }

    toadForOracleSearchCoordinates := RetrieveImageCoordinatesFromSegment("Toad for Oracle Search", "0-40", "80-100")
    PerformMouseActionAtCoordinates("Left", toadForOracleSearchCoordinates)
    Sleep(2000)

    SendEvent("{Tab}") ; TAB (Text to find:)
    Sleep(1200)

    PasteSearch(appName)

    Sleep(1200)
    SendEvent("{Enter}") ; ENTER (Search)
    Sleep(1200)
    SendEvent("+{Tab}") ; SHIFT+TAB (Item)
    Sleep(1200)
    SendEvent("+{F10}") ; SHIFT+F10 (Right-click)
    Sleep(1200)
    SendEvent("{Down}") ; DOWN ARROW (Goto Item)
    Sleep(1200)
    SendEvent("{Enter}") ; ENTER (Goto Item)
    Sleep(1200)
    toadForOraclePlayCoordinates := RetrieveImageCoordinatesFromSegment("Toad for Oracle Play", "0-40", "0-20")

    if runtimeDate !== "" {
        PerformMouseActionAtCoordinates("Move", toadForOraclePlayCoordinates)

        while (A_Now < DateAdd(runtimeDate, -1, "Seconds")) {
            Sleep(240)
        }

        while (A_Now < runtimeDate) {
            Sleep(16)
        }
    }

    PerformMouseActionAtCoordinates("Left", toadForOraclePlayCoordinates)
    Sleep(16)
    PerformMouseActionAtCoordinates("Move", (Round(A_ScreenWidth/2)) . "x" . (Round(A_ScreenHeight/1.2)))

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; ******************** ;
; Helper Methods       ;
; ******************** ;

CombineExcelCode(introCode, mainCode, outroCode := "") {
    combinedCode := introCode . "`r`n`r`n" . mainCode

    if outroCode !== "" {
        combinedCode := combinedCode . "`r`n`r`n" . outroCode
    } 

    return combinedCode
}