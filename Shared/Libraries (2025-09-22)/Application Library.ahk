#Requires AutoHotkey v2.0
#Include Base Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

global applicationRegistry := Map()

; **************************** ;
; Application Registry         ;
; **************************** ;

RegisterApplications() {
    global applicationRegistry
   
    applications := ConvertCsvToArrayOfMaps(ExtractParentDirectory(A_LineFile) . "Mappings\Applications.csv")

    for application in applications {
        applicationName := application["Name"]

        applicationRegistry[applicationName] := Map()

        applicationRegistry[applicationName]["Executable Path"] := ExecutablePathResolve(applicationName)

        if applicationRegistry[applicationName]["Executable Path"] = "" {
            applicationRegistry[applicationName]["Installed"] := false
        } else {
            applicationRegistry[applicationName]["Installed"] := true
            ResolveFactsForApplication(applicationName)
        }
    }

    installedApplications := []
    for outerKey, innerMap in applicationRegistry {
        if innermap["Installed"] = true {
            configuration := outerkey . "|" . innerMap["Executable Path"] . "|" . innerMap["Executable Hash"] . "|" . innerMap["Executable Version"] . "|" . innerMap["Binary Type"]

            switch outerKey
            {
                case "Excel":
                    configuration := configuration . "|" . "Personal Macro Workbook: " . innerMap["Personal Macro Workbook"] . "|" . "Code Execution: " . innerMap["Code Execution"]
            }

            installedApplications.Push(configuration)
            innerMap["Symbol Ledger Lookup"] := configuration . "|" . "A"
        }
    }

    SymbolLedgerBatchAppend("A", installedApplications)

    for outerKey, innerMap in applicationRegistry {
        if innermap["Installed"] = true {
            innerMap["Symbol"] := symbolLedger[innerMap["Symbol Ledger Lookup"]]["Symbol"]
        }
    }
}

ExecutablePathResolve(applicationName) {
    executablePath      := ""
    executableName      := ""
    executableDirectory := ""
    registryKeyPaths    := []

    static applicationExecutableDirectoryCandidates := ConvertCsvToArrayOfMaps(ExtractParentDirectory(A_LineFile) . "Mappings\Application Executable Directory Candidates.csv")
    static applicationRegistryPathCandidates        := ConvertCsvToArrayOfMaps(ExtractParentDirectory(A_LineFile) . "Mappings\Application Registry Path Candidates.csv")

    for applicationExecutableDirectoryCandidate in applicationExecutableDirectoryCandidates {
        if applicationExecutableDirectoryCandidate["Name"] = applicationName {
            executableName      := applicationExecutableDirectoryCandidate["Executable"]
            executableDirectory := applicationExecutableDirectoryCandidate["Directory"]

            for applicationRegistryPathCandidate in applicationRegistryPathCandidates {
                if applicationRegistryPathCandidate["Name"] = applicationName {
                    registryPath := applicationRegistryPathCandidate["Registry Path"]

                    registryKeyPaths.Push(registryPath)
                }
            }

            executablePath := ExecutablePathViaRegistry(executablePath, executableName, registryKeyPaths)
            executablePath := ExecutablePathViaUninstall(executablePath, executableName)
            executablePath := ExecutablePathViaDirectory(executablePath, executableName, executableDirectory)

            if executablePath !== "" {
                break
            }
        }
    }

    return executablePath
}

ExecutablePathViaDirectory(executablePath, executableName, directoryName) {
    if executablePath !== "" {
        return executablePath
    }

    static candidateBaseDirectories

    if !IsSet(candidateBaseDirectories) {
        candidateBaseDirectories := []

        programFilesDirectory := EnvGet("ProgramFiles")
        if programFilesDirectory {
            candidateBaseDirectories.Push(programFilesDirectory)
        }

        programFilesX86Directory := EnvGet("ProgramFiles(x86)")
        if programFilesX86Directory && programFilesX86Directory != programFilesDirectory {
            candidateBaseDirectories.Push(programFilesX86Directory)
        }

        programFilesW6432Directory := EnvGet("ProgramW6432")
        if programFilesW6432Directory && programFilesW6432Directory != programFilesDirectory && programFilesW6432Directory != programFilesX86Directory {
            candidateBaseDirectories.Push(programFilesW6432Directory)
        }

        localApplicationDataDirectory := EnvGet("LOCALAPPDATA")
        if localApplicationDataDirectory {
            candidateBaseDirectories.Push(localApplicationDataDirectory)
            candidateBaseDirectories.Push(localApplicationDataDirectory . "\Programs")
        }
    }

    for baseDirectory in candidateBaseDirectories {
        candidatePath := baseDirectory . "\" . directoryName . "\" . executableName

        if FileExist(candidatePath) {
            executablePath := candidatePath
            break
        }
    }

    return executablePath
}

ExecutablePathViaRegistry(executablePath, executableName, registryKeyPaths) {
    if executablePath !== "" {
        return executablePath
    }

    if registryKeyPaths.Length !== 0 {
        originalRegistryKeyPaths := registryKeyPaths.Clone()
        registryKeyPaths := []

        for registryKeyPath in originalRegistryKeyPaths {
            registryKeyPaths.Push("HKCU\Software\" . registryKeyPath)
            registryKeyPaths.Push("HKLM\Software\" . registryKeyPath)
            registryKeyPaths.Push("HKLM\Software\WOW6432Node\" . registryKeyPath)
        }
    }

    registryKeyPaths.Push("HKCU\Software\Microsoft\Windows\CurrentVersion\App Paths")
    registryKeyPaths.Push("HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths")
    registryKeyPaths.Push("HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths")

    for registryKeyPath in registryKeyPaths {
        installFolder := ""

        try {
            installFolder := RegRead(registryKeyPath, "Path")
        }

        if installFolder {
            candidatePath := RTrim(installFolder, "\/") . "\" . executableName

            if FileExist(candidatePath) {
                executablePath := candidatePath
                break
            }
        }
    }

    return executablePath
}

ExecutablePathViaUninstall(executablePath, executableName) {
    if executablePath !== "" {
        return executablePath
    }

    for uninstallBaseKeyPath in [
        "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ] {
        Loop Reg, uninstallBaseKeyPath, "K" {
            uninstallSubKeyPath := A_LoopRegKey . "\" . A_LoopRegName

            displayName := ""
            try {
                displayName := RegRead(uninstallSubKeyPath, "DisplayName")
            }

            if !(displayName && InStr(StrLower(displayName), StrLower(StrReplace(executableName, ".exe", "")))) {
                continue
            }

            displayIcon := ""
            try {
                displayIcon := RegRead(uninstallSubKeyPath, "DisplayIcon")
            }

            if displayIcon {
                pathToExecutable := Trim(StrSplit(displayIcon, ",")[1], ' "')
                if FileExist(pathToExecutable) && (SubStr(StrLower(pathToExecutable), -StrLen(executableName)) = StrLower(executableName)) {
                    executablePath := pathToExecutable
                }
            }

            installLocation := ""
            try {
                installLocation := RegRead(uninstallSubKeyPath, "InstallLocation")
            }

            if installLocation {
                pathToExecutable := RTrim(installLocation, "\/") . "\" . executableName
                if FileExist(pathToExecutable) {
                    executablePath := pathToExecutable
                }
            }
        }
    }

    return executablePath
}

ResolveFactsForApplication(applicationName) {
    global applicationRegistry

    applicationRegistry[applicationName]["Executable Hash"]     := EncodeSha256HexToBase80(Hash.File("SHA256", applicationRegistry[applicationName]["Executable Path"]))
    applicationRegistry[applicationName]["Executable Version"]  := FileGetVersion(applicationRegistry[applicationName]["Executable Path"])
    applicationRegistry[applicationName]["Binary Type"]         := DetermineWindowsBinaryType(applicationName)

    SplitPath(applicationRegistry[applicationName]["Executable Path"], &executableFilename)
    applicationRegistry[applicationName]["Executable Filename"] := executableFilename

    switch applicationName
    {
        case "Excel":
            personalMacroWorkbookPath := EnvGet("AppData") . "\Microsoft\Excel\XLSTART\PERSONAL.XLSB"
            if FileExist(personalMacroWorkbookPath) {
                applicationRegistry["Excel"]["Personal Macro Workbook"] := "Enabled"
            } else {
                applicationRegistry["Excel"]["Personal Macro Workbook"] := "Disabled"
            }

            Run('"' . applicationRegistry["Excel"]["Executable Path"] . '" /e', , , &excelProcessIdentifier)
            excelApplication := WaitForExcelToLoad(excelProcessIdentifier)
            excelWorkbook    := excelApplication.Workbooks.Add()

            excelMacroTestCode := 'Sub CellPing(): Range("A1").Value = "Cell": End Sub'

            ExcelScriptExecution(excelMacroTestCode, true)
            Sleep(160)
            SendEvent("^{F4}") ; CTRL+F4 (Close Window: Module)
            Sleep(160)

            if excelApplication.ActiveSheet.Range("A1").Value = "Cell" {
                applicationRegistry["Excel"]["Code Execution"] := "Basic"

                excelApplication.ActiveSheet.Range("A1").Value := ""

                ExcelScriptExecution(excelMacroTestCode)
                Sleep(160)

                if excelApplication.ActiveSheet.Range("A1").Value = "Cell" {
                    applicationRegistry["Excel"]["Code Execution"] := "Full"
                }
            } else {
                applicationRegistry["Excel"]["Code Execution"] := "Failed"
            }

            excelWorkbook.Close(false)
            excelApplication.DisplayAlerts := false
            excelApplication.Quit()

            excelWorkbook := 0
            excelApplication := 0
    }
}

DetermineWindowsBinaryType(applicationName) {
    static SCS_32BIT_BINARY := 0
    static SCS_DOS_BINARY   := 1
    static SCS_WOW_BINARY   := 2
    static SCS_PIF_BINARY   := 3
    static SCS_POSIX_BINARY := 4
    static SCS_OS2_BINARY   := 5
    static SCS_64BIT_BINARY := 6

    classificationResult := "N/A"
    scsCode := 0

    callSucceeded := DllCall("Kernel32\GetBinaryTypeW", "str", applicationRegistry[applicationName]["Executable Path"], "uint*", &scsCode, "int")

    if callSucceeded != 0 {
        switch scsCode {
            case SCS_32BIT_BINARY:
                classificationResult := "32-bit"
            case SCS_64BIT_BINARY:
                classificationResult := "64-bit"
            case SCS_DOS_BINARY:
                classificationResult := "DOS"
            case SCS_WOW_BINARY:
                classificationResult := "Windows 16-bit"
            case SCS_PIF_BINARY:
                classificationResult := "PIF"
            case SCS_POSIX_BINARY:
                classificationResult := "POSIX"
            case SCS_OS2_BINARY:
                classificationResult := "OS/2"
        }
    }

    return classificationResult
}

CloseApplication(applicationName) {
    static methodName := RegisterMethod("CloseApplication(applicationName As String [Type: Search Open])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Close Application (" . applicationName . ")", methodName, [applicationName])

    try {
        if !applicationRegistry.Has(applicationName) {
            throw Error("Application " . Chr(34) . applicationName . Chr(34) . " invalid.")
        }
    } catch as missingApplicationError {
        LogInformationConclusion("Failed", logValuesForConclusion, missingApplicationError)
    }

    SplitPath(applicationRegistry[applicationName]["Executable Path"], &executableName)
    ProcessClose(executableName)
    ProcessWaitClose(executableName, 4)

    LogInformationConclusion("Completed", logValuesForConclusion)
}

ValidateApplicationFact(applicationName, factName, factValue) {
    static methodName := RegisterMethod("ValidateApplicationFact(applicationName As String, factName As String, factValue As String)" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Validate Application Fact (" . applicationName . ", " . factName . ", " . factValue . ")", methodName, [applicationName, factName, factValue])

    try {
        if !applicationRegistry[applicationName].Has(factName) {
            throw Error("Application " . Chr(34) . applicationName . Chr(34) . " does not have a valid fact name: " . Chr(34) . factName . Chr(34))
        }
    } catch as factNameMissingError {
        LogInformationConclusion("Failed", logValuesForConclusion, factNameMissingError)
    }

    try {
        if applicationRegistry[applicationName][factName] !== factValue {
            throw Error("Application " . Chr(34) . applicationName . Chr(34) . " with fact name of " . Chr(34) . factName . Chr(34) . " does not match fact value of: " . Chr(34) . factValue . Chr(34))
        }
    } catch as factValueMissingError {
        LogInformationConclusion("Failed", logValuesForConclusion, factValueMissingError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Excel                        ;
; **************************** ;

ExcelExtensionRun(documentName, saveDirectory, code, displayName := "", aboutRange := "", aboutCondition := "") {
    static methodName := RegisterMethod("ExcelExtensionRun(documentName As String [Type: Search], saveDirectory As String [Type: Directory], code As String [Type: Code], displayName As String [Optional], aboutRange As String [Optional] [Type: Search Open], aboutCondition As String [Optional] [Type: Search Open])" . LibraryTag(A_LineFile), A_LineNumber + 7)
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

    Run('"' . applicationRegistry["Excel"]["Executable Path"] . '" "' . excelFilePath . '"', , , &excelProcessIdentifier)
    excelApplication := WaitForExcelToLoad(excelProcessIdentifier)
    excelWorkbook := excelApplication.ActiveWorkbook

    aboutWorksheet := ""
    aboutWorksheetFound := false

    for sheet in excelWorkbook.Worksheets {
        if sheet.Name = "About" {
            aboutWorksheet := sheet
            sheet := 0
            aboutWorksheetFound := true
            break
        }

        sheet := 0
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
            aboutWorksheet   := 0
            excelWorkbook    := 0
            excelApplication := 0

            ExcelScriptExecution(code)
            WaitForExcelToClose(excelProcessIdentifier)

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
                aboutWorksheet   := 0
                excelWorkbook    := 0
                excelApplication := 0

                ExcelScriptExecution(code)
                WaitForExcelToClose(excelProcessIdentifier)

                LogInformationConclusion("Completed", logValuesForConclusion)
            } else {
                activeWorkbook := excelApplication.ActiveWorkbook
                activeWorkbook.Close(false)
                excelApplication.DisplayAlerts := false
                excelApplication.Quit()

                aboutWorksheet   := 0
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

            aboutWorksheet   := 0
            activeWorkbook   := 0
            excelWorkbook    := 0
            excelApplication := 0

            LogInformationConclusion("Skipped", logValuesForConclusion)
        }
    } else {
        aboutWorksheet   := 0
        excelWorkbook    := 0
        excelApplication := 0

        ExcelScriptExecution(code)
        WaitForExcelToClose(excelProcessIdentifier)

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

        Run('"' . applicationRegistry["Excel"]["Executable Path"] . '" /e', , , &excelProcessIdentifier)
        excelApplication := WaitForExcelToLoad(excelProcessIdentifier)
        excelWorkbook := excelApplication.Workbooks.Add()
        excelWorkbook    := 0
        excelApplication := 0
        ExcelScriptExecution(code)
        WaitForExcelToClose(excelProcessIdentifier)

        DeleteFile(sidecarPath) ; Remove sidecar after a successful run.
        LogInformationConclusion("Completed", logValuesForConclusion)
    } else {
        LogInformationConclusion("Skipped", logValuesForConclusion)
    }
}

WaitForExcelToClose(excelProcessIdentifier, maxWaitMinutes := 240, mouseMoveIntervalSec := 120) {
    static methodName := RegisterMethod("WaitForExcelToClose(excelProcessIdentifier As Integer, maxWaitMinutes As Integer [Optional: 240], mouseMoveIntervalSec As Integer [Optional: 120])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Wait for Excel to Close", methodName, [excelProcessIdentifier, maxWaitMinutes, mouseMoveIntervalSec])

    totalSecondsToWait := maxWaitMinutes * 60
    secondsSinceLastMouseMove := 0

    sawProcessExit := false
    Loop totalSecondsToWait {
        if ProcessWaitClose(excelProcessIdentifier, 1) = 0 {
            sawProcessExit := true
            break
        }

        secondsSinceLastMouseMove += 1
        if secondsSinceLastMouseMove >= mouseMoveIntervalSec {
            ; Generate real input (0,0 is a no-op): nudge out and back.
            MouseMove 1, 0, 0, "R"
            MouseMove -1, 0, 0, "R"
            secondsSinceLastMouseMove := 0
        }
    }

    try {
        if sawProcessExit = false {
            throw Error("Excel did not close within " . maxWaitMinutes . " minutes.")
        }
    } catch as excelCloseError {
        LogInformationConclusion("Failed", logValuesForConclusion, excelCloseError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

WaitForExcelToLoad(excelProcessIdentifier) {
    static methodName := RegisterMethod("WaitForExcelToLoad(excelProcessIdentifier As Integer)" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Wait for Excel to Load", methodName, [excelProcessIdentifier])

    try {
        if !WinWait("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier, , 480) {
            throw Error("Excel document did not appear within 480 seconds.")
        }
    } catch as documentNotLoadedError {
        LogInformationConclusion("Failed", logValuesForConclusion, documentNotLoadedError)
    }

    WinActivate("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier)
    WinWaitActive("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier, , 10)

    excelApplication := ComObjActive("Excel.Application")

    LogInformationConclusion("Completed", logValuesForConclusion)
    return excelApplication
}

; **************************** ;
; SQL Server Management Studio ;
; **************************** ;

StartSqlServerManagementStudioAndConnect() {
    static methodName := RegisterMethod("StartSqlServerManagementStudioAndConnect()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Start SQL Server Management Studio and Connect", methodName)

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

    savePath := saveDirectory . filename . ".csv"

    SendInput("^n") ; CTRL+N (Query with Current Connection)
    Sleep(2000)

    PasteCode(code, "--")

    SendInput("!x") ; ALT+X (Execute)
    sqlQuerySuccessfulCoordinates := GetImageCoordinatesFromSegment("SQL Server Management Studio Query executed successfully", "6-26", "88-96", 360)
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

    maximumWaitMilliseconds := 10000
    pollIntervalMilliseconds := 100
    startTickCount := A_TickCount

    fileExistsAlready := !!FileExist(savePath)

    if fileExistsAlready = false {
        while !FileExist(savePath) && (A_TickCount - startTickCount) < maximumWaitMilliseconds {
            Sleep(pollIntervalMilliseconds)
        }
    }

    if fileExistsAlready = true {
        previousModifiedTime := FileGetTime(savePath, "M")

        Sleep(480)

        SendEvent("y") ; Y (Yes)

        startTickCount := A_TickCount
        Sleep(1000)
        while FileGetTime(savePath, "M") = previousModifiedTime && (A_TickCount - startTickCount) < maximumWaitMilliseconds {
            Sleep(pollIntervalMilliseconds)
        }

        try {
            if FileGetTime(savePath, "M") = previousModifiedTime {
                throw Error("Timed out waiting for overwrite: " . savePath)
            }
        } catch as timedOutError {
            LogInformationConclusion("Failed", logValuesForConclusion, timedOutError)
        }

        Sleep(240)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Toad for Oracle              ;
; **************************** ;

ExecuteAutomationApp(appName, runtimeDate := "") {
    static methodName := RegisterMethod("ExecuteAutomationApp(appName As String [Type: Search], runtimeDate As String [Optional] [Type: Raw Date Time])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Verify Toad for Oracle Works", methodName, [appName, runtimeDate])

    static toadExecutableFilename := applicationRegistry["Toad for Oracle"]["Executable Filename"]

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

        if dialogHasAppeared = false {
            ; Phase 1: waiting for the dialog to appear
            if dialogExists != false {
                dialogHasAppeared := true
                firstSeenTickCount := A_TickCount
            } else if A_TickCount - overallStartTickCount >= 2000 {
                ; No dialog ever appeared -> likely instant/local reconnect
                break
            }
        } else {
            ; Phase 2: dialog has appeared; wait until it closes (disappears)
            if dialogExists = false {
                ; Closed = reconnect finished
                break
            }
            try {
                if A_TickCount - firstSeenTickCount >= 30000 {
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

    toadForOracleSearchCoordinates := GetImageCoordinatesFromSegment("Toad for Oracle Search", "0-40", "80-100")
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
    toadForOraclePlayCoordinates := GetImageCoordinatesFromSegment("Toad for Oracle Run selected apps", "0-40", "0-20")

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

; **************************** ;
; Helper Methods               ;
; **************************** ;

CombineExcelCode(introCode, mainCode, outroCode := "") {
    combinedCode := introCode . "`r`n`r`n" . mainCode

    if outroCode !== "" {
        combinedCode := combinedCode . "`r`n`r`n" . outroCode
    } 

    return combinedCode
}