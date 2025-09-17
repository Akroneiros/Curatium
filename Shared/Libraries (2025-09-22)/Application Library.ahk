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

    for applicationName in [
        "7-Zip",
        "Chrome",
        "DevToys",
        "Edge",
        "Everything",
        "Excel",
        "Firefox",
        "KeePass",
        "Notepad++",
        "OBS Studio",
        "Paint Shop Pro",
        "PowerPoint",
        "qBittorrent",
        "SQL Server Management Studio",
        "Toad for Oracle",
        "TrueCrypt",
        "Visual Studio Code",
        "Workstation Pro",
        "WinSCP",
        "Word"
    ] {
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

    switch applicationName
    {
        case "7-Zip":
            executableName      := "7zFM.exe"
            executableDirectory := "7-Zip"
            registryKeyPaths    := [
                "HKCU\Software\7-Zip",
                "HKLM\Software\7-Zip",
                "HKLM\Software\WOW6432Node\7-Zip"
            ]
        case "Chrome":
            executableName      := "chrome.exe"
            executableDirectory := "Google\Chrome\Application"
        case "DevToys":
            executableName      := "DevToys.exe"
            executableDirectory := "DevToys Preview"
        case "Edge":
            executableName      := "msedge.exe"
            executableDirectory := "Microsoft\Edge\Application"
        case "Everything":
            executableName      := "Everything.exe"
            executableDirectory := "Everything"
            registryKeyPaths    := [
                "HKCU\Software\voidtools\Everything",
                "HKLM\Software\voidtools\Everything",
                "HKLM\Software\WOW6432Node\voidtools\Everything"
            ]
        case "Excel":
            executableName      := "EXCEL.EXE"
            executableDirectory := "Microsoft Office\root\Office16"
            registryKeyPaths    := [
                "HKLM\Software\Microsoft\Office\16.0\Common\InstallRoot",
                "HKLM\Software\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot",
                "HKLM\Software\Microsoft\Office\15.0\Common\InstallRoot",
                "HKLM\Software\WOW6432Node\Microsoft\Office\15.0\Common\InstallRoot"
            ]
        case "Firefox":
            executableName      := "firefox.exe"
            executableDirectory := "Mozilla Firefox"
        case "KeePass":
            executableName      := "KeePass.exe"
            executableDirectory := "KeePass Password Safe 2"
        case "Notepad++":
            executableName      := "notepad++.exe"
            executableDirectory := "Notepad++"
        case "OBS Studio":
            executableName      := "obs64.exe"
            executableDirectory := "obs-studio\bin\64bit"
        case "Paint Shop Pro":
            executableName      := "Paint Shop Pro 9.exe"
            executableDirectory := "Jasc Software Inc\Paint Shop Pro 9"
        case "PowerPoint":
            executableName      := "POWERPNT.EXE"
            executableDirectory := "Microsoft Office\root\Office16"
            registryKeyPaths    := [
                "HKLM\Software\Microsoft\Office\16.0\Common\InstallRoot",
                "HKLM\Software\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot",
                "HKLM\Software\Microsoft\Office\15.0\Common\InstallRoot",
                "HKLM\Software\WOW6432Node\Microsoft\Office\15.0\Common\InstallRoot"
            ]
        case "qBittorrent":
            executableName      := "qbittorrent.exe"
            executableDirectory := "qBittorrent"
        case "SQL Server Management Studio":
            executableName      := "SSMS.exe"
            executableDirectory := "Microsoft SQL Server Management Studio 21\Release\Common7\IDE"
            registryKeyPaths    := [
                "HKLM\Software\Microsoft\Microsoft SQL Server Management Studio",
                "HKLM\Software\WOW6432Node\Microsoft\Microsoft SQL Server Management Studio"
            ]
        case "Toad for Oracle":
            executableName      := "Toad.exe"
            executableDirectory := "Quest Software\Toad for Oracle Subscription Edition\Toad for Oracle Subscription"
            registryKeyPaths := [
                "HKLM\Software\Quest Software\Toad for Oracle",
                "HKLM\Software\WOW6432Node\Quest Software\Toad for Oracle",
                "HKLM\Software\Dell\Toad for Oracle",
                "HKLM\Software\WOW6432Node\Dell\Toad for Oracle"
            ]
        case "TrueCrypt":
            executableName      := "TrueCrypt.exe"
            executableDirectory := "TrueCrypt"
        case "Visual Studio Code":
            executableName      := "Code.exe"
            executableDirectory := "Microsoft VS Code"
        case "WinSCP":
            executableName      := "WinSCP.exe"
            executableDirectory := "WinSCP"
        case "Word":
            executableName      := "WINWORD.EXE"
            executableDirectory := "Microsoft Office\root\Office16"
            registryKeyPaths    := [
                "HKLM\Software\Microsoft\Office\16.0\Common\InstallRoot",
                "HKLM\Software\WOW6432Node\Microsoft\Office\16.0\Common\InstallRoot",
                "HKLM\Software\Microsoft\Office\15.0\Common\InstallRoot",
                "HKLM\Software\WOW6432Node\Microsoft\Office\15.0\Common\InstallRoot"
            ]
        case "Workstation Pro":
            executableName      := "vmware.exe"
            executableDirectory := "VMware\VMware Workstation"
    }

    executablePath := ExecutablePathViaDirectory(executablePath, executableName, executableDirectory)
    executablePath := ExecutablePathViaRegistry(executablePath, executableName, registryKeyPaths)
    executablePath := ExecutablePathViaUninstall(executablePath, executableName)

    return executablePath
}

ExecutablePathViaDirectory(executablePath, executableName, directoryName) {
    if executablePath !== "" {
        return executablePath
    }

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
        candidateBaseDirectories.Push(localApplicationDataDirectory . "\Programs")
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
        "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ] {
        Loop Reg, uninstallBaseKeyPath, "K" {
            uninstallSubKeyPath := A_LoopRegKey "\" A_LoopRegName

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
                if FileExist(pathToExecutable) && (SubStr(StrLower(pathToExecutable), -9) = "\" . StrLower(executableName)) {
                    executablePath := pathToExecutable
                }
            }

            installLocation := ""
            try {
                installLocation := RegRead(uninstallSubKeyPath, "InstallLocation")
            }
            if installLocation {
                pathToExecutable := RTrim(installLocation, "\/") "\" . executableName
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

    applicationRegistry[applicationName]["Executable Hash"]    := Hash.File("SHA256", applicationRegistry[applicationName]["Executable Path"])
    applicationRegistry[applicationName]["Executable Version"] := FileGetVersion(applicationRegistry[applicationName]["Executable Path"])
    applicationRegistry[applicationName]["Binary Type"]        := DetermineWindowsBinaryType(applicationName)

    switch applicationName
    {
        case "Excel":
            personalMacroWorkbookPath := EnvGet("AppData") . "\Microsoft\Excel\XLSTART\PERSONAL.XLSB"
            if FileExist(personalMacroWorkbookPath) {
                applicationRegistry["Excel"]["Personal Macro Workbook"] := "Enabled"
            } else {
                applicationRegistry["Excel"]["Personal Macro Workbook"] := "Disabled"
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

            activeWorkbook.Close(false)
            excelApplication.DisplayAlerts := false
            excelApplication.Quit()

            activeWorkbook := 0
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

; ******************** ;
; Excel                ;
; ******************** ;

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

    savePath := saveDirectory . filename . ".csv"

    SendInput("^n") ; CTRL+N (Query with Current Connection)
    Sleep(2000)

    PasteCode(code, "--")

    SendInput("!x") ; ALT+X (Execute)
    sqlQuerySuccessfulCoordinates := RetrieveImageCoordinatesFromSegment("SQL Server Management Studio Query Successful", "6-26", "88-96", 360)
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