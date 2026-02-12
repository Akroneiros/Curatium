#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Base Library.ahk
#Include File Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

global applicationRegistry := Map()

; **************************** ;
; Application Registry         ;
; **************************** ;

DefineApplicationRegistry() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Define Application Registry")

    global applicationRegistry

    mappedApplicationsFilePath := system["Mappings Directory"] . "Applications.csv"
    if !FileExist(mappedApplicationsFilePath) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Applications.csv not found in the directory for Mappings.")
    }
   
    applications := ConvertCsvToArrayOfMaps(mappedApplicationsFilePath)
    for application in applications {
        applicationName    := application["Name"]
        applicationCounter := application["Counter"] + 0

        applicationRegistry[applicationName] := Map()
        applicationRegistry[applicationName]["Counter"] := applicationCounter
    }

    LogConclusion("Completed", logValuesForConclusion)
}

RegisterApplications() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Register Applications")

    global applicationRegistry

    applicationWhitelist := system["Configuration"]["Application Whitelist"]
    applicationWhitelistLength := applicationWhitelist.Length
    if applicationWhitelistLength != 0 {
        for application in applicationWhitelist {
            applicationRegistry[application]["Whitelisted"] := true
        }

        for application in applicationRegistry {
            if !applicationRegistry[application].Has("Whitelisted") {
                applicationRegistry[application]["Whitelisted"] := false
                applicationRegistry[application]["Installed"]   := false
            }
        }
    } else {
        for application in applicationRegistry {
            applicationRegistry[application]["Whitelisted"] := true
        }
    }

    for application in applicationRegistry {
        if applicationWhitelistLength != 0 {
            if applicationRegistry[application]["Whitelisted"] = false {
                continue
            }
        }

        executablePathSearchResult := ExecutablePathResolve(application)
        if Type(executablePathSearchResult) = "Map" {
            applicationRegistry[application]["Executable Path"]   := executablePathSearchResult["Executable Path"]
            applicationRegistry[application]["Resolution Method"] := executablePathSearchResult["Resolution Method"]
            applicationRegistry[application]["Installed"]         := true
        } else {
            applicationRegistry[application]["Installed"]         := false
        }
    }

    CreateApplicationImages()

    for application in applicationRegistry {
        if applicationRegistry[application]["Installed"] = true {
             ResolveFactsForApplication(application, applicationRegistry[application]["Counter"])
        }
    }

    if applicationRegistry["DaVinci Resolve Studio"]["Installed"] && applicationRegistry["DaVinci Resolve"]["Installed"] {
        if applicationRegistry["DaVinci Resolve Studio"]["Executable Path"] = applicationRegistry["DaVinci Resolve"]["Executable Path"] {
            executableDirectory := ExtractDirectory(applicationRegistry["DaVinci Resolve Studio"]["Executable Path"])
            readMeFilePath      := executableDirectory . "Documents\ReadMe.html"

            if FileExist(readMeFilePath) {
                readMeFileContents := ReadFileOnHashMatch(readMeFilePath, Hash.File("SHA256", readMeFilePath))

                notInstalledApplication := "DaVinci Resolve Studio"
                if InStr(readMeFileContents, "About DaVinci Resolve Studio") {
                    notInstalledApplication := "DaVinci Resolve"
                }

                applicationRegistry[notInstalledApplication].Delete("Binary Type")
                applicationRegistry[notInstalledApplication].Delete("Executable Filename")
                applicationRegistry[notInstalledApplication].Delete("Executable Hash")
                applicationRegistry[notInstalledApplication].Delete("Executable Path")
                applicationRegistry[notInstalledApplication].Delete("Executable Version")
                applicationRegistry[notInstalledApplication].Delete("Resolution Method")

                applicationRegistry[notInstalledApplication]["Installed"] := false
            }
        }
    }

    installedApplications := []
    for outerKey, innerMap in applicationRegistry {
        if innerMap["Installed"] {
            configuration := outerKey . "|" . innerMap["Executable Path"] . "|" . innerMap["Executable Hash"] . "|" . innerMap["Executable Version"] . "|" . innerMap["Binary Type"]
            configuration := configuration . "|" . innerMap["Counter"] . "|" . SubStr(innerMap["Resolution Method"], 1, 1)
            installedApplications.Push(configuration)
            innerMap["Executable Hash"] := DecodeBaseToSha256Hex(innerMap["Executable Hash"], 86)
        }
    }

    BatchAppendExecutionLog("Application", installedApplications)

    LogConclusion("Completed", logValuesForConclusion)
}

ExecutablePathResolve(applicationName) {
    static methodName := RegisterMethod("applicationName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [applicationName])

    global applicationRegistry

    executablePathSearchResult := ""

    static combinedApplicationExecutableDirectoryCandidates := unset

    if !IsSet(combinedApplicationExecutableDirectoryCandidates) {
        combinedApplicationExecutableDirectoryCandidates := []

        for projectApplicationExecutableDirectoryCandidate in system["Configuration"]["Application Executable Directory Candidates"] {
            if applicationRegistry[projectApplicationExecutableDirectoryCandidate[1]]["Whitelisted"] = false {
                continue
            }

            combinedApplicationExecutableDirectoryCandidates.Push(Map(
                "Directory", projectApplicationExecutableDirectoryCandidate[3], "Executable", projectApplicationExecutableDirectoryCandidate[2], "Name", projectApplicationExecutableDirectoryCandidate[1], "Source", "Project"
            ))
        }

        sharedApplicationExecutableDirectoryCandidates := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Application Executable Directory Candidates.csv")
        for index, sharedApplicationExecutableDirectoryCandidate in sharedApplicationExecutableDirectoryCandidates {
            if applicationRegistry[sharedApplicationExecutableDirectoryCandidate["Name"]]["Whitelisted"] = false {
                continue
            }

            sharedApplicationExecutableDirectoryCandidates[index]["Source"] := "Shared"
            combinedApplicationExecutableDirectoryCandidates.Push(sharedApplicationExecutableDirectoryCandidate)
        }

        for outerKey in applicationRegistry {
            for applicationExecutableDirectoryCandidate in combinedApplicationExecutableDirectoryCandidates {
                if outerKey != applicationExecutableDirectoryCandidate["Name"] {
                    continue
                }

                for collisionCandidate in combinedApplicationExecutableDirectoryCandidates {
                    if applicationExecutableDirectoryCandidate["Executable"] = collisionCandidate["Executable"] && outerKey != collisionCandidate["Name"] {

                        applicationRegistry[outerKey]["Executable Collision"] := true
                        break 2
                    }
                }
            }
        }
    }

    projectEntryAvailable := false
    relevantApplicationExecutableDirectoryCandidates := []
    for applicationExecutableDirectoryCandidate in combinedApplicationExecutableDirectoryCandidates {
        if applicationExecutableDirectoryCandidate["Name"] = applicationName {
            if applicationExecutableDirectoryCandidate["Source"] = "Project" {
                projectEntryAvailable := true
            }

            filteredApplicationExecutableDirectoryCandidate := Map()
            filteredApplicationExecutableDirectoryCandidate["Directory"] := applicationExecutableDirectoryCandidate["Directory"]
            filteredApplicationExecutableDirectoryCandidate["Executable"] := applicationExecutableDirectoryCandidate["Executable"]

            relevantApplicationExecutableDirectoryCandidates.Push(filteredApplicationExecutableDirectoryCandidate)
        }
    }

    if projectEntryAvailable || applicationRegistry[applicationName].Has("Executable Collision") {
        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaReference(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }

        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaAppPaths(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }

        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaUninstall(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }
    } else {
        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaUninstall(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }

        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaAppPaths(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }

        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaReference(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }
    }

    return executablePathSearchResult
}

ExecutablePathViaReference(applicationName, applicationExecutableDirectoryCandidates) {
    static methodName := RegisterMethod("applicationName As String, applicationExecutableDirectoryCandidates As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [applicationName, applicationExecutableDirectoryCandidates])

    static candidateBaseDirectories := unset
    if !IsSet(candidateBaseDirectories) {
        candidateBaseDirectories := []

        configurationCandidateBaseDirectories := system["Configuration"]["Candidate Base Directories"]
        for configurationCandidateBaseDirectory in configurationCandidateBaseDirectories {
            if DirExist(configurationCandidateBaseDirectory) {
                candidateBaseDirectories.Push(configurationCandidateBaseDirectory)
            }
        }

        defaultCandidateBaseDirectories := [EnvGet("LOCALAPPDATA"), EnvGet("LOCALAPPDATA") . "\Programs", EnvGet("ProgramFiles"), EnvGet("ProgramFiles(x86)"), EnvGet("ProgramW6432"), EnvGet("SystemDrive") . "\", EnvGet("USERPROFILE")]
        for defaultCandidateBaseDirectory in defaultCandidateBaseDirectories {
            if DirExist(defaultCandidateBaseDirectory) {
                candidateBaseDirectories.Push(defaultCandidateBaseDirectory)
            }
        }

        candidateBaseDirectories := RemoveDuplicatesFromArray(candidateBaseDirectories)
    }

    executablePathSearchResult := ""
    for applicationExecutableDirectoryCandidate in applicationExecutableDirectoryCandidates {
        directoryName := applicationExecutableDirectoryCandidate["Directory"]
        executableName := applicationExecutableDirectoryCandidate["Executable"]

        baseDirectoriesWithDirectoryNames := []
        for candidateBaseDirectory in candidateBaseDirectories {
            pathToExecutable := candidateBaseDirectory . "\" . directoryName . "\" . executableName

            if DirExist(candidateBaseDirectory . "\" . directoryName . "\") {
                baseDirectoriesWithDirectoryNames.Push(candidateBaseDirectory . "\" . directoryName . "\")
            }

            if FileExist(pathToExecutable) {
                executablePathSearchResult := pathToExecutable
                break
            }
        }

        if executablePathSearchResult = "" {
            extensionPosition   := InStr(executableName, ".", , -1)
            executableExtension := SubStr(executableName, extensionPosition)
            extensionLength     := StrLen(executableExtension)
            
            for baseDirectoryWithName in baseDirectoriesWithDirectoryNames {
                if executablePathSearchResult != "" {
                    break
                }

                applicationDirectoryFileList := GetFilesFromDirectory(baseDirectoryWithName)

                if applicationDirectoryFileList.Length = 0 {
                    continue
                } else {
                    for filePath in applicationDirectoryFileList {
                        filename := ExtractFilename(filePath)

                        if InStr(filename, applicationName) && SubStr(filename, -extensionLength) = executableExtension {
                            executablePathSearchResult := filePath
                            break
                        }
                    }
                }
            }
        }

        StrReplace(directoryName, ".", "", , &dotOccurrencesInDirectoryNameCount)
        if executablePathSearchResult = "" && dotOccurrencesInDirectoryNameCount >= 2 {
            directoryNameSegments := StrSplit(directoryName, "\")
            versionSegmentIndex := 0
            versionSegment := ""

            for index, directoryNameSegment in directoryNameSegments {
                StrReplace(directoryNameSegment, ".", "", , &dotOccurrencesInDirectoryNameSegmentCount)

                if dotOccurrencesInDirectoryNameSegmentCount < 2 {
                    continue
                }

                versionSegmentIndex := index
                versionSegment := directoryNameSegment
                break
            }

            if versionSegmentIndex != 0 {
                relativePathBeforeVersionSegment := ""
                relativePathAfterVersionSegment  := ""

                for index, directoryNameSegment in directoryNameSegments {
                    if index < versionSegmentIndex {
                        if relativePathBeforeVersionSegment != "" {
                            relativePathBeforeVersionSegment .= "\"
                        }
                        relativePathBeforeVersionSegment .= directoryNameSegment
                    } else if index > versionSegmentIndex {
                        if relativePathAfterVersionSegment != "" {
                            relativePathAfterVersionSegment .= "\"
                        }
                        relativePathAfterVersionSegment .= directoryNameSegment
                    }
                }

                highestVersionKey := ""
                highestVersionExecutablePath := ""

                for candidateBaseDirectory in candidateBaseDirectories {
                    applicationRootDirectory := candidateBaseDirectory
                    if relativePathBeforeVersionSegment != "" {
                        applicationRootDirectory .= "\" . relativePathBeforeVersionSegment
                    }
                    applicationRootDirectory .= "\"

                    if !DirExist(applicationRootDirectory) {
                        continue
                    }

                    loop files, applicationRootDirectory . "*", "D" {
                        folderName := A_LoopFileName

                        StrReplace(folderName, ".", "", , &dotOccurrencesInFolderName)
                        if dotOccurrencesInFolderName < 2 {
                            continue
                        }

                        firstDigitPositionInFolderName := RegExMatch(folderName, "\d")
                        if firstDigitPositionInFolderName = 0 {
                            continue
                        }

                        versionText := SubStr(folderName, firstDigitPositionInFolderName)
                        if !RegExMatch(versionText, "^\d+(?:\.\d+)*$") {
                            continue
                        }

                        versionKey := ""
                        for versionPartIndex, versionPart in StrSplit(versionText, ".") {
                            versionKey .= Format("{:06}", Number(versionPart))
                        }

                        pathToExecutable := A_LoopFileFullPath
                        if relativePathAfterVersionSegment != "" {
                            pathToExecutable .= "\" . relativePathAfterVersionSegment
                        }
                        pathToExecutable .= "\" . executableName

                        if !FileExist(pathToExecutable) {
                            continue
                        }

                        if highestVersionKey = "" || StrCompare(versionKey, highestVersionKey) > 0 {
                            highestVersionKey := versionKey
                            highestVersionExecutablePath := pathToExecutable
                        }
                    }
                }

                if highestVersionExecutablePath != "" {
                    executablePathSearchResult := highestVersionExecutablePath
                }
            }
        }
    }

    if executablePathSearchResult != "" {
        executablePathSearchResult := Map(
            "Executable Path",   executablePathSearchResult,
            "Resolution Method", "Reference"
        )
    }

    return executablePathSearchResult
}

ExecutablePathViaAppPaths(applicationName, applicationExecutableDirectoryCandidates) {
    static methodName := RegisterMethod("applicationName As String, applicationExecutableDirectoryCandidates As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [applicationName, applicationExecutableDirectoryCandidates])

    static appPathsBaseRegistryKeys := [
        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths",
        "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths",
        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"
    ]

    executablePathSearchResult := ""
    for applicationExecutableDirectoryCandidate in applicationExecutableDirectoryCandidates {
        directoryName := applicationExecutableDirectoryCandidate["Directory"]
        executableName := applicationExecutableDirectoryCandidate["Executable"]

        for appPathsBaseRegistryKey in appPathsBaseRegistryKeys {
            subkeyPath := appPathsBaseRegistryKey . "\" . executableName

            pathToExecutable := ""
            try {
                pathToExecutable := RegRead(subkeyPath, "")
            }

            if pathToExecutable {
                if FileExist(pathToExecutable) && ExtractFilename(pathToExecutable) = executableName {
                    if applicationRegistry[applicationName].Has("Executable Collision") {
                        if !InStr(pathToExecutable, applicationName) {
                            continue
                        }
                    }

                    executablePathSearchResult := pathToExecutable
                    break
                }
            }
        }
    }

    if executablePathSearchResult != "" {
        executablePathSearchResult := Map(
            "Executable Path",   executablePathSearchResult,
            "Resolution Method", "App Paths"
        )
    }

    return executablePathSearchResult
}

ExecutablePathViaUninstall(applicationName, applicationExecutableDirectoryCandidates) {
    static methodName := RegisterMethod("applicationName As String, applicationExecutableDirectoryCandidates As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [applicationName, applicationExecutableDirectoryCandidates])

    static uninstallBaseKeyPaths := [
        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    executablePathSearchResult := ""
    for applicationExecutableDirectoryCandidate in applicationExecutableDirectoryCandidates {
        directoryName := applicationExecutableDirectoryCandidate["Directory"]
        executableName := applicationExecutableDirectoryCandidate["Executable"]

        if StrLen(applicationName) < 4 || StrLen(executableName) < 8 {
            continue
        } else {
            requiredLength := 4
            applicationNamePartiallyMatchesExecutableNameCondition := false
            
            executableNameWithoutExtension := ExtractFilename(executableName, true)
            shorterText   := StrLower(applicationName)
            longerText    := StrLower(executableNameWithoutExtension)
            shorterLength := StrLen(shorterText)
            longerLength  := StrLen(longerText)
            if shorterLength > longerLength {
                temporarySwapHolder := shorterText
                shorterText := longerText
                longerText  := temporarySwapHolder
                shorterLength := StrLen(shorterText)
                longerLength  := StrLen(longerText)
            }

            maximumStartIndex := shorterLength - requiredLength + 1
            loop maximumStartIndex {
                currentStartIndex := A_Index
                substringToSearch := SubStr(shorterText, currentStartIndex, requiredLength)
                if InStr(longerText, substringToSearch) {
                    applicationNamePartiallyMatchesExecutableNameCondition := true
                    break
                }
            }

            if !applicationNamePartiallyMatchesExecutableNameCondition {
                continue
            }

            for uninstallBaseKeyPath in uninstallBaseKeyPaths {
                loop reg, uninstallBaseKeyPath, "K" {
                    uninstallSubKeyPath := A_LoopRegKey . "\" . A_LoopRegName

                    displayName := ""
                    try {
                        displayName := RegRead(uninstallSubKeyPath, "DisplayName")
                    }

                    if !(displayName && InStr(StrLower(displayName), StrLower(executableNameWithoutExtension))) {
                        continue
                    }

                    displayIcon := ""
                    try {
                        displayIcon := RegRead(uninstallSubKeyPath, "DisplayIcon")
                    }

                    if displayIcon {
                        pathToExecutable := Trim(StrSplit(displayIcon, ",")[1], ' "')
                        if FileExist(pathToExecutable) && (SubStr(StrLower(pathToExecutable), -StrLen(executableName)) = StrLower(executableName)) {
                            if applicationRegistry[applicationName].Has("Executable Collision") {
                                if !InStr(pathToExecutable, applicationName) {
                                    continue
                                }
                            }

                            executablePathSearchResult := pathToExecutable
                            break 2
                        }
                    }

                    installLocation := ""
                    try {
                        installLocation := RegRead(uninstallSubKeyPath, "InstallLocation")
                    }

                    if installLocation {
                        pathToExecutable := RTrim(installLocation, "\/") . "\" . executableName
                        if FileExist(pathToExecutable) {
                            if applicationRegistry[applicationName].Has("Executable Collision") {
                                if !InStr(pathToExecutable, applicationName) {
                                    continue
                                }
                            }

                            executablePathSearchResult := pathToExecutable
                            break 2
                        }
                    }
                }
            }
        }
    }

    if executablePathSearchResult != "" {
        executablePathSearchResult := Map(
            "Executable Path",   executablePathSearchResult,
            "Resolution Method", "Uninstall"
        )
    }

    return executablePathSearchResult
}

ResolveFactsForApplication(applicationName, counter) {
    static methodName := RegisterMethod("applicationName As String, counter As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [applicationName, counter])

    global applicationRegistry

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Excel Tiny Delay", 16, false)
        SetMethodSetting(methodName, "Excel Short Delay", 208, false)
        SetMethodSetting(methodName, "Excel Medium Delay", 540, false)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    executableHash := Hash.File("SHA256", applicationRegistry[applicationName]["Executable Path"])
    executableHash := EncodeSha256HexToBase(executableHash, 86)

    executableVersion := "N/A"
    try {
        executableVersion := FileGetVersion(applicationRegistry[applicationName]["Executable Path"])
    }

    binaryType := DetermineWindowsBinaryType(applicationRegistry[applicationName]["Executable Path"])

    applicationRegistry[applicationName]["Executable Hash"]    := executableHash
    applicationRegistry[applicationName]["Executable Version"] := executableVersion
    applicationRegistry[applicationName]["Binary Type"]        := binaryType

    SplitPath(applicationRegistry[applicationName]["Executable Path"], &executableFilename)
    applicationRegistry[applicationName]["Executable Filename"] := executableFilename

    static commandLineExecutables := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Command Line Executables.csv")
    for commandLineExecutable in commandLineExecutables {
        if applicationName = commandLineExecutable["Name"] {
            directoryPath := ExtractDirectory(applicationRegistry[applicationName]["Executable Path"])

            if FileExist(directoryPath . commandLineExecutable["Command Line Executable"]) {
                applicationRegistry[applicationName]["Command Line Executable Path"] := directoryPath . commandLineExecutable["Command Line Executable"]
            }
        }
    }

    switch applicationName {
        case "CyberChef":
            cyberChefHtml := ReadFileOnHashMatch(applicationRegistry[applicationName]["Executable Path"], DecodeBaseToSha256Hex(applicationRegistry[applicationName]["Executable Hash"], 86))

            versionPattern := "i)CyberChef\s+version:\s*(\d+(?:\.\d+)*)"
            if RegExMatch(cyberChefHtml, versionPattern, &versionMatch) {
                applicationRegistry[applicationName]["Executable Version"] := versionMatch[1]
            }
        case "Excel":
            excelTinyDelay   := settings.Get("Excel Tiny Delay")
            excelShortDelay  := settings.Get("Excel Short Delay")
            excelMediumDelay := settings.Get("Excel Medium Delay")

            CloseApplication("Excel")

            excelApplication := ComObject("Excel.Application")
            excelWorkbook    := excelApplication.Workbooks.Add()
            excelWorksheet   := excelWorkbook.ActiveSheet
            excelApplication.Visible := true

            excelWindowHandle := excelApplication.Hwnd
            while !excelWindowHandle := excelApplication.Hwnd {
                Sleep(excelTinyDelay)
            }
            excelProcessIdentifier := WinGetPID("ahk_id " . excelWindowHandle)

            excelMainWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["Excel"]["Executable Filename"] . " ahk_class XLMAIN", 60)
            ActivateWindow(excelMainWindowSearchResults, true)

            personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
            if !FileExist(personalMacroWorkbookPath) {
                KeyboardShortcut("ALT", "Q") ; Microsoft Search
                Sleep(excelMediumDelay)
                PasteText("Record Macro")
                Sleep(excelMediumDelay)
                SendInput("{Down}") ; Record Macro: Select
                Sleep(excelShortDelay)
                SendInput("{Enter}") ; Record Macro: Apply
                excelRecordMacroWindowSearchResults := SearchForWindow("Record Macro ahk_exe " . applicationRegistry["Excel"]["Executable Filename"], 60)
                ActivateWindow(excelRecordMacroWindowSearchResults)
                SendInput("{Tab}") ; Shortcut key:
                Sleep(excelShortDelay)
                SendInput("{Tab}") ; Store macro in:
                Sleep(excelShortDelay)
                SendInput("{Up}") ; Activate list
                Sleep(excelShortDelay)
                SendInput("{Up}") ; This Workbook -> New Workbook
                Sleep(excelShortDelay)
                SendInput("{Up}") ; New Workbook -> Personal Macro Workbook
                Sleep(excelShortDelay)
                SendInput("{Enter}") ; Apply
                Sleep(excelShortDelay)
                SendInput("{Tab}") ; Description:
                Sleep(excelShortDelay)
                SendInput("{Tab}") ; OK
                Sleep(excelShortDelay)
                SendInput("{Enter}") ; OK: Apply
                Sleep(excelMediumDelay)
                ActivateWindow(excelMainWindowSearchResults)
                KeyboardShortcut("ALT", "Q") ; Microsoft Search
                Sleep(excelMediumDelay)
                PasteText("Record Macro")
                Sleep(excelMediumDelay)
                SendInput("{Down}") ; Record Macro: Select
                Sleep(excelShortDelay)
                SendInput("{Enter}") ; Record Macro: Stop Recording
                Sleep(excelShortDelay)
                KeyboardShortcut("ALT", "F11") ; Open the Visual Basic editor.
                visualBasicEditorWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["Excel"]["Executable Filename"] . " ahk_class wndclass_desked_gsk", 60, "Failed to open the Visual Basic editor via ALT+F11 in Excel.")
                ActivateWindow(visualBasicEditorWindowSearchResults)
                Sleep(excelShortDelay)
                KeyboardShortcut("CTRL", "R") ; Project Explorer
                Sleep(excelShortDelay)
                SendInput("{Down}") ; Sheet1 (Sheet1) -> ThisWorkbook
                Sleep(excelShortDelay)
                SendInput("{Down}") ; ThisWorkbook -> VBAProject (PERSONAL.XLSB)
                Sleep(excelShortDelay)
                SendInput("{Right}") ; VBAProject (PERSONAL.XLSB): Expand
                Sleep(excelShortDelay)
                SendInput("{Down}") ; Microsoft Excel Objects
                Sleep(excelShortDelay)
                SendInput("{Down}") ; Modules
                Sleep(excelShortDelay)
                SendInput("{Right}") ; Modules: Expand
                Sleep(excelShortDelay)
                SendInput("{Down}") ; Module1
                Sleep(excelShortDelay)
                SendInput("{Enter}") ; Module1: Open
                Sleep(excelShortDelay)
                PasteText("Sub Macro()" . "`r`n`r`n" . "End Sub", "'")

                Loop excelApplication.Workbooks.Count {
                    currentWorkbook := excelApplication.Workbooks.Item(A_Index)
                    if personalMacroWorkbookPath = currentWorkbook.FullName {
                        currentWorkbook.Save()
                        break
                    }
                }

                Sleep(excelTinyDelay)
                KeyboardShortcut("ALT", "Q") ; Close and Return to Microsoft Excel
                Sleep(excelShortDelay)
            }

            excelMacroCode := "Sub Run()" . "`r`n" . '    Range("A1").Value = "Cell"' . "`r`n" . "End Sub"
            OpenVisualBasicEditorAndRunCode(excelMacroCode, excelApplication)
            Sleep(excelTinyDelay + excelTinyDelay)

            if excelWorksheet.Range("A1").Value != "Cell" {
                LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to execute Excel Macro Code.")
            }

            applicationRegistry["Excel"]["International"] := Map()

            excelInternational := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Excel International (2025-09-26).csv")
            for international in excelInternational {
                applicationRegistry["Excel"]["International"][international["Label"]] := excelApplication.International[international["Value"]]
            }

            excelWorkbook.Close(false)
            excelApplication.DisplayAlerts := false
            excelApplication.Quit()

            excelWorksheet   := 0
            excelWorkbook    := 0
            excelApplication := 0
            ProcessWaitClose(excelProcessIdentifier, 2)
        case "Word":
            wordApplication := ComObject("Word.Application")

            applicationRegistry["Word"]["International"] := Map()

            wordInternational := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Word International (2025-09-26).csv")

            for international in wordInternational {
                applicationRegistry["Word"]["International"][international["Label"]] := wordApplication.International[international["Value"]]
            }

            wordApplication.Quit()
            wordApplication := 0
    }
}

ValidateApplicationFact(applicationName, factName, factValue) {
    static methodName := RegisterMethod("applicationName As String, factName As String, factValue As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [applicationName, factName, factValue], "Validate Application Fact (" . applicationName . ", " . factName . ", " . factValue . ")")

    if !applicationRegistry[applicationName].Has(factName) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, 'Application "' . applicationName . '" does not have a valid fact name: ' . factName)
    }

    if applicationRegistry[applicationName][factName] !== factValue {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, 'Application "' . applicationName . '" with fact name of "' . factName . '" does not match fact value of: ' . factValue)
    }

    LogConclusion("Completed", logValuesForConclusion)
}

ValidateApplicationInstalled(applicationName) {
    static methodName := RegisterMethod("applicationName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [applicationName], "Validate Application Installed (" . applicationName . ")")

    if !applicationRegistry.Has(applicationName) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Application doesn't exist: " . applicationName)
    }

    if !applicationRegistry[applicationName]["Installed"] {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Application not installed: " . applicationName)
    }

    applicationIsInstalled := true

    LogConclusion("Completed", logValuesForConclusion)
    return applicationIsInstalled
}

; **************************** ;
; Shared                       ;
; **************************** ;

CloseApplication(applicationName) {
    static methodName := RegisterMethod("applicationName As String [Constraint: Locator]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [applicationName], "Close Application (" . applicationName . ")")

    if !applicationRegistry.Has(applicationName) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Application not found: " . applicationName)
    }

    executableName := applicationRegistry[applicationName]["Executable Filename"]

    if !ProcessExist(executableName) {
        LogConclusion("Skipped", logValuesForConclusion)
    } else {
        ProcessClose(executableName)
        ProcessWaitClose(executableName, 4)

        LogConclusion("Completed", logValuesForConclusion)
    }
}

; **************************** ;
; Excel                        ;
; **************************** ;

ExcelExtensionRun(documentName, saveDirectory, code, displayName := "", aboutRange := "", aboutCondition := "") {
    static methodName := RegisterMethod("documentName As String [Constraint: Locator], saveDirectory As String [Constraint: Directory], code As String [Constraint: Summary], displayName As String [Optional], aboutRange As String [Optional] [Constraint: Locator], aboutCondition As String [Optional] [Constraint: Locator]", A_ThisFunc, A_LineFile, A_LineNumber + 7)
    overlayValue := ""
    if displayName = "" {
        overlayValue := documentName . " Excel Extension Run"
    } else {
        overlayValue := displayName . " Excel Extension Run"
    }
    logValuesForConclusion := LogBeginning(methodName, [documentName, saveDirectory, code, displayName, aboutRange, aboutCondition], overlayValue)

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Tiny Delay", 32, false)
        SetMethodSetting(methodName, "Short Delay", 260, false)

        defaultMethodSettingsSet := true
    }

    settings    := methodRegistry[methodName]["Settings"]
    tinyDelay   := settings.Get("Tiny Delay")
    shortDelay  := settings.Get("Short Delay")

    excelFilePath := FileExistsInDirectory(documentName, saveDirectory, "xlsx")
    if excelFilePath = "" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "documentName not found: " . documentName)
    }

    excelApplication := ComObject("Excel.Application")
    excelApplication.Workbooks.Open(excelFilePath, 0)
    excelWorkbook    := excelApplication.ActiveWorkbook
    excelApplication.Visible := true

    static personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
    excelApplication.Workbooks.Open(personalMacroWorkbookPath)

    excelApplication.CalculateUntilAsyncQueriesDone()
    while excelApplication.CalculationState != 0 {
        Sleep(tinyDelay + tinyDelay)
    }

    excelWindowHandle := excelApplication.Hwnd
    while !excelWindowHandle := excelApplication.Hwnd {
        Sleep(tinyDelay)
    }
    excelProcessIdentifier := WinGetPID("ahk_id " . excelWindowHandle)

    aboutWorksheet      := ""
    aboutWorksheetFound := false

    for worksheet in excelWorkbook.Worksheets {
        if worksheet.Name = "About" {
            aboutWorksheet      := worksheet
            worksheet           := 0
            aboutWorksheetFound := true
            break
        }

        worksheet := 0
    }

    if (aboutRange != "" || aboutCondition != "") && !aboutWorksheetFound {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Worksheet About not found with arguments passed in.")
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

    if aboutWorksheetFound && (aboutRange != "" || aboutCondition != "") {
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
            OpenVisualBasicEditorAndRunCode(code, excelApplication)
            WaitForExcelToClose(excelProcessIdentifier)
            aboutWorksheet   := 0
            excelWorkbook    := 0
            excelApplication := 0
            ProcessWaitClose(excelProcessIdentifier, 2)

            LogConclusion("Completed", logValuesForConclusion)
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
                OpenVisualBasicEditorAndRunCode(code, excelApplication)
                WaitForExcelToClose(excelProcessIdentifier)
                aboutWorksheet   := 0
                excelWorkbook    := 0
                excelApplication := 0
                ProcessWaitClose(excelProcessIdentifier, 2)

                LogConclusion("Completed", logValuesForConclusion)
            } else {
                activeWorkbook := excelApplication.ActiveWorkbook
                activeWorkbook.Close(false)
                excelApplication.DisplayAlerts := false
                excelApplication.Quit()

                aboutWorksheet   := 0
                activeWorkbook   := 0
                excelWorkbook    := 0
                excelApplication := 0
                Sleep(shortDelay * 4)

                LogConclusion("Skipped", logValuesForConclusion)
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
            Sleep(shortDelay * 4)

            LogConclusion("Skipped", logValuesForConclusion)
        }
    } else {
        OpenVisualBasicEditorAndRunCode(code, excelApplication)
        WaitForExcelToClose(excelProcessIdentifier)
        aboutWorksheet   := 0
        excelWorkbook    := 0
        excelApplication := 0
        ProcessWaitClose(excelProcessIdentifier, 2)

        LogConclusion("Completed", logValuesForConclusion)
    }
}

ExcelStartingRun(documentName, saveDirectory, code, displayName := "") {
    static methodName := RegisterMethod("documentName As String [Constraint: Locator], saveDirectory As String [Constraint: Directory], code As String [Constraint: Summary], displayName As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 7)
    overlayValue := ""
    if displayName = "" {
        overlayValue := documentName . " Excel Starting Run"
    } else {
        overlayValue := displayName . " Excel Starting Run"
    }
    logValuesForConclusion := LogBeginning(methodName, [documentName, saveDirectory, code, displayName], overlayValue)

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Run Attempts", 4, false)
        SetMethodSetting(methodName, "Tiny Delay", 32, false)
        SetMethodSetting(methodName, "Short Delay", 260, false)

        defaultMethodSettingsSet := true
    }

    settings    := methodRegistry[methodName]["Settings"]
    runAttempts := settings.Get("Run Attempts")
    tinyDelay   := settings.Get("Tiny Delay")
    shortDelay  := settings.Get("Short Delay")

    xlsxPath := FileExistsInDirectory(documentName, saveDirectory, "xlsx")
    txtPath  := FileExistsInDirectory(documentName, saveDirectory, "txt")

    if txtPath != "" && xlsxPath != "" {
        DeleteFile(txtPath)
        DeleteFile(xlsxPath)
        xlsxPath := ""
    } else {
        if txtPath != "" {
            DeleteFile(txtPath)
        }
    }

    if xlsxPath != "" {
        LogConclusion("Skipped", logValuesForConclusion)
    } else {
        sidecarPath := saveDirectory . documentName . ".txt"
        WriteTextIntoFile("", sidecarPath, "UTF-8")

        excelApplication := ComObject("Excel.Application")
        excelWorkbook    := excelApplication.Workbooks.Add()
        excelApplication.Visible := true

        excelWindowHandle := excelApplication.Hwnd
        while !excelWindowHandle := excelApplication.Hwnd {
            Sleep(tinyDelay)
        }
        excelProcessIdentifier := WinGetPID("ahk_id " . excelWindowHandle)

        OpenVisualBasicEditorAndRunCode(code, excelApplication)
        WaitForExcelToClose(excelProcessIdentifier)
        excelWorkbook    := 0
        excelApplication := 0
        ProcessWaitClose(excelProcessIdentifier, 2)

        DeleteFile(sidecarPath) ; Remove sidecar after a successful run.
        LogConclusion("Completed", logValuesForConclusion)
    }
}

OpenVisualBasicEditorAndRunCode(code, excelApplication) {
    static methodName := RegisterMethod("code As String [Constraint: Summary], excelApplication As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [code, excelApplication], "Open Visual Basic Editor and Run Code (Length: " . StrLen(code) . ")")

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Run Attempts", 4, false)
        SetMethodSetting(methodName, "Tiny Delay", 32, false)
        SetMethodSetting(methodName, "Short Delay", 320, false)

        defaultMethodSettingsSet := true
    }

    settings    := methodRegistry[methodName]["Settings"]
    runAttempts := settings.Get("Run Attempts")
    tinyDelay   := settings.Get("Tiny Delay")
    shortDelay  := settings.Get("Short Delay")

    static personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
    excelApplication.Workbooks.Open(personalMacroWorkbookPath)

    worksheets := []
    Loop excelApplication.Workbooks.Count {
        currentWorkbook := excelApplication.Workbooks.Item(A_Index)
        if currentWorkbook.Name = "PERSONAL.XLSB" {
            continue
        }

        for worksheet in currentWorkbook.Worksheets {
            worksheets.Push(worksheet.Name)
        }

        break
    }

    excelWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["Excel"]["Executable Filename"] . " ahk_class XLMAIN", 60)
    ActivateWindow(excelWindowSearchResults)

    Loop runAttempts {
        KeyboardShortcut("ALT", "F11") ; Open the Visual Basic editor.

        irregularWorksheetExists := false

        excelApplication.DisplayAlerts := false
        Loop excelApplication.Workbooks.Count {
            currentWorkbook := excelApplication.Workbooks.Item(A_Index)
            if currentWorkbook.Name = "PERSONAL.XLSB" {
                continue
            }

            for worksheet in currentWorkbook.Worksheets {
                worksheetMatch := false

                for worksheetName in worksheets {
                    if worksheet.Name = worksheetName {
                        worksheetMatch := true

                        break
                    }
                }

                if worksheetMatch {
                    continue
                }

                irregularWorksheetExists := true

                worksheet.Delete()
                Sleep(tinyDelay)
            }

            break
        }
        excelApplication.DisplayAlerts := true

        if irregularWorksheetExists {
            if methodRegistry["KeyboardShortcut"]["Settings"]["Tiny Delay"] < 160 {
                SetMethodSetting("KeyboardShortcut", "Tiny Delay", methodRegistry["KeyboardShortcut"]["Settings"]["Tiny Delay"] + 32, true)
            }

            continue
        }

        visualBasicEditorWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["Excel"]["Executable Filename"] . " ahk_class wndclass_desked_gsk", 60, "Failed to open the Visual Basic editor via ALT+F11 in Excel.")
        ActivateWindow(visualBasicEditorWindowSearchResults, true)
        Sleep(tinyDelay)
        PasteText(code, "'")
        Sleep(shortDelay)
        SendInput("{F5}") ; Run Sub/UserForm
        Sleep(tinyDelay)
        break
    }

    LogConclusion("Completed", logValuesForConclusion)
}

WaitForExcelToClose(excelProcessIdentifier) {
    static methodName := RegisterMethod("excelProcessIdentifier As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [excelProcessIdentifier], "Wait for Excel to Close (PID: " . excelProcessIdentifier . ")")

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Long Delay", 1000, false)
        SetMethodSetting(methodName, "Total Seconds to Wait", 240 * 60, false)
        SetMethodSetting(methodName, "Mouse Move Interval Seconds", 120, false)

        defaultMethodSettingsSet := true
    }

    settings                 := methodRegistry[methodName]["Settings"]
    longDelay                := settings.Get("Long Delay")
    totalSecondsToWait       := settings.Get("Total Seconds to Wait")
    mouseMoveIntervalSeconds := settings.Get("Mouse Move Interval Seconds")

    secondsSinceLastMouseMove := 0

    userInterfaceIsGone := false
    loop totalSecondsToWait {
        windowCount := WinGetList("ahk_pid " . excelProcessIdentifier).Length
        if windowCount = 0 {
            Sleep(longDelay)
            userInterfaceIsGone := true
            break
        }

        secondsSinceLastMouseMove += 1
        if secondsSinceLastMouseMove >= mouseMoveIntervalSeconds {
            MouseMove 1, 0, 0, "R"
            MouseMove -1, 0, 0, "R"
            secondsSinceLastMouseMove := 0
        }

        Sleep(longDelay)
    }

    if !userInterfaceIsGone {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Excel did not close within " . totalSecondsToWait . " seconds.")
    }

    LogConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; SQL Server Management Studio ;
; **************************** ;

StartSqlServerManagementStudioAndConnect() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Start SQL Server Management Studio and Connect")

    static sqlServerManagementStudioIsInstalled := ValidateApplicationInstalled("SQL Server Management Studio")

    Run('"' . applicationRegistry["SQL Server Management Studio"]["Executable Path"] . '"')
    sqlServerManagementStudioConnectToServerWindowSearchResults := SearchForWindow("Connect ahk_exe " . applicationRegistry["SQL Server Management Studio"]["Executable Filename"], 60, "Connect to Server Window not found.")
    ActivateWindow(sqlServerManagementStudioConnectToServerWindowSearchResults)

    SendInput("{Enter}") ; Connect

    if !WinWaitClose("Connect ahk id " . sqlServerManagementStudioConnectToServerWindowSearchResults["Window Handle"],, 40) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Connection failed.")
    }

    sqlServerManagementStudioMainWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["SQL Server Management Studio"]["Executable Filename"], 60)
    ActivateWindow(sqlServerManagementStudioMainWindowSearchResults, true)

    LogConclusion("Completed", logValuesForConclusion)
}

ExecuteSqlQueryAndSaveAsCsv(code, saveDirectory, filename) {
    static methodName := RegisterMethod("code As String [Constraint: Summary], saveDirectory As String [Constraint: Directory], filename As String [Constraint: Locator]",  A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [code, saveDirectory, filename], "Execute SQL Query and Save (" . filename . ")")

    static sqlServerManagementStudioIsInstalled := ValidateApplicationInstalled("SQL Server Management Studio")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Short Delay", 100, false)
        SetMethodSetting(methodName, "Medium Delay", 480, false)
        SetMethodSetting(methodName, "Long Delay", 1000, false)

        defaultMethodSettingsSet := true
    }

    settings    := methodRegistry[methodName]["Settings"]
    shortDelay  := settings.Get("Short Delay")
    mediumDelay := settings.Get("Medium Delay")
    longDelay   := settings.Get("Long Delay")

    savePath := saveDirectory . filename . ".csv"

    KeyboardShortcut("CTRL", "N") ; Query with Current Connection
    Sleep(longDelay)
    PasteText(code, "--")
    Sleep(mediumDelay)
    SendInput("{F5}") ; Run the selected portion of the query editor or the entire query editor if nothing is selected
    sqlServerManagementStudioQueryExecutedSuccessfullyImageSearchResults := SearchForDirectoryImage("SQL Server Management Studio", "Query executed successfully", 360)
    sqlServerManagentStudioQueryExecutedSuccessfullyImageCoordinates     := ExtractImageCoordinates(sqlServerManagementStudioQueryExecutedSuccessfullyImageSearchResults)
    sqlServerManagementStudioResultsWindowCoordinates                    := ModifyScreenCoordinates(80, -80, sqlServerManagentStudioQueryExecutedSuccessfullyImageCoordinates)
    PerformMouseActionAtCoordinates("Left", sqlServerManagementStudioResultsWindowCoordinates)
    Sleep(mediumDelay)
    PerformMouseActionAtCoordinates("Right", sqlServerManagementStudioResultsWindowCoordinates)
    Sleep(mediumDelay)
    SendInput("v") ; Save Results As...
    sqlServerManagementStudioSaveResultsWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["SQL Server Management Studio"]["Executable Filename"] . " ahk_class #32770", 60)
    ActivateWindow(sqlServerManagementStudioSaveResultsWindowSearchResults)
    KeyboardShortcut("ALT", "N") ; File name
    Sleep(mediumDelay)
    PasteText(savePath)
    Sleep(shortDelay)
    SendInput("{Enter}") ; Save

    maximumWaitMilliseconds := longDelay * 10
    startTickCount := A_TickCount

    fileExistsAlready := !!FileExist(savePath)

    if !fileExistsAlready {
        while !FileExist(savePath) && (A_TickCount - startTickCount) < maximumWaitMilliseconds {
            Sleep(shortDelay)
        }
    }

    if fileExistsAlready {
        previousModifiedTime := FileGetTime(savePath, "M")
        Sleep(longDelay)
        SendInput("y") ; Yes
        startTickCount := A_TickCount
        Sleep(longDelay)
        while FileGetTime(savePath, "M") = previousModifiedTime && (A_TickCount - startTickCount) < maximumWaitMilliseconds {
            Sleep(shortDelay)
        }

        if FileGetTime(savePath, "M") = previousModifiedTime {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Timed out waiting for overwrite: " . savePath)
        }

        Sleep(mediumDelay)
    }

    LogConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Toad for Oracle              ;
; **************************** ;

ExecuteAutomationApp(appName, runtimeDate := "") {
    static methodName := RegisterMethod("appName As String [Constraint: Locator], runtimeDate As String [Optional] [Constraint: Raw Date Time]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [appName, runtimeDate], "Execute Automation App (" . appName . ")")

    static toadForOracleIsInstalled := ValidateApplicationInstalled("Toad for Oracle")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Tiny Delay", 16, false)
        SetMethodSetting(methodName, "Short Delay", 400, false)
        SetMethodSetting(methodName, "Medium Delay", 880, false)
        SetMethodSetting(methodName, "Long Delay", 1280, false)
        SetMethodSetting(methodName, "Massive Delay", 30000, false)

        defaultMethodSettingsSet := true
    }

    settings     := methodRegistry[methodName]["Settings"]
    tinyDelay    := settings.Get("Tiny Delay")
    shortDelay   := settings.Get("Short Delay")
    mediumDelay  := settings.Get("Medium Delay")
    longDelay    := settings.Get("Long Delay")
    massiveDelay := settings.Get("Massive Delay")

    static toadForOracleExecutableFilename := applicationRegistry["Toad for Oracle"]["Executable Filename"]

    if !ProcessExist(toadForOracleExecutableFilename) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Toad for Oracle process is not running.")
    }

    toadForOracleDatabaseLoginWindowSearchResults := SearchForWindow("ahk_exe " . toadForOracleExecutableFilename . " ahk_class TfrmLogin", 1)
    if toadForOracleDatabaseLoginWindowSearchResults["Success"] = true {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "No server connection is active in Toad for Oracle (Database Login window is open).")
    }

    toadForOracleMainWindowSearchResults := SearchForWindow("ahk_exe " . toadForOracleExecutableFilename . " ahk_class TfrmMain", 60)
    ActivateWindow(toadForOracleMainWindowSearchResults, true)

    KeyboardShortcut("ALT", "S") ; Session
    Sleep(shortDelay)
    SendInput("t") ; Test All Connections (Reconnect) [OR] Test/Reconnect
    Sleep(shortDelay)
    SendInput("t") ; Test All Connections (Reconnect) [OR] t
    Sleep(shortDelay)
    SendInput("{Backspace}") ; Remove character in case present.
    Sleep(tinyDelay)
    SendInput("{Backspace}") ; Remove character in case present.
    Sleep(shortDelay)
    
    overallStartTickCount := A_TickCount
    firstSeenTickCount := 0
    dialogHasAppeared := false

    while true {
        dialogExists := WinExist("ahk_exe " . toadForOracleExecutableFilename . " ahk_class TReconnectForm")

        if !dialogHasAppeared {
            if dialogExists != false {
                dialogHasAppeared := true
                firstSeenTickCount := A_TickCount
            } else if A_TickCount - overallStartTickCount >= (longDelay + longDelay) {
                break
            }
        } else {
            if !dialogExists {
                break
            }

            if A_TickCount - firstSeenTickCount >= massiveDelay {
                LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Reconnect dialog did not close within " . Round(massiveDelay / 1000) . " seconds.")
            }
        }

        Sleep(tinyDelay)
    }

    Sleep(mediumDelay)

    KeyboardShortcut("ALT", "U") ; Utilities
    Sleep(mediumDelay)
    SendInput("{Enter}") ; Automation Designer
    toadForOracleSearchImageSearchResults := SearchForDirectoryImage("Toad for Oracle", "Search")
    toadForOracleSearchImageCoordinates   := ExtractImageCoordinates(toadForOracleSearchImageSearchResults)
    PerformMouseActionAtCoordinates("Left", toadForOracleSearchImageCoordinates)
    Sleep(mediumDelay + longDelay)
    SendInput("{Tab}") ; Text to find:
    Sleep(mediumDelay)
    PasteText(appName)
    Sleep(mediumDelay)
    SendInput("{Enter}") ; Search
    Sleep(longDelay)
    KeyboardShortcut("SHIFT", "TAB") ; Item

    if toadForOracleSearchImageSearchResults["Variant"] = "c" || toadForOracleSearchImageSearchResults["Variant"] = "d" {
        Sleep(mediumDelay)
        KeyboardShortcut("SHIFT", "TAB") ; Item
    }

    Sleep(mediumDelay)
    KeyboardShortcut("SHIFT", "F10") ; Right-click
    Sleep(mediumDelay)
    SendInput("{Down}") ; Goto Item
    Sleep(mediumDelay)
    SendInput("{Enter}") ; Goto Item
    Sleep(longDelay)
    toadForOracleRunSelectedAppsImageSearchResults := SearchForDirectoryImage("Toad for Oracle", "Run selected apps")
    toadForOracleRunSelectedAppsImageCoordinates   := ExtractImageCoordinates(toadForOracleRunSelectedAppsImageSearchResults)

    if runtimeDate != "" {
        PerformMouseActionAtCoordinates("Move", toadForOracleRunSelectedAppsImageCoordinates)

        while A_Now < DateAdd(runtimeDate, -1, "Seconds") {
            Sleep(shortDelay)
        }

        while A_Now < runtimeDate {
            Sleep(tinyDelay)
        }
    }

    PerformMouseActionAtCoordinates("Left", toadForOracleRunSelectedAppsImageCoordinates)
    Sleep(tinyDelay)
    PerformMouseActionAtCoordinates("Move", (Round(A_ScreenWidth/2)) . "x" . (Round(A_ScreenHeight/1.2)))

    LogConclusion("Completed", logValuesForConclusion)
}