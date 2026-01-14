#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Base Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

global applicationRegistry := unset

; **************************** ;
; Application Registry         ;
; **************************** ;

RegisterApplications() {
    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [], "Register Applications")

    global applicationRegistry
    applicationRegistry := Map()
   
    mappingApplications := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Applications.csv")
    for mappingApplication in mappingApplications {
        applicationName := mappingApplication["Name"]

        applicationRegistry[applicationName] := Map()

        applicationRegistry[applicationName]["Counter"] := mappingApplication["Counter"] + 0
    }

    applicationWhitelist := system["Configuration"]["Application Whitelist"]
    applicationWhitelistLength := applicationWhitelist.Length
    if applicationWhitelistLength != 0 {
        for application in applicationWhitelist {
            if !applicationRegistry.Has(application) {
                LogConclusion("Failed", logValuesForConclusion, A_LineNumber, 'Application "' . application . '" does not exist.')
            }

            applicationRegistry[application]["Whitelisted"] := true
        }

        for application in applicationRegistry {
            if !applicationRegistry[application].Has("Whitelisted") {
                applicationRegistry[application]["Whitelisted"] := false
                applicationRegistry[application]["Installed"]   := false
            }
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
            ResolveFactsForApplication(application, applicationRegistry[application]["Counter"])
        } else {
            applicationRegistry[application]["Installed"]         := false
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

            switch outerKey {
                case "Excel":
                    configuration := configuration . "|" . "Personal Macro Workbook: " . innerMap["Personal Macro Workbook"] . "|" . "Code Execution: " . innerMap["Code Execution"]
            }

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

        projectApplicationExecutableDirectoryCandidatesFilePath := system["Project Directory"] . "Application Executable Directory Candidates.csv"
        if FileExist(projectApplicationExecutableDirectoryCandidatesFilePath) {
            projectApplicationExecutableDirectoryCandidates := ConvertCsvToArrayOfMaps(projectApplicationExecutableDirectoryCandidatesFilePath)

            for index, projectApplicationExecutableDirectoryCandidate in projectApplicationExecutableDirectoryCandidates {
                projectApplicationExecutableDirectoryCandidates[index]["Source"] := "Project"
                combinedApplicationExecutableDirectoryCandidates.Push(projectApplicationExecutableDirectoryCandidate)
            }
        }

        sharedApplicationExecutableDirectoryCandidates := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Application Executable Directory Candidates.csv")

        for index, sharedApplicationExecutableDirectoryCandidate in sharedApplicationExecutableDirectoryCandidates {
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
    static methodName := RegisterMethod("applicationName As String, applicationExecutableDirectoryCandidates As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
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

                applicationDirectoryFileList := GetFilesFromDirectory(baseDirectoryWithName, true)

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
    static methodName := RegisterMethod("applicationName As String, applicationExecutableDirectoryCandidates As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
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
    static methodName := RegisterMethod("applicationName As String, applicationExecutableDirectoryCandidates As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
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
        SetMethodSetting(methodName, "Excel Short Delay", 360, false)

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
            excelTinyDelay  := settings.Get("Excel Tiny Delay")
            excelShortDelay := settings.Get("Excel Short Delay")

            CloseApplication("Excel")

            applicationRegistry["Excel"]["Code Execution"] := "Failed"
            excelMacroCode := 'Sub Run()' . '`r`n' . '    Range("A1").Value = "Cell"' . '`r`n' . 'End Sub'
            excelApplication := ComObject("Excel.Application")
            excelWorkbook := excelApplication.Workbooks.Add()
            excelApplication.Visible := true

            personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
            if FileExist(personalMacroWorkbookPath) {
                applicationRegistry["Excel"]["Personal Macro Workbook"] := "Enabled"
                excelApplication.Workbooks.Open(personalMacroWorkbookPath)
            } else {
                applicationRegistry["Excel"]["Personal Macro Workbook"] := "Disabled"
            }
            
            excelProcessIdentifier := ExcelActivateVisualBasicEditorAndPasteCode(excelMacroCode, excelApplication)
            Sleep(excelShortDelay)
            SendInput("{F5}") ; Run Sub/UserForm
            Sleep(excelShortDelay)

            if excelApplication.ActiveSheet.Range("A1").Value = "Cell" {
                applicationRegistry["Excel"]["Code Execution"] := "Basic"
                excelApplication.ActiveSheet.Range("A1").Value := ""
            }

            if applicationRegistry["Excel"]["Personal Macro Workbook"] = "Enabled" && applicationRegistry["Excel"]["Code Execution"] = "Basic" {
                applicationRegistry["Excel"]["Code Execution"] := "Partial"

                KeyboardShortcut("CTRL", "F4") ; Close Window: Module
                Sleep(excelShortDelay)
                KeyboardShortcut("CTRL", "A") ; Select All
                Sleep(excelShortDelay)
                SendInput("{Delete}") ; Delete
                A_Clipboard := excelMacroCode
                KeyboardShortcut("CTRL", "V") ; Paste
                Sleep(excelShortDelay)
                SendInput("{F5}") ; Run Sub/UserForm
                Sleep(excelShortDelay)
                SendInput("{Esc}") ; Close Window Macros
                Sleep(excelTinyDelay + excelTinyDelay)

                if excelApplication.ActiveSheet.Range("A1").Value = "Cell" {
                    applicationRegistry["Excel"]["Code Execution"] := "Full"
                }
            }

            applicationRegistry["Excel"]["International"] := Map()

            excelInternational := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Excel International (2025-09-26).csv")
            for international in excelInternational {
                applicationRegistry["Excel"]["International"][international["Label"]] := excelApplication.International[international["Value"]]
            }

            excelWorkbook.Close(false)
            excelApplication.DisplayAlerts := false
            excelApplication.Quit()

            excelWorkbook := 0
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

    settings   := methodRegistry[methodName]["Settings"]
    tinyDelay  := settings.Get("Tiny Delay")
    shortDelay := settings.Get("Short Delay")

    excelFilePath := FileExistsInDirectory(documentName, saveDirectory, "xlsx")
    if excelFilePath = "" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "documentName not found: " . documentName)
    }

    excelApplication := ComObject("Excel.Application")
    excelApplication.Workbooks.Open(excelFilePath, 0)
    excelWorkbook := excelApplication.ActiveWorkbook 
    excelApplication.Visible := true

    static personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
    if applicationRegistry["Excel"]["Personal Macro Workbook"] = "Enabled" {
        excelApplication.Workbooks.Open(personalMacroWorkbookPath)
    }

    excelApplication.CalculateUntilAsyncQueriesDone()
    while excelApplication.CalculationState != 0 {
        Sleep(tinyDelay + tinyDelay)
    }

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
            excelProcessIdentifier := ExcelActivateVisualBasicEditorAndPasteCode(code, excelApplication)
            Sleep(shortDelay)
            SendInput("{F5}") ; Run Sub/UserForm
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
                excelProcessIdentifier := ExcelActivateVisualBasicEditorAndPasteCode(code, excelApplication)
                Sleep(shortDelay)
                SendInput("{F5}") ; Run Sub/UserForm
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
        excelProcessIdentifier := ExcelActivateVisualBasicEditorAndPasteCode(code, excelApplication)
        Sleep(shortDelay)
        SendInput("{F5}") ; Run Sub/UserForm
        WaitForExcelToClose(excelProcessIdentifier)
        aboutWorksheet   := 0
        excelWorkbook    := 0
        excelApplication := 0
        ProcessWaitClose(excelProcessIdentifier, 2)

        LogConclusion("Completed", logValuesForConclusion)
    }
}

ExcelActivateVisualBasicEditorAndPasteCode(code, excelApplication) {
    static methodName := RegisterMethod("code As String [Constraint: Summary], excelApplication As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [code, excelApplication], "Excel Activate Visual Basic Editor and Paste Code (Length: " . StrLen(code) . ")")

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Tiny Delay", 32, false)
        SetMethodSetting(methodName, "Short Delay", 160, false)
        SetMethodSetting(methodName, "Medium Delay", 240, false)

        defaultMethodSettingsSet := true
    }

    settings    := methodRegistry[methodName]["Settings"]
    tinyDelay   := settings.Get("Tiny Delay")
    shortDelay  := settings.Get("Short Delay")
    mediumDelay := settings.Get("Medium Delay")

    excelWindowHandle := excelApplication.Hwnd
    while !excelWindowHandle := excelApplication.Hwnd {
        Sleep(tinyDelay)
    }

    excelProcessIdentifier := WinGetPID("ahk_id " . excelWindowHandle)
    excelWindowHandle      := ActivateWindow("ahk_id " . excelWindowHandle . " ahk_class XLMAIN")

    KeyboardShortcut("ALT", "F11") ; Microsoft Visual Basic for Applications

    visualBasicEditorWindowHandle := ActivateWindow("ahk_pid " . excelProcessIdentifier . " ahk_class wndclass_desked_gsk")
    if visualBasicEditorWindowHandle {
        if applicationRegistry["Excel"]["Code Execution"] != "Full" {
            KeyboardShortcut("ALT", "I") ; Insert
            Sleep(shortDelay)
            SendInput("m") ; Module
            Sleep(mediumDelay)
        }

        PasteText(code, "'")

        LogConclusion("Completed", logValuesForConclusion)
        return excelProcessIdentifier
    } else {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to open the Visual Basic Editor via ALT+F11 in Excel.")
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
        SetMethodSetting(methodName, "Tiny Delay", 32, false)
        SetMethodSetting(methodName, "Short Delay", 260, false)

        defaultMethodSettingsSet := true
    }

    settings   := methodRegistry[methodName]["Settings"]
    tinyDelay  := settings.Get("Tiny Delay")
    shortDelay := settings.Get("Short Delay")

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
        FileAppend("", sidecarPath, "UTF-8-RAW")

        excelApplication := ComObject("Excel.Application")
        excelWorkbook := excelApplication.Workbooks.Add()
        excelApplication.Visible := true

        static personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
        if applicationRegistry["Excel"]["Personal Macro Workbook"] = "Enabled" {
            excelApplication.Workbooks.Open(personalMacroWorkbookPath)
        }

        excelProcessIdentifier := ExcelActivateVisualBasicEditorAndPasteCode(code, excelApplication)
        Sleep(shortDelay)
        SendInput("{F5}") ; Run Sub/UserForm
        WaitForExcelToClose(excelProcessIdentifier)
        excelWorkbook    := 0
        excelApplication := 0
        ProcessWaitClose(excelProcessIdentifier, 2)

        DeleteFile(sidecarPath) ; Remove sidecar after a successful run.
        LogConclusion("Completed", logValuesForConclusion)
    }
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
    sqlServerManagementStudioConnectionWindowHandle := ActivateWindow("Connect ahk_exe " . applicationRegistry["SQL Server Management Studio"]["Executable Filename"], "Connect Dialog Window not found.")
    SendInput("{Enter}") ; Connect

    if WinWaitClose("Connect ahk_id " . sqlServerManagementStudioConnectionWindowHandle,, 40) {
    } else {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Connection failed.")
    }

    microsoftSqlServerManagementStudioWindowHandle := ActivateWindow("ahk_exe " . applicationRegistry["SQL Server Management Studio"]["Executable Filename"])
    WinMaximize("ahk_id " . microsoftSqlServerManagementStudioWindowHandle)

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
    sqlQuerySuccessfulResults := SearchForDirectoryImage("SQL Server Management Studio", "Query executed successfully", 360)
    sqlQuerySuccessfulCoordinates := ExtractScreenCoordinates(sqlQuerySuccessfulResults)
    sqlQueryResultsWindowCoordinates := ModifyScreenCoordinates(80, -80, sqlQuerySuccessfulCoordinates)
    PerformMouseActionAtCoordinates("Left", sqlQueryResultsWindowCoordinates)
    Sleep(mediumDelay)
    PerformMouseActionAtCoordinates("Right", sqlQueryResultsWindowCoordinates)
    Sleep(mediumDelay)
    SendInput("v") ; Save Results As...
    saveResultsAsWindowHandle := ActivateWindow("ahk_class #32770")
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

    if WinExist("ahk_exe " . toadForOracleExecutableFilename . " ahk_class TfrmLogin") {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "No server connection is active in Toad for Oracle (login dialog is open).")
    }

    windowCriteria := "ahk_exe " . toadForOracleExecutableFilename . " ahk_class TfrmMain"
    toadForOracleWindowHandle := ActivateWindow(windowCriteria)
    WinMaximize("ahk_id " . toadForOracleWindowHandle)
    Sleep(shortDelay)

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
    toadForOracleSearchResults := SearchForDirectoryImage("Toad for Oracle", "Search")
    toadForOracleSearchCoordinates := ExtractScreenCoordinates(toadForOracleSearchResults)
    PerformMouseActionAtCoordinates("Left", toadForOracleSearchCoordinates)
    Sleep(mediumDelay + longDelay)
    SendInput("{Tab}") ; Text to find:
    Sleep(mediumDelay)
    PasteText(appName)
    Sleep(mediumDelay)
    SendInput("{Enter}") ; Search
    Sleep(longDelay)
    KeyboardShortcut("SHIFT", "TAB") ; Item

    if toadForOracleSearchResults["Variant"] = "c" || toadForOracleSearchResults["Variant"] = "d" {
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
    toadForOracleRunSelectedAppsResults := SearchForDirectoryImage("Toad for Oracle", "Run selected apps")
    toadForOracleRunSelectedAppsCoordinates := ExtractScreenCoordinates(toadForOracleRunSelectedAppsResults)

    if runtimeDate != "" {
        PerformMouseActionAtCoordinates("Move", toadForOracleRunSelectedAppsCoordinates)

        while A_Now < DateAdd(runtimeDate, -1, "Seconds") {
            Sleep(shortDelay)
        }

        while A_Now < runtimeDate {
            Sleep(tinyDelay)
        }
    }

    PerformMouseActionAtCoordinates("Left", toadForOracleRunSelectedAppsCoordinates)
    Sleep(tinyDelay)
    PerformMouseActionAtCoordinates("Move", (Round(A_ScreenWidth/2)) . "x" . (Round(A_ScreenHeight/1.2)))

    LogConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

DetermineWindowsBinaryType(executablePath) {
    static methodName := RegisterMethod("executablePath As String [Constraint: Absolute Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [executablePath])

    static SCS_32BIT_BINARY := 0
    static SCS_DOS_BINARY   := 1
    static SCS_WOW_BINARY   := 2
    static SCS_PIF_BINARY   := 3
    static SCS_POSIX_BINARY := 4
    static SCS_OS2_BINARY   := 5
    static SCS_64BIT_BINARY := 6

    classificationResult := "N/A"
    scsCode := 0

    executableSubsystemRetrievedSuccessfully := DllCall("Kernel32\GetBinaryTypeW", "Str", executablePath, "UInt*", &scsCode, "Int")

    if executableSubsystemRetrievedSuccessfully {
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