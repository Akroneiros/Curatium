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
    static methodName := RegisterMethod("RegisterApplications()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Register Applications", methodName)

    global applicationRegistry
    applicationRegistry := Map()
   
    mappingApplications := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Applications.csv")
    for application in mappingApplications {
        applicationName := application["Name"]

        applicationRegistry[applicationName] := Map()

        applicationRegistry[applicationName]["Counter"] := application["Counter"] + 0
    }

    projectApplicationsFilePath := ""
    if FileExist(system["Project Directory"] . "Applications.csv") {
        projectApplicationsFilePath := system["Project Directory"] . "Applications.csv"
    }

    if projectApplicationsFilePath != "" {
        projectApplications := ConvertCsvToArrayOfMaps(projectApplicationsFilePath)

        for application in projectApplications {
            applicationName := application["Name"]

            try {
                if !applicationRegistry.Has(applicationName) {
                    throw Error("Application " . Chr(34) . applicationName . Chr(34) . " does not exist.")
                }
            } catch as applicationMissingError {
                LogInformationConclusion("Failed", logValuesForConclusion, applicationMissingError)
            }

            executablePathSearchResult := ExecutablePathResolve(applicationName)
            if Type(executablePathSearchResult) = "Map" {
                applicationRegistry[applicationName]["Executable Path"]   := executablePathSearchResult["Executable Path"]
                applicationRegistry[applicationName]["Resolution Method"] := executablePathSearchResult["Resolution Method"]
                applicationRegistry[applicationName]["Installed"]         := true
                ResolveFactsForApplication(applicationName, applicationRegistry[applicationName]["Counter"])
            } else {
                applicationRegistry[applicationName]["Installed"]         := false
            }
        }

        for application in mappingApplications {
            if !applicationRegistry[application["Name"]].Has("Installed") {
                applicationRegistry[application["Name"]]["Installed"]     := false
            }
        }
    } else {
        for application in mappingApplications {
            applicationName := application["Name"]

            executablePathSearchResult := ExecutablePathResolve(applicationName)
            if Type(executablePathSearchResult) = "Map" {
                applicationRegistry[applicationName]["Executable Path"]   := executablePathSearchResult["Executable Path"]
                applicationRegistry[applicationName]["Resolution Method"] := executablePathSearchResult["Resolution Method"]
                applicationRegistry[applicationName]["Installed"]         := true
                ResolveFactsForApplication(applicationName, applicationRegistry[applicationName]["Counter"])
            } else {
                applicationRegistry[applicationName]["Installed"]         := false
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

    LogInformationConclusion("Completed", logValuesForConclusion)
}

ExecutablePathResolve(applicationName) {
    static methodName := RegisterMethod("ExecutablePathResolve(applicationName As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [applicationName])

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

    if projectEntryAvailable {
        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaDirectory(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }

        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaRegistry(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }

        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaUninstall(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }
    } else {
        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaUninstall(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }

        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaRegistry(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }

        if Type(executablePathSearchResult) != "Map" {
            executablePathSearchResult := ExecutablePathViaDirectory(applicationName, relevantApplicationExecutableDirectoryCandidates)
        }
    }

    return executablePathSearchResult
}

ExecutablePathViaDirectory(applicationName, applicationExecutableDirectoryCandidates) {
    static methodName := RegisterMethod("ExecutablePathViaDirectory(applicationName As String, applicationExecutableDirectoryCandidates As Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [applicationName, applicationExecutableDirectoryCandidates])

    static candidateBaseDirectories := unset
    if !IsSet(candidateBaseDirectories) {
        candidateBaseDirectories := []

        portableFilesDirectory        := ExtractDirectory(A_WinDir) . "Portable Files"
        programFilesPortableDirectory := ExtractDirectory(A_WinDir) . "Program Files (Portable)"

        if DirExist(portableFilesDirectory) {
            candidateBaseDirectories.Push(portableFilesDirectory)
        }

        if DirExist(programFilesPortableDirectory) {
            candidateBaseDirectories.Push(programFilesPortableDirectory)
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
            "Resolution Method", "Directory"
        )
    }

    return executablePathSearchResult
}

ExecutablePathViaRegistry(applicationName, applicationExecutableDirectoryCandidates) {
    static methodName := RegisterMethod("ExecutablePathViaRegistry(applicationName As String, applicationExecutableDirectoryCandidates As Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [applicationName, applicationExecutableDirectoryCandidates])

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
            "Resolution Method", "Registry"
        )
    }

    return executablePathSearchResult
}

ExecutablePathViaUninstall(applicationName, applicationExecutableDirectoryCandidates) {
    static methodName := RegisterMethod("ExecutablePathViaUninstall(applicationName As String, applicationExecutableDirectoryCandidates As Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [applicationName, applicationExecutableDirectoryCandidates])

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
    static methodName := RegisterMethod("ResolveFactsForApplication(applicationName As String, counter As Integer)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [applicationName, counter])

    global applicationRegistry

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

    switch applicationName {
        case "CyberChef":
            cyberChefHtml := ReadFileOnHashMatch(applicationRegistry[applicationName]["Executable Path"], DecodeBaseToSha256Hex(applicationRegistry[applicationName]["Executable Hash"], 86))

            versionPattern := "i)CyberChef\s+version:\s*(\d+(?:\.\d+)*)"
            if RegExMatch(cyberChefHtml, versionPattern, &versionMatch) {
                applicationRegistry[applicationName]["Executable Version"] := versionMatch[1]
            }
        case "Excel":
            CloseApplication("Excel")

            applicationRegistry["Excel"]["Code Execution"] := "Failed"
            excelMacroCode := 'Sub Run(): Range("A1").Value = "Cell": End Sub'
            excelApplication := ComObject("Excel.Application")
            excelWorkbook := excelApplication.Workbooks.Add()
            excelApplication.Visible := true

            excelWindowHandle := excelApplication.Hwnd
            while !excelWindowHandle := excelApplication.Hwnd {
                Sleep(16)
            }

            personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
            if FileExist(personalMacroWorkbookPath) {
                applicationRegistry["Excel"]["Personal Macro Workbook"] := "Enabled"
                excelApplication.Workbooks.Open(personalMacroWorkbookPath)
            } else {
                applicationRegistry["Excel"]["Personal Macro Workbook"] := "Disabled"
            }
            
            excelProcessIdentifier := WinGetPID("ahk_id " . excelWindowHandle)
            WinActivate("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier)
            WinWaitActive("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier, , 10)
            ExcelActivateEditorAndPasteCode(excelMacroCode)
            SendEvent("{F5}") ; F5 (Run Sub/UserForm)
            Sleep(200)

            if excelApplication.ActiveSheet.Range("A1").Value = "Cell" {
                applicationRegistry["Excel"]["Code Execution"] := "Basic"
                excelApplication.ActiveSheet.Range("A1").Value := ""
            }

            if applicationRegistry["Excel"]["Personal Macro Workbook"] = "Enabled" && applicationRegistry["Excel"]["Code Execution"] = "Basic" {
                applicationRegistry["Excel"]["Code Execution"] := "Partial"

                SendEvent("^{F4}") ; CTRL+F4 (Close Window: Module)
                Sleep(200)
                SendEvent("^a") ; CTRL+A (Select All)
                Sleep(200)
                SendEvent("^a") ; CTRL+A (Select All)
                Sleep(200)
                SendEvent("{Delete}") ; Delete (Delete)
                A_Clipboard := excelMacroCode
                ClipWait(2)
                SendEvent("^v") ; CTRL+V (Paste)
                Sleep(200)
                SendEvent("{F5}") ; F5 (Run Sub/UserForm)
                Sleep(200)
                SendEvent("{Esc}") ; Escape (Close Window Macros)
                Sleep(32)

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
    static methodName := RegisterMethod("ValidateApplicationFact(applicationName As String, factName As String, factValue As String)", A_LineFile, A_LineNumber + 1)
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

ValidateApplicationInstalled(applicationName) {
    static methodName := RegisterMethod("ValidateApplicationInstalled(applicationName As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Validate Application Installed (" . applicationName . ")", methodName, [applicationName])

    try {
        if !applicationRegistry.Has(applicationName) {
            throw Error("Application doesn't exist: " . applicationName)
        }
    } catch as applicationMissingError {
        LogInformationConclusion("Failed", logValuesForConclusion, applicationMissingError)
    }

    try {
        if !applicationRegistry[applicationName]["Installed"] {
            throw Error("Application not installed: " . applicationName)
        }
    } catch as applicationNotInstalledError {
        LogInformationConclusion("Failed", logValuesForConclusion, applicationNotInstalledError)
    }

    applicationIsInstalled := true

    LogInformationConclusion("Completed", logValuesForConclusion)
    return applicationIsInstalled
}

; **************************** ;
; Shared                       ;
; **************************** ;

CloseApplication(applicationName) {
    static methodName := RegisterMethod("CloseApplication(applicationName As String [Constraint: Locator])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Close Application (" . applicationName . ")", methodName, [applicationName])

    try {
        if !applicationRegistry.Has(applicationName) {
            throw Error("Application not found: " . applicationName)
        }
    } catch as missingApplicationError {
        LogInformationConclusion("Failed", logValuesForConclusion, missingApplicationError)
    }

    executableName := applicationRegistry[applicationName]["Executable Filename"]

    if !ProcessExist(executableName) {
        LogInformationConclusion("Skipped", logValuesForConclusion)
    } else {
        ProcessClose(executableName)
        ProcessWaitClose(executableName, 4)

        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

; **************************** ;
; Excel                        ;
; **************************** ;

ExcelExtensionRun(documentName, saveDirectory, code, displayName := "", aboutRange := "", aboutCondition := "") {
    static methodName := RegisterMethod("ExcelExtensionRun(documentName As String [Constraint: Locator], saveDirectory As String [Constraint: Directory], code As String [Constraint: Code], displayName As String [Optional], aboutRange As String [Optional] [Constraint: Locator], aboutCondition As String [Optional] [Constraint: Locator])", A_LineFile, A_LineNumber + 7)
    overlayValue := ""
    if displayName = "" {
        overlayValue := documentName . " Excel Extension Run"
    } else {
        overlayValue := displayName . " Excel Extension Run"
    }
    logValuesForConclusion := LogInformationBeginning(overlayValue, methodName, [documentName, saveDirectory, code, displayName, aboutRange, aboutCondition])

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    excelFilePath := FileExistsInDirectory(documentName, saveDirectory, "xlsx")
    try {
        if excelFilePath = "" {
            throw Error("documentName not found: " . documentName)
        }
    } catch as documentNameNotFoundError {
        LogInformationConclusion("Failed", logValuesForConclusion, documentNameNotFoundError)
    }

    excelApplication := ComObject("Excel.Application")
    excelApplication.Workbooks.Open(excelFilePath, 0)
    excelWorkbook := excelApplication.ActiveWorkbook 
    excelApplication.Visible := true

    excelWindowHandle := excelApplication.Hwnd
    while !excelWindowHandle := excelApplication.Hwnd {
        Sleep(32)
    }

    static personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
    if applicationRegistry["Excel"]["Personal Macro Workbook"] = "Enabled" {
        excelApplication.Workbooks.Open(personalMacroWorkbookPath)
    }

    excelProcessIdentifier := WinGetPID("ahk_id " . excelWindowHandle)
    WinActivate("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier)
    WinWaitActive("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier, , 10)

    excelApplication.CalculateUntilAsyncQueriesDone()
    while excelApplication.CalculationState != 0 {
        Sleep(64)
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
            ExcelActivateEditorAndPasteCode(code)
            SendEvent("{F5}") ; F5 (Run Sub/UserForm)
            WaitForExcelToClose(excelProcessIdentifier)
            aboutWorksheet   := 0
            excelWorkbook    := 0
            excelApplication := 0
            ProcessWaitClose(excelProcessIdentifier, 2)

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
                ExcelActivateEditorAndPasteCode(code)
                SendEvent("{F5}") ; F5 (Run Sub/UserForm)
                WaitForExcelToClose(excelProcessIdentifier)
                aboutWorksheet   := 0
                excelWorkbook    := 0
                excelApplication := 0
                ProcessWaitClose(excelProcessIdentifier, 2)

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
                ProcessWaitClose(excelProcessIdentifier, 2)

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
            ProcessWaitClose(excelProcessIdentifier, 2)

            LogInformationConclusion("Skipped", logValuesForConclusion)
        }
    } else {
        ExcelActivateEditorAndPasteCode(code)
        SendEvent("{F5}") ; F5 (Run Sub/UserForm)
        WaitForExcelToClose(excelProcessIdentifier)
        aboutWorksheet   := 0
        excelWorkbook    := 0
        excelApplication := 0
        ProcessWaitClose(excelProcessIdentifier, 2)

        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

ExcelActivateEditorAndPasteCode(code) {
    static methodName := RegisterMethod("ExcelActivateEditorAndPasteCode(code As String [Constraint: Code])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Excel Activate Editor and Paste Code (Length: " . StrLen(code) . ")", methodName, [code])

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    SendEvent("!{F11}") ; F11 (Microsoft Visual Basic for Applications)
    WinWait("ahk_class wndclass_desked_gsk", , 10)
    WinActivate("ahk_class wndclass_desked_gsk")
    WinWaitActive("ahk_class wndclass_desked_gsk", , 2)

    if applicationRegistry["Excel"]["Code Execution"] != "Full" {
        SendEvent("!i") ; ALT+I (Insert)
        Sleep(560)
        SendEvent("m") ; M (Module)
        Sleep(440)
    }

    PasteCode(code, "'")

    LogInformationConclusion("Completed", logValuesForConclusion)
}

ExcelStartingRun(documentName, saveDirectory, code, displayName := "") {
    static methodName := RegisterMethod("ExcelStartingRun(documentName As String [Constraint: Locator], saveDirectory As String [Constraint: Directory], code As String [Constraint: Code], displayName As String [Optional])", A_LineFile, A_LineNumber + 7)
    overlayValue := ""
    if displayName = "" {
        overlayValue := documentName . " Excel Starting Run"
    } else {
        overlayValue := displayName . " Excel Starting Run"
    }
    logValuesForConclusion := LogInformationBeginning(overlayValue, methodName, [documentName, saveDirectory, code, displayName])

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

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
        LogInformationConclusion("Skipped", logValuesForConclusion)
    } else {
        sidecarPath := saveDirectory . documentName . ".txt"
        FileAppend("", sidecarPath, "UTF-8-RAW")

        excelApplication := ComObject("Excel.Application")
        excelWorkbook := excelApplication.Workbooks.Add()
        excelApplication.Visible := true

        excelWindowHandle := excelApplication.Hwnd
        while !excelWindowHandle := excelApplication.Hwnd {
            Sleep(32)
        }

        static personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
        if applicationRegistry["Excel"]["Personal Macro Workbook"] = "Enabled" {
            excelApplication.Workbooks.Open(personalMacroWorkbookPath)
        }

        excelProcessIdentifier := WinGetPID("ahk_id " . excelWindowHandle)
        WinActivate("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier)
        WinWaitActive("ahk_class XLMAIN ahk_pid " . excelProcessIdentifier, , 10)
        ExcelActivateEditorAndPasteCode(code)
        SendEvent("{F5}") ; F5 (Run Sub/UserForm)
        WaitForExcelToClose(excelProcessIdentifier)
        excelWorkbook    := 0
        excelApplication := 0
        ProcessWaitClose(excelProcessIdentifier, 2)

        DeleteFile(sidecarPath) ; Remove sidecar after a successful run.
        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

WaitForExcelToClose(excelProcessIdentifier) {
    static methodName := RegisterMethod("WaitForExcelToClose(excelProcessIdentifier As Integer)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Wait for Excel to Close (PID: " . excelProcessIdentifier . ")", methodName, [excelProcessIdentifier])

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    totalSecondsToWait := 240 * 60
    mouseMoveIntervalSec := 120
    secondsSinceLastMouseMove := 0

    userInterfaceIsGone := false
    loop totalSecondsToWait {
        windowCount := WinGetList("ahk_pid " excelProcessIdentifier).Length
        if windowCount = 0 {
            Sleep(1000)
            userInterfaceIsGone := true
            break
        }

        secondsSinceLastMouseMove += 1
        if secondsSinceLastMouseMove >= mouseMoveIntervalSec {
            MouseMove 1, 0, 0, "R"
            MouseMove -1, 0, 0, "R"
            secondsSinceLastMouseMove := 0
        }

        Sleep(1000)
    }

    try {
        if !userInterfaceIsGone {
            throw Error("Excel did not close within 240 minutes.")
        }
    } catch as excelCloseError {
        LogInformationConclusion("Failed", logValuesForConclusion, excelCloseError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; SQL Server Management Studio ;
; **************************** ;

StartSqlServerManagementStudioAndConnect() {
    static methodName := RegisterMethod("StartSqlServerManagementStudioAndConnect()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Start SQL Server Management Studio and Connect", methodName)

    static sqlServerManagementStudioIsInstalled := ValidateApplicationInstalled("SQL Server Management Studio")

    Run('"' . applicationRegistry["SQL Server Management Studio"]["Executable Path"] . '"')
    sqlServerManagementStudioExecutableFilename := applicationRegistry["SQL Server Management Studio"]["Executable Filename"]
    WinWaitActive("Connect to Server ahk_exe " . sqlServerManagementStudioExecutableFilename,, 20)
    SendInput("{Enter}")

    try {
        if WinWaitClose("Connect to Server ahk_exe " . sqlServerManagementStudioExecutableFilename,, 40) {
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
    static methodName := RegisterMethod("ExecuteSqlQueryAndSaveAsCsv(code As String [Constraint: Code], saveDirectory As String [Constraint: Directory], filename As String [Constraint: Locator])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Execute SQL Query and Save (" . filename . ")", methodName, [code, saveDirectory, filename])

    static sqlServerManagementStudioIsInstalled := ValidateApplicationInstalled("SQL Server Management Studio")

    savePath := saveDirectory . filename . ".csv"

    SendInput("^n") ; CTRL+N (Query with Current Connection)
    Sleep(800)
    PasteCode(code, "--")
    SendInput("{F5}") ; F5 (Run the selected portion of the query editor or the entire query editor if nothing is selected)
    sqlQuerySuccessfulResults := SearchForDirectoryImage("SQL Server Management Studio", "Query executed successfully", 360)
    sqlQuerySuccessfulCoordinates := ExtractScreenCoordinates(sqlQuerySuccessfulResults)
    sqlQueryResultsWindowCoordinates := ModifyScreenCoordinates(80, -80, sqlQuerySuccessfulCoordinates)
    PerformMouseActionAtCoordinates("Left", sqlQueryResultsWindowCoordinates)
    Sleep(480)
    PerformMouseActionAtCoordinates("Right", sqlQueryResultsWindowCoordinates)
    Sleep(480)
    SendEvent("v") ; V (Save Results As...)
    WinWaitActive("ahk_class #32770",, 2)
    SendEvent("!n") ; ALT+N (File name)
    Sleep(80)
    PastePath(savePath)
    SendInput("{Enter}") ; ENTER (Save)

    maximumWaitMilliseconds := 10000
    pollIntervalMilliseconds := 100
    startTickCount := A_TickCount

    fileExistsAlready := !!FileExist(savePath)

    if !fileExistsAlready {
        while !FileExist(savePath) && (A_TickCount - startTickCount) < maximumWaitMilliseconds {
            Sleep(pollIntervalMilliseconds)
        }
    }

    if fileExistsAlready {
        previousModifiedTime := FileGetTime(savePath, "M")
        Sleep(1000)
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
    static methodName := RegisterMethod("ExecuteAutomationApp(appName As String [Constraint: Locator], runtimeDate As String [Optional] [Constraint: Raw Date Time])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Execute Automation App (" . appName . ")", methodName, [appName, runtimeDate])

    static toadForOracleIsInstalled := ValidateApplicationInstalled("Toad for Oracle")

    static toadForOracleExecutableFilename := applicationRegistry["Toad for Oracle"]["Executable Filename"]

    windowCriteria := "ahk_exe " . toadForOracleExecutableFilename . " ahk_class TfrmMain"

    try {
        if !ProcessExist(toadForOracleExecutableFilename) {
            throw Error("Toad for Oracle process is not running.")
        }
    } catch as processNotRunningError {
        LogInformationConclusion("Failed", logValuesForConclusion, processNotRunningError)
    }

    try {
        if WinExist("ahk_exe " . toadForOracleExecutableFilename . " ahk_class TfrmLogin") {
            throw Error("No server connection is active in Toad for Oracle (login dialog is open).")
        }
    } catch as noActiveConnectionError {
        LogInformationConclusion("Failed", logValuesForConclusion, noActiveConnectionError)
    }

    WinActivate(windowCriteria)
    WinMaximize(windowCriteria)

    SendEvent("!s") ; ALT+S (Session)
    Sleep(400)
    SendEvent("t") ; T (Test All Connections (Reconnect) [OR] Test/Reconnect)
    Sleep(400)
    SendEvent("t") ; T (Test All Connections (Reconnect) [OR] t)
    Sleep(400)
    SendEvent("{Backspace}") ; Backspace (Remove t if present)
    
    overallStartTickCount := A_TickCount
    firstSeenTickCount := 0
    dialogHasAppeared := false

    while true {
        dialogExists := WinExist("ahk_exe " . toadForOracleExecutableFilename . " ahk_class TReconnectForm")

        if !dialogHasAppeared {
            if dialogExists != false {
                dialogHasAppeared := true
                firstSeenTickCount := A_TickCount
            } else if A_TickCount - overallStartTickCount >= 2000 {
                break
            }
        } else {
            if !dialogExists {
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
        submenuWindowHandle := WinWait("ahk_exe " toadForOracleExecutableFilename " ahk_class TdxBarSubMenuControl", , 1000)
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

    toadForOracleSearchResults := SearchForDirectoryImage("Toad for Oracle", "Search")
    toadForOracleSearchCoordinates := ExtractScreenCoordinates(toadForOracleSearchResults)
    PerformMouseActionAtCoordinates("Left", toadForOracleSearchCoordinates)
    Sleep(2000)
    SendEvent("{Tab}") ; TAB (Text to find:)
    Sleep(800)
    PasteSearch(appName)
    Sleep(800)
    SendEvent("{Enter}") ; ENTER (Search)
    Sleep(1200)
    SendEvent("+{Tab}") ; SHIFT+TAB (Item)

    if toadForOracleSearchResults["Variant"] = "c" || toadForOracleSearchResults["Variant"] = "d" {
        Sleep(800)
        SendEvent("+{Tab}") ; SHIFT+TAB (Item)
    }

    Sleep(800)
    SendEvent("+{F10}") ; SHIFT+F10 (Right-click)
    Sleep(800)
    SendEvent("{Down}") ; DOWN ARROW (Goto Item)
    Sleep(800)
    SendEvent("{Enter}") ; ENTER (Goto Item)
    Sleep(1200)
    toadForOracleRunSelectedAppsResults := SearchForDirectoryImage("Toad for Oracle", "Run selected apps")
    toadForOracleRunSelectedAppsCoordinates := ExtractScreenCoordinates(toadForOracleRunSelectedAppsResults)

    if runtimeDate != "" {
        PerformMouseActionAtCoordinates("Move", toadForOracleRunSelectedAppsCoordinates)

        while A_Now < DateAdd(runtimeDate, -1, "Seconds") {
            Sleep(240)
        }

        while A_Now < runtimeDate {
            Sleep(16)
        }
    }

    PerformMouseActionAtCoordinates("Left", toadForOracleRunSelectedAppsCoordinates)
    Sleep(16)
    PerformMouseActionAtCoordinates("Move", (Round(A_ScreenWidth/2)) . "x" . (Round(A_ScreenHeight/1.2)))

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

DetermineWindowsBinaryType(executablePath) {
    static methodName := RegisterMethod("DetermineWindowsBinaryType(executablePath As String [Constraint: Absolute Path])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [executablePath])

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