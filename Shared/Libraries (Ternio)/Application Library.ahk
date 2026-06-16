#Requires AutoHotkey v2.0
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

; **************************** ;
; Application Registry         ;
; **************************** ;

RegisterApplications(applications, applicationExecutableDirectoryCandidates) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("applications As Array, applicationExecutableDirectoryCandidates As Array", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [applications, applicationExecutableDirectoryCandidates], "Register Applications")

    global applicationRegistry

    for application in applications {
        applicationName         := application["Name"]
        applicationCounter      := application["Counter"] + 0
        applicationWhitelisted  := application["Whitelisted"]
        applicationSharedImages := false

        if application.Has("Shared Images") {
            applicationSharedImages := true
        }

        applicationRegistry[applicationName] := Map(
            "Counter",       applicationCounter,
            "Shared Images", applicationSharedImages,
            "Whitelisted",   applicationWhitelisted
        )
    }

    combinedApplicationExecutableDirectoryCandidates := []
    for outerKey, innerValue in applicationRegistry {
        if innerValue["Whitelisted"] {
            applicationRegistry[outerKey]["Application Executable Directory Candidates"] := []
            for applicationExecutableDirectoryCandidate in applicationExecutableDirectoryCandidates {
                if outerKey = applicationExecutableDirectoryCandidate["Name"] {
                    if applicationExecutableDirectoryCandidate["Source"] = "Project" {
                        applicationRegistry[outerKey]["Application Executable Directory Candidates"].Push(applicationExecutableDirectoryCandidate)
                        combinedApplicationExecutableDirectoryCandidates.Push(applicationExecutableDirectoryCandidate)
                    }
                }
            }

            for applicationExecutableDirectoryCandidate in applicationExecutableDirectoryCandidates {
                if outerKey = applicationExecutableDirectoryCandidate["Name"] {
                    if applicationExecutableDirectoryCandidate["Source"] = "Shared" {
                        applicationRegistry[outerKey]["Application Executable Directory Candidates"].Push(applicationExecutableDirectoryCandidate)
                        combinedApplicationExecutableDirectoryCandidates.Push(applicationExecutableDirectoryCandidate)
                    }
                }
            }
        }
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

    installedApplicationsWithImageLibraryDataCount := 0
    for outerKey, innerValue in applicationRegistry {
        if !innerValue.Has("Executable Collision") {
            innerValue["Executable Collision"] := false
        }

        projectSourcePresent := false
        for applicationExecutableDirectoryCandidate in innerValue["Application Executable Directory Candidates"] {
            if applicationExecutableDirectoryCandidate["Source"] = "Project" {
                projectSourcePresent := true
                break
            }
        }

        executablePathSearchResult := unset
        if projectSourcePresent || innerValue["Executable Collision"] {
            executablePathSearchResult := ExecutablePathViaReference(innerValue)

            if !executablePathSearchResult["Success"] {
                executablePathSearchResult := ExecutablePathViaAppPaths(innerValue)
            }

            if !executablePathSearchResult["Success"] {
                executablePathSearchResult := ExecutablePathViaUninstall(innerValue)
            }
        } else {
            executablePathSearchResult := ExecutablePathViaUninstall(innerValue)

            if !executablePathSearchResult["Success"] {
                executablePathSearchResult := ExecutablePathViaAppPaths(innerValue)
            }

            if !executablePathSearchResult["Success"] {
                executablePathSearchResult := ExecutablePathViaReference(innerValue)
            }
        }

        if executablePathSearchResult["Success"] {
            if innerValue["Shared Images"] {
                installedApplicationsWithImageLibraryDataCount++
            }

            applicationRegistry[outerKey]["Executable Path"]   := executablePathSearchResult["Executable Path"]
            applicationRegistry[outerKey]["Resolution Method"] := executablePathSearchResult["Resolution Method"]
            applicationRegistry[outerKey]["Installed"]         := true
        } else {
            applicationRegistry[outerKey]["Installed"]         := false
        }
    }

    if installedApplicationsWithImageLibraryDataCount != 0 {
        switch system["Environment"]["Display Resolution"] {
            case "1920x1080":
                switch system["Environment"]["DPI Scale"] {
                    case "100%", "125%", "150%":
                        CreateImagesFromCatalog("Full High Definition")
                }
            case "2560x1440":
                switch system["Environment"]["DPI Scale"] {
                    case "100%", "125%", "150%":
                        CreateImagesFromCatalog("Quad High Definition")
                }
            case "3840x2160":
                switch system["Environment"]["DPI Scale"] {
                    case "100%", "125%", "150%", "175%":
                        CreateImagesFromCatalog("Ultra High Definition")
                }
        }
    }

    if system["Directories"].Has("Application Image Override Directory") {
        applicationFolders := GetFoldersFromDirectory(system["Configuration"]["Settings"]["Application Image Override Directory"])
        for applicationFolder in applicationFolders {
            SplitPath(RTrim(applicationFolder, "\"), &applicationName)

            actionImageDirectories := GetFoldersFromDirectory(applicationFolder)
            for actionFolderPath in actionImageDirectories {
                SplitPath(RTrim(actionFolderPath, "\/"), &actionDirectoryName)

                if !RegExMatch(actionDirectoryName, "^\s*(.+?)\s*\(([a-p])\)\s*$", &matchResults) {
                    LogConclusion("Failed", logConclusionData, A_LineNumber, "Folder does not match format of Action Name (a...p): " . actionDirectoryName)
                }

                if !imageRegistry[applicationName].Has(matchResults[1]) {
                    LogConclusion("Failed", logConclusionData, A_LineNumber, "Can't be overriden as it doesn't exist for the application " . applicationName . " and Action Name: " . matchResults[1])
                }

                variantFound := false
                overridePath := actionFolderPath . system["Environment"]["Display Resolution"] . " @ " . system["Environment"]["DPI Scale"] . "."
                for variant in imageRegistry[applicationName][matchResults[1]] {
                    if variant["Variant"] = matchResults[2] {
                        overridePath := overridePath . variant["Extension"]
                        if FileExist(overridePath) {
                            variantFound := true
                            break
                        }
                    }
                }

                if !variantFound {
                    LogConclusion("Failed", logConclusionData, A_LineNumber, "Can't be overriden as variant " . matchResults[2] . " doesn't exist for the application " . applicationName . " and Action Name: " . matchResults[1])
                }

                for variant in imageRegistry[applicationName][matchResults[1]] {
                    if variant["Variant"] = matchResults[2] {
                        variant["Path"] := overridePath

                        imageDimensions   := StrSplit(GetImageDimensions(variant["Path"]), "x")
                        variant["Width"]  := imageDimensions[1] + 0
                        variant["Height"] := imageDimensions[2] + 0
                    }
                }
            }
        }
    }

    for application in applicationRegistry {
        if applicationRegistry[application]["Installed"] = true {
             ResolveFactsForApplication(application, applicationRegistry[application]["Counter"])
        }
    }

    if applicationRegistry["DaVinci Resolve Studio"]["Installed"] && applicationRegistry["DaVinci Resolve"]["Installed"] {
        if applicationRegistry["DaVinci Resolve Studio"]["Executable Path"] = applicationRegistry["DaVinci Resolve"]["Executable Path"] {
            executableDirectory := GetPathComponents(applicationRegistry["DaVinci Resolve Studio"]["Executable Path"])["Directory"]
            readMeFilePath      := executableDirectory . "Documents\ReadMe.html"

            if FileExist(readMeFilePath) {
                readMeFileContents := ReadFileOnHashMatch(readMeFilePath, GetFileHash(readMeFilePath, "SHA-256"))

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
    for outerKey, innerValue in applicationRegistry {
        if innerValue["Installed"] {
            configuration := outerKey . "|" . innerValue["Executable Path"] . "|" . innerValue["Executable Hash"] . "|" . innerValue["Executable Version"] . "|" . innerValue["Binary Type"]
            configuration := configuration . "|" . innerValue["Counter"] . "|" . SubStr(innerValue["Resolution Method"], 1, 1)
            installedApplications.Push(configuration)
            innerValue["Executable Hash"] := DecodeBaseToSha256Hex(innerValue["Executable Hash"], 86)
        }
    }

    BatchAppendExecutionLog("Application", installedApplications)

    LogConclusion("Completed", logConclusionData)
}

ExecutablePathViaReference(application) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("application As Map", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [application])

    applicationName := application["Application Executable Directory Candidates"][1]["Name"]
    executablePathSearchResult := Map(
        "Resolution Method", "Reference",
        "Success",           false
    )

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

    for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
        directoryName  := applicationExecutableDirectoryCandidate["Directory"]
        executableName := applicationExecutableDirectoryCandidate["Executable"]

        baseDirectoriesWithDirectoryNames := []
        for candidateBaseDirectory in candidateBaseDirectories {
            executablePath := candidateBaseDirectory . "\" . directoryName . "\" . executableName

            if DirExist(candidateBaseDirectory . "\" . directoryName . "\") {
                baseDirectoriesWithDirectoryNames.Push(candidateBaseDirectory . "\" . directoryName . "\")
            }

            if FileExist(executablePath) {
                executablePathSearchResult["Executable Path"] := executablePath
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
                        filename := GetPathComponents(filePath)["Filename"]

                        if InStr(filename, applicationName) && SubStr(filename, -extensionLength) = executableExtension {
                            executablePathSearchResult["Executable Path"] := filePath
                            break
                        }
                    }
                }
            }
        }

        StrReplace(directoryName, ".", "", , &dotOccurrencesInDirectoryNameCount)
        if executablePathSearchResult = "" && dotOccurrencesInDirectoryNameCount >= 2 {
            directoryNameSegments := StrSplit(directoryName, "\")
            versionSegmentIndex   := 0
            versionSegment        := ""

            for index, directoryNameSegment in directoryNameSegments {
                StrReplace(directoryNameSegment, ".", "", , &dotOccurrencesInDirectoryNameSegmentCount)

                if dotOccurrencesInDirectoryNameSegmentCount < 2 {
                    continue
                }

                versionSegmentIndex := index
                versionSegment      := directoryNameSegment
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

                highestVersionKey            := ""
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

                    Loop Files, applicationRootDirectory . "*", "D" {
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
                        for versionPart in StrSplit(versionText, ".") {
                            versionKey .= Format("{:06}", Number(versionPart))
                        }

                        executablePath := A_LoopFileFullPath
                        if relativePathAfterVersionSegment != "" {
                            executablePath .= "\" . relativePathAfterVersionSegment
                        }
                        executablePath .= "\" . executableName

                        if !FileExist(executablePath) {
                            continue
                        }

                        if highestVersionKey = "" || StrCompare(versionKey, highestVersionKey) > 0 {
                            highestVersionKey := versionKey
                            highestVersionExecutablePath := executablePath
                        }
                    }
                }

                if highestVersionExecutablePath != "" {
                    executablePathSearchResult["Executable Path"] := highestVersionExecutablePath
                }
            }
        }
    }

    if executablePathSearchResult.Has("Executable Path") {
        executablePathSearchResult["Success"] := true
    }

    return executablePathSearchResult
}

ExecutablePathViaAppPaths(application) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("application As Map", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [application])

    applicationName := application["Application Executable Directory Candidates"][1]["Name"]
    executablePathSearchResult := Map(
        "Resolution Method", "App Paths",
        "Success",           false
    )

    static appPathsBaseRegistryKeys := [
        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths",
        "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths",
        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"
    ]

    for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
        executableName := applicationExecutableDirectoryCandidate["Executable"]

        for appPathsBaseRegistryKey in appPathsBaseRegistryKeys {
            subkeyPath := appPathsBaseRegistryKey . "\" . executableName

            executablePath := ""
            try {
                executablePath := RegRead(subkeyPath, "")
            }

            if executablePath {
                if FileExist(executablePath) && GetPathComponents(executablePath)["Filename"] = executableName {
                    if applicationRegistry[applicationName].Has("Executable Collision") {
                        if !InStr(executablePath, applicationName) {
                            continue
                        }
                    }

                    executablePathSearchResult["Executable Path"] := executablePath
                    break
                }
            }
        }
    }

    if executablePathSearchResult.Has("Executable Path") {
        executablePathSearchResult["Success"] := true
    }

    return executablePathSearchResult
}

ExecutablePathViaUninstall(application) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("application As Map", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [application])

    applicationName := application["Application Executable Directory Candidates"][1]["Name"]
    executablePathSearchResult := Map(
        "Resolution Method", "Uninstall",
        "Success",           false
    )

    static uninstallBaseKeyPaths := [
        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
        executableName := applicationExecutableDirectoryCandidate["Executable"]

        if StrLen(applicationName) < 4 || StrLen(executableName) < 8 {
            continue
        }

        requiredLength := 4
        applicationNamePartiallyMatchesExecutableNameCondition := false
        
        SplitPath(executableName, , , , &executableNameNoExtension)
        shorterText   := StrLower(applicationName)
        longerText    := StrLower(executableNameNoExtension)
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
        Loop maximumStartIndex {
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
            Loop Reg, uninstallBaseKeyPath, "K" {
                uninstallSubKeyPath := A_LoopRegKey . "\" . A_LoopRegName

                displayName := ""
                try {
                    displayName := RegRead(uninstallSubKeyPath, "DisplayName")
                }

                if !displayName || !InStr(displayName, executableNameNoExtension) {
                    continue
                }

                displayIcon := ""
                try {
                    displayIcon := RegRead(uninstallSubKeyPath, "DisplayIcon")
                }

                if displayIcon {
                    executablePath := RegExReplace(displayIcon, ",-?\d+$")
                    executablePath := StrReplace(executablePath, "/", "\")
                    
                    if FileExist(executablePath) && (SubStr(StrLower(executablePath), -StrLen(executableName)) = StrLower(executableName)) {
                        if applicationRegistry[applicationName]["Executable Collision"] {
                            if !InStr(executablePath, applicationName) {
                                continue
                            }
                        }

                        executablePathSearchResult["Executable Path"] := executablePath
                        break 2
                    }
                }

                installLocation := ""
                try {
                    installLocation := RegRead(uninstallSubKeyPath, "InstallLocation")
                }

                if installLocation {
                    executablePath := RTrim(installLocation, "\/")
                    executablePath := executablePath . "\" . executableName

                    if FileExist(executablePath) {
                        if applicationRegistry[applicationName]["Executable Collision"] {
                            if !InStr(executablePath, applicationName) {
                                continue
                            }
                        }

                        executablePathSearchResult["Executable Path"] := executablePath
                        break 2
                    }
                }
            }
        }
    }

    if executablePathSearchResult.Has("Executable Path") {
        executablePathSearchResult["Success"] := true
    }

    return executablePathSearchResult
}

ResolveFactsForApplication(applicationName, counter) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("applicationName As String, counter As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [applicationName, counter])

    global applicationRegistry

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Excel Tiny Delay", 16, 16, 128)
        ConfigureMethodSetting(methodName, "Excel Short Delay", 256, 64, 2048)
        ConfigureMethodSetting(methodName, "Excel Medium Delay", 640, 160, 5120)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    executableHash := GetFileHash(applicationRegistry[applicationName]["Executable Path"], "SHA-256")
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

    static commandLineExecutablesHash    := GetFileHash(system["Directories"]["Mappings"] . "Command Line Executables.csv", "SHA-256")
    static commandLineExecutablesContent := ReadFileOnHashMatch(system["Directories"]["Mappings"] . "Command Line Executables.csv", commandLineExecutablesHash)
    static commandLineExecutablesArray   := ParseDelimitedRowsToArrayOfMaps(commandLineExecutablesContent)
    for commandLineExecutable in commandLineExecutablesArray {
        if applicationName = commandLineExecutable["Name"] {
            directoryPath := GetPathComponents(applicationRegistry[applicationName]["Executable Path"])["Directory"]

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
            excelTinyDelay   := settings["Excel Tiny Delay"].Get("Value")
            excelShortDelay  := settings["Excel Short Delay"].Get("Value")
            excelMediumDelay := settings["Excel Medium Delay"].Get("Value")

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
                LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to execute Excel Macro Code.")
            }

            applicationRegistry["Excel"]["International"] := Map()

            excelInternationalHash    := GetFileHash(system["Directories"]["Constants"] . "Excel International (2025-09-26).csv", "SHA-256")
            excelInternationalContent := ReadFileOnHashMatch(system["Directories"]["Constants"] . "Excel International (2025-09-26).csv", excelInternationalHash)
            excelInternationalArray   := ParseDelimitedRowsToArrayOfMaps(excelInternationalContent)
            for international in excelInternationalArray {
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

            wordInternationalHash    := GetFileHash(system["Directories"]["Constants"] . "Word International (2025-09-26).csv", "SHA-256")
            wordInternationalContent := ReadFileOnHashMatch(system["Directories"]["Constants"] . "Word International (2025-09-26).csv", wordInternationalHash)
            wordInternationalArray   := ParseDelimitedRowsToArrayOfMaps(wordInternationalContent)

            for international in wordInternationalArray {
                applicationRegistry["Word"]["International"][international["Label"]] := wordApplication.International[international["Value"]]
            }

            wordApplication.Quit()
            wordApplication := 0
    }
}

ValidateApplicationFact(applicationName, factName, factValue) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("applicationName As String, factName As String, factValue As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [applicationName, factName, factValue])

    if !applicationRegistry[applicationName].Has(factName) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, 'Application "' . applicationName . '" does not have a valid fact name: ' . factName)
    }

    if applicationRegistry[applicationName][factName] !== factValue {
        LogConclusion("Failed", logConclusionData, A_LineNumber, 'Application "' . applicationName . '" with fact name of "' . factName . '" does not match fact value of: ' . factValue)
    }
}

ValidateApplicationInstalled(applicationName) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("applicationName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [applicationName])

    if !applicationRegistry.Has(applicationName) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Application doesn't exist: " . applicationName)
    }

    if !applicationRegistry[applicationName]["Installed"] {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Application not installed: " . applicationName)
    }

    applicationIsInstalled := true

    return applicationIsInstalled
}

; **************************** ;
; Shared                       ;
; **************************** ;

CloseApplication(applicationName) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("applicationName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [applicationName], "Close Application (" . applicationName . ")")

    if !applicationRegistry.Has(applicationName) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Application not found: " . applicationName)
    }

    executableName := applicationRegistry[applicationName]["Executable Filename"]

    if !ProcessExist(executableName) {
        LogConclusion("Skipped", logConclusionData)
    } else {
        ProcessClose(executableName)
        ProcessWaitClose(executableName, 4)

        LogConclusion("Completed", logConclusionData)
    }
}

; **************************** ;
; Excel                        ;
; **************************** ;

ExcelExtensionRun(documentName, saveDirectory, code, displayName := "", aboutRange := "", aboutCondition := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    overlayValue      := (displayName = "" ? documentName : displayName) . " Excel Extension Run"
    static methodName := RegisterMethod("documentName As String, saveDirectory As String [Constraint: Directory], code As String, displayName As String [Optional], aboutRange As String [Optional], aboutCondition As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [documentName, saveDirectory, code, displayName, aboutRange, aboutCondition], overlayValue)

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Tiny Delay", 32, 16, 128)
        ConfigureMethodSetting(methodName, "Short Delay", 256, 128, 1536)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    tinyDelay  := settings["Tiny Delay"].Get("Value")
    shortDelay := settings["Short Delay"].Get("Value")

    excelFilePath := FileExistsInDirectory(documentName, saveDirectory, "xlsx")
    if excelFilePath = "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "documentName not found: " . documentName)
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
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Worksheet About not found with arguments passed in.")
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
        aboutValues["RetrievedDate"] := SubStr(StrReplace(aboutValues["RetrievedDate"], "Retrieved Date: ", ""), 1, -1)
        aboutValues["EditionName"] := StrReplace(aboutValues["EditionName"], "Edition Name: ", "")

        if aboutValues[aboutRange] = aboutCondition {
            OpenVisualBasicEditorAndRunCode(code, excelApplication)
            WaitForExcelToClose(excelProcessIdentifier)
            aboutWorksheet   := 0
            excelWorkbook    := 0
            excelApplication := 0
            ProcessWaitClose(excelProcessIdentifier, 2)

            LogConclusion("Completed", logConclusionData)
        } else if aboutRange = "ProgressionStatus" {
            conditionParts := StrSplit(aboutCondition, ", ")
            builtPrefix    := ""
            matchedIndex   := 0

            for index, currentPart in conditionParts {
                if index = 1 {
                    builtPrefix := currentPart
                } else {
                    builtPrefix := builtPrefix . ", " . currentPart
                }

                candidateValue := builtPrefix . "."
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

                LogConclusion("Completed", logConclusionData)
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

                LogConclusion("Skipped", logConclusionData)
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

            LogConclusion("Skipped", logConclusionData)
        }
    } else {
        OpenVisualBasicEditorAndRunCode(code, excelApplication)
        WaitForExcelToClose(excelProcessIdentifier)
        aboutWorksheet   := 0
        excelWorkbook    := 0
        excelApplication := 0
        ProcessWaitClose(excelProcessIdentifier, 2)

        LogConclusion("Completed", logConclusionData)
    }
}

ExcelStartingRun(documentName, saveDirectory, code, displayName := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    overlayValue      := (displayName = "" ? documentName : displayName) . " Excel Starting Run"
    static methodName := RegisterMethod("documentName As String, saveDirectory As String [Constraint: Directory], code As String, displayName As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [documentName, saveDirectory, code, displayName], overlayValue)

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Tiny Delay", 32, 16, 128)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    tinyDelay := settings["Tiny Delay"].Get("Value")

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
        LogConclusion("Skipped", logConclusionData)
    } else {
        sidecarPath := saveDirectory . documentName . ".txt"
        WriteTextToFile("", sidecarPath, "UTF-8")

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

        DeleteFile(sidecarPath)
        LogConclusion("Completed", logConclusionData)
    }
}

OpenVisualBasicEditorAndRunCode(code, excelApplication) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("code As String, excelApplication As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [code, excelApplication], "Open Visual Basic Editor and Run Code (Length: " . StrLen(code) . ")")

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Max Attempts", 4, 1, 16, 1)
        ConfigureMethodSetting(methodName, "Tiny Delay", 64, 16, 192, 32)
        ConfigureMethodSetting(methodName, "Short Delay", 384, 128, 1280, 64)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    maxAttempts := settings["Max Attempts"].Get("Value")
    tinyDelay   := settings["Tiny Delay"].Get("Value")
    shortDelay  := settings["Short Delay"].Get("Value")

    attempts          := 0
    loopWasSuccessful := false

    static personalMacroWorkbookPath := excelApplication.StartupPath . "\PERSONAL.XLSB"
    excelApplication.Workbooks.Open(personalMacroWorkbookPath)

    excelMainWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["Excel"]["Executable Filename"] . " ahk_class XLMAIN", 60)
    ActivateWindow(excelMainWindowSearchResults)
    excelApplication.DisplayAlerts := false

    activeWorkbook := unset
    for workbook in excelApplication.Workbooks {
        if workbook.Name = "PERSONAL.XLSB" {
            continue
        }

        activeWorkbook := workbook

        break
    }

    originalSheets := []
    for sheet in activeWorkbook.Sheets {
        originalSheets.Push(sheet.Name)
    }

    while attempts < maxAttempts {
        attempts++

        if attempts >= 2 {
            tinyDelay  := tinyDelay + (attempts * methodRegistry[methodName]["Settings"]["Tiny Delay"]["Delta"])
            shortDelay := shortDelay + (attempts * methodRegistry[methodName]["Settings"]["Short Delay"]["Delta"])

            logConclusionData["Context"] := "Failed on attempt " . attempts . " of " . maxAttempts . ". Tiny delay was " . tinyDelay . " milliseconds. Short delay was " . shortDelay . " milliseconds."
            
            IncreaseMethodSetting("KeyboardShortcut", "Tiny Delay")
        }

        KeyboardShortcut("ALT", "F11") ; Open the Visual Basic editor.

        sheetsBoundForDeletion := []
        for sheet in activeWorkbook.Sheets {
            isOriginalSheet := false

            for originalSheet in originalSheets {
                if originalSheet = sheet.Name {
                    isOriginalSheet := true
                }
            }

            if !isOriginalSheet {
                sheetsBoundForDeletion.Push(sheet.Name)
            }
        }

        if sheetsBoundForDeletion.Length != 0 {
            for sheetBoundForDeletion in sheetsBoundForDeletion {
                activeWorkbook.Sheets(sheetBoundForDeletion).Delete()
                Sleep(tinyDelay)
            }

            continue
        }

        excelApplication.DisplayAlerts := true

        visualBasicEditorWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["Excel"]["Executable Filename"] . " ahk_class wndclass_desked_gsk", 60, "Failed to open the Visual Basic editor via ALT+F11 in Excel.")
        ActivateWindow(visualBasicEditorWindowSearchResults, true)
        Sleep(tinyDelay + tinyDelay)
        PasteText(code, "'")
        Sleep(shortDelay)
        SendInput("{F5}") ; Run Sub/UserForm
        Sleep(shortDelay)

        visualBasicEditorMacroWindowSearchResults := SearchForWindow("Macros ahk_exe " . applicationRegistry["Excel"]["Executable Filename"] . " ahk_class #32770", 1)
        if visualBasicEditorMacroWindowSearchResults["Success"] {
            IncreaseMethodSetting(methodName, "Tiny Delay")
            IncreaseMethodSetting(methodName, "Short Delay")
            SendInput("{Esc}") ; Close Macros Window.
            Sleep(shortDelay)

            continue
        }

        loopWasSuccessful := true
        if attempts >= 2 {
            logConclusionData["Context"] := "Succeeded on attempt " . attempts . " of " . maxAttempts . ". Tiny delay is " . tinyDelay . " milliseconds. Short delay is " . shortDelay . " milliseconds."

            IncreaseMethodSetting(methodName, "Tiny Delay")
            IncreaseMethodSetting(methodName, "Short Delay")
        }

        break
    }

    if !loopWasSuccessful {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to open visual basic editor and run code in " . maxAttempts . " attempts.")
    }

    LogConclusion("Completed", logConclusionData)
}

WaitForExcelToClose(excelProcessIdentifier) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("excelProcessIdentifier As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [excelProcessIdentifier], "Wait for Excel to Close (PID: " . excelProcessIdentifier . ")")

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Total Seconds to Wait", 14400, 10, 43200)
        ConfigureMethodSetting(methodName, "Mouse Move Interval Seconds", 120, 1, 840)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    totalSecondsToWait       := settings["Total Seconds to Wait"].Get("Value")
    mouseMoveIntervalSeconds := settings["Mouse Move Interval Seconds"].Get("Value")

    secondDelay               := 1000
    secondsSinceLastMouseMove := 0

    userInterfaceIsGone := false
    Loop totalSecondsToWait {
        windowCount := WinGetList("ahk_pid " . excelProcessIdentifier).Length
        if windowCount = 0 {
            Sleep(secondDelay)
            userInterfaceIsGone := true
            break
        }

        secondsSinceLastMouseMove += 1
        if secondsSinceLastMouseMove >= mouseMoveIntervalSeconds {
            MouseMove 1, 0, 0, "R"
            MouseMove -1, 0, 0, "R"
            secondsSinceLastMouseMove := 0
        }

        Sleep(secondDelay)
    }

    if !userInterfaceIsGone {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Excel did not close within " . totalSecondsToWait . " seconds.")
    }

    LogConclusion("Completed", logConclusionData)
}

; **************************** ;
; SQL Server Management Studio ;
; **************************** ;

StartSqlServerManagementStudioAndConnect() {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [], "Start SQL Server Management Studio and Connect")

    static sqlServerManagementStudioIsInstalled := ValidateApplicationInstalled("SQL Server Management Studio")

    Run('"' . applicationRegistry["SQL Server Management Studio"]["Executable Path"] . '"')
    sqlServerManagementStudioConnectToServerWindowSearchResults := SearchForWindow("Connect ahk_exe " . applicationRegistry["SQL Server Management Studio"]["Executable Filename"], 60, "Connect to Server Window not found.")
    ActivateWindow(sqlServerManagementStudioConnectToServerWindowSearchResults)

    SendInput("{Enter}") ; Connect

    if !WinWaitClose("Connect ahk id " . sqlServerManagementStudioConnectToServerWindowSearchResults["Window Handle"],, 40) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Connection failed.")
    }

    sqlServerManagementStudioMainWindowSearchResults := SearchForWindow("ahk_exe " . applicationRegistry["SQL Server Management Studio"]["Executable Filename"], 60)
    ActivateWindow(sqlServerManagementStudioMainWindowSearchResults, true)

    LogConclusion("Completed", logConclusionData)
}

ExecuteSqlQueryAndSaveAsCsv(code, saveDirectory, filename) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("code As String, saveDirectory As String [Constraint: Directory], filename As String [Constraint: Filename]",  A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [code, saveDirectory, filename], "Execute SQL Query and Save (" . filename . ")")

    static sqlServerManagementStudioIsInstalled := ValidateApplicationInstalled("SQL Server Management Studio")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Max Attempts", 4, 1, 16, 1)
        ConfigureMethodSetting(methodName, "Times to Attempt", 120, 1, 7200)
        ConfigureMethodSetting(methodName, "Short Delay", 128, 32, 1280, 32)
        ConfigureMethodSetting(methodName, "Medium Delay", 512, 128, 3072, 48)
        ConfigureMethodSetting(methodName, "Long Delay", 1024, 256, 6144, 96)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    maxAttempts    := settings["Max Attempts"].Get("Value")
    timesToAttempt := settings["Times to Attempt"].Get("Value")
    shortDelay     := settings["Short Delay"].Get("Value")
    mediumDelay    := settings["Medium Delay"].Get("Value")
    longDelay      := settings["Long Delay"].Get("Value")

    attempts          := 0
    loopWasSuccessful := false
    savePath          := saveDirectory . filename . ".csv"

    while attempts < maxAttempts {
        attempts++

        if attempts >= 2 {
            shortDelay  := shortDelay + (attempts * methodRegistry[methodName]["Settings"]["Short Delay"]["Delta"])
            mediumDelay := mediumDelay + (attempts * methodRegistry[methodName]["Settings"]["Medium Delay"]["Delta"])
            longDelay   := longDelay + (attempts * methodRegistry[methodName]["Settings"]["Long Delay"]["Delta"])

            logConclusionData["Context"] := "Failed on attempt " . attempts . " of " . maxAttempts . ". Short delay was " . shortDelay . " milliseconds. Medium delay was " . mediumDelay . " milliseconds. Long delay was " . longDelay . " milliseconds."
            
            IncreaseMethodSetting("KeyboardShortcut", "Tiny Delay")
        }

        KeyboardShortcut("CTRL", "N") ; Query with Current Connection
        Sleep(longDelay)

        sqlServerManagementStudioConnectedImageSearchResults := SearchForDirectoryImage("SQL Server Management Studio", "Connected", 4)
        if !sqlServerManagementStudioConnectedImageSearchResults["Success"] {
            continue ; Failed to open a New Query window with the current connection, go to next attempt.
        }

        PasteText(code, "--")
        Sleep(mediumDelay)
        SendInput("{F5}") ; Run the selected portion of the query editor or the entire query editor if nothing is selected
        Sleep(longDelay)

        sqlServerManagementStudioConnectedImageSearchResults := SearchForDirectoryImage("SQL Server Management Studio", "Connected", 2)
        if sqlServerManagementStudioConnectedImageSearchResults["Success"] {
            continue ; Failed to run the query, go to next attempt.
        }

        sqlServerManagementStudioQueryExecutedSuccessfullyImageCoordinates := unset
        Loop timesToAttempt {
            sqlServerManagementStudioQueryExecutedSuccessfullyImageSearchResults := SearchForDirectoryImage("SQL Server Management Studio", "Query executed successfully", 2)
            if sqlServerManagementStudioQueryExecutedSuccessfullyImageSearchResults["Success"] {
                sqlServerManagementStudioQueryExecutedSuccessfullyImageCoordinates := ExtractImageCoordinates(sqlServerManagementStudioQueryExecutedSuccessfullyImageSearchResults)

                break
            }

            sqlServerManagementStudioQueryCompletedWithErrorsImageSearchResults := SearchForDirectoryImage("SQL Server Management Studio", "Query completed with errors", 2)
            if sqlServerManagementStudioQueryCompletedWithErrorsImageSearchResults["Success"] {
                break
            }
        }

        if !IsSet(sqlServerManagementStudioQueryExecutedSuccessfullyImageCoordinates) {
            continue ; Query failed, go to next attempt.
        }
        
        sqlServerManagementStudioResultsWindowCoordinates := ModifyScreenCoordinates(80, -80, sqlServerManagementStudioQueryExecutedSuccessfullyImageCoordinates)
        
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
        startTickCount          := A_TickCount

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
                LogConclusion("Failed", logConclusionData, A_LineNumber, "Timed out waiting for overwrite: " . savePath)
            }

            Sleep(mediumDelay)
        }

        loopWasSuccessful := true
        if attempts >= 2 {
            logConclusionData["Context"] := "Succeeded on attempt " . attempts . " of " . maxAttempts . ". Short delay is " . shortDelay . " milliseconds. Medium delay is " . mediumDelay . " milliseconds. Long delay is " . longDelay . " milliseconds."

            IncreaseMethodSetting(methodName, "Short Delay")
            IncreaseMethodSetting(methodName, "Medium Delay")
            IncreaseMethodSetting(methodName, "Long Delay")
        }
        
        break
    }

    if !loopWasSuccessful {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to execute SQL query and save as CSV in " . maxAttempts . " attempts.")
    }

    LogConclusion("Completed", logConclusionData)
}

; **************************** ;
; Toad for Oracle              ;
; **************************** ;

ExecuteAutomationApp(appName, runtimeDate := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("appName As String, runtimeDate As String [Optional] [Constraint: Raw Date Time]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [appName, runtimeDate], "Execute Automation App (" . appName . ")")

    static toadForOracleIsInstalled := ValidateApplicationInstalled("Toad for Oracle")

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Tiny Delay", 16, 16, 128)
        ConfigureMethodSetting(methodName, "Short Delay", 448, 128, 1536)
        ConfigureMethodSetting(methodName, "Medium Delay", 896, 256, 3584)
        ConfigureMethodSetting(methodName, "Long Delay", 1280, 640, 5120)
        ConfigureMethodSetting(methodName, "Massive Delay", 30000, 10000, 60000)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]
    
    tinyDelay    := settings["Tiny Delay"].Get("Value")
    shortDelay   := settings["Short Delay"].Get("Value")
    mediumDelay  := settings["Medium Delay"].Get("Value")
    longDelay    := settings["Long Delay"].Get("Value")
    massiveDelay := settings["Massive Delay"].Get("Value")

    static toadForOracleExecutableFilename := applicationRegistry["Toad for Oracle"]["Executable Filename"]

    if !ProcessExist(toadForOracleExecutableFilename) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Toad for Oracle process is not running.")
    }

    toadForOracleDatabaseLoginWindowSearchResults := SearchForWindow("ahk_exe " . toadForOracleExecutableFilename . " ahk_class TfrmLogin", 1)
    if toadForOracleDatabaseLoginWindowSearchResults["Success"] = true {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "No server connection is active in Toad for Oracle (Database Login window is open).")
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
                LogConclusion("Failed", logConclusionData, A_LineNumber, "Reconnect dialog did not close within " . Round(massiveDelay / 1000) . " seconds.")
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

    LogConclusion("Completed", logConclusionData)
}