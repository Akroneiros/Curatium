#Requires AutoHotkey v2.0
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

; **************************** ;
; Application Registry         ;
; **************************** ;

RegisterApplications() {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 5, Map(
        "Excel Tiny Delay", Map("Default", 16, "Floor", 16, "Ceiling", 128),
        "Excel Short Delay", Map("Default", 256, "Floor", 64, "Ceiling", 2048),
        "Excel Medium Delay", Map("Default", 640, "Floor", 160, "Ceiling", 5120),
        "Log to Execution Log", Map("Default", 1, "Floor", 0, "Ceiling", 1)))
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [], "Register Applications")

    global applicationRegistry

    newLine := system["Constants"]["New Line"]

    settings := methodRegistry[methodName]["Settings"]

    excelTinyDelay    := settings["Excel Tiny Delay"]["Value"]
    excelShortDelay   := settings["Excel Short Delay"]["Value"]
    excelMediumDelay  := settings["Excel Medium Delay"]["Value"]
    logToExecutionLog := settings["Log to Execution Log"]["Value"]

    applications                             := system["Mappings"]["Applications"]
    applicationExecutableDirectoryCandidates := system["Mappings"]["Application Executable Directory Candidates"]

    for application in applications {
        applicationName         := application["Name"]
        applicationCounter      := application["Counter"]
        applicationWhitelisted  := application["Whitelisted"]

        applicationRegistry[applicationName] := Map(
            "Counter",       applicationCounter,
            "Whitelisted",   applicationWhitelisted
        )

        if application.Has("Command Line Executable") {
            applicationRegistry[applicationName]["Command Line Executable"] := application["Command Line Executable"]
        }

        if application.Has("Shared Images") {
            applicationRegistry[applicationName]["Shared Images"] := application["Shared Images"]
        }
    }

    combinedApplicationExecutableDirectoryCandidates := []
    for applicationName, application in applicationRegistry {
        if application["Whitelisted"] {
            applicationRegistry[applicationName]["Application Executable Directory Candidates"] := []
            for applicationExecutableDirectoryCandidate in applicationExecutableDirectoryCandidates {
                if applicationName = applicationExecutableDirectoryCandidate["Name"] {
                    if applicationExecutableDirectoryCandidate["Source"] = "Project" {
                        applicationRegistry[applicationName]["Application Executable Directory Candidates"].Push(applicationExecutableDirectoryCandidate)
                        combinedApplicationExecutableDirectoryCandidates.Push(applicationExecutableDirectoryCandidate)
                    }
                }
            }

            for applicationExecutableDirectoryCandidate in applicationExecutableDirectoryCandidates {
                if applicationName = applicationExecutableDirectoryCandidate["Name"] {
                    if applicationExecutableDirectoryCandidate["Source"] = "Shared" {
                        applicationRegistry[applicationName]["Application Executable Directory Candidates"].Push(applicationExecutableDirectoryCandidate)
                        combinedApplicationExecutableDirectoryCandidates.Push(applicationExecutableDirectoryCandidate)
                    }
                }
            }
        }
    }

    for applicationName, application in applicationRegistry {
        if application.Has("Application Executable Directory Candidates") {
            for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
                applicationExecutableDirectoryCandidate["Root Directory"] := StrSplit(applicationExecutableDirectoryCandidate["Directory"], "\")[1]
            }
        }
    }

    for applicationName in applicationRegistry {
        for applicationExecutableDirectoryCandidate in combinedApplicationExecutableDirectoryCandidates {
            if applicationName != applicationExecutableDirectoryCandidate["Name"] {
                continue
            }

            for collisionCandidate in combinedApplicationExecutableDirectoryCandidates {
                if applicationExecutableDirectoryCandidate["Executable"] = collisionCandidate["Executable"] && applicationName != collisionCandidate["Name"] {

                    applicationRegistry[applicationName]["Executable Collision"] := true
                    break 2
                }
            }
        }
    }

    applicationRootDirectories := []
    for applicationName, application in applicationRegistry {
        if application["Whitelisted"] {
            for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
                applicationRootDirectories.Push(applicationExecutableDirectoryCandidate["Root Directory"])
            }
        }
    }

    applicationRootDirectories := RemoveDuplicatesFromArray(applicationRootDirectories)

    applicationDirectories := []
    for directoryPath in system["Mappings"]["Candidate Base Directories"] {
        Loop Files, directoryPath "\*", "D" {
            SplitPath(A_LoopFileFullPath, &candidateRootDirectory, &candidateParentDirectory)
            for applicationRootDirectory in applicationRootDirectories {
                if applicationRootDirectory = candidateRootDirectory {
                    if StrLen(candidateParentDirectory) = 2 {
                        candidateParentDirectory := candidateParentDirectory . "\"
                    }
                    applicationDirectories.Push(Map(
                        "Parent", candidateParentDirectory,
                        "Root",   candidateRootDirectory,
                        "Path",   A_LoopFileFullPath
                    ))
                    break
                }
            }
        }
    }

    appPathsBaseRegistryKeys := [
        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths",
        "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths",
        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"
    ]

    uninstallBaseRegistryKeys := [
        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]

    for applicationName, application in applicationRegistry {
        if !application["Whitelisted"] {
            continue
        }

        projectSourcePresent := false
        for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
            if applicationExecutableDirectoryCandidate["Source"] = "Project" {
                projectSourcePresent := true
                break
            }
        }

        dispatchTypes := ["Uninstall", "App Paths", "Reference"]
        if projectSourcePresent || application.Has("Executable Collision") {
            dispatchTypes := ["Reference", "Uninstall", "App Paths"]
        }

        for dispatchType in dispatchTypes {
            if application.Has("Executable Path") {
                break
            }

            for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
                executableDirectory       := applicationExecutableDirectoryCandidate["Directory"]
                executableName            := applicationExecutableDirectoryCandidate["Executable"]
                executableNameNoExtension := unset
                executablePath            := ""
                executableRootDirectory   := applicationExecutableDirectoryCandidate["Root Directory"]
                resolutionMethod          := unset

                if dispatchType != "Reference" {
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
                        substringToSearch := SubStr(shorterText, A_Index, requiredLength)
                        if InStr(longerText, substringToSearch) {
                            applicationNamePartiallyMatchesExecutableNameCondition := true
                            break
                        }
                    }

                    if !applicationNamePartiallyMatchesExecutableNameCondition {
                        continue
                    }
                }

                switch dispatchType {
                    case "App Paths":
                        for appPathsBaseRegistryKey in appPathsBaseRegistryKeys {
                            subkeyPath := appPathsBaseRegistryKey . "\" . executableName

                            try {
                                executablePath := RegRead(subkeyPath, "")
                            }

                            if executablePath {
                                if FileExist(executablePath) && GetPathComponents(executablePath)["Filename"] = executableName {
                                    if application.Has("Executable Collision") {
                                        if !InStr(executablePath, applicationName) {
                                            continue
                                        }
                                    }

                                    resolutionMethod := "App Paths"
                                    break
                                }
                            }
                        }
                    case "Reference":
                        for applicationDirectory in applicationDirectories {
                            if executableRootDirectory = applicationDirectory["Root"] {
                                applicationDirectory := StrReplace(applicationDirectory["Parent"] . "\", "\\", "\") . executableDirectory . "\"
                                executablePath       := applicationDirectory . executableName
                                if FileExist(executablePath) {
                                    application["Executable Path"]   := executablePath
                                    application["Resolution Method"] := "Reference"
                                    break 2
                                } else {
                                    application["Application Directory"] := applicationDirectory
                                }
                            }
                        }
                    case "Uninstall":
                        for uninstallBaseRegistryKey in uninstallBaseRegistryKeys {
                            Loop Reg, uninstallBaseRegistryKey, "K" {
                                uninstallSubRegistryKey := A_LoopRegKey . "\" . A_LoopRegName

                                displayName := ""
                                try {
                                    displayName := RegRead(uninstallSubRegistryKey, "DisplayName")
                                }

                                if !displayName || !InStr(displayName, executableNameNoExtension) {
                                    continue
                                }

                                displayIcon := ""
                                try {
                                    displayIcon := RegRead(uninstallSubRegistryKey, "DisplayIcon")
                                }

                                if displayIcon {
                                    executablePath := RegExReplace(displayIcon, ",-?\d+$")
                                    executablePath := StrReplace(executablePath, "/", "\")
                                    
                                    if FileExist(executablePath) && (SubStr(StrLower(executablePath), -StrLen(executableName)) = StrLower(executableName)) {
                                        if application.Has("Executable Collision") {
                                            if !InStr(executablePath, applicationName) {
                                                continue
                                            }
                                        }

                                        resolutionMethod := "Uninstall Display Icon"
                                        break 2
                                    }
                                }

                                installLocation := ""
                                try {
                                    installLocation := RegRead(uninstallSubRegistryKey, "InstallLocation")
                                }

                                if installLocation {
                                    executablePath := RTrim(installLocation, "\/")
                                    executablePath := executablePath . "\" . executableName

                                    if FileExist(executablePath) {
                                        if application.Has("Executable Collision") {
                                            if !InStr(executablePath, applicationName) {
                                                continue
                                            }
                                        }

                                        resolutionMethod := "Uninstall Install Location"
                                        break 2
                                    }
                                }
                            }
                        }
                }

                if IsSet(resolutionMethod) {
                    SplitPath(executablePath, , &executableDirectory)

                    requiredLength := 4
                    directoryPartiallyMatchesApplicationName := false

                    shorterText   := StrLower(applicationName)
                    longerText    := StrLower(executableDirectory)
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
                    if maximumStartIndex > 0 {
                        Loop maximumStartIndex {
                            substringToSearch := SubStr(shorterText, A_Index, requiredLength)
                            if InStr(longerText, substringToSearch) {
                                directoryPartiallyMatchesApplicationName := true
                                break
                            }
                        }
                    }

                    if !directoryPartiallyMatchesApplicationName {
                        continue
                    }

                    application["Executable Path"]   := executablePath
                    application["Resolution Method"] := resolutionMethod
                    break 2
                }
            }
        }
        
        if !application.Has("Executable Path") && application.Has("Application Directory") {
            if application["Application Executable Directory Candidates"].Length != 1 {
                totalDotOccurencesInDirectoryName := 0
                for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
                    dotCount := StrLen(applicationExecutableDirectoryCandidate["Directory"]) - StrLen(StrReplace(applicationExecutableDirectoryCandidate["Directory"], "."))
                    totalDotOccurencesInDirectoryName += dotCount
                }

                if totalDotOccurencesInDirectoryName / application["Application Executable Directory Candidates"].Length >= 2 {
                    for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
                        if InStr(application["Application Directory"], applicationExecutableDirectoryCandidate["Directory"]) {
                            applicationDirectory := application["Application Directory"]
                            executableName       := applicationExecutableDirectoryCandidate["Executable"]
                            
                            directoryNameSegments := StrSplit(applicationDirectory, "\")
                            versionSegmentIndex   := 0
                            
                            for index, directoryNameSegment in directoryNameSegments {
                                if directoryNameSegment = "" {
                                    continue
                                }
                                
                                StrReplace(directoryNameSegment, ".", "", , &dotOccurrencesInDirectoryNameSegment)
                                if dotOccurrencesInDirectoryNameSegment >= 2 {
                                    firstDigitPositionInSegment := RegExMatch(directoryNameSegment, "\d")
                                    if firstDigitPositionInSegment > 0 {
                                        versionSegmentIndex := index
                                        break
                                    }
                                }
                            }
                            
                            if versionSegmentIndex = 0 {
                                continue
                            }

                            relativePathBeforeVersionSegment := ""
                            relativePathAfterVersionSegment  := ""
                            
                            for index, directoryNameSegment in directoryNameSegments {
                                if directoryNameSegment = "" {
                                    continue
                                }

                                
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
                            
                            parentDirectory := ""
                            if relativePathBeforeVersionSegment != "" {
                                parentDirectory := relativePathBeforeVersionSegment . "\"
                            } else {
                                parentDirectory := RegExReplace(applicationDirectory, "[^\\]+\\$", "")
                            }
                            
                            highestVersionKey            := ""
                            highestVersionExecutablePath := ""
                            
                            Loop Files, parentDirectory . "*", "D" {
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
                                
                                executablePath := parentDirectory . folderName
                                if relativePathAfterVersionSegment != "" {
                                    executablePath .= "\" . relativePathAfterVersionSegment
                                }
                                executablePath .= "\" . executableName
                                
                                if FileExist(executablePath) && (highestVersionKey = "" || StrCompare(versionKey, highestVersionKey) > 0) {
                                    highestVersionKey            := versionKey
                                    highestVersionExecutablePath := executablePath
                                }
                            }
                            
                            if highestVersionExecutablePath != "" {
                                application["Executable Path"]   := highestVersionExecutablePath
                                application["Resolution Method"] := "Reference Multiple Dots"
                                break
                            }
                        }
                    }
                }
            }
        }
    }

    for applicationName, application in applicationRegistry {
        if application.Has("Executable Path") {
            application["Installed"] := true

            if application.Has("Application Directory") {
                application.Delete("Application Directory")
            }

            SplitPath(application["Executable Path"], &executableFilename, &directoryPath)

            if application.Has("Command Line Executable") {
                if FileExist(directoryPath . "\" . application["Command Line Executable"]) {
                    application["Command Line Executable Path"] := directoryPath . "\" . application["Command Line Executable"]
                }
            }

            executableVersion := "N/A"
            try {
                executableVersion := FileGetVersion(application["Executable Path"])
            }

            application["Executable Binary Type"] := DetermineWindowsBinaryType(application["Executable Path"])
            application["Executable Filename"]    := executableFilename
            application["Executable Hash"]        := GetFileHash(application["Executable Path"], "SHA-256")
            application["Executable Version"]     := executableVersion
        } else {
            application["Installed"] := false
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

                applicationRegistry[notInstalledApplication].Delete("Executable Binary Type")
                applicationRegistry[notInstalledApplication].Delete("Executable Filename")
                applicationRegistry[notInstalledApplication].Delete("Executable Hash")
                applicationRegistry[notInstalledApplication].Delete("Executable Path")
                applicationRegistry[notInstalledApplication].Delete("Executable Version")
                applicationRegistry[notInstalledApplication].Delete("Resolution Method")

                applicationRegistry[notInstalledApplication]["Installed"] := false
            }
        }
    }

    installedApplicationsWithImageLibraryDataCount := 0
    for applicationName, application in applicationRegistry {
        if application["Installed"] {
            if application.Has("Shared Images") {
                installedApplicationsWithImageLibraryDataCount++
            }
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

    applicationExecutableVersionsFileHash := GetFileHash(system["Paths"]["Application Executable Versions"], "SHA-256")
    applicationExecutableVersionsContent  := ReadFileOnHashMatch(system["Paths"]["Application Executable Versions"], applicationExecutableVersionsFileHash)
    applicationExecutableVersionsMapping  := ParseDelimitedRowsToArrayOfMaps(applicationExecutableVersionsContent)

    installedApplications := []
    for applicationName, application in applicationRegistry {
        if application["Installed"] {
            for applicationExecutableVersion in applicationExecutableVersionsMapping {
                if applicationName = applicationExecutableVersion["Name"] {
                    if application["Executable Hash"] = applicationExecutableVersion["Executable Hash"] {
                        application["Executable Version"] := applicationExecutableVersion["Executable Version"]
                    }
                }
            }

            switch applicationName {
                case "Cura":
                    if application["Executable Version"] = "N/A" {
                        SplitPath(application["Executable Path"], , &directoryName)
                        if RegExMatch(directoryName, "i)UltiMaker Cura ([\d\.]+)", &versionMatch) {
                            application["Executable Version"] := versionMatch[1]
                        }
                    }
                case "CyberChef":
                    if application["Executable Version"] = "N/A" {
                        filename := GetPathComponents(application["Executable Path"])["Filename"]
                        if RegExMatch(filename, "_v(\d+\.\d+(?:\.\d+)*)", &versionMatch) {
                            application["Executable Version"] := versionMatch[1]
                        }
                    }
                case "Excel":
                    CloseApplication("Excel")

                    excelApplication := ComObject("Excel.Application")
                    excelWorkbook    := excelApplication.Workbooks.Add()
                    excelWorksheet   := excelWorkbook.ActiveSheet
                    excelApplication.Visible := true

                    excelWindowHandle := excelApplication.Hwnd
                    while !excelWindowHandle := excelApplication.Hwnd {
                        Sleep(excelTinyDelay)
                    }
                    SetApplicationRegistryValue("Excel", "Process Identifier", WinGetPID("ahk_id " . excelWindowHandle))

                    excelMainWindowSearchResults := SearchForWindow("ahk_exe " . application["Executable Filename"] . " ahk_class XLMAIN", 60)
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
                        excelRecordMacroWindowSearchResults := SearchForWindow("Record Macro ahk_exe " . application["Executable Filename"], 60)
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
                        visualBasicEditorWindowSearchResults := SearchForWindow("ahk_exe " . application["Executable Filename"] . " ahk_class wndclass_desked_gsk", 60, "Failed to open the Visual Basic editor via ALT+F11 in Excel.")
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

                    application["Personal Macro Workbook"] := personalMacroWorkbookPath

                    excelMacroCode := "Sub Run()" . newLine . '    Range("A1").Value = "Cell"' . newLine . "End Sub"
                    OpenVisualBasicEditorAndRunCode(excelApplication, excelMacroCode)
                    Sleep(excelTinyDelay + excelTinyDelay)

                    if excelWorksheet.Range("A1").Value != "Cell" {
                        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to execute Excel Macro Code.")
                    }

                    userInterfaceLCID := "User Interface Language Code Identifier"
                    application["Environment"] := Map(
                        "Computer Name",      system["Environment"]["Computer Name"],
                        "Display Language",   system["Environment"]["Display Language"],
                        "Display Resolution", system["Environment"]["Display Resolution"],
                        "DPI Scale",          system["Environment"]["DPI Scale"],
                        "Excel Version",      application["Executable Version"],
                        "Input Language",     system["Environment"]["Input Language"],
                        "Keyboard Layout",    system["Environment"]["Keyboard Layout"],
                        "Operating System",   system["Environment"]["Operating System"]["Full Name"],
                        "QPC Frequency",      system["Environment"]["QPC Frequency"],
                        "Regional Format",    system["Environment"]["Regional Format"],
                        "Run Identifier",     system["Runtime"]["Run Identifier"],
                        "Time Zone",          system["Environment"]["Time Zone"]["Key Name"],
                        userInterfaceLCID,    excelApplication.LanguageSettings.LanguageID(2),
                        "Username",           system["Environment"]["Username"]
                    )

                    excelInternationalFileHash := "f22a6b4c3a81f479bb7844429d5effff494023ae29fdd414bed848d54143f0f0"
                    excelInternationalContent  := ReadFileOnHashMatch(system["Paths"]["Excel International"], excelInternationalFileHash)
                    excelInternationalConstant := ParseDelimitedRowsToArrayOfMaps(excelInternationalContent)

                    for rowMap in excelInternationalConstant {
                        rowMap["Value"] := rowMap["Value"] + 0
                    }

                    application["International Constant"] := excelInternationalConstant

                    application["International"] := Map()
                    for international in application["International Constant"] {
                        application["International"][international["Label"]] := excelApplication.International[international["Value"]]
                    }

                    for international, value in application["International"] {
                        if Type(value) = "Float" {
                            application["International"][international] := Round(value)
                        }
                    }

                    excelDefaultCellStylesFileHash := GetFileHash(system["Paths"]["Excel Default Cell Styles"], "SHA-256")
                    excelDefaultCellStylesContent  := ReadFileOnHashMatch(system["Paths"]["Excel Default Cell Styles"], excelDefaultCellStylesFileHash)
                    excelDefaultCellStylesMapping  := ParseDelimitedRowsToArrayOfMaps(excelDefaultCellStylesContent)

                    for excelDefaultCellStyle in excelDefaultCellStylesMapping {
                        excelDefaultCellStyle[userInterfaceLCID] := excelDefaultCellStyle[userInterfaceLCID] + 0
                    }

                    application["Default Cell Styles Mapping"] := excelDefaultCellStylesMapping

                    application["Default Cell Styles"] := ExtractRowFromArrayOfMapsOnHeaderCondition(excelDefaultCellStylesMapping, userInterfaceLCID, application["Environment"][userInterfaceLCID])
                    application["Default Cell Styles"].Delete(userInterfaceLCID)

                    excelWorkbook.Close(false)
                    excelApplication.DisplayAlerts := false
                    excelApplication.Quit()

                    excelWorksheet   := 0
                    excelWorkbook    := 0
                    excelApplication := 0
                    ProcessWaitClose(applicationRegistry["Excel"]["Process Identifier"], 2)
                case "SoapUI":
                    if application["Executable Version"] = "N/A" {
                        if RegExMatch(application["Executable Filename"], "i)SoapUI-(\d+\.\d+(?:\.\d+)*)\.exe$", &versionMatch) {
                            application["Executable Version"] := versionMatch[1]
                        }
                    }
                case "Word":
                    wordApplication := ComObject("Word.Application")

                    userInterfaceLCID := "User Interface Language Code Identifier"
                    application["Environment"] := Map(
                        "Computer Name",      system["Environment"]["Computer Name"],
                        "Display Language",   system["Environment"]["Display Language"],
                        "Display Resolution", system["Environment"]["Display Resolution"],
                        "DPI Scale",          system["Environment"]["DPI Scale"],
                        "Input Language",     system["Environment"]["Input Language"],
                        "Keyboard Layout",    system["Environment"]["Keyboard Layout"],
                        "Operating System",   system["Environment"]["Operating System"]["Full Name"],
                        "QPC Frequency",      system["Environment"]["QPC Frequency"],
                        "Regional Format",    system["Environment"]["Regional Format"],
                        "Run Identifier",     system["Runtime"]["Run Identifier"],
                        "Time Zone",          system["Environment"]["Time Zone"]["Key Name"],
                        userInterfaceLCID,    wordApplication.LanguageSettings.LanguageID(2),
                        "Username",           system["Environment"]["Username"],
                        "Word Version",       application["Executable Version"]
                    )

                    wordInternationalFileHash := "d586eccccd709b85ebabbcd09a339a828fc46945df05e680c6ca52403dae8755"
                    wordInternationalContent  := ReadFileOnHashMatch(system["Paths"]["Word International"], wordInternationalFileHash)
                    wordInternationalConstant := ParseDelimitedRowsToArrayOfMaps(wordInternationalContent)

                    for rowMap in wordInternationalConstant {
                        rowMap["Value"] := rowMap["Value"] + 0
                    }

                    application["International"] := Map()
                    for international in wordInternationalConstant {
                        application["International"][international["Label"]] := wordApplication.International[international["Value"]]
                    }

                    for international, value in application["International"] {
                        if Type(value) = "Float" {
                            application["International"][international] := Round(value)
                        }
                    }

                    wordApplication.Quit()
                    wordApplication := 0
            }

            resolutionMethodInitialism := unset
            switch application["Resolution Method"] {
                case "App Paths":
                    resolutionMethodInitialism := "AP"
                case "Reference":
                    resolutionMethodInitialism := "R"
                case "Reference Multiple Dots":
                    resolutionMethodInitialism := "RMD"
                case "Uninstall Display Icon":
                    resolutionMethodInitialism := "UDI"
                case "Uninstall Install Location":
                    resolutionMethodInitialism := "UIL"
            }

            configuration := application["Counter"] . "|" . application["Executable Path"] . "|" . EncodeSha256HexToBase(application["Executable Hash"], 86) . "|" . application["Executable Version"] . "|" . application["Executable Binary Type"]
            configuration := configuration . "|" . resolutionMethodInitialism
            installedApplications.Push(configuration)
        }
    }

    if logToExecutionLog {
        BatchAppendExecutionLog("Application", installedApplications)
    }

    LogConclusion("Completed", logConclusionData)
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

    static methodName := RegisterMethod("applicationName As String [Constraint: Application Name]", A_ThisFunc, A_LineFile, A_LineNumber + 2, Map(
        "Timeout", Map("Default", 4, "Floor", 1, "Ceiling", 120)))
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [applicationName], "Close Application (" . applicationName . ")")

    settings := methodRegistry[methodName]["Settings"]

    timeout := settings["Timeout"]["Value"]

    executableName := applicationRegistry[applicationName]["Executable Filename"]

    if !ProcessExist(executableName) {
        LogConclusion("Skipped", logConclusionData)
        return
    }

    closedProcessIdentifier := ProcessClose(executableName)
    if !closedProcessIdentifier {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to close process for application.")
    }

    if ProcessWaitClose(executableName, timeout) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Process did not close within the timeout of " . timeout . " seconds.")
    }

    LogConclusion("Completed", logConclusionData)
}

SetApplicationRegistryValue(applicationName, propertyName, propertyValue) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("applicationName As String [Constraint: Application Name], propertyName As String, propertyValue As Object", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [applicationName, propertyName, propertyValue])

    global applicationRegistry

    applicationRegistry[applicationName][propertyName] := propertyValue
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
; Excel                        ;
; **************************** ;

ExcelStartingRun(documentName, saveDirectory, code, spreadsheetOperationsTemplate, displayName := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    overlayValue      := (displayName = "" ? documentName : displayName) . " Excel Starting Run"
    static methodName := RegisterMethod("documentName As String, saveDirectory As String [Constraint: Directory], code As String, spreadsheetOperationsTemplate As Map, displayName As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [documentName, saveDirectory, code, spreadsheetOperationsTemplate, displayName], overlayValue)

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    excelFilePath    := SearchForUniqueFileInDirectory(documentName, saveDirectory, "xlsx")
    sentinelFilePath := SearchForUniqueFileInDirectory(documentName . " (Sentinel)", saveDirectory, "txt")

    if excelFilePath != "" && sentinelFilePath != "" {
        try {
            FileDelete(excelFilePath)
        } catch as excelFileDeleteFailedError {
            LogConclusion("Failed", logConclusionData, excelFileDeleteFailedError.Line, excelFileDeleteFailedError.Message)
        }

        try {
            FileDelete(sentinelFilePath)
        } catch as sentinelFileDeleteFailedError {
            LogConclusion("Failed", logConclusionData, sentinelFileDeleteFailedError.Line, sentinelFileDeleteFailedError.Message)
        }

        excelFilePath := ""
    } else if excelFilePath = "" && sentinelFilePath != "" {
        try {
            FileDelete(sentinelFilePath)
        } catch as sentinelFileDeleteFailedError {
            LogConclusion("Failed", logConclusionData, sentinelFileDeleteFailedError.Line, sentinelFileDeleteFailedError.Message)
        }
    }

    if excelFilePath != "" {
        LogConclusion("Skipped", logConclusionData)
        return
    }

    sentinelFilePath := saveDirectory . documentName . " (Sentinel)" . ".txt"

    try {
        sentinelFileHandle := FileOpen(sentinelFilePath, "w", "UTF-8-RAW")
        sentinelFileHandle.Write(system["Runtime"]["Run Identifier"])
        sentinelFileHandle.Close()
    } catch as sentinelFileWriteError {
        LogConclusion("Failed", logConclusionData, sentinelFileWriteError.Line, sentinelFileWriteError.Message)
    }

    excelApplication  := StartExcel()
    combinedExcelCode := CombineExcelCode(code, spreadsheetOperationsTemplate, excelApplication)

    OpenVisualBasicEditorAndRunCode(excelApplication, combinedExcelCode)
    WaitForExcelToClose()
    excelApplication := 0
    ProcessWaitClose(applicationRegistry["Excel"]["Process Identifier"], 2)

    try {
        FileDelete(sentinelFilePath)
    } catch as sentinelFileDeleteFailedError {
        LogConclusion("Failed", logConclusionData, sentinelFileDeleteFailedError.Line, sentinelFileDeleteFailedError.Message)
    }

    LogConclusion("Completed", logConclusionData)
}

ExcelExtensionRun(documentName, saveDirectory, code, spreadsheetOperationsTemplate, displayName := "", progressionStatusCondition := "", augmentationModulesCondition := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    overlayValue      := (displayName = "" ? documentName : displayName) . " Excel Extension Run"
    static methodName := RegisterMethod("documentName As String, saveDirectory As String [Constraint: Directory], code As String, spreadsheetOperationsTemplate As Map, displayName As String [Optional], " . 
        "progressionStatusCondition As String [Optional], augmentationModulesCondition As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 2, Map(
            "Medium Delay", Map("Default", 1024, "Floor", 256, "Ceiling", 4096)))
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"),
        [documentName, saveDirectory, code, spreadsheetOperationsTemplate, displayName, progressionStatusCondition, augmentationModulesCondition], overlayValue)

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    settings := methodRegistry[methodName]["Settings"]

    mediumDelay := settings["Medium Delay"]["Value"]

    excelFilePath := SearchForUniqueFileInDirectory(documentName, saveDirectory, "xlsx")
    if excelFilePath = "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, 'Parameter "' . "documentName" . '" failed validation. ' . "Unable to find file.")
    }

    excelApplication  := StartExcel(excelFilePath)
    combinedExcelCode := CombineExcelCode(code, spreadsheetOperationsTemplate, excelApplication)

    aboutWorksheet      := unset
    aboutWorksheetFound := false

    for worksheet in excelApplication.ActiveWorkbook.Worksheets {
        if worksheet.Name = "About" {
            aboutWorksheetFound := true
            break
        }
    }

    if !aboutWorksheetFound && (progressionStatusCondition != "" || augmentationModulesCondition != "") {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Worksheet About not found with conditions passed in.")
    }

    if progressionStatusCondition = "" && augmentationModulesCondition != "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "If conditions for Augmentation Modules are present it also requires conditions for Progression Status.")
    }

    if progressionStatusCondition = "" && augmentationModulesCondition = "" {
        OpenVisualBasicEditorAndRunCode(excelApplication, combinedExcelCode)
        WaitForExcelToClose()
        excelApplication := 0
        ProcessWaitClose(applicationRegistry["Excel"]["Process Identifier"], 2)

        LogConclusion("Completed", logConclusionData)
        return
    }

    aboutValues := unset
    if aboutWorksheetFound {
        aboutWorksheet := excelApplication.ActiveWorkbook.Worksheets("About")

        aboutValues := Map(
            "ProgressionStatus",   "A3",
            "AugmentationModules", "A4"
        )

        for fieldName, cellAddress in aboutValues {
            aboutValues[fieldName] := aboutWorksheet.Range(cellAddress).Value
        }

        aboutValues["ProgressionStatus"]   := StrReplace(aboutValues["ProgressionStatus"], "Progression Status: ", "")
        aboutValues["AugmentationModules"] := StrReplace(aboutValues["AugmentationModules"], "Augmentation Modules: ", "")
    }

    progressionConditionParts := StrSplit(progressionStatusCondition, ", ")
    progressionMatchedIndex   := 0
    progressionBuiltPrefix    := ""

    for index, currentPart in progressionConditionParts {
        if progressionBuiltPrefix = "" {
            progressionBuiltPrefix := currentPart
        } else {
            progressionBuiltPrefix .= ", " . currentPart
        }

        if aboutValues["ProgressionStatus"] = progressionBuiltPrefix . "." {
            progressionMatchedIndex := index
            break
        }
    }

    if progressionStatusCondition != "" && augmentationModulesCondition = "" {
        if progressionMatchedIndex > 0 {
            OpenVisualBasicEditorAndRunCode(excelApplication, combinedExcelCode)
            WaitForExcelToClose()
            aboutWorksheet   := 0
            excelApplication := 0
            ProcessWaitClose(applicationRegistry["Excel"]["Process Identifier"], 2)

            LogConclusion("Completed", logConclusionData)
        } else {
            activeWorkbook := excelApplication.ActiveWorkbook
            activeWorkbook.Close(false)
            excelApplication.DisplayAlerts := false
            excelApplication.Quit()

            aboutWorksheet   := 0
            activeWorkbook   := 0
            excelApplication := 0
            Sleep(mediumDelay)

            LogConclusion("Skipped", logConclusionData)
        }
    } else if progressionStatusCondition != "" && augmentationModulesCondition != "" {
        augmentationConditionParts := StrSplit(augmentationModulesCondition, ", ")
        augmentationMatchedIndex   := 0
        augmentationBuiltPrefix    := ""

        for index, currentPart in augmentationConditionParts {
            if augmentationBuiltPrefix = "" {
                augmentationBuiltPrefix := currentPart
            } else {
                augmentationBuiltPrefix .= ", " . currentPart
            }

            if aboutValues["AugmentationModules"] = augmentationBuiltPrefix . "." {
                augmentationMatchedIndex := index
                break
            }
        }

        if progressionMatchedIndex > 0 && augmentationMatchedIndex > 0 {
            OpenVisualBasicEditorAndRunCode(excelApplication, combinedExcelCode)
            WaitForExcelToClose()
            aboutWorksheet   := 0
            excelApplication := 0
            ProcessWaitClose(applicationRegistry["Excel"]["Process Identifier"], 2)

            LogConclusion("Completed", logConclusionData)
        } else {
            activeWorkbook := excelApplication.ActiveWorkbook
            activeWorkbook.Close(false)
            excelApplication.DisplayAlerts := false
            excelApplication.Quit()

            aboutWorksheet   := 0
            activeWorkbook   := 0
            excelApplication := 0
            Sleep(mediumDelay)

            LogConclusion("Skipped", logConclusionData)
        }
    }
}

OpenVisualBasicEditorAndRunCode(excelApplication, code) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("excelApplication As Object, code As String", A_ThisFunc, A_LineFile, A_LineNumber + 4, Map(
        "Max Attempts", Map("Default", 4, "Floor", 1, "Ceiling", 16, "Delta", 1),
        "Tiny Delay", Map("Default", 64, "Floor", 16, "Ceiling", 192, "Delta", 32),
        "Short Delay", Map("Default", 384, "Floor", 128, "Ceiling", 1280, "Delta", 64)))
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [excelApplication, code], "Open Visual Basic Editor and Run Code (Length: " . StrLen(code) . ")")

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    settings := methodRegistry[methodName]["Settings"]

    maxAttempts := settings["Max Attempts"]["Value"]
    tinyDelay   := settings["Tiny Delay"]["Value"]
    shortDelay  := settings["Short Delay"]["Value"]

    attempts          := 0
    loopWasSuccessful := false

    excelApplication.Workbooks.Open(applicationRegistry["Excel"]["Personal Macro Workbook"])

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

StartExcel(excelFilePath := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("excelFilePath As String [Optional] [Constraint: Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1, Map(
        "Tiny Delay", Map("Default", 32, "Floor", 16, "Ceiling", 128)))
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [excelFilePath])

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    settings := methodRegistry[methodName]["Settings"]

    tinyDelay := settings["Tiny Delay"]["Value"]

    If excelFilePath != "" {
        If !InStr(excelFilePath, ".xlsx") {
            LogConclusion("Failed", logConclusionData, A_LineNumber, 'Parameter "' . "excelFilePath" . '" failed validation. ' . "Path lacks the required extension of xlsx.")
        }
    }

    excelApplication := ComObject("Excel.Application")
    if excelFilePath = "" {
        excelApplication.Workbooks.Add()
    } else {
        excelApplication.Workbooks.Open(excelFilePath, 0)
    }

    excelApplication.Visible := true

    excelWindowHandle := excelApplication.Hwnd
    while !excelWindowHandle := excelApplication.Hwnd {
        Sleep(tinyDelay)
    }

    SetApplicationRegistryValue("Excel", "Process Identifier", WinGetPID("ahk_id " . excelWindowHandle))

    return excelApplication
}

WaitForExcelToClose() {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 3, Map(
        "Total Seconds to Wait", Map("Default", 14400, "Floor", 10, "Ceiling", 43200),
        "Mouse Move Interval Seconds", Map("Default", 120, "Floor", 1, "Ceiling", 840)))
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [], "Wait for Excel to Close")

    static excelIsInstalled := ValidateApplicationInstalled("Excel")

    settings := methodRegistry[methodName]["Settings"]

    totalSecondsToWait       := settings["Total Seconds to Wait"]["Value"]
    mouseMoveIntervalSeconds := settings["Mouse Move Interval Seconds"]["Value"]

    secondDelay               := 1000
    secondsSinceLastMouseMove := 0

    userInterfaceIsGone := false
    Loop totalSecondsToWait {
        windowCount := WinGetList("ahk_pid " . applicationRegistry["Excel"]["Process Identifier"]).Length
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

    static methodName := RegisterMethod("code As String, saveDirectory As String [Constraint: Directory], filename As String [Constraint: Filename]",  A_ThisFunc, A_LineFile, A_LineNumber + 6, Map(
        "Max Attempts", Map("Default", 4, "Floor", 1, "Ceiling", 16, "Delta", 1),
        "Times to Attempt", Map("Default", 120, "Floor", 1, "Ceiling", 7200),
        "Short Delay", Map("Default", 128, "Floor", 32, "Ceiling", 1280, "Delta", 32),
        "Medium Delay", Map("Default", 512, "Floor", 128, "Ceiling", 3072, "Delta", 48),
        "Long Delay", Map("Default", 1024, "Floor", 256, "Ceiling", 6144, "Delta", 96)))
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [code, saveDirectory, filename], "Execute SQL Query and Save (" . filename . ")")

    static sqlServerManagementStudioIsInstalled := ValidateApplicationInstalled("SQL Server Management Studio")

    settings := methodRegistry[methodName]["Settings"]

    maxAttempts    := settings["Max Attempts"]["Value"]
    timesToAttempt := settings["Times to Attempt"]["Value"]
    shortDelay     := settings["Short Delay"]["Value"]
    mediumDelay    := settings["Medium Delay"]["Value"]
    longDelay      := settings["Long Delay"]["Value"]

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
        startTickCount          := DllCall("Kernel32\GetTickCount64", "UInt64")

        fileExistsAlready := !!FileExist(savePath)

        if !fileExistsAlready {
            while !FileExist(savePath) && (DllCall("Kernel32\GetTickCount64", "UInt64") - startTickCount) < maximumWaitMilliseconds {
                Sleep(shortDelay)
            }
        }

        if fileExistsAlready {
            previousModifiedTime := FileGetTime(savePath, "M")
            Sleep(longDelay)
            SendInput("y") ; Yes
            startTickCount := DllCall("Kernel32\GetTickCount64", "UInt64")
            Sleep(longDelay)
            while FileGetTime(savePath, "M") = previousModifiedTime && (DllCall("Kernel32\GetTickCount64", "UInt64") - startTickCount) < maximumWaitMilliseconds {
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

    static methodName := RegisterMethod("appName As String, runtimeDate As String [Optional] [Constraint: Raw Date Time]", A_ThisFunc, A_LineFile, A_LineNumber + 6, Map(
        "Tiny Delay", Map("Default", 16, "Floor", 16, "Ceiling", 128),
        "Short Delay", Map("Default", 448, "Floor", 128, "Ceiling", 1536),
        "Medium Delay", Map("Default", 896, "Floor", 256, "Ceiling", 3584),
        "Long Delay", Map("Default", 1280, "Floor", 640, "Ceiling", 5120),
        "Massive Delay", Map("Default", 30000, "Floor", 10000, "Ceiling", 60000)))
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [appName, runtimeDate], "Execute Automation App (" . appName . ")")

    static toadForOracleIsInstalled := ValidateApplicationInstalled("Toad for Oracle")

    settings := methodRegistry[methodName]["Settings"]
    
    tinyDelay    := settings["Tiny Delay"]["Value"]
    shortDelay   := settings["Short Delay"]["Value"]
    mediumDelay  := settings["Medium Delay"]["Value"]
    longDelay    := settings["Long Delay"]["Value"]
    massiveDelay := settings["Massive Delay"]["Value"]

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
    
    overallStartTickCount := DllCall("Kernel32\GetTickCount64", "UInt64")
    firstSeenTickCount    := 0
    dialogHasAppeared     := false

    while true {
        dialogExists := WinExist("ahk_exe " . toadForOracleExecutableFilename . " ahk_class TReconnectForm")

        if !dialogHasAppeared {
            if dialogExists != false {
                dialogHasAppeared := true
                firstSeenTickCount := DllCall("Kernel32\GetTickCount64", "UInt64")
            } else if DllCall("Kernel32\GetTickCount64", "UInt64") - overallStartTickCount >= (longDelay + longDelay) {
                break
            }
        } else {
            if !dialogExists {
                break
            }

            if DllCall("Kernel32\GetTickCount64", "UInt64") - firstSeenTickCount >= massiveDelay {
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