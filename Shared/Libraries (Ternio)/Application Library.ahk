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

    static methodName := RegisterMethod("", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [], "Register Applications")

    global applicationRegistry

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Excel Tiny Delay", 16, 16, 128)
        ConfigureMethodSetting(methodName, "Excel Short Delay", 256, 64, 2048)
        ConfigureMethodSetting(methodName, "Excel Medium Delay", 640, 160, 5120)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]

    excelTinyDelay   := settings["Excel Tiny Delay"].Get("Value")
    excelShortDelay  := settings["Excel Short Delay"].Get("Value")
    excelMediumDelay := settings["Excel Medium Delay"].Get("Value")

    applications                             := system["Mappings"]["Applications"]
    applicationExecutableDirectoryCandidates := system["Mappings"]["Application Executable Directory Candidates"]

    for application in applications {
        applicationName         := application["Name"]
        applicationCounter      := application["Counter"] + 0
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
            switch dispatchType {
                case "App Paths":
                    if application.Has("Executable Path") {
                        break
                    }

                    static appPathsBaseRegistryKeys := [
                        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths",
                        "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths",
                        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"
                    ]

                    for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
                        executableName  := applicationExecutableDirectoryCandidate["Executable"]

                        for appPathsBaseRegistryKey in appPathsBaseRegistryKeys {
                            subkeyPath := appPathsBaseRegistryKey . "\" . executableName

                            executablePath := ""
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

                                    application["Executable Path"]   := executablePath
                                    application["Resolution Method"] := "App Paths"
                                    break
                                }
                            }
                        }
                    }
                case "Reference":
                    if application.Has("Executable Path") {
                        break
                    }

                    for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
                        applicationExecutableDirectory     := applicationExecutableDirectoryCandidate["Directory"]
                        applicationExecutableFilename      := applicationExecutableDirectoryCandidate["Executable"]
                        applicationExecutableRootDirectory := applicationExecutableDirectoryCandidate["Root Directory"]

                        for applicationDirectory in applicationDirectories {
                            if applicationExecutableRootDirectory = applicationDirectory["Root"] {
                                applicationDirectory := applicationDirectory["Parent"] . "\" . applicationExecutableDirectory . "\"
                                executablePath       := applicationDirectory . applicationExecutableFilename
                                if FileExist(executablePath) {
                                    application["Executable Path"]   := executablePath
                                    application["Resolution Method"] := "Reference"
                                    break
                                } else {
                                    application["Application Directory"] := applicationDirectory
                                }
                            }
                        }
                    }
                case "Uninstall":
                    if application.Has("Executable Path") {
                        break
                    }

                    static uninstallBaseKeyPaths := [
                        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                        "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                        "HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                    ]

                    for applicationExecutableDirectoryCandidate in application["Application Executable Directory Candidates"] {
                        executableName  := applicationExecutableDirectoryCandidate["Executable"]

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
                                        if application.Has("Executable Collision") {
                                            if !InStr(executablePath, applicationName) {
                                                continue
                                            }
                                        }

                                        application["Executable Path"]   := executablePath
                                        application["Resolution Method"] := "Uninstall"
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
                                        if application.Has("Executable Collision") {
                                            if !InStr(executablePath, applicationName) {
                                                continue
                                            }
                                        }

                                        application["Executable Path"]   := executablePath
                                        application["Resolution Method"] := "Uninstall"
                                        break 2
                                    }
                                }
                            }
                        }
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
                                application["Resolution Method"] := "Reference"
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
    installedApplications := []
    installedApplicationsWithImageLibraryDataCount := 0

    for applicationName, application in applicationRegistry {
        if application["Installed"] {
            if application.Has("Shared Images") {
                installedApplicationsWithImageLibraryDataCount++
            }

            switch applicationName {
                case "Capture2Text":
                    if application["Executable Hash"] = "d90d3684ccd34128556d33623a2d079400754ee9732bd1df7274f00a8e4fbb72" || application["Executable Hash"] = "320da826e0eddf763fd423c497ce5cd11703ff894b39da2f7b7db3ee78b1890f" {
                        application["Executable Version"] := "4.6.3"
                    }
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
                case "Exact Audio Copy":
                    if application["Executable Hash"] = "a169a5cad41cc341f2b108d8a38eaf957c817df1f971ac13aa1fb731f957a204" {
                        application["Executable Version"] := "1.8"
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
                    excelProcessIdentifier := WinGetPID("ahk_id " . excelWindowHandle)

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

                    excelMacroCode := "Sub Run()" . "`r`n" . '    Range("A1").Value = "Cell"' . "`r`n" . "End Sub"
                    OpenVisualBasicEditorAndRunCode(excelMacroCode, excelApplication)
                    Sleep(excelTinyDelay + excelTinyDelay)

                    if excelWorksheet.Range("A1").Value != "Cell" {
                        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to execute Excel Macro Code.")
                    }

                    application["Personal Macro Workbook"] := personalMacroWorkbookPath

                    application["International"] := Map()
                    for international in system["Constants"]["Excel International"] {
                        application["International"][international["Label"]] := excelApplication.International[international["Value"]]
                    }

                    excelWorkbook.Close(false)
                    excelApplication.DisplayAlerts := false
                    excelApplication.Quit()

                    excelWorksheet   := 0
                    excelWorkbook    := 0
                    excelApplication := 0
                    ProcessWaitClose(excelProcessIdentifier, 2)
                case "GtkHash":
                    if application["Executable Hash"] = "bf4ee99fac496949a6619d90994fed19b2a199b7ea6af126cb8ab84555c73928" {
                        application["Executable Version"] := "1.5"
                    }
                case "SoapUI":
                    if application["Executable Version"] = "N/A" {
                        if RegExMatch(application["Executable Filename"], "i)SoapUI-(\d+\.\d+(?:\.\d+)*)\.exe$", &versionMatch) {
                            application["Executable Version"] := versionMatch[1]
                        }
                    }
                case "Word":
                    wordApplication := ComObject("Word.Application")

                    application["International"] := Map()
                    for international in system["Constants"]["Word International"] {
                        application["International"][international["Label"]] := wordApplication.International[international["Value"]]
                    }

                    wordApplication.Quit()
                    wordApplication := 0
            }

            configuration := applicationName . "|" . application["Executable Path"] . "|" . application["Executable Hash"] . "|" . application["Executable Version"] . "|" . application["Executable Binary Type"]
            configuration := configuration . "|" . application["Counter"] . "|" . SubStr(application["Resolution Method"], 1, 1)
            installedApplications.Push(configuration)
        }
    }

    BatchAppendExecutionLog("Application", installedApplications)

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