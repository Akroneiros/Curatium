#Requires AutoHotkey v2.0
#Include Shared\Libraries (2025-08-30)

#Include Base Library.ahk
#Include Application Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

global projectDirectory := ExtractDirectory(A_ScriptFullPath) . "Projects\" . RegExReplace(LibraryTag(A_LineFile), "^ @ (.*?) \(.*\)$", "$1") . "\"

Main() {
    OverlayUpdateCustomLine(overlaySummaryKey := OverlayGenerateNextKey("[[Custom]]"), "Overlay Summary: " . StrReplace(LibraryTag(A_LineFile), " @ ", ""))
    OverlayInsertSpacer()

    ; ******************** ;
    ; Variables            ;
    ; ******************** ;
     
    OverlayUpdateCustomLine(overlayVariablesKey := OverlayGenerateNextKey("[[Custom]]"), "Initializing Variables" . overlayStatus["Beginning"])

    ; SQL queries from AdventureWorks2022.bak: https://learn.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver17&tabs=ssms
    adventureWorksSqlQueries := [
        ["Locations (v1, 2025-09-02)",    "5b282f1971ad80d92b4b0b92d268b2882070ba85ec8dbf29459938869474e26a"],
        ["Unit Measure (v1, 2025-09-02)", "1fd42a6843bbf663a3ac62857439e9e9fa1b2a5a0c0332816fa46bc96e6c07b8"]
    ]

    OverlayUpdateCustomLine(overlayVariablesKey, "Initializing Variables" . overlayStatus["Completed"])

    ; ******************** ;
    ; Requirements         ;
    ; ******************** ;

    OverlayUpdateCustomLine(overlayRequirementsKey := OverlayGenerateNextKey("[[Custom]]"), "Verifying Requirements" . overlayStatus["Beginning"])

    CreateSharedImages()
    RegisterApplications()
    
    ValidateApplicationFact("Excel", "Installed", "Yes")
    ValidateApplicationFact("Excel", "Personal Macro Workbook", "Enabled")
    ValidateApplicationFact("Excel", "Code Execution", "Full")
    ; ValidateApplicationFact("Notepad++", "Installed", "Yes")
    ValidateApplicationFact("SQL Server Management Studio", "Installed", "Yes")
    ; ValidateApplicationFact("Toad for Oracle", "Installed", "Yes")

    uniqueDirectories := []
    uniqueDirectories.Push("C:\Import\")
    uniqueDirectories.Push("C:\Export\")
    uniqueDirectories.Push(projectDirectory)

    SymbolLedgerBatchAppend("D", uniqueDirectories)

    for index, uniqueDirectoryValue in uniqueDirectories {
        if !InStr(uniqueDirectoryValue, A_UserName) {
            EnsureDirectoryExists(uniqueDirectoryValue)
            CleanOfficeLocksInFolder(uniqueDirectoryValue)
        }
    }

    projectFiles := GetFileListFromDirectory(projectDirectory)
    for index, entry in projectFiles {
        projectFiles[index] := [entry, ExtractTrailingDateAsIso(RegExReplace(ExtractFilename(entry, true), ".*\(([^)]*)\).*", "$1"), "Year-Month-Day")]
    }

    for index, entry in projectFiles {
        if AssignFileTimesAsLocalIso(entry[1], "Created") !== entry[2] . " 12:00:00" || AssignFileTimesAsLocalIso(entry[1], "Modified") !== entry[2] . " 12:00:00" {
            SetFileTimeFromLocalIsoDateTime(entry[1], entry[2] . " 12:00:00", "Created")
            SetFileTimeFromLocalIsoDateTime(entry[1], entry[2] . " 12:00:00", "Modified")
        }
    }

    OverlayUpdateCustomLine(overlayRequirementsKey, "Verifying Requirements" . overlayStatus["Completed"])

    ; ******************** ;
    ; Code                 ;
    ; ******************** ;

    OverlayUpdateCustomLine(overlayCodeKey := OverlayGenerateNextKey("[[Custom]]"), "Loading Code to Memory" .  overlayStatus["Beginning"])

    spreadsheetOperationsTemplate := AssignSpreadsheetOperationsTemplateCombined("v0.39")
    introCode := spreadsheetOperationsTemplate["Intro Code"]
    outroCode := spreadsheetOperationsTemplate["Outro Code"]
    frameworkDemonstrationCode := ReadFileOnHashMatch(projectDirectory . "Framework Demonstration (v1, 2025-09-02)" . ".txt", "4804034afb9166a0677233f01266d8523637812e503eaadab0d0e7f158c6035a")

    for index, entry in adventureWorksSqlQueries {
        entry.Push("C:\Import\")
        adventureWorksSqlQueries[index][2] := ReadFileOnHashMatch(projectDirectory . adventureWorksSqlQueries[index][1] . ".sql", adventureWorksSqlQueries[index][2])
    }

    dateOfToday := FormatTime(A_Now, "yyyyMMdd")
    allSQLFilesUpToDate := true
    filteredSQLQueries := []

    for query in adventureWorksSqlQueries {
        queryName  := query[1]
        targetFile := query[3] . queryName . ".csv"

        needsRun := false

        if !FileExist(targetFile) {
            needsRun := true
        } else {
            fileTimeRaw := FileGetTime(targetFile, "M")
            fileDate := FormatTime(fileTimeRaw, "yyyyMMdd")
            if fileDate != dateOfToday {
                needsRun := true
            }
        }

        if needsRun {
            filteredSQLQueries.Push(query)
            allSqlFilesUpToDate := false
        }
    }

    OverlayUpdateCustomLine(overlayCodeKey, "Loading Code to Memory" . overlayStatus["Completed"])

    ; ******************** ;
    ; Configuration        ;
    ; ******************** ;
    
    OverlayUpdateCustomLine(overlayConfigurationKey := OverlayGenerateNextKey("[[Custom]]"), "Selecting Configuration" . overlayStatus["Beginning"])

    ; Configuration here later.

    OverlayUpdateCustomLine(overlayConfigurationKey, "Selecting Configuration" . overlayStatus["Completed"])

    ; ******************** ;
    ; Main                 ;
    ; ******************** ;
    
    ; SSMS: Tools → Options... → Query Results → SQL Server → Results to Grid → Enable: Include column headers when copying or saving the results
    OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey := OverlayGenerateNextKey("[[Custom]]"), "AdventureWorks SQL Queries" . overlayStatus["Beginning"])
    if allSqlFilesUpToDate = false {
        StartMicrosoftSQLServerManagementStudioAndConnect()

        for index, query in filteredSqlQueries {
            OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey, "AdventureWorks SQL Queries (" . index . "/" . filteredSqlQueries.Length . ")" . overlayStatus["Beginning"])
            ExecuteSQLQueryAndSaveAsCsv(query[2], query[3], query[1])

            if filteredSqlQueries.Length = index {
                OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey, "AdventureWorks SQL Queries (" . index . "/" . filteredSqlQueries.Length . ")" . overlayStatus["Completed"])
            }
        }

        CloseMicrosoftSqlServerManagementStudio()
    } else {
        OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey, "AdventureWorks SQL Queries (Already Done)" . overlayStatus["Skipped"])
    }

    ExcelStartingRun("Framework Demonstration (v1, 2025-09-02)", "C:\Export\", CombineExcelCode(introCode, frameworkDemonstrationCode, outroCode))
}

Launcher() {
    LogEngine("Beginning")
    OverlayStart()

    for methodName in [
        ; "ValidateApplicationFact",
        ; "ExcelExtensionRun",
        ; "ExcelScriptExecution",
        "ExcelStartingRun",
        ; "WaitForExcelToClose",
        ; "WaitForExcelToLoad",
        ; "StartMicrosoftSQLServerManagementStudioAndConnect",
        ; "ExecuteSQLQueryAndSaveAsCsv",
        ; "CloseMicrosoftSqlServerManagementStudio",
        ; "ExecuteAutomationApp",
        ; "AssignSpreadsheetOperationsTemplateCombined",
        ; "AssignHeroAliases",
        ; "ModifyScreenCoordinates",
        ; "PerformMouseActionAtCoordinates",
        ; "AssignFileTimesAsLocalIso",
        ; "ExtractTrailingDateAsIso",
        ; "PreventSystemGoingIdleUntilRuntime",
        ; "SetFileTimeFromLocalIsoDateTime",
        ; "ValidateRuntimeDate",
        ; "WaitUntilFileIsModifiedToday",
        ; "CleanOfficeLocksInFolder",
        ; "CopyFileToTarget",
        ; "DeleteFile",
        ; "EnsureDirectoryExists",
        ; "FileExistsInDirectory",
        ; "GetFileListFromDirectory",
        ; "MoveFileToDirectory",
        ; "ReadFileOnHashMatch",
        ; "WriteBase64IntoImageFileWithHash",
        ; "AssignSharedImages",
        ; "CreateSharedImages",
        ; "RetrieveImageCoordinatesFromSegment",
        ; "AbortExecution",
        ; "OverlayChangeTransparency",
        ; "OverlayChangeVisibility",
        ; "OverlayHideLogForMethod",
        ; "OverlayShowLogForMethod",
        ; "OverlayInsertSpacer",
        ; "OverlayUpdateCustomLine"
    ] {
        OverlayShowLogForMethod(methodName)
    }

    Main()
    LogEngine("Completed")
}

Launcher()