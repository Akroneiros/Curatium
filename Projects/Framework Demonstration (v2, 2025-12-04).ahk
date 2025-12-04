#Requires AutoHotkey v2.0
#Include ..\Shared\Libraries (2025-12-04)

#Include Base Library.ahk
#Include Application Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

global mainWindow := unset

Main() {
    LogEngine("Beginning")

    OverlayStart()

    for methodName in [
        "ExcelStartingRun"
    ] {
        OverlayShowLogForMethod(methodName)
    }

    OverlayUpdateCustomLine(overlaySummaryKey := OverlayGenerateNextKey("[[Custom]]"), "Overlay Summary: " . system["Project Name"])
    OverlayInsertSpacer()

    ; ******************** ;
    ; Variables            ;
    ; ******************** ;
     
    OverlayUpdateCustomLine(overlayVariablesKey := OverlayGenerateNextKey("[[Custom]]"), "Initializing Variables" . overlayStatus["Beginning"])

    ; SQL queries from AdventureWorks2022.bak: https://learn.microsoft.com/en-us/sql/samples/adventureworks-install-configure?view=sql-server-ver17&tabs=ssms
    adventureWorksSqlQueries := [
        ["Locations (v1, 2025-09-04)",    "5b282f1971ad80d92b4b0b92d268b2882070ba85ec8dbf29459938869474e26a"],
        ["Unit Measure (v1, 2025-09-04)", "1fd42a6843bbf663a3ac62857439e9e9fa1b2a5a0c0332816fa46bc96e6c07b8"]
    ]

    OverlayUpdateCustomLine(overlayVariablesKey, "Initializing Variables" . overlayStatus["Completed"])

    ; ******************** ;
    ; Requirements         ;
    ; ******************** ;

    OverlayUpdateCustomLine(overlayRequirementsKey := OverlayGenerateNextKey("[[Custom]]"), "Verifying Requirements" . overlayStatus["Beginning"])

    RegisterApplications()
    CreateApplicationImages()
    ValidateDisplayScaling()
   
    uniqueDirectories := []
    uniqueDirectories.Push("C:\Import\")
    uniqueDirectories.Push("C:\Export\")
    uniqueDirectories.Push(system["Project Directory"])

    BatchAppendSymbolLedger("D", uniqueDirectories)

    for index, uniqueDirectoryValue in uniqueDirectories {
        if !InStr(uniqueDirectoryValue, A_UserName) {
            EnsureDirectoryExists(uniqueDirectoryValue)
            CleanOfficeLocksInFolder(uniqueDirectoryValue)
        }
    }

    projectFiles := GetFilesFromDirectory(system["Project Directory"])
    for index, entry in projectFiles {
        if InStr(entry, ".csv") {
            projectFiles[index] := [entry, ""]
            continue
        }

        projectFiles[index] := [entry, ExtractTrailingDateAsIso(RegExReplace(ExtractFilename(entry, true), ".*\(([^)]*)\).*", "$1"), "Year-Month-Day")]
    }

    for index, entry in projectFiles {
        if InStr(entry[1], ".csv") {
            continue
        }

        if AssignFileTimeAsLocalIso(entry[1], "Created") !== entry[2] . " 12:00:00" || AssignFileTimeAsLocalIso(entry[1], "Modified") !== entry[2] . " 12:00:00" {
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
    frameworkDemonstrationCode := ReadFileOnHashMatch(system["Project Directory"] . "Framework Demonstration (v1, 2025-09-04)" . ".vba", "9cb31a09306cc07b11e05f7d94b2442fdc9c83cedaba7cd82b2bf3c242ad7cf2")

    for index, entry in adventureWorksSqlQueries {
        entry.Push("C:\Import\")
        adventureWorksSqlQueries[index][2] := ReadFileOnHashMatch(system["Project Directory"] . adventureWorksSqlQueries[index][1] . ".tsql", adventureWorksSqlQueries[index][2])
    }

    dateOfToday := FormatTime(A_Now, "yyyyMMdd")
    allSqlFilesUpToDate := true
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
    OverlayInsertSpacer()

    ; ******************** ;
    ; Main                 ;
    ; ******************** ;

    ; SSMS: Tools -> Options... -> Query Results -> SQL Server -> Results to Grid -> Enable: Include column headers when copying or saving the results.
    OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey := OverlayGenerateNextKey("[[Custom]]"), "AdventureWorks SQL Queries" . overlayStatus["Beginning"])

    if applicationRegistry["SQL Server Management Studio"]["Installed"] = true {
        if allSqlFilesUpToDate = false {
            StartSqlServerManagementStudioAndConnect()

            for index, query in filteredSqlQueries {
                OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey, "AdventureWorks SQL Queries (" . index . "/" . filteredSqlQueries.Length . ")" . overlayStatus["Beginning"])
                ExecuteSqlQueryAndSaveAsCsv(query[2], query[3], query[1])

                if filteredSqlQueries.Length = index {
                    OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey, "AdventureWorks SQL Queries (" . index . "/" . filteredSqlQueries.Length . ")" . overlayStatus["Completed"])
                }
            }

            CloseApplication("SQL Server Management Studio")
        } else {
            OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey, "AdventureWorks SQL Queries (Already Done)" . overlayStatus["Skipped"])
        }
    } else {
        OverlayUpdateCustomLine(adventureWorksSqlQueriesStatusKey, "AdventureWorks SQL Queries (Not Installed)" . overlayStatus["Skipped"])
    }

    ExcelStartingRun("Framework Demonstration (v1, 2025-09-04)", "C:\Export\", CombineCode(introCode, frameworkDemonstrationCode, outroCode))

    LogEngine("Completed")
}

BaseHashConverter() {
    LogEngine("Beginning")

    selectedDirectory := ""

    childWindow := Gui(, "")
    childWindow.MarginX := 12
    childWindow.MarginY := 12

    hashText := childWindow.Add("Edit", "y+10 w480 r1 Center -VScroll -HScroll Limit64", "")
    hashText.SetFont("s10", "Consolas")
    convertHashButton := childWindow.Add("Button", "w160 Center", "Convert Hash")
    
    childWindow.Show()
    childWindow.GetClientPos(, , &clientWidth, &clientHeight)
    controlWidthPixels := 160
    newXPosition := (clientWidth - controlWidthPixels) // 2

    convertHashButton.Move(newXPosition)

    convertHashButton.OnEvent("Click", (*) => (
        hashText.Value := ConvertHashValue(hashText.Value)
    ))

    childWindow.OnEvent("Escape", (*) => childWindow.Destroy())

    childWindow.Show("AutoSize")
    WinWaitClose("ahk_id " . childWindow.Hwnd)
}

ConvertHashValue(hashValue) {
    LogEngine("Beginning")

    if StrLen(hashValue) = 64 {
        hashValue := EncodeSha256HexToBase(hashValue, 86)
    } else if StrLen(hashValue) = 40 {
        hashValue := DecodeBaseToSha256Hex(hashValue, 86)
    }

    return hashValue
}

ImagesToBase64Converter() {
    selectedDirectory := ""

    childWindow := Gui(, "")
    childWindow.MarginX := 12
    childWindow.MarginY := 12

    selectDirectoryButton := childWindow.Add("Button", "y+10 w160", "Select Directory")
    selectedPathText := childWindow.Add("Edit", "y+10 w160 r4 ReadOnly Center -VScroll -HScroll", "Selected: (none)")
    processImagesButton := childWindow.Add("Button", "w160 Disabled", "Process Images")

    processImagesButton.OnEvent("Click", (*) => (
        selectDirectoryButton.Enabled := false,
        processImagesButton.Enabled := false,
        LogEngine("Beginning"),
        ConvertImagesToBase64ImageLibrary(selectedDirectory),
        childWindow.Destroy(),
        LogEngine("Completed")
    ))

    selectDirectoryButton.OnEvent("Click", (*) => (
        (selectedDirectory := DirSelect() . "\") ? ( selectedPathText.Value := selectedDirectory, selectedPathText.Focus(),
        Send("{End}"),
        selectDirectoryButton.Focus(),
        processImagesButton.Enabled := true): 0
    ))

    childWindow.OnEvent("Escape", (*) => childWindow.Destroy())

    childWindow.Show("AutoSize")
    WinWaitClose("ahk_id " . childWindow.Hwnd)
}

ImageRangeConverter() {
    try {
        DllCall("SetProcessDpiAwarenessContext", "ptr", -4, "int")
    }
    CoordMode("Mouse", "Screen")

    primaryMonitorIndex := MonitorGetPrimary()
    monitorLeft := monitorTop := monitorRight := monitorBottom := 0
    MonitorGet(primaryMonitorIndex, &monitorLeft, &monitorTop, &monitorRight, &monitorBottom)
    if monitorRight <= monitorLeft || monitorBottom <= monitorTop {
        MsgBox "Could not read primary monitor bounds. Exiting."
        ExitApp
    }
    monitorWidth  := monitorRight  - monitorLeft
    monitorHeight := monitorBottom - monitorTop

    overlayWindow := ""

    resetKey := "PrintScreen"
    Loop {
        KeyWait("LButton", "D") ; First click.
        Sleep(16)
        MouseGetPos(&x1, &y1)
        if x1 < monitorLeft || x1 >= monitorRight || y1 < monitorTop || y1 >= monitorBottom {
            KeyWait("LButton")
            continue
        }

        ; Wait for either: Print Screen to cancel first point, or LButton for second point.
        cancelFirstPoint := false
        KeyWait("LButton") ; Ensure the button from the first click is released.

        Loop {
            if GetKeyState(resetKey, "P") {
                KeyWait(resetKey)
                cancelFirstPoint := true
                break
            }
            if GetKeyState("LButton", "P") {
                KeyWait("LButton", "D") ; Second click pressed.
                Sleep(16)
                MouseGetPos(&x2, &y2)
                break
            }
            Sleep(16)
        }

        if cancelFirstPoint {
            ToolTip("First point cleared. Select again.")
            Sleep(800)
            ToolTip()
            continue
        }

        ; ---- validate second point ----
        if x2 < monitorLeft || x2 >= monitorRight || y2 < monitorTop || y2 >= monitorBottom {
            continue
        }

        leftPixel   := (x1 < x2) ? x1 : x2
        rightPixel  := (x1 > x2) ? x1 : x2
        topPixel    := (y1 < y2) ? y1 : y2
        bottomPixel := (y1 > y2) ? y1 : y2
        rectangleWidthPixels  := rightPixel  - leftPixel
        rectangleHeightPixels := bottomPixel - topPixel
        if rectangleWidthPixels <= 0 || rectangleHeightPixels <= 0 {
            continue
        }

        ; Replace previous overlay with a new pixel-perfect one.
        if IsObject(overlayWindow) {
            try {
                overlayWindow.Destroy()
            }
        }
        overlayWindow := Gui("-Caption +ToolWindow +AlwaysOnTop +E0x20 -DPIScale")
        overlayWindow.BackColor := "Lime"
        overlayWindow.Show(Format("x{} y{} w{} h{}", leftPixel, topPixel, rectangleWidthPixels, rectangleHeightPixels))
        WinSetTransparent 64, overlayWindow.Hwnd

        horizontalStartPercent := Round(((leftPixel   - monitorLeft) / monitorWidth)  * 100, 1)
        horizontalEndPercent   := Round(((rightPixel  - monitorLeft) / monitorWidth)  * 100, 1)
        verticalStartPercent   := Round(((topPixel    - monitorTop)  / monitorHeight) * 100, 1)
        verticalEndPercent     := Round(((bottomPixel - monitorTop)  / monitorHeight) * 100, 1)

        resultString := Format("{}-{}, {}-{}",
            (Mod(horizontalStartPercent, 1) = 0) ? Format("{:.0f}", horizontalStartPercent) : Format("{:.1f}", horizontalStartPercent),
            (Mod(horizontalEndPercent,   1) = 0) ? Format("{:.0f}", horizontalEndPercent)   : Format("{:.1f}", horizontalEndPercent),
            (Mod(verticalStartPercent,   1) = 0) ? Format("{:.0f}", verticalStartPercent)   : Format("{:.1f}", verticalStartPercent),
            (Mod(verticalEndPercent,     1) = 0) ? Format("{:.0f}", verticalEndPercent)     : Format("{:.1f}", verticalEndPercent)
        )

        A_Clipboard := resultString
        ToolTip("Copied: " resultString "`nPress " . resetKey . " to select again.`nPress Escape to quit.", leftPixel + 20, topPixel + 20)

        KeyWait(resetKey, "D")
        KeyWait(resetKey)

        ToolTip() ; Clear UI before next loop.
        if IsObject(overlayWindow) {
            try {
                overlayWindow.Destroy()
            }
            overlayWindow := ""
        }
    }
}

Button_Click(guiCtrlObject, *) {
    selectedButton := guiCtrlObject.Name

    global mainWindow

    if IsSet(mainWindow) && mainWindow {
        mainWindow.Destroy()
    }

    switch selectedButton
    {
        case "buttonMain":
            Main()
        case "buttonBaseHashConverter":
            BaseHashConverter()
        case "buttonImagesToBase64Converter":
            ImagesToBase64Converter()
        case "buttonImageRangeConverter":
            ImageRangeConverter()
    }

    switch selectedButton
    {
        case "buttonMain":
        case "buttonImageRangeConverter":
        default:
            ExitApp()
    }    
}

Launcher() {
    global mainWindow

    mainWindow := Gui(, "Framework Demonstration")
    mainWindow.MarginX := 6
    mainWindow.MarginY := 6

    buttonMainScript              := mainWindow.Add("Button", "vbuttonMain w360 h32", "Main Script")
    buttonBaseHashConverter       := mainWindow.Add("Button", "vbuttonBaseHashConverter w360 h32", "Base Hash Converter")
    buttonImagesToBase64Converter := mainWindow.Add("Button", "vbuttonImagesToBase64Converter w360 h32", "Images to Base64 Converter")
    buttonImageRangeConverter     := mainWindow.Add("Button", "vbuttonImageRangeConverter w360 h32", "Image Range Converter")

    buttonMainScript.OnEvent("Click", Button_Click)
    buttonBaseHashConverter.OnEvent("Click", Button_Click)
    buttonImagesToBase64Converter.OnEvent("Click", Button_Click)
    buttonImageRangeConverter.OnEvent("Click", Button_Click)

    mainWindow.OnEvent("Close", (*) => ExitApp())
    mainWindow.Show("AutoSize")
}

Launcher()