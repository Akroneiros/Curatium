#Requires AutoHotkey v2.0
#Include Application Library.ahk
#Include Base Library.ahk
#Include File Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

PreventSystemGoingIdleUntilTimestamp(utcTimestamp) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    utcTimestampLabel := "N/A"
    if ValidateDataUsingConstraint(utcTimestamp, "String", "ISO Date Time") = "" {
        utcTimestampOverlay := utcTimestamp

        utcTimestampOverlayLength := StrLen(utcTimestampOverlay)
        if utcTimestampOverlayLength = 26 {
            utcTimestampOverlay := SubStr(utcTimestampOverlay, 1, -7)
        } else if utcTimestampOverlayLength = 23 {
            utcTimestampOverlay := SubStr(utcTimestampOverlay, 1, -4)
        }

        utcTimestampLabel := Format("{} @ {} UTC", ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"][FormatTime(StrReplace(StrReplace(StrReplace(utcTimestampOverlay, "-"), " "), ":"), "WDay") + 0], SubStr(utcTimestampOverlay, 12))
    }

    overlayValue := Format("Prevent System Going Idle Until Date Time ({})", utcTimestampLabel)
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("utcTimestamp As String [Constraint: ISO Date Time]", methodName, A_LineFile, A_LineNumber + 5, Map(
            "Milliseconds Between Mouse Moves", Map("Default", 4000, "Floor", 2000, "Ceiling", 598000),
            "Milliseconds Between Function Key Presses", Map("Default", 60000, "Floor", 2000, "Ceiling", 898000)
        ))
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [utcTimestamp], overlayValue)

    settings := methodRegistry[methodName]["Settings"]

    millisecondsBetweenMouseMoves         := settings["Milliseconds Between Mouse Moves"]["Value"]
    millisecondsBetweenFunctionKeyPresses := settings["Milliseconds Between Function Key Presses"]["Value"]

    utcTimestampLength := StrLen(utcTimestamp)
    if utcTimestampLength = 26 {
        utcTimestamp := SubStr(utcTimestamp, 1, -7)
    } else if utcTimestampLength = 23 {
        utcTimestamp := SubStr(utcTimestamp, 1, -4)
    }

    utcTimestamp                  := RegExReplace(utcTimestamp, "\D")
    lastFunctionKeyPressTickCount := DllCall("Kernel32\GetTickCount64", "UInt64")

    previousMouseCoordinateMode := A_CoordModeMouse
    CoordMode("Mouse", "Screen")

    while DateDiff(utcTimestamp, A_NowUTC, "Seconds") > 1 {
        secondsRemainingBeforeFinalSecond := DateDiff(utcTimestamp, A_NowUTC, "Seconds") - 1
        millisecondsUntilNextMouseMove    := millisecondsBetweenMouseMoves

        if secondsRemainingBeforeFinalSecond * 1000 < millisecondsUntilNextMouseMove {
            millisecondsUntilNextMouseMove := secondsRemainingBeforeFinalSecond * 1000
        }

        if millisecondsUntilNextMouseMove > 0 {
            Sleep(millisecondsUntilNextMouseMove)
        }

        if DateDiff(utcTimestamp, A_NowUTC, "Seconds") <= 1 {
            break
        }

        screenWidth  := A_ScreenWidth
        screenHeight := A_ScreenHeight

        direction := Random(1, 4)
        deltaX    := 0
        deltaY    := 0

        MouseGetPos(&mouseX, &mouseY)
        if direction = 1 && mouseX >= screenWidth - 1 {
            direction := 2
        } else if direction = 2 && mouseX <= 0 {
            direction := 1
        } else if direction = 3 && mouseY >= screenHeight - 1 {
            direction := 4
        } else if direction = 4 && mouseY <= 0 {
            direction := 3
        }

        if direction = 1 {
            deltaX := 1
        } else if direction = 2 {
            deltaX := -1
        } else if direction = 3 {
            deltaY := 1
        } else {
            deltaY := -1
        }

        MouseMove(deltaX, deltaY, 0, "R")
        Sleep(Random(64, 128))
        MouseMove(-deltaX, -deltaY, 0, "R")

        if DllCall("Kernel32\GetTickCount64", "UInt64") - lastFunctionKeyPressTickCount >= millisecondsBetweenFunctionKeyPresses {
            functionKeyNumber := Random(13, 24)
            Send("{F" . functionKeyNumber . "}")
            lastFunctionKeyPressTickCount := DllCall("Kernel32\GetTickCount64", "UInt64")
        }
    }

    CoordMode("Mouse", previousMouseCoordinateMode)

    while A_NowUTC < utcTimestamp {
    }

    LogConclusion("Completed", logConclusionData)
}

SetDirectoryTimestamp(directoryPath, utcTimestamp, timeType) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}"', "Accessed", "Created", "Modified")
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("directoryPath As String [Constraint: Directory], utcTimestamp As String [Constraint: ISO Date Time], timeType As String [Whitelist: " . timeTypeWhitelist . "]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [directoryPath, utcTimestamp, timeType], "Set Directory Timestamp")

    utcTimestampLength := StrLen(utcTimestamp)
    if utcTimestampLength = 26 {
        utcTimestamp := SubStr(utcTimestamp, 1, -7)
    } else if utcTimestampLength = 23 {
        utcTimestamp := SubStr(utcTimestamp, 1, -4)
    }

    currentDirectoryTimestamp := GetDirectoryTimeAsTimestamp(directoryPath, timeType)
    if currentDirectoryTimestamp = utcTimestamp {
        LogConclusion("Skipped", logConclusionData)
        return
    }

    utcTimestamp  := RegExReplace(utcTimestamp, "\D")
    directoryPath := RTrim(directoryPath, "\")

    utcSystemTimeBuffer := Buffer(16, 0)
    NumPut("UShort", SubStr(utcTimestamp, 1, 4), utcSystemTimeBuffer, 0)
    NumPut("UShort", SubStr(utcTimestamp, 5, 2), utcSystemTimeBuffer, 2)
    NumPut("UShort", 0,                          utcSystemTimeBuffer, 4)
    NumPut("UShort", SubStr(utcTimestamp, 7, 2), utcSystemTimeBuffer, 6)
    NumPut("UShort", SubStr(utcTimestamp, 9, 2), utcSystemTimeBuffer, 8)
    NumPut("UShort", SubStr(utcTimestamp, 11, 2), utcSystemTimeBuffer, 10)
    NumPut("UShort", SubStr(utcTimestamp, 13, 2), utcSystemTimeBuffer, 12)
    NumPut("UShort", 0,                           utcSystemTimeBuffer, 14)

    utcFileTimeBuffer := Buffer(8, 0)
    if !DllCall("Kernel32\SystemTimeToFileTime", "Ptr", utcSystemTimeBuffer, "Ptr", utcFileTimeBuffer, "Int") {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to convert system time to file time format. [Kernel32\SystemTimeToFileTime" . ", System Error Code: " . A_LastError . "]")
    }

    accessMode := 0x100
    shareMode  := 0x7
    flags      := 0x80 | 0x02000000

    directoryHandle := DllCall("Kernel32\CreateFileW", "WStr", directoryPath, "UInt", accessMode, "UInt", shareMode, "Ptr", 0, "UInt", 3, "UInt", flags, "Ptr", 0, "Ptr")
    if directoryHandle = -1 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to open directory. [Kernel32\CreateFileW" . ", System Error Code: " . A_LastError . "]")
    }

    accessTimePointer   := 0
    creationTimePointer := 0
    modifiedTimePointer := 0
    switch timeType {
        case "Accessed": accessTimePointer   := utcFileTimeBuffer.Ptr
        case "Created":  creationTimePointer := utcFileTimeBuffer.Ptr
        case "Modified": modifiedTimePointer := utcFileTimeBuffer.Ptr
    }

    dateAndTimeSetToDirectorySuccessfully := DllCall("Kernel32\SetFileTime", "Ptr", directoryHandle, "Ptr", creationTimePointer, "Ptr", accessTimePointer, "Ptr", modifiedTimePointer, "Int")
    if !dateAndTimeSetToDirectorySuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "SetFileTime failed")
    }

    DllCall("Kernel32\CloseHandle", "Ptr", directoryHandle)

    LogConclusion("Completed", logConclusionData)
}

SetFileTimestamp(filePath, utcTimestamp, timeType) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}"', "Accessed", "Created", "Modified")
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("filePath As String [Constraint: Path], utcTimestamp As String [Constraint: ISO Date Time], timeType As String [Whitelist: " . timeTypeWhitelist . "]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [filePath, utcTimestamp, timeType], "Set File Timestamp")

    utcTimestampLength := StrLen(utcTimestamp)
    if utcTimestampLength = 26 {
        utcTimestamp := SubStr(utcTimestamp, 1, -7)
    } else if utcTimestampLength = 23 {
        utcTimestamp := SubStr(utcTimestamp, 1, -4)
    }

    currentFileTimestamp := GetFileTimeAsTimestamp(filePath, timeType)
    if currentFileTimestamp = utcTimestamp {
        LogConclusion("Skipped", logConclusionData)
        return
    }

    utcTimestamp := RegExReplace(utcTimestamp, "\D")

    utcSystemTimeBuffer := Buffer(16, 0)
    NumPut("UShort", SubStr(utcTimestamp, 1, 4), utcSystemTimeBuffer, 0)
    NumPut("UShort", SubStr(utcTimestamp, 5, 2), utcSystemTimeBuffer, 2)
    NumPut("UShort", 0,                          utcSystemTimeBuffer, 4)
    NumPut("UShort", SubStr(utcTimestamp, 7, 2), utcSystemTimeBuffer, 6)
    NumPut("UShort", SubStr(utcTimestamp, 9, 2), utcSystemTimeBuffer, 8)
    NumPut("UShort", SubStr(utcTimestamp, 11, 2), utcSystemTimeBuffer, 10)
    NumPut("UShort", SubStr(utcTimestamp, 13, 2), utcSystemTimeBuffer, 12)
    NumPut("UShort", 0,                           utcSystemTimeBuffer, 14)

    utcFileTimeBuffer := Buffer(8, 0)

    convertedSystemTimeToFileTimeSuccessfully := DllCall("Kernel32\SystemTimeToFileTime", "Ptr", utcSystemTimeBuffer, "Ptr", utcFileTimeBuffer, "Int")
    if !convertedSystemTimeToFileTimeSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to convert system time to file time format. [Kernel32\SystemTimeToFileTime" . ", System Error Code: " . A_LastError . "]")
    }

    accessMode := 0x100
    shareMode  := 0x7
    flags      := 0x80 | 0x02000000

    fileHandle := DllCall("Kernel32\CreateFileW", "WStr", filePath, "UInt", accessMode, "UInt", shareMode, "Ptr", 0, "UInt", 3, "UInt", flags, "Ptr", 0, "Ptr")
    if fileHandle = -1 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to open file. [Kernel32\CreateFileW" . ", System Error Code: " . A_LastError . "]")
    }

    accessTimePointer   := 0
    creationTimePointer := 0
    modifiedTimePointer := 0

    switch timeType {
        case "Accessed": accessTimePointer   := utcFileTimeBuffer.Ptr
        case "Created":  creationTimePointer := utcFileTimeBuffer.Ptr
        case "Modified": modifiedTimePointer := utcFileTimeBuffer.Ptr
    }

    dateAndTimeSetToFileSuccessfully := DllCall("Kernel32\SetFileTime", "Ptr", fileHandle, "Ptr", creationTimePointer, "Ptr", accessTimePointer, "Ptr", modifiedTimePointer, "Int")
    if !dateAndTimeSetToFileSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "SetFileTime failed")
    }

    DllCall("Kernel32\CloseHandle", "Ptr", fileHandle)

    LogConclusion("Completed", logConclusionData)
}

WaitUntilFileIsModifiedToday(filePath) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("filePath As String [Constraint: Valid Path]", methodName, A_LineFile, A_LineNumber + 7, Map(
            "Milliseconds Between Mouse Moves", Map("Default", 4000, "Floor", 2000, "Ceiling", 598000),
            "Milliseconds Between File Checks", Map("Default", 2000, "Floor", 1000, "Ceiling", 10000),
            "Milliseconds Between Function Key Presses", Map("Default", 60000, "Floor", 2000, "Ceiling", 898000),
            "Max Wait Minutes", Map("Default", 360, "Floor", 1, "Ceiling", 1438)
        ))
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [filePath], "Wait Until File is Modified Today: " . filePath)

    settings := methodRegistry[methodName]["Settings"]

    millisecondsBetweenMouseMoves         := settings["Milliseconds Between Mouse Moves"]["Value"]
    millisecondsBetweenFileChecks         := settings["Milliseconds Between File Checks"]["Value"]
    millisecondsBetweenFunctionKeyPresses := settings["Milliseconds Between Function Key Presses"]["Value"]
    maxWaitMinutes                        := settings["Max Wait Minutes"]["Value"]

    dateOfToday        := FormatTime(A_Now, "yyyy-MM-dd")
    maxLoops           := (maxWaitMinutes * 60000) // millisecondsBetweenFileChecks
    timeSinceLastMouse := 0
    lastFunctionKeyPressTickCount := DllCall("Kernel32\GetTickCount64", "UInt64")

    previousMouseCoordinateMode := A_CoordModeMouse
    CoordMode("Mouse", "Screen")

    Loop maxLoops {
        if FileExist(filePath) {
            fileModifiedDate := FileGetTime(filePath, "M")
            fileModifiedDate := FormatTime(fileModifiedDate, "yyyy-MM-dd")

            if dateOfToday = fileModifiedDate {
                break
            }
        }

        Sleep(millisecondsBetweenFileChecks)
        timeSinceLastMouse += millisecondsBetweenFileChecks

        if timeSinceLastMouse >= millisecondsBetweenMouseMoves {
            screenWidth  := A_ScreenWidth
            screenHeight := A_ScreenHeight

            direction := Random(1, 4)
            deltaX    := 0
            deltaY    := 0

            MouseGetPos(&mouseX, &mouseY)
            if direction = 1 && mouseX >= screenWidth - 1 {
                direction := 2
            } else if direction = 2 && mouseX <= 0 {
                direction := 1
            } else if direction = 3 && mouseY >= screenHeight - 1 {
                direction := 4
            } else if direction = 4 && mouseY <= 0 {
                direction := 3
            }

            if direction = 1 {
                deltaX := 1
            } else if direction = 2 {
                deltaX := -1
            } else if direction = 3 {
                deltaY := 1
            } else {
                deltaY := -1
            }

            MouseMove(deltaX, deltaY, 0, "R")
            Sleep(Random(64, 128))
            MouseMove(-deltaX, -deltaY, 0, "R")

            timeSinceLastMouse := 0
        }

        if DllCall("Kernel32\GetTickCount64", "UInt64") - lastFunctionKeyPressTickCount >= millisecondsBetweenFunctionKeyPresses {
            functionKeyNumber := Random(13, 24)
            Send("{F" . functionKeyNumber . "}")
            lastFunctionKeyPressTickCount := DllCall("Kernel32\GetTickCount64", "UInt64")
        }
    }

    CoordMode("Mouse", previousMouseCoordinateMode)

    LogConclusion("Completed", logConclusionData)
}

; **************************** ;
; Core Methods                 ;
; **************************** ;

TelemetryTimestamp(durationInMilliseconds) {
    result := Map(
        "Duration in Milliseconds",  durationInMilliseconds
    )

    kernel32ModuleHandle   := DllCall("GetModuleHandle", "Str", "Kernel32", "Ptr")
    preciseFunctionAddress := DllCall("GetProcAddress", "Ptr", kernel32ModuleHandle, "AStr", "GetSystemTimePreciseAsFileTime", "Ptr")

    originalAutoHotkeyThreadHandle   := DllCall("GetCurrentThread", "Ptr")
    originalAutoHotkeyThreadPriority := DllCall("GetThreadPriority", "Ptr", originalAutoHotkeyThreadHandle, "Int")

    if originalAutoHotkeyThreadPriority = 0 {
        DllCall("SetThreadPriority", "Ptr", originalAutoHotkeyThreadHandle, "Int", 2) ; Change to Highest.
    }

    tickCountPre     := unset
    timingBuffer     := Buffer(24, 0)
    systemTimeBuffer := Buffer(16, 0)
    qpcPrePointer    := timingBuffer.Ptr
    fileTimePointer  := timingBuffer.Ptr + 8
    qpcPostPointer   := timingBuffer.Ptr + 16
    qpcPreValue      := unset
    qpcLoopStopValue := unset
    pairedReadings   := []

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    qpcPreValue      := NumGet(qpcPrePointer, "Int64")
    qpcLoopStopValue := qpcPreValue + (system["Environment"]["QPC Frequency"] * durationInMilliseconds // 1000)

    if preciseFunctionAddress {
        while qpcPreValue < qpcLoopStopValue {
            tickCountPre := DllCall("Kernel32\GetTickCount64", "UInt64")
            DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
            DllCall("Kernel32\GetSystemTimePreciseAsFileTime", "Ptr", fileTimePointer)
            DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")
            pairedReadings.Push([tickCountPre, qpcPreValue := NumGet(qpcPrePointer, "Int64"), NumGet(fileTimePointer, "UInt64"), NumGet(qpcPostPointer, "Int64"), DllCall("Kernel32\GetTickCount64", "UInt64")])
        }
    } else {
        while qpcPreValue < qpcLoopStopValue {
            tickCountPre := DllCall("Kernel32\GetTickCount64", "UInt64")
            DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
            DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", fileTimePointer)
            DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")
            pairedReadings.Push([tickCountPre, qpcPreValue := NumGet(qpcPrePointer, "Int64"), NumGet(fileTimePointer, "UInt64"), NumGet(qpcPostPointer, "Int64"), DllCall("Kernel32\GetTickCount64", "UInt64")])
        }
    }

    if originalAutoHotkeyThreadPriority = 0 {
        DllCall("SetThreadPriority", "Ptr", originalAutoHotkeyThreadHandle, "Int", 0) ; Change to Normal.
    }

    result["Number of Readings"]   := pairedReadings.Length
    result["QPC Earliest Reading"] := pairedReadings[1][2]
    result["QPC Latest Reading"]   := pairedReadings[-1][4]

    structuredPairedReadings := []
    for pairedReading in pairedReadings {
        structuredPairReading := Map(
            "Tick Count Pre",                 pairedReading[1],
            "Query Performance Counter Pre",  pairedReading[2],
            "Timestamp",                      pairedReading[3],
            "Query Performance Counter Post", pairedReading[4],
            "Tick Count Post",                pairedReading[5]
        )

        structuredPairReading["Query Performance Counter Delta"] := structuredPairReading["Query Performance Counter Post"] - structuredPairReading["Query Performance Counter Pre"]
        structuredPairReading["Tick Count Delta"]                := structuredPairReading["Tick Count Post"] - structuredPairReading["Tick Count Pre"]

        structuredPairedReadings.Push(structuredPairReading)
    }

    bestPairedReading := unset
    for structuredPairedReading in structuredPairedReadings {
        if structuredPairedReading["Tick Count Delta"] = 0 {
            continue
        }

        if !IsSet(bestPairedReading) {
            bestPairedReading := structuredPairedReading
        } else {
            if bestPairedReading["Query Performance Counter Delta"] > structuredPairedReading["Query Performance Counter Delta"] {
                bestPairedReading := structuredPairedReading
            }
        }
    }

    if !IsSet(bestPairedReading) {
        bestPairedReading := structuredPairedReadings[1]

        for structuredPairedReading in structuredPairedReadings {
            if bestPairedReading["Query Performance Counter Delta"] > structuredPairedReading["Query Performance Counter Delta"] {
                bestPairedReading := structuredPairedReading
            }
        }
    }

    NumPut("UInt64", bestPairedReading["Timestamp"], fileTimePointer)
    DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimePointer, "Ptr", systemTimeBuffer.Ptr, "Int")

    year        := NumGet(systemTimeBuffer,  0, "UShort")
    month       := NumGet(systemTimeBuffer,  2, "UShort")
    day         := NumGet(systemTimeBuffer,  6, "UShort")
    hour        := NumGet(systemTimeBuffer,  8, "UShort")
    minute      := NumGet(systemTimeBuffer, 10, "UShort")
    second      := NumGet(systemTimeBuffer, 12, "UShort")
    millisecond := NumGet(systemTimeBuffer, 14, "UShort")

    result["Tick Count"]              := bestPairedReading["Tick Count Post"]
    result["Tick Count Pre"]          := bestPairedReading["Tick Count Pre"]
    result["QPC Pre Timestamp"]       := bestPairedReading["Query Performance Counter Pre"]
    result["QPC Post Timestamp"]      := bestPairedReading["Query Performance Counter Post"]
    result["QPC Delta Timestamp"]     := bestPairedReading["Query Performance Counter Delta"]
    result["QPC Midpoint Timestamp"]  := result["QPC Pre Timestamp"] + (result["QPC Delta Timestamp"] // 2)
    result["UTC Timestamp File Time"] := bestPairedReading["Timestamp"]
    result["UTC Timestamp Integer"]   := Format("{:04}{:02}{:02}{:02}{:02}{:02}{:03}", year, month, day, hour, minute, second, millisecond) + 0
    if preciseFunctionAddress {
        fileTimeTicks     := NumGet(fileTimePointer, "Int64")
        ticksWithinSecond := Mod(fileTimeTicks, 10000000)
        microsecond       := ticksWithinSecond // 10

        result["UTC Timestamp Precise"] := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:06}", year, month, day, hour, minute, second, microsecond)
    } else {
        result["UTC Timestamp Precise"] := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}", year, month, day, hour, minute, second, millisecond)
    }

    return result
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

AdjustTimestampByTimeAmount(isoTimestamp, timeUnit, time) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static timeUnitWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Days", "Hours", "Minutes", "Seconds", "Milliseconds", "Microseconds")
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("isoTimestamp As String [Constraint: ISO Date Time], timeUnit As String [Whitelist: " . timeUnitWhitelist . "], time as Integer", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [isoTimestamp, timeUnit, time])

    subseconds         := ""
    isoTimestampLength := StrLen(isoTimestamp)

    if isoTimestampLength = 26 {
        subseconds   := SubStr(isoTimestamp, -7)
        isoTimestamp := SubStr(isoTimestamp, 1, -7)
    } else if isoTimestampLength = 23 {
        subseconds   := SubStr(isoTimestamp, -4)
        isoTimestamp := SubStr(isoTimestamp, 1, -4)
    }

    if timeUnit = "Milliseconds" || timeUnit = "Microseconds" {
        inputPrecision := 0
        if StrLen(subseconds) = 7 {
            inputPrecision := 6
        } else if StrLen(subseconds) = 4 {
            inputPrecision := 3
        }

        adjustmentPrecision := 3
        if timeUnit = "Microseconds" {
            adjustmentPrecision := 6
        }
        precision := Max(inputPrecision, adjustmentPrecision)

        if timeUnit = "Milliseconds" {
            wholeSecondsToAdd := time // 1000
            remainderUnits := Mod(time, 1000) * 1000
        } else {
            wholeSecondsToAdd := time // 1000000
            remainderUnits := Mod(time, 1000000)
        }

        if remainderUnits < 0 {
            remainderUnits += 1000000
            wholeSecondsToAdd -= 1
        }

        currentUnits := 0
        if subseconds != "" {
            currentUnits := SubStr(subseconds, 2) + 0
            if inputPrecision = 3 {
                currentUnits *= 1000
            }
        }

        totalUnits := currentUnits + remainderUnits
        carrySeconds := totalUnits // 1000000
        remainingUnits := Mod(totalUnits, 1000000)
        if remainingUnits < 0 {
            remainingUnits += 1000000
            carrySeconds -= 1
        }

        wholeSecondsToAdd += carrySeconds

        subseconds := "." . SubStr(Format("{:06}", remainingUnits), 1, precision)

        time     := wholeSecondsToAdd
        timeUnit := "Seconds"
    }

    adjustedIsoDateTime := DateAdd(RegExReplace(isoTimestamp, "\D"), time, timeUnit)

    if adjustedIsoDateTime = "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Resulting timestamp is outside the supported range of 1601-01-01 00:00:00 through 30827-12-31 23:59:59.")
    }

    if (adjustedIsoDateTime + 0) > 99991231235959 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Resulting timestamp is higher than 9999-12-31 23:59:59.")
    }

    adjustedIsoDateTime := FormatTime(adjustedIsoDateTime, "yyyy-MM-dd HH:mm:ss") . subseconds

    return adjustedIsoDateTime
}

ConvertIntegerToUtcTimestamp(integerValue) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("integerValue As Integer [Constraint: Integer UTC Timestamp]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [integerValue])

    integerValueText       := integerValue . ""
    integerValueTextLength := StrLen(integerValueText)

    year        := SubStr(integerValueText,  1, 4)
    month       := SubStr(integerValueText,  5, 2)
    day         := SubStr(integerValueText,  7, 2)
    hour        := SubStr(integerValueText,  9, 2)
    minute      := SubStr(integerValueText, 11, 2)
    second      := SubStr(integerValueText, 13, 2)

    utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)
    if integerValueTextLength = 17 {
        utcTimestamp .= "." . SubStr(integerValueText, 15)
    }

    validation := ValidateDataUsingConstraint(utcTimestamp, "String", "ISO Date Time")
    if validation != "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Integer UTC Timestamp converted to " . validation)
    }

    return utcTimestamp
}

ConvertLocalTimestampToUtcTimestampUsingTimeZoneKey(localTimestamp, timeZoneKeyName) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("localTimestamp As String [Constraint: ISO Date Time], timeZoneKeyName As String", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [localTimestamp, timeZoneKeyName])

    localTimestampLength := StrLen(localTimestamp)
    if localTimestampLength = 26 {
        localTimestamp := SubStr(localTimestamp, 1, -3)
    }

    localTimestamp       := RegExReplace(localTimestamp, "\D")
    localTimestampLength := StrLen(localTimestamp)

    year   := SubStr(localTimestamp,  1, 4) + 0
    month  := SubStr(localTimestamp,  5, 2) + 0
    day    := SubStr(localTimestamp,  7, 2) + 0
    hour   := SubStr(localTimestamp,  9, 2) + 0
    minute := SubStr(localTimestamp, 11, 2) + 0
    second := SubStr(localTimestamp, 13, 2) + 0

    milliseconds := unset
    if localTimestampLength = 17 {
        milliseconds := SubStr(localTimestamp, 15) + 0
    } else {
        milliseconds := 0
    }

    localSystemTimeBuffer := Buffer(16, 0)
    NumPut("UShort", year, localSystemTimeBuffer, 0)
    NumPut("UShort", month, localSystemTimeBuffer, 2)
    NumPut("UShort", day, localSystemTimeBuffer, 6)
    NumPut("UShort", hour, localSystemTimeBuffer, 8)
    NumPut("UShort", minute, localSystemTimeBuffer, 10)
    NumPut("UShort", second, localSystemTimeBuffer, 12)
    NumPut("UShort", milliseconds, localSystemTimeBuffer, 14)

    static DYNAMIC_TIME_ZONE_INFORMATION_SIZE    := 432
    static DYNAMIC_DAYLIGHT_TIME_DISABLED_OFFSET := 428
    static TIME_ZONE_KEY_NAME_OFFSET             := 172

    static timezoneCache := Map()
    if !timezoneCache.Has(timeZoneKeyName) {
        preparedBuffer := Buffer(DYNAMIC_TIME_ZONE_INFORMATION_SIZE, 0)
        StrPut(timeZoneKeyName, preparedBuffer.Ptr + TIME_ZONE_KEY_NAME_OFFSET, 128, "UTF-16")
        NumPut("UChar", 0, preparedBuffer, DYNAMIC_DAYLIGHT_TIME_DISABLED_OFFSET)
        timezoneCache[timeZoneKeyName] := preparedBuffer
    }
    dynamicTimeZoneInformationBuffer := timezoneCache[timeZoneKeyName]

    static utcSystemTimeBuffer := Buffer(16, 0)

    convertedLocalTimeToUtcSuccessfully := DllCall("Kernel32\TzSpecificLocalTimeToSystemTimeEx", "Ptr", dynamicTimeZoneInformationBuffer.Ptr, "Ptr", localSystemTimeBuffer.Ptr, "Ptr", utcSystemTimeBuffer.Ptr, "Int")
    if !convertedLocalTimeToUtcSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to convert local timestamp to UTC. [Kernel32\TzSpecificLocalTimeToSystemTimeEx, Error: " A_LastError "]")
    }

    year         := NumGet(utcSystemTimeBuffer, 0,  "UShort")
    month        := NumGet(utcSystemTimeBuffer, 2,  "UShort")
    day          := NumGet(utcSystemTimeBuffer, 6,  "UShort")
    hour         := NumGet(utcSystemTimeBuffer, 8,  "UShort")
    minute       := NumGet(utcSystemTimeBuffer, 10, "UShort")
    second       := NumGet(utcSystemTimeBuffer, 12, "UShort")
    milliseconds := NumGet(utcSystemTimeBuffer, 14, "UShort")

    utcTimestamp := unset
    if localTimestampLength = 17 {
        utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}", year, month, day, hour, minute, second, milliseconds)
    } else {
        utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)
    }

    return utcTimestamp
}

ConvertUnixTimeToUtcTimestamp(unixSeconds) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("unixSeconds As Integer", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [unixSeconds])

    if unixSeconds < -11644473600 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Unix Seconds predates 1601-01-01 UTC and cannot be represented as FILETIME: " . unixSeconds)
    }

    fileTimeTicks := (unixSeconds + 11644473600) * 10000000

    static fileTimeBuffer := Buffer(8, 0)
    NumPut("UInt64", fileTimeTicks, fileTimeBuffer)

    static systemTimeBuffer := Buffer(16, 0)
    convertedSuccessfully := DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")
    if !convertedSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to convert a file time to system time format. [Kernel32\FileTimeToSystemTime" . ", System Error Code: " . A_LastError . "]")
    }

    year   := NumGet(systemTimeBuffer,  0, "UShort")
    month  := NumGet(systemTimeBuffer,  2, "UShort")
    day    := NumGet(systemTimeBuffer,  6, "UShort")
    hour   := NumGet(systemTimeBuffer,  8, "UShort")
    minute := NumGet(systemTimeBuffer, 10, "UShort")
    second := NumGet(systemTimeBuffer, 12, "UShort")

    utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)

    return utcTimestamp
}

ConvertUtcTimestampToInteger(utcTimestamp) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("utcTimestamp As String [Constraint: ISO Date Time]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [utcTimestamp])

    utcTimestampLength := StrLen(utcTimestamp)
    if utcTimestampLength = 26 {
        utcTimestamp := SubStr(utcTimestamp, 1, -3)
    }

    utcTimestampInteger := RegExReplace(utcTimestamp, "\D") + 0

    return utcTimestampInteger
}

ConvertUtcTimestampToLocalTimestampUsingTimeZoneKey(utcTimestamp, timeZoneKeyName) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("utcTimestamp As String [Constraint: ISO Date Time], timeZoneKeyName As String", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [utcTimestamp, timeZoneKeyName])

    utcTimestampLength := StrLen(utcTimestamp)
    if utcTimestampLength = 26 {
        utcTimestamp := SubStr(utcTimestamp, 1, -3)
    }

    utcTimestamp       := RegExReplace(utcTimestamp, "\D")
    utcTimestampLength := StrLen(utcTimestamp)

    year   := SubStr(utcTimestamp,  1, 4) + 0
    month  := SubStr(utcTimestamp,  5, 2) + 0
    day    := SubStr(utcTimestamp,  7, 2) + 0
    hour   := SubStr(utcTimestamp,  9, 2) + 0
    minute := SubStr(utcTimestamp, 11, 2) + 0
    second := SubStr(utcTimestamp, 13, 2) + 0

    milliseconds := unset
    if utcTimestampLength = 17 {
        milliseconds := SubStr(utcTimestamp, 15) + 0
    } else {
        milliseconds := 0
    }

    utcSystemTimeBuffer := Buffer(16, 0)
    NumPut("UShort", year, utcSystemTimeBuffer, 0)
    NumPut("UShort", month, utcSystemTimeBuffer, 2)
    NumPut("UShort", day, utcSystemTimeBuffer, 6)
    NumPut("UShort", hour, utcSystemTimeBuffer, 8)
    NumPut("UShort", minute, utcSystemTimeBuffer, 10)
    NumPut("UShort", second, utcSystemTimeBuffer, 12)
    NumPut("UShort", milliseconds, utcSystemTimeBuffer, 14)

    static DYNAMIC_TIME_ZONE_INFORMATION_SIZE        := 432
    static DYNAMIC_DAYLIGHT_TIME_DISABLED_OFFSET     := 428
    static TIME_ZONE_KEY_NAME_OFFSET                 := 172

    static timezoneCache := Map()
    if !timezoneCache.Has(timeZoneKeyName) {
        preparedBuffer := Buffer(DYNAMIC_TIME_ZONE_INFORMATION_SIZE, 0)

        StrPut(timeZoneKeyName, preparedBuffer.Ptr + TIME_ZONE_KEY_NAME_OFFSET, 128, "UTF-16")
        NumPut("UChar", 0, preparedBuffer, DYNAMIC_DAYLIGHT_TIME_DISABLED_OFFSET)

        timezoneCache[timeZoneKeyName] := preparedBuffer
    }

    dynamicTimeZoneInformationBuffer := timezoneCache[timeZoneKeyName]

    static localSystemTimeBuffer := Buffer(16, 0)

    convertedUtcToLocalSuccessfully := DllCall("Kernel32\SystemTimeToTzSpecificLocalTimeEx", "Ptr", dynamicTimeZoneInformationBuffer.Ptr, "Ptr", utcSystemTimeBuffer.Ptr, "Ptr", localSystemTimeBuffer.Ptr, "Int")
    if !convertedUtcToLocalSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to convert UTC timestamp to local timestamp. [Kernel32\SystemTimeToTzSpecificLocalTimeEx, Error: " A_LastError "]")
    }

    year         := NumGet(localSystemTimeBuffer, 0,  "UShort")
    month        := NumGet(localSystemTimeBuffer, 2,  "UShort")
    day          := NumGet(localSystemTimeBuffer, 6,  "UShort")
    hour         := NumGet(localSystemTimeBuffer, 8,  "UShort")
    minute       := NumGet(localSystemTimeBuffer, 10, "UShort")
    second       := NumGet(localSystemTimeBuffer, 12, "UShort")
    milliseconds := NumGet(localSystemTimeBuffer, 14, "UShort")

    localTimestamp := unset
    if utcTimestampLength = 17 {
        localTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}", year, month, day, hour, minute, second, milliseconds)
    } else {
        localTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)
    }

    return localTimestamp
}

ExtractTrailingDateAsIso(inputValue, dateOrder) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static dateOrderWhitelist := Format('"{1}", "{2}", "{3}"', "Day-Month-Year", "Month-Day-Year", "Year-Month-Day")
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("inputValue As String, dateOrder As String [Whitelist: " . dateOrderWhitelist . "]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [inputValue, dateOrder])

    isoDate         := ""
    lastMatch       := unset

    currentPosition := 1
    pattern         := unset

    switch dateOrder {
        case "Day-Month-Year":
            pattern := "(?<!\d)(0[1-9]|[12]\d|3[01])([-./ ])(0[1-9]|1[0-2])\2((?:19|20)\d{2})"
        case "Month-Day-Year":
            pattern := "(?<!\d)(0[1-9]|1[0-2])([-./ ])(0[1-9]|[12]\d|3[01])\2((?:19|20)\d{2})"
        case "Year-Month-Day":
            pattern := "(?<!\d)((?:19|20)\d{2})([-./ ])(0[1-9]|1[0-2])\2(0[1-9]|[12]\d|3[01])"
    }

    while currentPosition := RegExMatch(inputValue, pattern, &matchObject, currentPosition) {
        lastMatch := Map()
        for index, value in matchObject {
            lastMatch[index] := value
        }
        currentPosition += StrLen(matchObject[0])
    }

    if IsSet(lastMatch) {
        switch dateOrder {
            case "Day-Month-Year":
                isoDate := lastMatch[4] . "-" . lastMatch[3] . "-" . lastMatch[1]
            case "Month-Day-Year":
                isoDate := lastMatch[4] . "-" . lastMatch[1] . "-" . lastMatch[3]
            case "Year-Month-Day":
                isoDate := lastMatch[1] . "-" . lastMatch[3] . "-" . lastMatch[4]
        }
    } else {
        switch dateOrder {
            case "Day-Month-Year":
                pattern := "(?<!\d)(0[1-9]|[12]\d|3[01])(0[1-9]|1[0-2])((?:19|20)\d{2})"
            case "Month-Day-Year":
                pattern := "(?<!\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])((?:19|20)\d{2})"
            case "Year-Month-Day":
                pattern := "(?<!\d)((?:19|20)\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])"
        }

        currentPosition := 1
        while currentPosition := RegExMatch(inputValue, pattern, &matchObject, currentPosition) {
            lastMatch := Map()
            for index, value in matchObject {
                lastMatch[index] := value
            }
            currentPosition += StrLen(matchObject[0])
        }

        if IsSet(lastMatch) {
            switch dateOrder {
                case "Day-Month-Year":
                    isoDate := lastMatch[3] . "-" . lastMatch[2] . "-" . lastMatch[1]
                case "Month-Day-Year":
                    isoDate := lastMatch[3] . "-" . lastMatch[1] . "-" . lastMatch[2]
                case "Year-Month-Day":
                    isoDate := lastMatch[1] . "-" . lastMatch[2] . "-" . lastMatch[3]
            }
        }
    }

    extractionResults := Map(
        "Input Value", inputValue,
        "Date Order",  dateOrder,
        "Success",     false
    )

    if isoDate != "" {
        extractionResults["Extracted Date"] := isoDate

        validation := ValidateDataUsingConstraint(isoDate, "String", "ISO Date")
        if validation != "" {
            extractionResults["Validation"] := validation
        } else {
            extractionResults["Success"] := true
        }
    }

    return extractionResults
}

GetDirectoryTimeAsTimestamp(directoryPath, timeType) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}"', "Accessed", "Created", "Modified")
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("directoryPath As String [Constraint: Directory], timeType As String [Whitelist: " . timeTypeWhitelist . "]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [directoryPath, timeType])

    directoryPath := RTrim(directoryPath, "\")

    offset := unset
    switch timeType {
        case "Accessed": offset := 8
        case "Created":  offset := 0
        case "Modified": offset := 16
    }

    directoryHandle := DllCall("Kernel32\CreateFileW", "WStr", directoryPath, "UInt", 0x80000000, "UInt", 0x1, "Ptr", 0, "UInt", 3, "UInt", 0x02000000, "Ptr", 0, "Ptr")
    if directoryHandle = -1 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to open directory. [Kernel32\CreateFileW" . ", System Error Code: " . A_LastError . "]")
    }

    fileTimesBuffer := Buffer(24, 0)

    retrievedDirectoryDateAndTimeSuccessfully := DllCall("Kernel32\GetFileTime", "Ptr", directoryHandle, "Ptr", fileTimesBuffer.Ptr, "Ptr", fileTimesBuffer.Ptr + 8, "Ptr", fileTimesBuffer.Ptr + 16, "Int")
    if !retrievedDirectoryDateAndTimeSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve date and time for when directory was created, last accessed and last modified. [Kernel32\GetFileTime" . ", System Error Code: " . A_LastError . "]")
    }

    utcFileTimeBuffer := Buffer(8, 0)

    DllCall("RtlMoveMemory", "Ptr", utcFileTimeBuffer.Ptr, "Ptr", fileTimesBuffer.Ptr + offset, "UPtr", 8)

    systemTimeBuffer := Buffer(16, 0)

    convertedFileTimeToSystemTimeSuccessfully := DllCall("Kernel32\FileTimeToSystemTime", "Ptr", utcFileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")
    if !convertedFileTimeToSystemTimeSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to convert file time to system time. [Kernel32\FileTimeToSystemTime" . ", System Error Code: " . A_LastError . "]")
    }

    year   := NumGet(systemTimeBuffer, 0, "UShort")
    month  := NumGet(systemTimeBuffer, 2, "UShort")
    day    := NumGet(systemTimeBuffer, 6, "UShort")
    hour   := NumGet(systemTimeBuffer, 8, "UShort")
    minute := NumGet(systemTimeBuffer, 10, "UShort")
    second := NumGet(systemTimeBuffer, 12, "UShort")

    DllCall("Kernel32\CloseHandle", "Ptr", directoryHandle)

    utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)

    return utcTimestamp
}

GetFileTimeAsTimestamp(filePath, timeType) {
    static timingBuffer     := Buffer(24, 0)
    static qpcPrePointer    := timingBuffer.Ptr
    static timestampPointer := timingBuffer.Ptr + 8
    static qpcPostPointer   := timingBuffer.Ptr + 16

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPrePointer, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampPointer)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostPointer, "Int")

    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}"', "Accessed", "Created", "Modified")
    static methodName := A_ThisFunc
    if !(methodRegistry.Has(methodName) && methodRegistry[methodName].Has("Registered")) {
        RegisterMethod("filePath As String [Constraint: Path], timeType As String [Whitelist: " . timeTypeWhitelist . "]", methodName, A_LineFile, A_LineNumber + 2, Map())
    }
    logConclusionData := LogBeginning(methodName, NumGet(qpcPrePointer, "Int64"), NumGet(timestampPointer, "Int64"), NumGet(qpcPostPointer, "Int64"), [filePath, timeType])

    offset := unset
    switch timeType {
        case "Accessed": offset := 8
        case "Created":  offset := 0
        case "Modified": offset := 16
    }

    fileHandle := DllCall("Kernel32\CreateFileW", "WStr", filePath, "UInt", 0x80000000, "UInt", 0x1, "Ptr", 0, "UInt", 3, "UInt", 0x02000000, "Ptr", 0, "Ptr")
    if fileHandle = -1 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to open file. [Kernel32\CreateFileW" . ", System Error Code: " . A_LastError . "]")
    }

    fileTimesBuffer := Buffer(24, 0)

    retrievedFileDateAndTimeSuccessfully := DllCall("Kernel32\GetFileTime", "Ptr", fileHandle, "Ptr", fileTimesBuffer.Ptr, "Ptr", fileTimesBuffer.Ptr + 8, "Ptr", fileTimesBuffer.Ptr + 16, "Int")
    if !retrievedFileDateAndTimeSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to retrieve date and time for when file was created, last accessed and last modified. [Kernel32\GetFileTime" . ", System Error Code: " . A_LastError . "]")
    }

    utcFileTimeBuffer := Buffer(8, 0)

    DllCall("RtlMoveMemory", "Ptr", utcFileTimeBuffer.Ptr, "Ptr", fileTimesBuffer.Ptr + offset, "UPtr", 8)

    systemTimeBuffer := Buffer(16, 0)

    convertedFileTimeToSystemTimeSuccessfully := DllCall("Kernel32\FileTimeToSystemTime", "Ptr", utcFileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")
    if !convertedFileTimeToSystemTimeSuccessfully {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to convert file time to system time. [Kernel32\FileTimeToSystemTime" . ", System Error Code: " . A_LastError . "]")
    }

    year   := NumGet(systemTimeBuffer, 0, "UShort")
    month  := NumGet(systemTimeBuffer, 2, "UShort")
    day    := NumGet(systemTimeBuffer, 6, "UShort")
    hour   := NumGet(systemTimeBuffer, 8, "UShort")
    minute := NumGet(systemTimeBuffer, 10, "UShort")
    second := NumGet(systemTimeBuffer, 12, "UShort")

    DllCall("Kernel32\CloseHandle", "Ptr", fileHandle)

    fileTimeAsUtc := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)

    return fileTimeAsUtc
}