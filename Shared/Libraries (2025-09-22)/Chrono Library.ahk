#Requires AutoHotkey v2.0
#Include File Library.ahk
#Include Logging Library.ahk

AssignFileTimeAsLocalIso(filePath, timeType) {
    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Accessed", "A", "Created", "C", "Modified", "M")
    static methodName := RegisterMethod("AssignFileTimeAsLocalIso(filePath As String [Type: Absolute Path], timeType As String [Whitelist: " . timeTypeWhitelist . "])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Assign File Times As Local ISO", methodName, [filePath, timeType])

    switch StrLower(timeType) {
        case "created", "c":
            offset := 0
        case "accessed", "a":
            offset := 8
        case "modified", "m":
            offset := 16
        default:
    }

    fileHandle := DllCall("Kernel32\CreateFileW", "wstr", filePath, "uint", 0x80000000, "uint", 0x1, "ptr", 0, "uint", 3, "uint", 0x02000000, "ptr", 0, "ptr")

    try {
        if fileHandle = -1 {
            throw Error("Failed to open file for reading: " . filePath)
        }
    } catch as failedToOpenFileForReadingError {
        LogInformationConclusion("Failed", logValuesForConclusion, failedToOpenFileForReadingError)
    }

    fileTimeBuffer := Buffer(24, 0)
    try {
        if !DllCall("Kernel32\GetFileTime", "ptr", fileHandle, "ptr", fileTimeBuffer.Ptr, "ptr", fileTimeBuffer.Ptr + 8, "ptr", fileTimeBuffer.Ptr + 16, "int")
        {
            throw Error("GetFileTime failed for: " . filePath)
        }
    } catch as getFileTimeFailedError {
        LogInformationConclusion("Failed", logValuesForConclusion, getFileTimeFailedError)
    }

    utcFileTime := Buffer(8, 0)
    DllCall("RtlMoveMemory", "ptr", utcFileTime.Ptr, "ptr", fileTimeBuffer.Ptr + offset, "uptr", 8)

    systemTime := Buffer(16, 0)
    try {
        if !DllCall("Kernel32\FileTimeToSystemTime", "ptr", utcFileTime.Ptr, "ptr", systemTime.Ptr, "int") {
            throw Error("FileTimeToSystemTime failed")
        }
    } catch as fileTimeToSystemTimeFailedError {
        LogInformationConclusion("Failed", logValuesForConclusion, fileTimeToSystemTimeFailedError)
    }

    localTime := Buffer(16, 0)
    try {
        if !DllCall("Kernel32\SystemTimeToTzSpecificLocalTime", "ptr", 0, "ptr", systemTime.Ptr, "ptr", localTime.Ptr, "int") {
            throw Error("SystemTimeToTzSpecificLocalTime failed")
        }
    } catch as systemTimeToTzSpecificLocalTimeFailedError {
        LogInformationConclusion("Failed", logValuesForConclusion, systemTimeToTzSpecificLocalTimeFailedError)
    }

    year   := NumGet(localTime, 0, "UShort")
    month  := NumGet(localTime, 2, "UShort")
    day    := NumGet(localTime, 6, "UShort")
    hour   := NumGet(localTime, 8, "UShort")
    minute := NumGet(localTime, 10, "UShort")
    second := NumGet(localTime, 12, "UShort")

    try {
        if IsSet(fileHandle) && fileHandle != -1 {
            DllCall("Kernel32\CloseHandle", "ptr", fileHandle)
        }
    } catch as fileHandleError {
        LogInformationConclusion("Failed", logValuesForConclusion, fileHandleError)
    }

    result := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)

    LogInformationConclusion("Completed", logValuesForConclusion)
    return result
}

ExtractTrailingDateAsIso(inputValue, dateOrder) {
    static dateOrderWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Day-Month-Year", "DMY", "Month-Day-Year", "MDY", "Year-Month-Day", "YMD")
    static methodName := RegisterMethod("ExtractTrailingDateAsIso(inputValue As String, dateOrder As String [Whitelist: " . dateOrderWhitelist . "])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Extract Trailing Date as ISO (" . inputValue . ")", methodName, [inputValue, dateOrder])

    isoDate := ""
    lastMatch := unset

    switch StrLower(dateOrder) {
        case "day-month-year", "dmy":
            currentPosition := 1
            pattern := "(?<!\d)(0[1-9]|[12]\d|3[01])([-./ ])(0[1-9]|1[0-2])\2((?:19|20)\d{2})"
            while currentPosition := RegExMatch(inputValue, pattern, &matchObject, currentPosition) {
                lastMatch := Map()
                for index, value in matchObject {
                    lastMatch[index] := value
                }
                currentPosition += StrLen(matchObject[0])
            }
            if IsSet(lastMatch) {
                isoDate := lastMatch[4] "-" lastMatch[3] "-" lastMatch[1]
            } else {
                lastMatch := unset
                currentPosition := 1
                pattern := "(?<!\d)(0[1-9]|[12]\d|3[01])(0[1-9]|1[0-2])((?:19|20)\d{2})"
                while currentPosition := RegExMatch(inputValue, pattern, &matchObject, currentPosition) {
                    lastMatch := Map()
                    for index, value in matchObject {
                        lastMatch[index] := value
                    }
                    currentPosition += StrLen(matchObject[0])
                }
                if IsSet(lastMatch) {
                    isoDate := lastMatch[3] "-" lastMatch[2] "-" lastMatch[1]
                }
            }
        case "month-day-year", "mdy":
            currentPosition := 1
            pattern := "(?<!\d)(0[1-9]|1[0-2])([-./ ])(0[1-9]|[12]\d|3[01])\2((?:19|20)\d{2})"
            while currentPosition := RegExMatch(inputValue, pattern, &matchObject, currentPosition) {
                lastMatch := Map()
                for index, value in matchObject {
                    lastMatch[index] := value
                }
                currentPosition += StrLen(matchObject[0])
            }
            if IsSet(lastMatch) {
                isoDate := lastMatch[4] "-" lastMatch[1] "-" lastMatch[3]
            } else {
                lastMatch := unset
                currentPosition := 1
                pattern := "(?<!\d)(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])((?:19|20)\d{2})"
                while currentPosition := RegExMatch(inputValue, pattern, &matchObject, currentPosition) {
                    lastMatch := Map()
                    for index, value in matchObject {
                        lastMatch[index] := value
                    }
                    currentPosition += StrLen(matchObject[0])
                }
                if IsSet(lastMatch) {
                    isoDate := lastMatch[3] "-" lastMatch[1] "-" lastMatch[2]
                }
            }
        case "year-month-day", "ymd":
            currentPosition := 1
            pattern := "(?<!\d)((?:19|20)\d{2})([-./ ])(0[1-9]|1[0-2])\2(0[1-9]|[12]\d|3[01])"
            while currentPosition := RegExMatch(inputValue, pattern, &matchObject, currentPosition) {
                lastMatch := Map()
                for index, value in matchObject {
                    lastMatch[index] := value
                }
                currentPosition += StrLen(matchObject[0])
            }
            if IsSet(lastMatch) {
                isoDate := lastMatch[1] "-" lastMatch[3] "-" lastMatch[4]
            } else {
                lastMatch := unset
                currentPosition := 1
                pattern := "(?<!\d)((?:19|20)\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])"
                while currentPosition := RegExMatch(inputValue, pattern, &matchObject, currentPosition) {
                    lastMatch := Map()
                    for index, value in matchObject {
                        lastMatch[index] := value
                    }
                    currentPosition += StrLen(matchObject[0])
                }
                if IsSet(lastMatch) {
                    isoDate := lastMatch[1] "-" lastMatch[2] "-" lastMatch[3]
                }
            }
    }

    if isoDate = "" {
        try {
            throw Error("No date found in input: " inputValue)
        } catch as noDateFoundError {
            LogInformationConclusion("Failed", logValuesForConclusion, noDateFoundError)
        }
    } else {
        try {
            dateParts := StrSplit(isoDate, "-")
            year      := dateParts[1] + 0
            month     := dateParts[2] + 0
            day       := dateParts[3] + 0

            validation := ValidateIsoDate(year, month, day)
            if validation !== "" {
                throw Error(validation)
            }
        } catch as validationError {
            LogInformationConclusion("Failed", logValuesForConclusion, validationError)
        }

        LogInformationConclusion("Completed", logValuesForConclusion)
        return isoDate
    }
}

PreventSystemGoingIdleUntilRuntime(runtimeDate, randomizePixelMovement := false) {
    static methodName := RegisterMethod("PreventSystemGoingIdleUntilRuntime(runtimeDate As String [Type: Raw Date Time], randomizePixelMovement As Boolean [Optional: false])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Prevent System Going Idle Until Runtime (" . FormatTime(runtimeDate, "yyyy-MM-dd HH:mm:ss") . ")", methodName, [runtimeDate, randomizePixelMovement])
    
    counter := 0

    if randomizePixelMovement = false {
        while (DateDiff(runtimeDate, A_Now, "Seconds") > 60) {
            counter += 1
            if counter >= 48 {
                MouseMove(0, 0, 0, "R")
                counter := 0
            }

            Sleep(10000)
        }
    } else {
        while (DateDiff(runtimeDate, A_Now, "Seconds") > 60) {
            counter += 1
            if counter >= 48 {
                MouseGetPos(&mouseX, &mouseY)
                screenWidth  := A_ScreenWidth
                screenHeight := A_ScreenHeight

                direction := Random(1, 4)

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
                    MouseMove 1, 0, 0, "R"
                } else if direction = 2 {
                    MouseMove -1, 0, 0, "R"
                } else if direction = 3 {
                    MouseMove 0, 1, 0, "R"
                } else {
                    MouseMove 0, -1, 0, "R"
                }

                Sleep(Random(200, 800))
                if direction = 1 {
                    MouseMove -1, 0, 0, "R"
                } else if direction = 2 {
                    MouseMove 1, 0, 0, "R"
                } else if direction = 3 {
                    MouseMove 0, -1, 0, "R"
                } else {
                    MouseMove 0, 1, 0, "R"
                }

                counter := 0
            }

            Sleep(10000)
        }
    }

    while (A_Now < DateAdd(runtimeDate, -1, "Seconds")) {
        Sleep(240)
    }

    while (A_Now < runtimeDate) {
        Sleep(16)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

SetDirectoryTimeFromLocalIsoDateTime(directoryPath, localIsoDateTime, timeType) {
    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Accessed", "A", "Created", "C", "Modified", "M")
    static methodName := RegisterMethod("SetDirectoryTimeFromLocalIsoDateTime(directoryPath As String [Type: Directory], localIsoDateTime As String [Type: ISO Date Time], timeType As String [Whitelist: " . timeTypeWhitelist . "])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Set Directory Time From Local ISO Date Time", methodName, [directoryPath, localIsoDateTime, timeType])

    directoryPath := RTrim(directoryPath, "\")

    numericString := RegExReplace(localIsoDateTime, "[^0-9]")
    localSystemTime := Buffer(16, 0)
    NumPut "UShort", SubStr(numericString, 1, 4),  localSystemTime,  0
    NumPut "UShort", SubStr(numericString, 5, 2),  localSystemTime,  2
    NumPut "UShort", 0,                            localSystemTime,  4
    NumPut "UShort", SubStr(numericString, 7, 2),  localSystemTime,  6
    NumPut "UShort", SubStr(numericString, 9, 2),  localSystemTime,  8
    NumPut "UShort", SubStr(numericString,11, 2),  localSystemTime, 10
    NumPut "UShort", SubStr(numericString,13, 2),  localSystemTime, 12
    NumPut "UShort", 0,                            localSystemTime, 14

    utcSystemTime := Buffer(16, 0)
    try {
        if !DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "ptr", 0, "ptr", localSystemTime, "ptr", utcSystemTime, "int") {
            throw Error("TzSpecificLocalTimeToSystemTime failed (input may not exist in current time zone)")
        }
    } catch as tzSpecificLocalTimeToSystemTimeError {
        LogInformationConclusion("Failed", logValuesForConclusion, tzSpecificLocalTimeToSystemTimeError)
    }

    utcFileTime := Buffer(8, 0)
    try {
        if !DllCall("Kernel32\SystemTimeToFileTime", "ptr", utcSystemTime, "ptr", utcFileTime, "int") {
            throw Error("SystemTimeToFileTime failed")
        }
    } catch as systemTimeToFileTimeError {
        LogInformationConclusion("Failed", logValuesForConclusion, systemTimeToFileTimeError)
    }

    accessMode := 0x100
    shareMode  := 0x7
    flags      := 0x80 | 0x02000000
    handle     := DllCall("Kernel32\CreateFileW", "wstr", directoryPath, "uint", accessMode, "uint", shareMode, "ptr", 0, "uint", 3, "uint", flags, "ptr", 0, "ptr")

    try {
        if handle = -1 {
            throw Error("CreateFileW failed")
        }
    } catch as createFileWFailedError {
        LogInformationConclusion("Failed", logValuesForConclusion, createFileWFailedError)
    }

    pointerCreation := 0
    pointerAccessed := 0
    pointerModified := 0
    switch StrLower(timeType) {
        case "accessed", "a":
            pointerAccessed := utcFileTime.Ptr
        case "created", "c":
            pointerCreation := utcFileTime.Ptr
        case "modified", "m":
            pointerModified := utcFileTime.Ptr
        default:
    }

    success := DllCall("Kernel32\SetFileTime"
        , "ptr", handle
        , "ptr", pointerCreation
        , "ptr", pointerAccessed
        , "ptr", pointerModified
        , "int")

    DllCall("Kernel32\CloseHandle", "ptr", handle)

    try { 
        if !success {
            throw Error("SetFileTime failed")
        }
    } catch as setFileTimeFailedError {
        LogInformationConclusion("Failed", logValuesForConclusion, setFileTimeFailedError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

SetFileTimeFromLocalIsoDateTime(filePath, localIsoDateTime, timeType) {
    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Accessed", "A", "Created", "C", "Modified", "M")
    static methodName := RegisterMethod("SetFileTimeFromLocalIsoDateTime(filePath As String [Type: Absolute Path], localIsoDateTime As String [Type: ISO Date Time], timeType As String [Whitelist: " . timeTypeWhitelist . "])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Set File Time From Local ISO Date Time", methodName, [filePath, localIsoDateTime, timeType])

    if AssignFileTimeAsLocalIso(filePath, timeType) = localIsoDateTime {
        LogInformationConclusion("Skipped", logValuesForConclusion)
    } else {
        numericString := RegExReplace(localIsoDateTime, "[^0-9]")
        localSystemTime := Buffer(16, 0)
        NumPut "UShort", SubStr(numericString, 1, 4),  localSystemTime,  0
        NumPut "UShort", SubStr(numericString, 5, 2),  localSystemTime,  2
        NumPut "UShort", 0,                            localSystemTime,  4
        NumPut "UShort", SubStr(numericString, 7, 2),  localSystemTime,  6
        NumPut "UShort", SubStr(numericString, 9, 2),  localSystemTime,  8
        NumPut "UShort", SubStr(numericString,11, 2),  localSystemTime, 10
        NumPut "UShort", SubStr(numericString,13, 2),  localSystemTime, 12
        NumPut "UShort", 0,                            localSystemTime, 14

        utcSystemTime := Buffer(16, 0)
        try {
            if !DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "ptr", 0, "ptr", localSystemTime, "ptr", utcSystemTime, "int") {
                throw Error("TzSpecificLocalTimeToSystemTime failed (input may not exist in current time zone)")
            }
        } catch as tzSpecificLocalTimeToSystemTimeError {
            LogInformationConclusion("Failed", logValuesForConclusion, tzSpecificLocalTimeToSystemTimeError)
        }

        utcFileTime := Buffer(8, 0)
        try {
            if !DllCall("Kernel32\SystemTimeToFileTime", "ptr", utcSystemTime, "ptr", utcFileTime, "int") {
                throw Error("SystemTimeToFileTime failed")
            }
        } catch as systemTimeToFileTimeError {
            LogInformationConclusion("Failed", logValuesForConclusion, systemTimeToFileTimeError)
        }

        accessMode := 0x100
        shareMode  := 0x7
        flags      := 0x80 | 0x02000000
        handle     := DllCall("Kernel32\CreateFileW", "wstr", filePath, "uint", accessMode, "uint", shareMode, "ptr", 0, "uint", 3, "uint", flags, "ptr", 0, "ptr")

        try {
            if handle = -1 {
                throw Error("CreateFileW failed")
            }
        } catch as createFileWFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, createFileWFailedError)
        }

        pointerCreation := 0
        pointerAccessed := 0
        pointerModified := 0
        switch StrLower(timeType) {
            case "accessed", "a":
                pointerAccessed := utcFileTime.Ptr
            case "created", "c":
                pointerCreation := utcFileTime.Ptr
            case "modified", "m":
                pointerModified := utcFileTime.Ptr
            default:
        }

        success := DllCall("Kernel32\SetFileTime", "ptr", handle, "ptr", pointerCreation, "ptr", pointerAccessed, "ptr", pointerModified, "int")

        DllCall("Kernel32\CloseHandle", "ptr", handle)

        try { 
            if !success {
                throw Error("SetFileTime failed")
            }
        } catch as setFileTimeFailedError {
            LogInformationConclusion("Failed", logValuesForConclusion, setFileTimeFailedError)
        }

        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

ValidateRuntimeDate(runtimeDate, minimumStartupInSeconds) {
    static methodName := RegisterMethod("ValidateRuntimeDate(runtimeDate As String [Type: Raw Date Time], minimumStartupInSeconds As Integer)" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Validate Runtime Date (" . runtimeDate . ")", methodName, [runtimeDate, minimumStartupInSeconds])

    try {
        if runtimeDate <= A_Now {
            throw Error("runtimeDate is in the past.")
        }
    } catch as runtimeInPastError {
        LogInformationConclusion("Failed", logValuesForConclusion, runtimeInPastError)
    }

    try {
        timeUntilStart := DateDiff(runtimeDate, A_Now, "Seconds")
        if timeUntilStart < minimumStartupInSeconds && SubStr(runtimeDate, 1, 8) = SubStr(A_Now, 1, 8) {
            throw Error("runtimeDate must be at least " . minimumStartupInSeconds . " seconds into the future. Current difference: " . timeUntilStart . " seconds.")
        }
    } catch as runtimeTooEarlyError {
        LogInformationConclusion("Failed", logValuesForConclusion, runtimeTooEarlyError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

WaitUntilFileIsModifiedToday(filePath) {
    static methodName := RegisterMethod("WaitUntilFileIsModifiedToday(filePath As String [Type: Absolute Path])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Wait Until File is Modified Today: " . ExtractFilename(filePath, true), methodName, [filePath])

    dateOfToday := FormatTime(A_Now, "yyyy-MM-dd")
    checkInterval := 4000   ; Check every 4 seconds (in milliseconds)
    mouseInterval := 120000 ; Move mouse every 2 minutes (in milliseconds)
    maxWaitMinutes := 360   ; Maximum wait time = 6 hours
    ; Calculate how many times to loop based on max wait time and check interval.
    ; Example: (360 minutes × 60,000 ms) ÷ 4,000 ms = 5,400 loops (i.e. 6 hours total)
    maxLoops := (maxWaitMinutes * 60000) // checkInterval
    timeSinceLastMouse := 0

    Loop maxLoops {
        fileModifiedDate := FileGetTime(filePath, "M") ; Get modified date for file.
        fileModifiedDate := FormatTime(fileModifiedDate, "yyyy-MM-dd")

        if dateOfToday = fileModifiedDate {
            break ; Break if file modified today.
        }

        Sleep(checkInterval)
        timeSinceLastMouse += checkInterval

        if timeSinceLastMouse >= mouseInterval {
            MouseMove(0, 0, 0, "R") ; For preventing screen saver from activating.
            timeSinceLastMouse := 0
        }
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

; ******************** ;
; Helper Methods       ;
; ******************** ;

CaptureTimeAnchor() {
    maximumSpinIterations := 2000000
    yieldEveryNSpins      := 50000
    timeoutMilliseconds   := 100

    ; Initial samples before the next coarse tick boundary
    tickBeforeChange                      := A_TickCount
    preciseUtcFileTimeBefore              := GetSystemTimePreciseAsFileTime64()
    queryPerformanceCounterTicksBefore    := QueryPerformanceCounterNow()
    queryPerformanceCounterTicksPerSecond := QueryPerformanceCounterFrequency()
    timeoutStartTickCount                 := A_TickCount

    spinIterations := 0

    loop maximumSpinIterations {
        spinIterations += 1

        currentTick := A_TickCount
        if currentTick != tickBeforeChange {
            tickAfterChange                 := currentTick
            preciseUtcFileTimeAfter         := GetSystemTimePreciseAsFileTime64()
            queryPerformanceCounterTicksAfter := QueryPerformanceCounterNow()

            preciseUtcFileTimeMidpoint := (preciseUtcFileTimeBefore + preciseUtcFileTimeAfter) // 2
            queryPerformanceCounterTicksMidpoint := (queryPerformanceCounterTicksBefore + queryPerformanceCounterTicksAfter) // 2

            ; Human-readable snapshots (for sanity checks)
            utcDateTimeIso   := FormatTime(A_NowUTC, "yyyy-MM-dd HH:mm:ss")
            localDateTimeIso := FormatTime(A_Now,    "yyyy-MM-dd HH:mm:ss")
            millisecondsPart := Format("{:03}", A_MSec)

            return Map(
                ; Coarse monotonic (ticks)
                "Tick Before Change",                       tickBeforeChange,
                "Tick After Change",                        tickAfterChange,
                "Tick Delta",                               tickAfterChange - tickBeforeChange,

                ; Precise wall-clock (UTC FILETIME, 100-ns)
                "Precise UTC FileTime Before",              preciseUtcFileTimeBefore,
                "Precise UTC FileTime After",               preciseUtcFileTimeAfter,
                "Precise UTC FileTime Midpoint",            preciseUtcFileTimeMidpoint,

                ; High-resolution monotonic (QueryPerformanceCounter)
                "QueryPerformanceCounter Ticks Before",     queryPerformanceCounterTicksBefore,
                "QueryPerformanceCounter Ticks After",      queryPerformanceCounterTicksAfter,
                "QueryPerformanceCounter Ticks Midpoint",   queryPerformanceCounterTicksMidpoint,
                "QueryPerformanceCounter Ticks Per Second", queryPerformanceCounterTicksPerSecond,

                ; Human-readable checks
                "UTC Date Time ISO",                        utcDateTimeIso,
                "Local Date Time ISO",                      localDateTimeIso,
                "Milliseconds Part",                        millisecondsPart,

                ; Diagnostics
                "Spin Iterations",                          spinIterations,
                "Timed Out",                                false
            )
        }

        ; Refresh "before" anchors while spinning toward the boundary
        preciseUtcFileTimeBefore           := GetSystemTimePreciseAsFileTime64()
        queryPerformanceCounterTicksBefore := QueryPerformanceCounterNow()

        ; Yield periodically so we do not hog the scheduler
        if Mod(spinIterations, yieldEveryNSpins) = 0 {
            Sleep(0)
        }

        ; Hard timeout guard: Best-effort snapshot without observing a boundary
        if A_TickCount - timeoutStartTickCount > timeoutMilliseconds {
            return Map(
                "Tick Before Change",                       tickBeforeChange,
                "Tick After Change",                        tickBeforeChange,
                "Tick Delta",                               0,

                "Precise UTC FileTime Before",              preciseUtcFileTimeBefore,
                "Precise UTC FileTime After",               preciseUtcFileTimeBefore,
                "Precise UTC FileTime Midpoint",            preciseUtcFileTimeBefore,

                "QueryPerformanceCounter Ticks Before",     queryPerformanceCounterTicksBefore,
                "QueryPerformanceCounter Ticks After",      queryPerformanceCounterTicksBefore,
                "QueryPerformanceCounter Ticks Midpoint",   queryPerformanceCounterTicksBefore,
                "QueryPerformanceCounter Ticks Per Second", queryPerformanceCounterTicksPerSecond,

                "UTC Date Time ISO",                        FormatTime(A_NowUTC, "yyyy-MM-dd HH:mm:ss"),
                "Local Date Time ISO",                      FormatTime(A_Now,    "yyyy-MM-dd HH:mm:ss"),
                "Milliseconds Part",                        Format("{:03}", A_MSec),

                "Spin Iterations",                          spinIterations,
                "Timed Out",                                true
            )
        }
    }
}

ConvertIsoToRawDateTime(isoDateTime) {
    isoDateTime := RegExReplace(isoDateTime, "[-: ]", "")

    return isoDateTime
}

GetSystemTimePreciseAsFileTime64() {
    fileTimeBuffer := Buffer(8, 0)
    DllCall("Kernel32\GetSystemTimePreciseAsFileTime", "ptr", fileTimeBuffer.Ptr)
    lowPart  := NumGet(fileTimeBuffer, 0, "UInt")
    highPart := NumGet(fileTimeBuffer, 4, "UInt")

    return (highPart << 32) | lowPart
}

IsValidGregorianDay(year, month, day) {
    isLeap := (Mod(year, 400) = 0) || (Mod(year, 4) = 0 && Mod(year, 100) != 0)
    daysInMonth := [31, (isLeap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    if day <= daysInMonth[month] {
        return true
    } else {
        return false
    }
}

LocalIsoWithUtcTag(localIsoString) {
    if !RegExMatch(localIsoString, "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$") {
        return localIsoString
    }

    dateTimeParts := StrSplit(localIsoString, " ")
    dateParts     := StrSplit(dateTimeParts[1], "-")
    timeParts     := StrSplit(dateTimeParts[2], ":")

    year   := dateParts[1] + 0
    month  := dateParts[2] + 0
    day    := dateParts[3] + 0
    hour   := timeParts[1] + 0
    minute := timeParts[2] + 0
    second := timeParts[3] + 0

    localSystemTime := Buffer(16, 0)
    NumPut "UShort", year,   localSystemTime, 0
    NumPut "UShort", month,  localSystemTime, 2
    NumPut "UShort", 0,      localSystemTime, 4
    NumPut "UShort", day,    localSystemTime, 6
    NumPut "UShort", hour,   localSystemTime, 8
    NumPut "UShort", minute, localSystemTime, 10
    NumPut "UShort", second, localSystemTime, 12
    NumPut "UShort", 0,      localSystemTime, 14

    utcSystemTime := Buffer(16, 0)
    utcSuccess := DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "ptr", 0, "ptr", localSystemTime, "ptr", utcSystemTime, "int")

    if !utcSuccess {
        return localIsoString
    }

    utcYear   := NumGet(utcSystemTime, 0, "UShort")
    utcMonth  := NumGet(utcSystemTime, 2, "UShort")
    utcDay    := NumGet(utcSystemTime, 6, "UShort")
    utcHour   := NumGet(utcSystemTime, 8, "UShort")
    utcMinute := NumGet(utcSystemTime, 10, "UShort")
    utcSecond := NumGet(utcSystemTime, 12, "UShort")
    utcIso := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", utcYear, utcMonth, utcDay, utcHour, utcMinute, utcSecond)

    return localIsoString " <UTC " utcIso ">"
}

QueryPerformanceCounterFrequency() {
    static cachedFrequency := 0
    if cachedFrequency = 0 {
        localBuffer := Buffer(8, 0)
        if !DllCall("Kernel32\QueryPerformanceFrequency", "ptr", localBuffer.Ptr, "int") {
            throw Error("QueryPerformanceFrequency failed.")
        }

        cachedFrequency := NumGet(localBuffer, 0, "Int64")
    }

    return cachedFrequency
}

QueryPerformanceCounterNow() {
    localBuffer := Buffer(8, 0)
    if !DllCall("Kernel32\QueryPerformanceCounter", "ptr", localBuffer.Ptr, "int") {
        throw Error("QueryPerformanceCounter failed.")
    }

    return NumGet(localBuffer, 0, "Int64")
}

ValidateIsoDate(year, month, day, hour := 0, minute := 0, second := 0) {
    validationResults := ""

    year   := Number(year)
    month  := Number(month)
    day    := Number(day)
    hour   := Number(hour)
    minute := Number(minute)
    second := Number(second)

    if year < 0 || year > 9999 {
        validationResults := "Invalid ISO 8601 Date: " . year . " (year out of range 0000–9999)."
    } else if month < 1 || month > 12 {
        validationResults := "Invalid ISO 8601 Date: " . month . " (month must be 01–12)."
    } else if day < 1 || day > 31 {
        validationResults := "Invalid ISO 8601 Date: " . day . " (day must be 01–31)."
    } else if !IsValidGregorianDay(year, month, day) {
        validationResults := "Invalid ISO 8601 Date: " . day . " (day out of range for month)."
    }
    
    if validationResults = "" && !(hour = 0 && minute = 0 && second = 0) {
        if hour < 0 || hour > 23 {
            validationResults := "Invalid ISO 8601 Date Time: " . hour . " (hour must be 00–23)."
        } else if minute < 0 || minute > 59 {
            validationResults := "Invalid ISO 8601 Date Time: " . minute . " (minute must be 00–59)."
        } else if second < 0 || second > 59 {
            validationResults := "Invalid ISO 8601 Date Time: " . second . " (second must be 00–59)."
        } else {
            localSystemTime := Buffer(16, 0)
            NumPut "UShort", year,   localSystemTime, 0
            NumPut "UShort", month,  localSystemTime, 2
            NumPut "UShort", 0,      localSystemTime, 4
            NumPut "UShort", day,    localSystemTime, 6
            NumPut "UShort", hour,   localSystemTime, 8
            NumPut "UShort", minute, localSystemTime, 10
            NumPut "UShort", second, localSystemTime, 12
            NumPut "UShort", 0,      localSystemTime, 14

            utcSystemTime := Buffer(16, 0)
            if !DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "ptr", 0, "ptr", localSystemTime, "ptr", utcSystemTime, "int") {
                validationResults := Format("Invalid ISO 8601 Date Time: {:04}-{:02}-{:02} {:02}:{:02}:{:02} (nonexistent local time, DST gap or system restriction).", year, month, day, hour, minute, second)
            } else {
                systemTimeLocal := Buffer(16, 0)
                DllCall("Kernel32\SystemTimeToTzSpecificLocalTime", "ptr", 0, "ptr", utcSystemTime, "ptr", systemTimeLocal, "int")
                roundTripYear   := NumGet(systemTimeLocal, 0, "UShort")
                roundTripMonth  := NumGet(systemTimeLocal, 2, "UShort")
                roundTripDay    := NumGet(systemTimeLocal, 6, "UShort")
                roundTripHour   := NumGet(systemTimeLocal, 8, "UShort")
                roundTripMinute := NumGet(systemTimeLocal, 10, "UShort")
                roundTripSecond := NumGet(systemTimeLocal, 12, "UShort")
                if roundTripYear != year || roundTripMonth != month || roundTripDay != day || roundTripHour != hour || roundTripMinute != minute || roundTripSecond != second {
                    validationResults := Format("Invalid ISO 8601 Date Time: {:04}-{:02}-{:02} {:02}:{:02}:{:02} (nonexistent local time after round-trip, DST gap).", year, month, day, hour, minute, second)
                }
            }
        }
    }

    return validationResults
}