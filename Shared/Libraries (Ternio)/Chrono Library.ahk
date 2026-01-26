#Requires AutoHotkey v2.0
#Include File Library.ahk
#Include Logging Library.ahk

AssignFileTimeAsLocalIso(filePath, timeType) {
    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Accessed", "A", "Created", "C", "Modified", "M")
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Path], timeType As String [Whitelist: " . timeTypeWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath, timeType], "Assign File Times As Local ISO")

    switch StrLower(timeType) {
        case "created", "c":
            offset := 0
        case "accessed", "a":
            offset := 8
        case "modified", "m":
            offset := 16
    }

    fileHandle := DllCall("Kernel32\CreateFileW", "WStr", filePath, "UInt", 0x80000000, "UInt", 0x1, "Ptr", 0, "UInt", 3, "UInt", 0x02000000, "Ptr", 0, "Ptr")

    if fileHandle = -1 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to open file for reading: " . filePath)
    }

    fileTimeBuffer := Buffer(24, 0)
    if !DllCall("Kernel32\GetFileTime", "Ptr", fileHandle, "Ptr", fileTimeBuffer.Ptr, "Ptr", fileTimeBuffer.Ptr + 8, "Ptr", fileTimeBuffer.Ptr + 16, "Int") {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "GetFileTime failed for: " . filePath)
    }

    utcFileTime := Buffer(8, 0)
    DllCall("RtlMoveMemory", "Ptr", utcFileTime.Ptr, "Ptr", fileTimeBuffer.Ptr + offset, "UPtr", 8)

    systemTime := Buffer(16, 0)
    if !DllCall("Kernel32\FileTimeToSystemTime", "Ptr", utcFileTime.Ptr, "Ptr", systemTime.Ptr, "Int") {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "FileTimeToSystemTime failed")
    }

    localTime := Buffer(16, 0)
    if !DllCall("Kernel32\SystemTimeToTzSpecificLocalTime", "Ptr", 0, "Ptr", systemTime.Ptr, "Ptr", localTime.Ptr, "Int") {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "SystemTimeToTzSpecificLocalTime failed")
    }

    year   := NumGet(localTime, 0, "UShort")
    month  := NumGet(localTime, 2, "UShort")
    day    := NumGet(localTime, 6, "UShort")
    hour   := NumGet(localTime, 8, "UShort")
    minute := NumGet(localTime, 10, "UShort")
    second := NumGet(localTime, 12, "UShort")

    try {
        if IsSet(fileHandle) && fileHandle != -1 {
            DllCall("Kernel32\CloseHandle", "Ptr", fileHandle)
        }
    } catch as fileHandleError {
        LogConclusion("Failed", logValuesForConclusion, fileHandleError)
    }

    result := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)

    LogConclusion("Completed", logValuesForConclusion)
    return result
}

ExtractTrailingDateAsIso(inputValue, dateOrder) {
    static dateOrderWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Day-Month-Year", "DMY", "Month-Day-Year", "MDY", "Year-Month-Day", "YMD")
    static methodName := RegisterMethod("inputValue As String, dateOrder As String [Whitelist: " . dateOrderWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [inputValue, dateOrder], "Extract Trailing Date as ISO (" . inputValue . ")")

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
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "No date found in input: " inputValue)
    } else {
        dateParts := StrSplit(isoDate, "-")
        year      := dateParts[1] + 0
        month     := dateParts[2] + 0
        day       := dateParts[3] + 0

        validation := ""
        if validation = "" {
            validation := ValidateDataUsingSpecification(year, "Integer", "Year")
        }

        if validation = "" {
            validation := ValidateDataUsingSpecification(month, "Integer", "Month")
        }

        if validation = "" {
            validation := ValidateDataUsingSpecification(day, "Integer", "Day")
        }

        if validation != "" {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, validation)
        }

        LogConclusion("Completed", logValuesForConclusion)
        return isoDate
    }
}

PreventSystemGoingIdleUntilRuntime(runtimeDate, randomizePixelMovement := false) {
    static methodName := RegisterMethod("runtimeDate As String [Constraint: Raw Date Time], randomizePixelMovement As Boolean [Optional: false]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [runtimeDate, randomizePixelMovement], "Prevent System Going Idle Until Runtime (" . FormatTime(runtimeDate, "yyyy-MM-dd HH:mm:ss") . ")")
    
    counter := 0

    if !randomizePixelMovement {
        while DateDiff(runtimeDate, A_Now, "Seconds") > 60 {
            counter += 1
            if counter >= 48 {
                MouseMove(0, 0, 0, "R")
                counter := 0
            }

            Sleep(10000)
        }
    } else {
        while DateDiff(runtimeDate, A_Now, "Seconds") > 60 {
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

    while A_Now < DateAdd(runtimeDate, -1, "Seconds") {
        Sleep(240)
    }

    while A_Now < runtimeDate {
        Sleep(16)
    }

    LogConclusion("Completed", logValuesForConclusion)
}

SetDirectoryTimeFromLocalIsoDateTime(directoryPath, localIsoDateTime, timeType) {
    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Accessed", "A", "Created", "C", "Modified", "M")
    static methodName := RegisterMethod("directoryPath As String [Constraint: Directory], localIsoDateTime As String [Constraint: ISO Date Time], timeType As String [Whitelist: " . timeTypeWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [directoryPath, localIsoDateTime, timeType], "Set Directory Time From Local ISO Date Time")

    directoryPath := RTrim(directoryPath, "\")

    numericString := RegExReplace(localIsoDateTime, "[^0-9]")
    localSystemTime := Buffer(16, 0)
    NumPut("UShort", SubStr(numericString, 1, 4), localSystemTime,  0)
    NumPut("UShort", SubStr(numericString, 5, 2), localSystemTime,  2)
    NumPut("UShort", 0,                           localSystemTime,  4)
    NumPut("UShort", SubStr(numericString, 7, 2), localSystemTime,  6)
    NumPut("UShort", SubStr(numericString, 9, 2), localSystemTime,  8)
    NumPut("UShort", SubStr(numericString,11, 2), localSystemTime, 10)
    NumPut("UShort", SubStr(numericString,13, 2), localSystemTime, 12)
    NumPut("UShort", 0,                           localSystemTime, 14)

    utcSystemTime := Buffer(16, 0)
    if !DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "Ptr", 0, "Ptr", localSystemTime, "Ptr", utcSystemTime, "Int") {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "TzSpecificLocalTimeToSystemTime failed (input may not exist in current time zone)")
    }

    utcFileTime := Buffer(8, 0)
    if !DllCall("Kernel32\SystemTimeToFileTime", "Ptr", utcSystemTime, "Ptr", utcFileTime, "Int") {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "SystemTimeToFileTime failed")
    }

    accessMode := 0x100
    shareMode  := 0x7
    flags      := 0x80 | 0x02000000
    handle     := DllCall("Kernel32\CreateFileW", "WStr", directoryPath, "UInt", accessMode, "UInt", shareMode, "Ptr", 0, "UInt", 3, "UInt", flags, "Ptr", 0, "Ptr")

    if handle = -1 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "CreateFileW failed")
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
    }

    success := DllCall("Kernel32\SetFileTime", "Ptr", handle, "Ptr", pointerCreation, "Ptr", pointerAccessed, "Ptr", pointerModified, "Int")

    DllCall("Kernel32\CloseHandle", "Ptr", handle)

    if !success {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "SetFileTime failed")
    }

    LogConclusion("Completed", logValuesForConclusion)
}

SetFileTimeFromLocalIsoDateTime(filePath, localIsoDateTime, timeType) {
    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Accessed", "A", "Created", "C", "Modified", "M")
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Path], localIsoDateTime As String [Constraint: ISO Date Time], timeType As String [Whitelist: " . timeTypeWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath, localIsoDateTime, timeType], "Set File Time From Local ISO Date Time")

    if AssignFileTimeAsLocalIso(filePath, timeType) = localIsoDateTime {
        LogConclusion("Skipped", logValuesForConclusion)
    } else {
        numericString := RegExReplace(localIsoDateTime, "[^0-9]")
        localSystemTime := Buffer(16, 0)
        NumPut("UShort", SubStr(numericString, 1, 4),  localSystemTime,  0)
        NumPut("UShort", SubStr(numericString, 5, 2),  localSystemTime,  2)
        NumPut("UShort", 0,                            localSystemTime,  4)
        NumPut("UShort", SubStr(numericString, 7, 2),  localSystemTime,  6)
        NumPut("UShort", SubStr(numericString, 9, 2),  localSystemTime,  8)
        NumPut("UShort", SubStr(numericString,11, 2),  localSystemTime, 10)
        NumPut("UShort", SubStr(numericString,13, 2),  localSystemTime, 12)
        NumPut("UShort", 0,                            localSystemTime, 14)

        utcSystemTime := Buffer(16, 0)
        if !DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "Ptr", 0, "Ptr", localSystemTime, "Ptr", utcSystemTime, "Int") {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "TzSpecificLocalTimeToSystemTime failed (input may not exist in current time zone)")
        }

        utcFileTime := Buffer(8, 0)
        if !DllCall("Kernel32\SystemTimeToFileTime", "Ptr", utcSystemTime, "Ptr", utcFileTime, "Int") {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "SystemTimeToFileTime failed")
        }

        accessMode := 0x100
        shareMode  := 0x7
        flags      := 0x80 | 0x02000000
        handle     := DllCall("Kernel32\CreateFileW", "WStr", filePath, "UInt", accessMode, "UInt", shareMode, "Ptr", 0, "UInt", 3, "UInt", flags, "Ptr", 0, "Ptr")

        if handle = -1 {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "CreateFileW failed")
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

        success := DllCall("Kernel32\SetFileTime", "Ptr", handle, "Ptr", pointerCreation, "Ptr", pointerAccessed, "Ptr", pointerModified, "Int")

        DllCall("Kernel32\CloseHandle", "Ptr", handle)

        if !success {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "SetFileTime failed")
        }

        LogConclusion("Completed", logValuesForConclusion)
    }
}

ValidateRuntimeDate(runtimeDate, minimumStartupInSeconds) {
    static methodName := RegisterMethod("runtimeDate As String [Constraint: Raw Date Time], minimumStartupInSeconds As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [runtimeDate, minimumStartupInSeconds], "Validate Runtime Date (" . runtimeDate . ")")

    if runtimeDate <= A_Now {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "runtimeDate is in the past.")
    }

    timeUntilStart := DateDiff(runtimeDate, A_Now, "Seconds")
    if timeUntilStart < minimumStartupInSeconds && SubStr(runtimeDate, 1, 8) = SubStr(A_Now, 1, 8) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "runtimeDate must be at least " . minimumStartupInSeconds . " seconds into the future. Current difference: " . timeUntilStart . " seconds.")
    }

    LogConclusion("Completed", logValuesForConclusion)
}

WaitUntilFileIsModifiedToday(filePath) {
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Save Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath], "Wait Until File is Modified Today: " . ExtractFilename(filePath, true))

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        SetMethodSetting(methodName, "Check Interval", 4000, false)
        SetMethodSetting(methodName, "Mouse Interval", 120000, false)
        SetMethodSetting(methodName, "Max Wait Minutes", 360, false)

        defaultMethodSettingsSet := true
    }

    settings    := methodRegistry[methodName]["Settings"]

    checkInterval  := settings.Get("Check Interval")
    mouseInterval  := settings.Get("Mouse Interval")
    maxWaitMinutes := settings.Get("Max Wait Minutes")

    dateOfToday := FormatTime(A_Now, "yyyy-MM-dd")
    maxLoops := (maxWaitMinutes * 60000) // checkInterval
    timeSinceLastMouse := 0

    loop maxLoops {
        if FileExist(filePath) {
            fileModifiedDate := FileGetTime(filePath, "M")
            fileModifiedDate := FormatTime(fileModifiedDate, "yyyy-MM-dd")

            if dateOfToday = fileModifiedDate {
                break
            }
        }

        Sleep(checkInterval)
        timeSinceLastMouse += checkInterval

        if timeSinceLastMouse >= mouseInterval {
            MouseMove(0, 0, 0, "R") ; For preventing screen saver from activating.
            timeSinceLastMouse := 0
        }
    }

    LogConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Core Methods                 ;
; **************************** ;

GetQueryPerformanceCounter() {
    static queryPerformanceCounterBuffer := Buffer(8, 0)

    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", queryPerformanceCounterBuffer.Ptr, "Int")

    queryPerformanceCounter := NumGet(queryPerformanceCounterBuffer, 0, "Int64")

    return queryPerformanceCounter
}

GetUtcTimestamp() {
    static systemTime := Buffer(16, 0)

    DllCall("Kernel32\GetSystemTime", "Ptr", systemTime.Ptr)

    year        := NumGet(systemTime,  0, "UShort")
    month       := NumGet(systemTime,  2, "UShort")
    day         := NumGet(systemTime,  6, "UShort")
    hour        := NumGet(systemTime,  8, "UShort")
    minute      := NumGet(systemTime, 10, "UShort")
    second      := NumGet(systemTime, 12, "UShort")
    millisecond := NumGet(systemTime, 14, "UShort")

    utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}", year, month, day, hour, minute, second, millisecond)

    return utcTimestamp
}

GetUtcTimestampInteger() {
    static systemTime := Buffer(16, 0)

    DllCall("Kernel32\GetSystemTime", "Ptr", systemTime.Ptr)

    year        := NumGet(systemTime,  0, "UShort")
    month       := NumGet(systemTime,  2, "UShort")
    day         := NumGet(systemTime,  6, "UShort")
    hour        := NumGet(systemTime,  8, "UShort")
    minute      := NumGet(systemTime, 10, "UShort")
    second      := NumGet(systemTime, 12, "UShort")
    millisecond := NumGet(systemTime, 14, "UShort")

    utcTimestampInteger := Format("{:04}{:02}{:02}{:02}{:02}{:02}{:03}", year, month, day, hour, minute, second, millisecond) + 0

    return utcTimestampInteger
}

GetUtcTimestampPrecise() {
    static fileTimeBuffer   := Buffer(8, 0)
    static systemTimeBuffer := Buffer(16, 0)

    static usePreciseSystemTimeFunction := unset

    if !IsSet(usePreciseSystemTimeFunction) {
        kernel32ModuleHandle   := DllCall("GetModuleHandle", "Str", "Kernel32", "Ptr")
        preciseFunctionAddress := DllCall("GetProcAddress", "Ptr", kernel32ModuleHandle, "AStr", "GetSystemTimePreciseAsFileTime", "Ptr")

        if preciseFunctionAddress != 0 {
            usePreciseSystemTimeFunction := true
        } else {
            usePreciseSystemTimeFunction := false
        }
    }

    if usePreciseSystemTimeFunction {
        DllCall("Kernel32\GetSystemTimePreciseAsFileTime", "Ptr", fileTimeBuffer.Ptr)       
        DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")

        year       := NumGet(systemTimeBuffer,  0, "UShort")
        month      := NumGet(systemTimeBuffer,  2, "UShort")
        dayOfMonth := NumGet(systemTimeBuffer,  6, "UShort")
        hour       := NumGet(systemTimeBuffer,  8, "UShort")
        minute     := NumGet(systemTimeBuffer, 10, "UShort")
        second     := NumGet(systemTimeBuffer, 12, "UShort")

        fileTimeTicks     := NumGet(fileTimeBuffer, 0, "Int64")
        ticksWithinSecond := Mod(fileTimeTicks, 10000000)
        microsecond       := ticksWithinSecond // 10

        utcTimestampPrecise := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:06}", year, month, dayOfMonth, hour, minute, second, microsecond)
    } else {
        utcTimestampPrecise := GetUtcTimestamp()
    }

    return utcTimestampPrecise
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

ConvertIntegerToUtcTimestamp(integerValue) {
    static methodName := RegisterMethod("integerValue As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [integerValue])

    digitText := integerValue . ""
    if !RegExMatch(digitText, "^\d+$") {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Input must contain only digits. Got: " . digitText)
    }

    digitTextLength := StrLen(digitText)
    if digitTextLength != 14 && digitTextLength != 17 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Expected 14 digits (no fractional) or 17 digits. Got length: " . digitTextLength)
    }

    year        := SubStr(digitText,  1, 4)
    month       := SubStr(digitText,  5, 2)
    day         := SubStr(digitText,  7, 2)
    hour        := SubStr(digitText,  9, 2)
    minute      := SubStr(digitText, 11, 2)
    second      := SubStr(digitText, 13, 2)
    millisecond := unset

    if digitTextLength = 17 {
        millisecond := SubStr(digitText, 15)
    }

    validation := ValidateDataUsingSpecification(year . "-" . month . "-" . day . " " . hour . ":" . minute . ":" . second, "String", "ISO Date Time")
    if validation != "" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, validation)
    }

    utcTimestamp := unset
    if digitTextLength = 14 {
        utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)
    } else if digitTextLength = 17 {
        utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}", year, month, day, hour, minute, second, millisecond)
    }

    return utcTimestamp
}

ConvertUnixTimeToUtcTimestamp(unixSeconds) {
    static methodName := RegisterMethod("unixSeconds As Integer", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [unixSeconds])

    if unixSeconds < -11644473600 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Unix Seconds predates 1601-01-01 UTC and cannot be represented as FILETIME: " . unixSeconds)
    }

    fileTimeTicks := (unixSeconds + 11644473600) * 10000000

    static fileTimeBuffer := Buffer(8, 0)
    NumPut("UInt64", fileTimeTicks, fileTimeBuffer)

    static systemTimeBuffer := Buffer(16, 0)
    convertedSuccessfully := DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")
    if !convertedSuccessfully {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to convert a file time to system time format. [Kernel32\FileTimeToSystemTime" . ", System Error Code: " . A_LastError . "]")
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
    static methodName := RegisterMethod("utcTimestamp As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [utcTimestamp])
    
    utcTimestampLength := StrLen(utcTimestamp)
    if utcTimestampLength != 19 && utcTimestampLength != 23 && utcTimestampLength != 26 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Expected length of 19, 23 or 26 but got: " . utcTimestampLength)
    }

    year        := SubStr(utcTimestamp,  1, 4)
    month       := SubStr(utcTimestamp,  6, 2)
    day         := SubStr(utcTimestamp,  9, 2)
    hour        := SubStr(utcTimestamp, 12, 2)
    minute      := SubStr(utcTimestamp, 15, 2)
    second      := SubStr(utcTimestamp, 18, 2)
    millisecond := unset

    if utcTimestampLength >= 23 {
        millisecond := SubStr(utcTimestamp, 21, 3)
    }

    validation := ValidateDataUsingSpecification(year . "-" . month . "-" . day . " " . hour . ":" . minute . ":" . second, "String", "ISO Date Time")
    if validation != "" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, validation)
    }

    if utcTimestampLength = 19 {
        utcTimestampCombinedDigits := year . month . day . hour . minute . second
    } else {
        utcTimestampCombinedDigits := year . month . day . hour . minute . second . millisecond
    }
    
    utcTimestampInteger := utcTimestampCombinedDigits + 0

    return utcTimestampInteger
}

ConvertUtcTimestampToLocalTimestampWithTimeZoneKey(utcTimestamp, timeZoneKeyName) {
    static methodName := RegisterMethod("utcTimestamp As String, timeZoneKeyName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [utcTimestamp, timeZoneKeyName])

    parts := StrSplit(utcTimestamp, " ")
    dateParts := StrSplit(parts[1], "-")

    timeAndFractionParts := StrSplit(parts[2], ".")
    timeParts := StrSplit(timeAndFractionParts[1], ":")

    hasMilliseconds := timeAndFractionParts.Length = 2
    milliseconds := 0
    if hasMilliseconds {
        fraction := timeAndFractionParts[2]
        if StrLen(fraction) > 3 {
            fraction := SubStr(fraction, 1, 3)
        } else if StrLen(fraction) < 3 {
            fraction := fraction . SubStr("000", 1, 3 - StrLen(fraction))
        }
        milliseconds := fraction + 0
    }

    utcSystemTime := Buffer(16, 0)
    NumPut("UShort", dateParts[1] + 0, utcSystemTime, 0)
    NumPut("UShort", dateParts[2] + 0, utcSystemTime, 2)
    NumPut("UShort", dateParts[3] + 0, utcSystemTime, 6)
    NumPut("UShort", timeParts[1] + 0, utcSystemTime, 8)
    NumPut("UShort", timeParts[2] + 0, utcSystemTime, 10)
    NumPut("UShort", timeParts[3] + 0, utcSystemTime, 12)
    NumPut("UShort", milliseconds,         utcSystemTime, 14)

    static dynamicTimeZoneInformationBuffer := Buffer(432, 0)
    StrPut(timeZoneKeyName, dynamicTimeZoneInformationBuffer.Ptr + 172, 128, "UTF-16")

    static localSystemTime := Buffer(16, 0)
    convertedUtcToLocalSuccessfully := DllCall("Kernel32\SystemTimeToTzSpecificLocalTimeEx", "Ptr", dynamicTimeZoneInformationBuffer.Ptr, "Ptr", utcSystemTime.Ptr, "Ptr", localSystemTime.Ptr, "Int")
    if !convertedUtcToLocalSuccessfully {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Failed to convert UTC timestamp to local timestamp. [Kernel32\SystemTimeToTzSpecificLocalTimeEx" . ", System Error Code: " . A_LastError . "]")
    }

    year        := NumGet(localSystemTime, 0,  "UShort")
    month       := NumGet(localSystemTime, 2,  "UShort")
    day         := NumGet(localSystemTime, 6,  "UShort")
    hour        := NumGet(localSystemTime, 8,  "UShort")
    minute      := NumGet(localSystemTime, 10, "UShort")
    second      := NumGet(localSystemTime, 12, "UShort")
    millisecond := NumGet(localSystemTime, 14, "UShort")

    localTimeText := ""
    if hasMilliseconds {
        localTimeText := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}", year, month, day, hour, minute, second, millisecond)
    } else {
        localTimeText := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", year, month, day, hour, minute, second)
    }

    return localTimeText
}