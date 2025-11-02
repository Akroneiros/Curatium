#Requires AutoHotkey v2.0
#Include File Library.ahk
#Include Logging Library.ahk

AssignFileTimeAsLocalIso(filePath, timeType) {
    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Accessed", "A", "Created", "C", "Modified", "M")
    static methodName := RegisterMethod("AssignFileTimeAsLocalIso(filePath As String [Type: Absolute Path], timeType As String [Whitelist: " . timeTypeWhitelist . "])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Assign File Times As Local ISO", methodName, [filePath, timeType])

    switch StrLower(timeType) {
        case "created", "c":
            offset := 0
        case "accessed", "a":
            offset := 8
        case "modified", "m":
            offset := 16
    }

    fileHandle := DllCall("Kernel32\CreateFileW", "WStr", filePath, "UInt", 0x80000000, "UInt", 0x1, "Ptr", 0, "UInt", 3, "UInt", 0x02000000, "Ptr", 0, "Ptr")

    try {
        if fileHandle = -1 {
            throw Error("Failed to open file for reading: " . filePath)
        }
    } catch as failedToOpenFileForReadingError {
        LogInformationConclusion("Failed", logValuesForConclusion, failedToOpenFileForReadingError)
    }

    fileTimeBuffer := Buffer(24, 0)
    try {
        if !DllCall("Kernel32\GetFileTime", "Ptr", fileHandle, "Ptr", fileTimeBuffer.Ptr, "Ptr", fileTimeBuffer.Ptr + 8, "Ptr", fileTimeBuffer.Ptr + 16, "Int")
        {
            throw Error("GetFileTime failed for: " . filePath)
        }
    } catch as getFileTimeFailedError {
        LogInformationConclusion("Failed", logValuesForConclusion, getFileTimeFailedError)
    }

    utcFileTime := Buffer(8, 0)
    DllCall("RtlMoveMemory", "Ptr", utcFileTime.Ptr, "Ptr", fileTimeBuffer.Ptr + offset, "UPtr", 8)

    systemTime := Buffer(16, 0)
    try {
        if !DllCall("Kernel32\FileTimeToSystemTime", "Ptr", utcFileTime.Ptr, "Ptr", systemTime.Ptr, "Int") {
            throw Error("FileTimeToSystemTime failed")
        }
    } catch as fileTimeToSystemTimeFailedError {
        LogInformationConclusion("Failed", logValuesForConclusion, fileTimeToSystemTimeFailedError)
    }

    localTime := Buffer(16, 0)
    try {
        if !DllCall("Kernel32\SystemTimeToTzSpecificLocalTime", "Ptr", 0, "Ptr", systemTime.Ptr, "Ptr", localTime.Ptr, "Int") {
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
            DllCall("Kernel32\CloseHandle", "Ptr", fileHandle)
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
    static methodName := RegisterMethod("ExtractTrailingDateAsIso(inputValue As String, dateOrder As String [Whitelist: " . dateOrderWhitelist . "])", A_LineFile, A_LineNumber + 1)
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
            if validation != "" {
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
    static methodName := RegisterMethod("PreventSystemGoingIdleUntilRuntime(runtimeDate As String [Type: Raw Date Time], randomizePixelMovement As Boolean [Optional: false])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Prevent System Going Idle Until Runtime (" . FormatTime(runtimeDate, "yyyy-MM-dd HH:mm:ss") . ")", methodName, [runtimeDate, randomizePixelMovement])
    
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

    LogInformationConclusion("Completed", logValuesForConclusion)
}

SetDirectoryTimeFromLocalIsoDateTime(directoryPath, localIsoDateTime, timeType) {
    static timeTypeWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}"', "Accessed", "A", "Created", "C", "Modified", "M")
    static methodName := RegisterMethod("SetDirectoryTimeFromLocalIsoDateTime(directoryPath As String [Type: Directory], localIsoDateTime As String [Type: ISO Date Time], timeType As String [Whitelist: " . timeTypeWhitelist . "])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Set Directory Time From Local ISO Date Time", methodName, [directoryPath, localIsoDateTime, timeType])

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
    try {
        if !DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "Ptr", 0, "Ptr", localSystemTime, "Ptr", utcSystemTime, "Int") {
            throw Error("TzSpecificLocalTimeToSystemTime failed (input may not exist in current time zone)")
        }
    } catch as tzSpecificLocalTimeToSystemTimeError {
        LogInformationConclusion("Failed", logValuesForConclusion, tzSpecificLocalTimeToSystemTimeError)
    }

    utcFileTime := Buffer(8, 0)
    try {
        if !DllCall("Kernel32\SystemTimeToFileTime", "Ptr", utcSystemTime, "Ptr", utcFileTime, "Int") {
            throw Error("SystemTimeToFileTime failed")
        }
    } catch as systemTimeToFileTimeError {
        LogInformationConclusion("Failed", logValuesForConclusion, systemTimeToFileTimeError)
    }

    accessMode := 0x100
    shareMode  := 0x7
    flags      := 0x80 | 0x02000000
    handle     := DllCall("Kernel32\CreateFileW", "WStr", directoryPath, "UInt", accessMode, "UInt", shareMode, "Ptr", 0, "UInt", 3, "UInt", flags, "Ptr", 0, "Ptr")

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
    }

    success := DllCall("Kernel32\SetFileTime", "Ptr", handle, "Ptr", pointerCreation, "Ptr", pointerAccessed, "Ptr", pointerModified, "Int")

    DllCall("Kernel32\CloseHandle", "Ptr", handle)

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
    static methodName := RegisterMethod("SetFileTimeFromLocalIsoDateTime(filePath As String [Type: Absolute Path], localIsoDateTime As String [Type: ISO Date Time], timeType As String [Whitelist: " . timeTypeWhitelist . "])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Set File Time From Local ISO Date Time", methodName, [filePath, localIsoDateTime, timeType])

    if AssignFileTimeAsLocalIso(filePath, timeType) = localIsoDateTime {
        LogInformationConclusion("Skipped", logValuesForConclusion)
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
        try {
            if !DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "Ptr", 0, "Ptr", localSystemTime, "Ptr", utcSystemTime, "Int") {
                throw Error("TzSpecificLocalTimeToSystemTime failed (input may not exist in current time zone)")
            }
        } catch as tzSpecificLocalTimeToSystemTimeError {
            LogInformationConclusion("Failed", logValuesForConclusion, tzSpecificLocalTimeToSystemTimeError)
        }

        utcFileTime := Buffer(8, 0)
        try {
            if !DllCall("Kernel32\SystemTimeToFileTime", "Ptr", utcSystemTime, "Ptr", utcFileTime, "Int") {
                throw Error("SystemTimeToFileTime failed")
            }
        } catch as systemTimeToFileTimeError {
            LogInformationConclusion("Failed", logValuesForConclusion, systemTimeToFileTimeError)
        }

        accessMode := 0x100
        shareMode  := 0x7
        flags      := 0x80 | 0x02000000
        handle     := DllCall("Kernel32\CreateFileW", "WStr", filePath, "UInt", accessMode, "UInt", shareMode, "Ptr", 0, "UInt", 3, "UInt", flags, "Ptr", 0, "Ptr")

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

        success := DllCall("Kernel32\SetFileTime", "Ptr", handle, "Ptr", pointerCreation, "Ptr", pointerAccessed, "Ptr", pointerModified, "Int")

        DllCall("Kernel32\CloseHandle", "Ptr", handle)

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
    static methodName := RegisterMethod("ValidateRuntimeDate(runtimeDate As String [Type: Raw Date Time], minimumStartupInSeconds As Integer)", A_LineFile, A_LineNumber + 1)
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
    static methodName := RegisterMethod("WaitUntilFileIsModifiedToday(filePath As String [Type: Absolute Path])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Wait Until File is Modified Today: " . ExtractFilename(filePath, true), methodName, [filePath])

    dateOfToday := FormatTime(A_Now, "yyyy-MM-dd")
    checkInterval := 4000   ; Check every 4 seconds (in milliseconds)
    mouseInterval := 120000 ; Move mouse every 2 minutes (in milliseconds)
    maxWaitMinutes := 360   ; Maximum wait time = 6 hours
    ; Calculate how many times to loop based on max wait time and check interval.
    ; Example: (360 minutes × 60,000 ms) ÷ 4,000 ms = 5,400 loops (i.e. 6 hours total)
    maxLoops := (maxWaitMinutes * 60000) // checkInterval
    timeSinceLastMouse := 0

    loop maxLoops {
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

; **************************** ;
; Helper Methods               ;
; **************************** ;

ConvertIntegerToUtcTimestamp(integerValue) {
    static methodName := RegisterMethod("ConvertIntegerToUtcTimestamp(integerValue As Integer)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [integerValue])

    digitText := integerValue . ""
    if !RegExMatch(digitText, "^\d+$") {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Input must contain only digits. Got: " . digitText)
    }

    digitTextLength := StrLen(digitText)
    if digitTextLength != 14 && digitTextLength != 17 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Expected 14 digits (no fractional) or 17 digits. Got length: " . digitTextLength)
    }

    year        := SubStr(digitText,  1, 4)
    month       := SubStr(digitText,  5, 2)
    day         := SubStr(digitText,  7, 2)
    hour        := SubStr(digitText,  9, 2)
    minute      := SubStr(digitText, 11, 2)
    second      := SubStr(digitText, 13, 2)
    millisecond := unset

    yearNumber        := year + 0
    monthNumber       := month + 0
    dayNumber         := day + 0
    hourNumber        := hour + 0
    minuteNumber      := minute + 0
    secondNumber      := second + 0
    millisecondNumber := unset

    if digitTextLength = 17 {
        millisecond := SubStr(digitText, 15)
        millisecondNumber := millisecond + 0
    }

    dateValidationResults := ValidateIsoDate(yearNumber, monthNumber, dayNumber, hourNumber, minuteNumber, secondNumber)

    if dateValidationResults != "" {
        LogHelperError(logValuesForConclusion, A_LineNumber, dateValidationResults)
    }

    utcTimestamp := unset
    if digitTextLength = 14 {
        utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}", yearNumber, monthNumber, dayNumber, hourNumber, minuteNumber, secondNumber)
    } else if digitTextLength = 17 {
        utcTimestamp := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:03}", yearNumber, monthNumber, dayNumber, hourNumber, minuteNumber, secondNumber, millisecondNumber)
    }

    return utcTimestamp
}

ConvertUnixTimeToUtcTimestamp(unixSeconds) {
    static methodName := RegisterMethod("ConvertUnixTimeToUtcTimestamp(unixSeconds As Integer)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [unixSeconds])

    if unixSeconds < -11644473600 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "unixSeconds predates 1601-01-01 UTC and cannot be represented as FILETIME: " . unixSeconds)
    }

    ; Unix seconds (1970; negatives allowed) → FILETIME ticks (since 1601, 100 ns): (unixSeconds + 11644473600) * 10000000
    fileTimeTicks := (unixSeconds + 11644473600) * 10000000

    static fileTimeBuffer := Buffer(8, 0)
    NumPut("UInt64", fileTimeTicks, fileTimeBuffer)

    static systemTimeBuffer := Buffer(16, 0)
    convertedSuccessfully := DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")
    if !convertedSuccessfully {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to convert a file time to system time format. [Kernel32\FileTimeToSystemTime" . ", System Error Code: " . A_LastError . "]")
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
    static methodName := RegisterMethod("ConvertUtcTimestampToInteger(utcTimestamp As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [utcTimestamp])
    
    utcTimestampLength := StrLen(utcTimestamp)
    if utcTimestampLength != 19 && utcTimestampLength != 23 && utcTimestampLength != 26 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Expected length of 19, 23 or 26 but got: " . utcTimestampLength)
    }

    year        := SubStr(utcTimestamp,  1, 4)
    month       := SubStr(utcTimestamp,  6, 2)
    day         := SubStr(utcTimestamp,  9, 2)
    hour        := SubStr(utcTimestamp, 12, 2)
    minute      := SubStr(utcTimestamp, 15, 2)
    second      := SubStr(utcTimestamp, 18, 2)
    millisecond := unset

    yearNumber   := year + 0
    monthNumber  := month + 0
    dayNumber    := day + 0
    hourNumber   := hour + 0
    minuteNumber := minute + 0
    secondNumber := second + 0

    if utcTimestampLength >= 23 {
        millisecond := SubStr(utcTimestamp, 21, 3)
    }

    dateValidationResults := ValidateIsoDate(yearNumber, monthNumber, dayNumber, hourNumber, minuteNumber, secondNumber)

    if dateValidationResults != "" {
        LogHelperError(logValuesForConclusion, A_LineNumber, dateValidationResults)
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
    static methodName := RegisterMethod("ConvertUtcTimestampToLocalTimestampWithTimeZoneKey(utcTimestamp As String, timeZoneKeyName As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [utcTimestamp, timeZoneKeyName])

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
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to convert UTC timestamp to local timestamp. [Kernel32\SystemTimeToTzSpecificLocalTimeEx" . ", System Error Code: " . A_LastError . "]")
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

GetQueryPerformanceCounter() {
    static methodName := RegisterMethod("GetQueryPerformanceCounter()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    static queryPerformanceCounterBuffer := Buffer(8, 0)
    queryPerformanceCounterRetrievedSuccessfully := DllCall("Kernel32\QueryPerformanceCounter", "Ptr", queryPerformanceCounterBuffer.Ptr, "Int")
    if !queryPerformanceCounterRetrievedSuccessfully {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to retrieve the current value of the performance counter which is a high resolution time stamp that can be used for time-interval measurements. [Kernel32\QueryPerformanceCounter" . ", System Error Code: " . A_LastError . "]")
    }

    queryPerformanceCounter := NumGet(queryPerformanceCounterBuffer, 0, "Int64")

    return queryPerformanceCounter
}

GetUtcTimestamp() {
    static methodName := RegisterMethod("GetUtcTimestamp()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetUtcTimestampInteger()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

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
    static methodName := RegisterMethod("GetUtcTimestampPrecise()", A_LineFile, A_LineNumber + 1)
    static logValuesForConclusion := LogHelperValidation(methodName)

    static fileTimeBuffer   := Buffer(8, 0)
    static systemTimeBuffer := Buffer(16, 0)

    static usePreciseSystemTimeFunction := unset

    if !IsSet(usePreciseSystemTimeFunction) {
        kernel32ModuleHandle := DllCall("GetModuleHandle", "Str", "Kernel32", "Ptr")
        preciseFunctionAddress := DllCall("GetProcAddress", "Ptr", kernel32ModuleHandle, "AStr", "GetSystemTimePreciseAsFileTime", "Ptr")

        if preciseFunctionAddress != 0 {
            usePreciseSystemTimeFunction := true
        } else {
            usePreciseSystemTimeFunction := false
        }
    }

    if usePreciseSystemTimeFunction {
        DllCall("Kernel32\GetSystemTimePreciseAsFileTime", "Ptr", fileTimeBuffer.Ptr)
        fileTimeTicks := NumGet(fileTimeBuffer, 0, "Int64")
        
        convertedFileTimeToSystemTimeSuccessfully := DllCall("Kernel32\FileTimeToSystemTime", "Ptr", fileTimeBuffer.Ptr, "Ptr", systemTimeBuffer.Ptr, "Int")
        if !convertedFileTimeToSystemTimeSuccessfully {
            LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to convert a file time to system time format. [Kernel32\FileTimeToSystemTime" . ", System Error Code: " . A_LastError . "]")
        }

        year       := NumGet(systemTimeBuffer,  0, "UShort")
        month      := NumGet(systemTimeBuffer,  2, "UShort")
        dayOfMonth := NumGet(systemTimeBuffer,  6, "UShort")
        hour       := NumGet(systemTimeBuffer,  8, "UShort")
        minute     := NumGet(systemTimeBuffer, 10, "UShort")
        second     := NumGet(systemTimeBuffer, 12, "UShort")

        ticksWithinSecond := Mod(fileTimeTicks, 10000000)
        microsecond       := ticksWithinSecond // 10

        utcTimestampPrecise := Format("{:04}-{:02}-{:02} {:02}:{:02}:{:02}.{:06}", year, month, dayOfMonth, hour, minute, second, microsecond)
    } else {
        utcTimestampPrecise := GetUtcTimestamp()
    }

    return utcTimestampPrecise
}

IsValidGregorianDay(year, month, day) {
    static methodName := RegisterMethod("IsValidGregorianDay(year As Integer, month As Integer, day As Integer)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [year, month, day])

    isLeap := (Mod(year, 400) = 0) || (Mod(year, 4) = 0 && Mod(year, 100) != 0)
    daysInMonth := [31, (isLeap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    if day <= daysInMonth[month] {
        return true
    } else {
        return false
    }
}

LocalIsoWithUtcTag(localIsoString) {
    static methodName := RegisterMethod("LocalIsoWithUtcTag(localIsoString As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [localIsoString])

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
    NumPut("UShort", year,   localSystemTime, 0)
    NumPut("UShort", month,  localSystemTime, 2)
    NumPut("UShort", 0,      localSystemTime, 4)
    NumPut("UShort", day,    localSystemTime, 6)
    NumPut("UShort", hour,   localSystemTime, 8)
    NumPut("UShort", minute, localSystemTime, 10)
    NumPut("UShort", second, localSystemTime, 12)
    NumPut("UShort", 0,      localSystemTime, 14)

    utcSystemTime := Buffer(16, 0)
    utcSuccess := DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "Ptr", 0, "Ptr", localSystemTime, "Ptr", utcSystemTime, "Int")

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

ValidateIsoDate(year, month, day, hour := unset, minute := unset, second := unset, checkLocalTime := unset) {
    static methodName := RegisterMethod("ValidateIsoDate(year As Integer, month As Integer, day As Integer, hour As Integer [Optional], minute As Integer [Optional], second As Integer [Optional], checkLocalTime As Boolean [Optional])", A_LineFile, A_LineNumber + 10)
    arrayValidation := [year, month, day]
    timeIsSet := IsSet(hour) && IsSet(minute) && IsSet(second)
    if timeIsSet {
        if !IsSet(checkLocalTime) {
            checkLocalTime := false
        }

        arrayValidation.Push(hour, minute, second, checkLocalTime)
    }
    logValuesForConclusion := LogHelperValidation(methodName, arrayValidation)

    validationResults := ""

    year   := Number(year)
    month  := Number(month)
    day    := Number(day)

    if timeIsSet {
        hour   := Number(hour)
        minute := Number(minute)
        second := Number(second)
    }

    if year < 0 || year > 9999 {
        validationResults := "Invalid ISO 8601 Date: " . year . " (year out of range 0000–9999)."
    } else if month < 1 || month > 12 {
        validationResults := "Invalid ISO 8601 Date: " . month . " (month must be 01–12)."
    } else if day < 1 || day > 31 {
        validationResults := "Invalid ISO 8601 Date: " . day . " (day must be 01–31)."
    } else if !IsValidGregorianDay(year, month, day) {
        validationResults := "Invalid ISO 8601 Date: " . day . " (day out of range for month)."
    }
    
    if validationResults = "" && timeIsSet {
        if hour < 0 || hour > 23 {
            validationResults := "Invalid ISO 8601 Date Time: " . hour . " (hour must be 00–23)."
        } else if minute < 0 || minute > 59 {
            validationResults := "Invalid ISO 8601 Date Time: " . minute . " (minute must be 00–59)."
        } else if second < 0 || second > 59 {
            validationResults := "Invalid ISO 8601 Date Time: " . second . " (second must be 00–59)."
        } else {
            if checkLocalTime {
                static localSystemTimeBuffer := Buffer(16, 0)
                NumPut("UShort", year,   localSystemTimeBuffer, 0)
                NumPut("UShort", month,  localSystemTimeBuffer, 2)
                NumPut("UShort", 0,      localSystemTimeBuffer, 4)
                NumPut("UShort", day,    localSystemTimeBuffer, 6)
                NumPut("UShort", hour,   localSystemTimeBuffer, 8)
                NumPut("UShort", minute, localSystemTimeBuffer, 10)
                NumPut("UShort", second, localSystemTimeBuffer, 12)
                NumPut("UShort", 0,      localSystemTimeBuffer, 14)

                static utcSystemTimeBuffer := Buffer(16, 0)
                if !DllCall("Kernel32\TzSpecificLocalTimeToSystemTime", "Ptr", 0, "Ptr", localSystemTimeBuffer, "Ptr", utcSystemTimeBuffer, "Int") {
                    validationResults := Format("Invalid ISO 8601 Date Time: {:04}-{:02}-{:02} {:02}:{:02}:{:02} (nonexistent local time, DST gap or system restriction).", year, month, day, hour, minute, second)
                } else {
                    systemTimeLocal := Buffer(16, 0)
                    DllCall("Kernel32\SystemTimeToTzSpecificLocalTime", "Ptr", 0, "Ptr", utcSystemTimeBuffer, "Ptr", systemTimeLocal, "Int")
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
    }

    return validationResults
}