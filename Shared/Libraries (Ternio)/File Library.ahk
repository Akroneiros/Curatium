#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Application Library.ahk
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include Image Library.ahk
#Include Logging Library.ahk

CleanOfficeLocksInFolder(directoryPath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("directoryPath As String [Constraint: Directory]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [directoryPath], "Clean Office Locks in Folder (" . directoryPath . ")")

    deletedCount     := 0
    filesInDirectory := GetFilesFromDirectory(directoryPath)

    if filesInDirectory.Length = 0 {
        LogConclusion("Skipped", logConclusionData)
    } else {
        for filePath in filesInDirectory {
            SplitPath(filePath, &fileName)

            if SubStr(fileName, 1, 2) = "~$" {
                try {
                    size := FileGetSize(filePath)

                    if size >= 0 && size <= 8192 {
                        FileDelete(filePath)
                        deletedCount++
                        logConclusionData["Context"] := "Office lock files deleted: " . deletedCount
                    }
                } catch {
                    continue
                }
            }
        }

        LogConclusion("Completed", logConclusionData)
    }
}

ConvertCsvToArrayOfMaps(filePath, delimiter := "|") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String [Constraint: Path], delimiter As String [Optional: |]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath, delimiter], "Convert CSV to Array of Maps (" . ExtractFilename(filePath) . ")")

    fileHash := GetFileHash(filePath, "SHA-256")
    fileText := ReadFileOnHashMatch(filePath, fileHash)

    fileText := StrReplace(StrReplace(fileText, "`r`n", "`n"), "`r", "`n")
    allLines := StrSplit(fileText, "`n")

    if allLines[1] = "" {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Header line is empty.")
    }

    headerNames := StrSplit(allLines[1], delimiter)

    rowsAsMaps := []
    loop allLines.Length - 1 {
        currentLine := allLines[1 + A_Index]

        if RegExMatch(currentLine, "^[ \t]*$") {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Found an empty line on line #" . (A_Index + 1) . ".")
        }

        fieldValues := StrSplit(currentLine, delimiter)
        rowMap := Map()

        loop headerNames.Length {
            headerName := headerNames[A_Index]
            valueText := (A_Index <= fieldValues.Length) ? fieldValues[A_Index] : ""
            rowMap[headerName] := valueText
        }

        rowsAsMaps.Push(rowMap)
    }
    
    LogConclusion("Completed", logConclusionData)
    return rowsAsMaps
}

CopyFileToTarget(filePath, targetDirectory, findValue := "", replaceValue := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String [Constraint: Path], targetDirectory As String [Constraint: Directory], findValue As String [Optional], replaceValue As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath, targetDirectory, findValue, replaceValue], "Copy File to Target (" . ExtractFilename(filePath) . ")")

    sourceFilename := ExtractFilename(filePath)

    if (findValue = "" && replaceValue != "") || (findValue != "" && replaceValue = "") {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Invalid find/replace combo.")
    }

    targetPath := targetDirectory . sourceFilename
    if findValue != "" && replaceValue != "" {
        targetPath := targetDirectory . StrReplace(sourceFilename, findValue, replaceValue)
    }

    if FileExist(targetPath) {
        LogConclusion("Skipped", logConclusionData)
    } else {
        try {
            FileCopy(filePath, targetPath)
        }

        if !FileExist(targetPath) {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Copy did not produce target.")
        }

        LogConclusion("Completed", logConclusionData)
    }
}

DeleteFile(filePath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String [Constraint: Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath], "Delete File (" . ExtractFilename(filePath) . ")")

    try {
        FileDelete(filePath)
    } catch as fileDeleteFailedError {
        LogConclusion("Failed", logConclusionData, fileDeleteFailedError.Line, "File delete failed: " . filePath)
    }

    if FileExist(filePath) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "File still exists after deletion attempt: " . filePath)
    }

    LogConclusion("Completed", logConclusionData)
}

EnsureDirectoryExists(directoryPath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("directoryPath As String [Constraint: Valid Directory]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [directoryPath], "Ensure Directory Exists (" . directoryPath . ")")

    if !DirExist(directoryPath) {
        try {
            DirCreate(directoryPath)
        } catch as directoryError {
            LogConclusion("Failed", logConclusionData, directoryError.Line, directoryError.Message)
        }

        LogConclusion("Completed", logConclusionData)
    } else {
        LogConclusion("Skipped", logConclusionData)
    }
}

MoveFileToDirectory(filePath, directoryPath, overwrite := false) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String [Constraint: Path], directoryPath As String [Constraint: Directory], overwrite As Boolean [Optional: false]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath, directoryPath, overwrite], "Move File to Directory (" . ExtractFilename(filePath) . ")")

    filename   := ExtractFilename(filePath)
    targetPath := directoryPath . filename

    if !DirExist(directoryPath) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Destination directory (" . directoryPath . ") does not exist.")
    }

    if !overwrite {
        if !FileExist(targetPath) || InStr(FileExist(targetPath), "D") {
            try {
                FileMove(filePath, targetPath, overwrite)
            } catch as moveError {
                LogConclusion("Failed", logConclusionData, moveError.Line, moveError.Message)
            }

            LogConclusion("Completed", logConclusionData)
        } else {
            LogConclusion("Skipped", logConclusionData)
        }
    } else {
        if filePath = targetPath {
            LogConclusion("Skipped", logConclusionData)
        } else {
            try {
                FileMove(filePath, targetPath, overwrite)
            } catch as moveError {
                LogConclusion("Failed", logConclusionData, moveError.Line, moveError.Message)
            }

            LogConclusion("Completed", logConclusionData)
        }
    }
}

ReadFileOnHashMatch(filePath, expectedHash) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String [Constraint: Path], expectedHash As String [Constraint: SHA-256]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath, expectedHash], "Read File on Hash Match (" . ExtractFilename(filePath) . ")")

    fileHash := GetFileHash(filePath, "SHA-256")
    if fileHash != expectedHash {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Hash mismatch in " . filePath . ". Expected: " . expectedHash . ". Results: " . fileHash)
    }

    fileBuffer := FileRead(filePath, "RAW")
    totalSize  := fileBuffer.Size
    fileText   := ""
    
    if totalSize = 0 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "File is empty: " . filePath)
    }

    byte1 := NumGet(fileBuffer.Ptr, 0, "UChar")
    byte2 := totalSize > 1 ? NumGet(fileBuffer.Ptr, 1, "UChar") : 0
    byte3 := totalSize > 2 ? NumGet(fileBuffer.Ptr, 2, "UChar") : 0
    byte4 := totalSize > 3 ? NumGet(fileBuffer.Ptr, 3, "UChar") : 0

    if totalSize >= 3 && byte1=0xEF && byte2=0xBB && byte3=0xBF {
        fileText := StrGet(fileBuffer.Ptr + 3, totalSize - 3, "UTF-8")
    } else if totalSize >= 2 && byte1=0xFF && byte2=0xFE {
        fileText := StrGet(fileBuffer.Ptr + 2, (totalSize - 2) // 2, "UTF-16")
    } else if totalSize >= 2 && byte1=0xFE && byte2=0xFF {
        beSize    := totalSize - 2
        swapped   := Buffer(beSize)
        sourcePtr := fileBuffer.Ptr + 2

        loop beSize // 2 {
            offset := (A_Index - 1) * 2
            NumPut("UChar", NumGet(sourcePtr + offset, 1, "UChar"), swapped.Ptr + offset, 0)
            NumPut("UChar", NumGet(sourcePtr + offset, 0, "UChar"), swapped.Ptr + offset, 1)
        }

        fileText := StrGet(swapped.Ptr, beSize // 2, "UTF-16")
    } else if (totalSize >= 4 && ((byte1=0x00 && byte2=0x00 && byte3=0xFE && byte4=0xFF) || (byte1=0xFF && byte2=0xFE && byte3=0x00 && byte4=0x00))) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "UTF-32 encoded text for file " . filePath . " is not supported.")
    } else {
        fileText := StrGet(fileBuffer.Ptr, totalSize, "UTF-8")
    }

    if SubStr(fileText, 1, 1) = Chr(0xFEFF) {
        fileText := SubStr(fileText, 2)
    }

    LogConclusion("Completed", logConclusionData)
    return fileText
}

WriteBase64IntoFileWithHash(base64Text, filePath, expectedHash) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("base64Text As String [Constraint: Base64], filePath As String [Constraint: Valid Path], expectedHash As String [Constraint: SHA-256]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [base64Text, filePath, expectedHash], "Write Base64 into File with Hash" . " (" . ExtractFilename(filePath) . ")")
    
    requiredSizeInBytes   := 0
    decodedByteCount      := 0
    cryptStringBase64Flag := 0x1
    decodedBinaryBuffer   := unset

    needsWrite := true
    if FileExist(filePath) {
        fileHash := GetFileHash(filePath, "SHA-256")

        if !(StrUpper(fileHash) = StrUpper(expectedHash)) {
            DeleteFile(filePath)
        } else {
            needsWrite := false
        }
    }

    if !needsWrite {
        LogConclusion("Skipped", logConclusionData)
    } else {       
        primaryDecodeSucceeded := true
        try {
            sizeProbeSucceeded := DllCall("Crypt32\CryptStringToBinaryW", "WStr", base64Text, "UInt", 0, "UInt", cryptStringBase64Flag, "Ptr", 0, "UInt*", &requiredSizeInBytes, "Ptr", 0, "Ptr", 0, "Int")

            if !sizeProbeSucceeded || requiredSizeInBytes <= 0 {
                logConclusionData["Context"] := "CryptStringToBinaryW size probe failed or returned zero size."
            }

            decodedBinaryBuffer := Buffer(requiredSizeInBytes, 0)
            decodeCallSucceeded := DllCall("Crypt32\CryptStringToBinaryW", "WStr", base64Text, "UInt", 0, "UInt", cryptStringBase64Flag, "Ptr", decodedBinaryBuffer.Ptr, "UInt*", requiredSizeInBytes, "Ptr", 0, "Ptr", 0, "Int")

            if !decodeCallSucceeded {
                logConclusionData["Context"] := "CryptStringToBinaryW decode call failed."
            }

            decodedByteCount := requiredSizeInBytes
            if decodedByteCount <= 0 {
                logConclusionData["Context"] := "Decoded buffer is empty after decode."
            }
        } catch {
            primaryDecodeSucceeded := false
        }

        ; MSXML fallback if the primary attempt fails.
        if !primaryDecodeSucceeded {
            try {
                xmlDocument := ComObject("MSXML2.DOMDocument.6.0")
                base64Element := xmlDocument.createElement("b64")
                base64Element.dataType := "bin.base64"
                base64Element.text := base64Text
                byteSafeArray := base64Element.nodeTypedValue
                decodedByteCount := byteSafeArray.MaxIndex(1) + 1

                if decodedByteCount <= 0 {
                    throw Error("Decoded byte array is empty after MSXML fallback.")
                }
            } catch as base64DecodingError {
                LogConclusion("Failed", logConclusionData, base64DecodingError.Line, base64DecodingError.Message)
            }

            decodedBinaryBuffer := Buffer(decodedByteCount, 0)
            index := 0
            while index < decodedByteCount {
                NumPut("UChar", byteSafeArray[index], decodedBinaryBuffer, index)
                index += 1
            }
        }

        temporaryFilePath := filePath . ".part"
        try {
            fileHandle := FileOpen(temporaryFilePath, "w")
            if !IsObject(fileHandle) {
                throw Error("Failed to open temporary file for writing: " . temporaryFilePath)
            }

            bytesWritten := fileHandle.RawWrite(decodedBinaryBuffer.Ptr, decodedByteCount)
            fileHandle.Close()

            if bytesWritten != decodedByteCount {
                throw Error("Incomplete write. Expected " . decodedByteCount . " bytes but wrote " . bytesWritten . " bytes.")
            }

            FileMove(temporaryFilePath, filePath, 1)
        } catch as fileWriteOrMoveError {
            try {
                FileDelete(temporaryFilePath)
            }

            LogConclusion("Failed", logConclusionData, fileWriteOrMoveError.Line, fileWriteOrMoveError.Message)
        }

        LogConclusion("Completed", logConclusionData)
    }
}

WriteTextToFile(text, filePath, encoding := "UTF-8-BOM", mode := "Overwrite") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static encodingWhitelist := Format('"{1}", "{2}", "{3}"', "UTF-8", "UTF-8-BOM", "UTF-16 LE BOM")
    static modeWhitelist := Format('"{1}", "{2}", "{3}", "{4}"', "Append", "Append Break", "Create", "Overwrite")
    static methodName := RegisterMethod("text As String [Optional], filePath As String [Constraint: Valid Path], encoding As String [Whitelist: " . encodingWhitelist . "], Mode as String [Whitelist: " . modeWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [text, filePath, encoding, mode], "Write Text Into File" . " (" . ExtractFilename(filePath) . ") with Mode: " . mode)

    if mode = "Create" && FileExist(filePath) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "File already exists.")
    }

    if !FileExist(filePath) && (mode = "Append" || mode = "Append Break") {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "File doesn't exist.")
    }

    switch encoding {
        case "UTF-8": encoding := "UTF-8-RAW"
        case "UTF-8-BOM": encoding := "UTF-8"
        case "UTF-16 LE BOM": encoding := "UTF-16"
    }
    
    if mode = "Append Break" {
        text := "`r`n" . text
    }

    switch mode {
        case "Append", "Append Break": mode := "a"
        case "Create", "Overwrite": mode := "w"
    }

    try {
        fileHandle := FileOpen(filePath, mode, encoding)
        fileHandle.Write(text)
        fileHandle.Close()
    } catch as fileWriteError {
        LogConclusion("Failed", logConclusionData, fileWriteError.Line, fileWriteError.Message)
    }

    LogConclusion("Completed", logConclusionData)
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

DetermineWindowsBinaryType(executablePath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("executablePath As String [Constraint: Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [executablePath])

    static SCS_32BIT_BINARY := 0
    static SCS_DOS_BINARY   := 1
    static SCS_WOW_BINARY   := 2
    static SCS_PIF_BINARY   := 3
    static SCS_POSIX_BINARY := 4
    static SCS_OS216_BINARY := 5
    static SCS_64BIT_BINARY := 6

    classificationResult := "N/A"
    binaryType           := 0

    binaryTypeRetrievedSuccessfully := DllCall("Kernel32\GetBinaryTypeW", "Str", executablePath, "UInt*", &binaryType, "Int")
    if binaryTypeRetrievedSuccessfully {
        switch binaryType {
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
            case SCS_OS216_BINARY:
                classificationResult := "OS/2"
        }
    }

    return classificationResult
}

ExtractDirectory(filePath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath])

    SplitPath(filePath, , &directoryPath)

    if directoryPath != "" && SubStr(directoryPath, -1) != "\" {
        directoryPath .= "\"
    }

    return directoryPath
}

ExtractFilename(filePath, removeFileExtension := false) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String, removeFileExtension As Boolean [Optional: false]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath, removeFileExtension])

    SplitPath(filePath, &filenameWithExtension, , , &filenameWithoutExtension)

    filename := filenameWithExtension
    if removeFileExtension {
        filename := filenameWithoutExtension
    }

    return filename
}

ExtractParentDirectory(filePath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath])

    SplitPath(filePath, , &directoryPath)
    SplitPath(directoryPath, , &parentFolderPath)

    if parentFolderPath != "" && SubStr(parentFolderPath, -1) != "\" {
        parentFolderPath .= "\"
    }

    return parentFolderPath
}

FileExistsInDirectory(filename, directoryPath, fileExtension := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filename As String [Constraint: Filename], directoryPath As String [Constraint: Directory], fileExtension As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filename, directoryPath, fileExtension])

    filesInDirectory := GetFilesFromDirectory(directoryPath)
    if filesInDirectory.Length = 0 {
        return ""
    }

    index := filesInDirectory.Length
    while index >= 1 {
        filePath := filesInDirectory[index]
        SplitPath(filePath, , , &loopFileExtension, &nameWithoutExtension)

        if ((fileExtension != "" && loopFileExtension != fileExtension) || !InStr(nameWithoutExtension, filename)) {
            filesInDirectory.RemoveAt(index)
        }

        index -= 1
    }

    if filesInDirectory.Length = 0 {
        return ""
    } else if filesInDirectory.Length = 1  {
        return filesInDirectory[1]
    } else {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Too many files match the filename (" . filename . ") in the directory: " . directoryPath)
    }
}

GetFilesFromDirectory(directoryPath, filterValue := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("directoryPath As String [Constraint: Directory], filterValue As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [directoryPath])

    files := []
    pattern := RTrim(directoryPath, "\/") . "\*"

    if filterValue = "" {
        loop files, pattern, "F" {
            files.Push(A_LoopFileFullPath)
        }
    } else {
        loop files, pattern, "F" {
            if InStr(A_LoopFileFullPath, filterValue) {
                files.Push(A_LoopFileFullPath)
            }
        }
    }

    return files
}

GetFileHash(filePath, algorithm) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static algorithmWhitelist := Format('"{1}", "{2}", "{3}", "{4}", "{5}", "{6}", "{7}"', "MD2", "MD4", "MD5", "SHA-1", "SHA-256", "SHA-384", "SHA-512")
    static methodName := RegisterMethod("filePath as String, algorithm As String [Whitelist: " . algorithmWhitelist . "]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath, algorithm])

    switch algorithm {
        case "MD2": algorithm := "MD2"
        case "MD4": algorithm := "MD4"
        case "MD5": algorithm := "MD5"
        case "SHA-1": algorithm := "SHA1"
        case "SHA-256": algorithm := "SHA256"
        case "SHA-384": algorithm := "SHA384"
        case "SHA-512": algorithm := "SHA512"
    }

    try {
        fileHash := Hash.File(algorithm, filePath)
    } catch as fileHashError {
        LogConclusion("Failed", logConclusionData, fileHashError.Line, fileHashError.Message)
    }

    return fileHash
}

GetFoldersFromDirectory(directoryPath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("directoryPath As String [Constraint: Directory]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [directoryPath])

    folders := []
    pattern := RTrim(directoryPath, "\/") . "\*"

    loop files, pattern, "D" {
        folders.Push(A_LoopFileFullPath . "\")
    }

    return folders
}

GetTextFileLineCount(filePath) {    
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("filePath As String [Constraint: Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [filePath])
 
    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Max Fast Size", 100000000, 10, 1000000000)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]
    
    maxFastSize := settings["Max Fast Size"].Get("Value")

    totalLineCount := 0

    fileSizeInBytes := 0
    try {
        fileSizeInBytes := FileGetSize(filePath)
    } catch as fileSizeError {
        LogConclusion("Failed", logConclusionData, fileSizeError.Line, fileSizeError.Message)
    }

    if fileSizeInBytes != 0 {
        try {
            if fileSizeInBytes < maxFastSize {
                content := FileRead(filePath)

                StrReplace(content, "`n", "", false, &totalLineCount)
                totalLineCount += 1
            } else {
                chunkSize := 4 * 1024 * 1024

                fileReader := FileOpen(filePath, "r")
                if !IsObject(fileReader) {
                    LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to open file.")
                }

                while !fileReader.AtEOF {
                    fileContentBuffer := fileReader.Read(chunkSize)

                    newlineCount := 0
                    StrReplace(fileContentBuffer, "`n", "", false, &newlineCount)
                    totalLineCount += newlineCount
                }
                fileReader.Close()

                totalLineCount += 1
            }
        } catch as fileError {
            LogConclusion("Failed", logConclusionData, fileError.Line, fileError.Message)
        }
    }

    return totalLineCount
}