#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Logging Library.ahk

CleanOfficeLocksInFolder(directoryPath) {
    static methodName := RegisterMethod("directoryPath As String [Constraint: Directory]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [directoryPath], "Clean Office Locks in Folder (" . directoryPath . ")")

    deletedCount     := 0
    filesInDirectory := GetFilesFromDirectory(directoryPath)

    if filesInDirectory.Length = 0 {
        LogConclusion("Skipped", logValuesForConclusion)
    } else {
        for index, filePath in filesInDirectory {
            SplitPath(filePath, &fileName)

            if SubStr(fileName, 1, 2) = "~$" {
                try {
                    size := FileGetSize(filePath)

                    if size >= 0 && size <= 8192 {
                        FileDelete(filePath)
                        deletedCount++
                        logValuesForConclusion["Context"] := "Office lock files deleted: " . deletedCount
                    }
                } catch {
                    continue
                }
            }
        }

        LogConclusion("Completed", logValuesForConclusion)
    }
}

ConvertCsvToArrayOfMaps(filePath, delimiter := "|") {
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Path], delimiter As String [Optional: |]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath], "Convert CSV to Array of Maps (" . ExtractFilename(filePath) . ")")

    hashValue := Hash.File("SHA256", filePath)
    fileText  := ReadFileOnHashMatch(filePath, hashValue)

    fileText := StrReplace(StrReplace(fileText, "`r`n", "`n"), "`r", "`n")
    allLines := StrSplit(fileText, "`n")

    if allLines[1] = "" {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Header line is empty.")
    }

    headerNames := StrSplit(allLines[1], delimiter)

    rowsAsMaps := []
    loop allLines.Length - 1 {
        currentLine := allLines[1 + A_Index]

        if RegExMatch(currentLine, "^[ \t]*$") {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Found an empty line on line #" . (A_Index + 1) . ".")
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
    
    LogConclusion("Completed", logValuesForConclusion)
    return rowsAsMaps
}

CopyFileToTarget(filePath, targetDirectory, findValue := "", replaceValue := "") {
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Path], targetDirectory As String [Constraint: Directory], findValue As String [Optional], replaceValue As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath, targetDirectory, findValue, replaceValue], "Copy File to Target (" . ExtractFilename(filePath) . ")")

    if ((findValue = "" && replaceValue != "") || (findValue != "" && replaceValue = "")) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Invalid find/replace combo.")
    }

    SplitPath(filePath, &sourceFilename, &sourceDirectoryPath, &sourceExtension, &sourceFilenameWithoutExtension)
    targetPath := ""

    if findValue = "" && replaceValue = "" {
        targetPath := targetDirectory . sourceFilename
    } else {
        if RegExMatch(replaceValue, '[<>:"/\\|?*]') {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "replaceValue contains invalid characters.")
        }

        targetPath := targetDirectory . StrReplace(sourceFilename, findValue, replaceValue)
    }

    if FileExist(targetPath) {
        LogConclusion("Skipped", logValuesForConclusion)
    } else {
        fileTimeCreated := AssignFileTimeAsLocalIso(filePath, "Created")
        FileCopy(filePath, targetPath)

        if !FileExist(targetPath) {
            LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Copy did not produce target.")
        }

        SetFileTimeFromLocalIsoDateTime(targetPath, fileTimeCreated, "Created")

        LogConclusion("Completed", logValuesForConclusion)
    }
}

DeleteFile(filePath) {
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath], "Delete File (" . ExtractFilename(filePath) . ")")

    try {
        FileDelete(filePath)
    } catch as fileDeleteFailedError {
        LogConclusion("Failed", logValuesForConclusion, fileDeleteFailedError.Line, "File delete failed: " . filePath)
    }

    if FileExist(filePath) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "File still exists after deletion attempt: " . filePath)
    }

    LogConclusion("Completed", logValuesForConclusion)
}

EnsureDirectoryExists(directoryPath) {
    static methodName := RegisterMethod("directoryPath As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [directoryPath], "Ensure Directory Exists (" . directoryPath . ")")

    if !DirExist(directoryPath) {
        try {
            DirCreate(directoryPath)

            if !DirExist(directoryPath) {
                throw Error("Failed to create directory: " . directoryPath)
            }
        } catch as directoryError {
            LogConclusion("Failed", logValuesForConclusion, directoryError.Line, directoryError.Message)
        }

        LogConclusion("Completed", logValuesForConclusion)
    } else {
        LogConclusion("Skipped", logValuesForConclusion)
    }
}

MoveFileToDirectory(filePath, directoryPath, overwrite := false) {
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Path], directoryPath As String [Constraint: Directory], overwrite As Boolean [Optional: false]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath, directoryPath, overwrite], "Move File to Directory (" . ExtractFilename(filePath) . ")")

    filename   := ExtractFilename(filePath)
    targetPath := directoryPath . filename

    if !DirExist(directoryPath) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Destination directory (" . directoryPath . ") does not exist.")
    }

    if !overwrite {
        if !FileExist(targetPath) || InStr(FileExist(targetPath), "D") {
            try {
                FileMove(filePath, targetPath, overwrite)
            } catch as moveError {
                LogConclusion("Failed", logValuesForConclusion, moveError.Line, moveError.Message)
            }

            LogConclusion("Completed", logValuesForConclusion)
        } else {
            LogConclusion("Skipped", logValuesForConclusion)
        }
    } else {
        if filePath = targetPath {
            LogConclusion("Skipped", logValuesForConclusion)
        } else {
            try {
                FileMove(filePath, targetPath, overwrite)
            } catch as moveError {
                LogConclusion("Failed", logValuesForConclusion, moveError.Line, moveError.Message)
            }

            LogConclusion("Completed", logValuesForConclusion)
        }
    }
}

ReadFileOnHashMatch(filePath, expectedHash) {
    static methodName := RegisterMethod("filePath As String [Constraint: Absolute Path], expectedHash As String [Constraint: SHA-256]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath, expectedHash], "Read File on Hash Match (" . ExtractFilename(filePath) . ")")

    if !FileExist(filePath) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Missing file: " . filePath)
    }

    if StrLen(expectedHash) != 64 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Invalid SHA-256 hash length (" . StrLen(expectedHash) . " characters): " . filePath)
    }

    try {
        fileHash := Hash.File("SHA256", filePath)
        if fileHash != expectedHash {
            throw Error("Hash mismatch in " . filePath . "`n`nExpected: " . expectedHash . "`nResults: " . fileHash)
        }
    } catch as hashMismatchError {
        LogConclusion("Failed", logValuesForConclusion, hashMismatchError.Line, hashMismatchError.Message)
    }

    fileBuffer := FileRead(filePath, "RAW")
    totalSize  := fileBuffer.Size
    fileText   := ""
    
    if totalSize = 0 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "File is empty: " . filePath)
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
        beSize := totalSize - 2
        swapped := Buffer(beSize)
        sourcePtr := fileBuffer.Ptr + 2

        loop beSize // 2 {
            offset := (A_Index - 1) * 2
            NumPut("UChar", NumGet(sourcePtr + offset, 1, "UChar"), swapped.Ptr + offset, 0)
            NumPut("UChar", NumGet(sourcePtr + offset, 0, "UChar"), swapped.Ptr + offset, 1)
        }
        fileText := StrGet(swapped.Ptr, beSize // 2, "UTF-16")
    } else if (totalSize >= 4 && ((byte1=0x00 && byte2=0x00 && byte3=0xFE && byte4=0xFF) || (byte1=0xFF && byte2=0xFE && byte3=0x00 && byte4=0x00))) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "UTF-32 encoded text for file " . filePath . " is not supported.")
    } else {
        fileText := StrGet(fileBuffer.Ptr, totalSize, "UTF-8")
    }

    if SubStr(fileText, 1, 1) = Chr(0xFEFF) {
        fileText := SubStr(fileText, 2)
    }

    LogConclusion("Completed", logValuesForConclusion)
    return fileText
}

WriteBase64IntoFileWithHash(base64Text, filePath, expectedHash) {
    static methodName := RegisterMethod("base64Text As String [Constraint: Base64], filePath As String [Constraint: Absolute Save Path], expectedHash As String [Constraint: SHA-256]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [base64Text, filePath, expectedHash], "Write Base64 into File with Hash" . " (" . ExtractFilename(filePath) . ")")
    
    requiredSizeInBytes := 0
    decodedByteCount := 0
    cryptStringBase64Flag := 0x1
    decodedBinaryBuffer := unset

    needsWrite := true
    if FileExist(filePath) {
        fileHash := Hash.File("SHA256", filePath)

        if !(StrUpper(fileHash) = StrUpper(expectedHash)) {
            DeleteFile(filePath)
        } else {
            needsWrite := false
        }
    }

    if !needsWrite {
        LogConclusion("Skipped", logValuesForConclusion)
    } else {       
        primaryDecodeSucceeded := true
        try {
            sizeProbeSucceeded := DllCall("Crypt32\CryptStringToBinaryW", "WStr", base64Text, "UInt", 0, "UInt", cryptStringBase64Flag, "Ptr", 0, "UInt*", &requiredSizeInBytes, "Ptr", 0, "Ptr", 0, "Int")

            if !sizeProbeSucceeded || requiredSizeInBytes <= 0 {
                logValuesForConclusion["Context"] := "CryptStringToBinaryW size probe failed or returned zero size."
            }

            decodedBinaryBuffer := Buffer(requiredSizeInBytes, 0)
            decodeCallSucceeded := DllCall("Crypt32\CryptStringToBinaryW", "WStr", base64Text, "UInt", 0, "UInt", cryptStringBase64Flag, "Ptr", decodedBinaryBuffer.Ptr, "UInt*", requiredSizeInBytes, "Ptr", 0, "Ptr", 0, "Int")

            if !decodeCallSucceeded {
                logValuesForConclusion["Context"] := "CryptStringToBinaryW decode call failed."
            }

            decodedByteCount := requiredSizeInBytes
            if decodedByteCount <= 0 {
                logValuesForConclusion["Context"] := "Decoded buffer is empty after decode."
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
                LogConclusion("Failed", logValuesForConclusion, base64DecodingError.Line, base64DecodingError.Message)
            }

            decodedBinaryBuffer := Buffer(decodedByteCount, 0)
            index := 0
            while index < decodedByteCount {
                NumPut("UChar", byteSafeArray[index], decodedBinaryBuffer, index)
                index += 1
            }
        }

        temporaryFilePath := filePath ".part"
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

            LogConclusion("Failed", logValuesForConclusion, fileWriteOrMoveError.Line, fileWriteOrMoveError.Message)
        }

        LogConclusion("Completed", logValuesForConclusion)
    }
}

WriteTextIntoFile(text, filePath, encoding := "UTF-8-BOM", overwrite := true) {
    static encodingWhitelist := Format('"{1}", "{2}", "{3}"', "UTF-8", "UTF-8-BOM", "UTF-16 LE BOM")
    static methodName := RegisterMethod("text As String [Constraint: Summary], filePath As String [Constraint: Absolute Save Path], encoding As String [Whitelist: " . encodingWhitelist . "], overwrite as Boolean [Optional: true]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [text, filePath, encoding, overwrite], "Write Text Into File" . " (" . ExtractFilename(filePath) . ")")

    if overwrite = false && FileExist(filePath) {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "File already exists and overwrite parameter is set to false.")
    }

    switch encoding {
        case "UTF-8": encoding := "UTF-8-RAW"
        case "UTF-8-BOM": encoding := "UTF-8"
        case "UTF-16 LE BOM": encoding := "UTF-16"
    }
    
    fileHandle := unset
    try {
        fileHandle := FileOpen(filePath, "w", encoding)
        fileHandle.Write(text)
        fileHandle.Close()
    } catch as fileWriteError {
        LogConclusion("Failed", logValuesForConclusion, fileWriteError.Line, fileWriteError.Message)
    }    

    LogConclusion("Completed", logValuesForConclusion)
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

ExtractDirectory(filePath) {
    static methodName := RegisterMethod("filePath As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath])

    SplitPath(filePath, , &directoryPath)

    if directoryPath != "" && SubStr(directoryPath, -1) != "\" {
        directoryPath .= "\"
    }

    return directoryPath
}

ExtractFilename(filePath, removeFileExtension := false) {
    static methodName := RegisterMethod("filePath As String, removeFileExtension As Boolean [Optional: false]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath, removeFileExtension])

    SplitPath(filePath, &filenameWithExtension, , , &filenameWithoutExtension)

    filename := filenameWithExtension
    if removeFileExtension {
        filename := filenameWithoutExtension
    }

    return filename
}

ExtractParentDirectory(filePath) {
    static methodName := RegisterMethod("filePath As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filePath])

    SplitPath(filePath, , &directoryPath)
    SplitPath(directoryPath, , &parentFolderPath)

    if parentFolderPath != "" && SubStr(parentFolderPath, -1) != "\" {
        parentFolderPath .= "\"
    }

    return parentFolderPath
}

FileExistsInDirectory(filename, directoryPath, fileExtension := "") {
    static methodName := RegisterMethod("filename As String [Constraint: Locator], directoryPath As String [Constraint: Directory], fileExtension As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [filename, directoryPath, fileExtension])

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
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Too many files match the filename (" . filename . ") in the directory: " . directoryPath)
    }
}

GetFilesFromDirectory(directoryPath, filterValue := "") {
    static methodName := RegisterMethod("directoryPath As String [Constraint: Directory], filterValue As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [directoryPath])

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

GetFoldersFromDirectory(directoryPath, emptyDirectoryAllowed := false) {
    static methodName := RegisterMethod("directoryPath As String [Constraint: Directory], emptyDirectoryAllowed As Boolean [Optional: false]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogBeginning(methodName, [directoryPath, emptyDirectoryAllowed])

    folders := []
    pattern := RTrim(directoryPath, "\/") . "\*"

    loop files, pattern, "D" {
        folders.Push(A_LoopFileFullPath . "\")
    }

    if !emptyDirectoryAllowed && folders.Length = 0 {
        LogConclusion("Failed", logValuesForConclusion, A_LineNumber, "Directory exists but contains no folders: " . directoryPath)
    }

    return folders
}