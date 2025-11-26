#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Logging Library.ahk

CleanOfficeLocksInFolder(directoryPath) {
    static methodName := RegisterMethod("CleanOfficeLocksInFolder(directoryPath As String [Constraint: Directory])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Clean Office Locks in Folder (" . directoryPath . ")", methodName, [directoryPath])

    deletedCount     := 0
    filesInDirectory := GetFilesFromDirectory(directoryPath, true)

    if filesInDirectory.Length = 0 {
        LogInformationConclusion("Skipped", logValuesForConclusion)
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

        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

ConvertCsvToArrayOfMaps(filePath, delimiter := "|") {
    static methodName := RegisterMethod("ConvertCsvToArrayOfMaps(filePath As String [Constraint: Absolute Path], delimiter As String [Optional: |])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Convert CSV to Array of Maps (" . ExtractFilename(filePath) . ")", methodName, [filePath])

    hashValue := Hash.File("SHA256", filePath)
    fileText  := ReadFileOnHashMatch(filePath, hashValue)

    fileText := StrReplace(StrReplace(fileText, "`r`n", "`n"), "`r", "`n")
    allLines := StrSplit(fileText, "`n")

    try {
        if allLines[1] = "" {
            throw Error("Header line is empty.")
        }
    } catch as headerLineEmptyError {
        LogInformationConclusion("Failed", logValuesForConclusion, headerLineEmptyError)
    }

    headerNames := StrSplit(allLines[1], delimiter)

    rowsAsMaps := []
    loop allLines.Length - 1 {
        currentLine := allLines[1 + A_Index]

        try { 
            if RegExMatch(currentLine, "^[ \t]*$") {
                throw Error("Found an empty line on line #" . (A_Index + 1) . ".")
            }
        } catch as emptyLineError {
            LogInformationConclusion("Failed", logValuesForConclusion, emptyLineError)
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
    
    LogInformationConclusion("Completed", logValuesForConclusion)
    return rowsAsMaps
}

CopyFileToTarget(filePath, targetDirectory, findValue := "", replaceValue := "") {
    static methodName := RegisterMethod("CopyFileToTarget(filePath As String [Constraint: Absolute Path], targetDirectory As String [Constraint: Directory], findValue As String [Optional], replaceValue As String [Optional])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Copy File to Target (" . ExtractFilename(filePath) . ")", methodName, [filePath, targetDirectory, findValue, replaceValue])

    try {
        if ((findValue = "" && replaceValue != "") || (findValue != "" && replaceValue = "")) {
            throw Error("Invalid find/replace combo.")
        }
    } catch as invalidArgumentCombinationError {
        LogInformationConclusion("Failed", logValuesForConclusion, invalidArgumentCombinationError)
    }

    SplitPath(filePath, &sourceFilename, &sourceDirectoryPath, &sourceExtension, &sourceFilenameWithoutExtension)
    targetPath := ""

    if findValue = "" && replaceValue = "" {
        targetPath := targetDirectory . sourceFilename
    } else {
        try {
            if RegExMatch(replaceValue, '[<>:"/\\|?*]') {
                throw Error("replaceValue contains invalid characters.")
            }
        } catch as invalidCharactersError {
            LogInformationConclusion("Failed", logValuesForConclusion, invalidCharactersError)
        }

        targetPath := targetDirectory . StrReplace(sourceFilename, findValue, replaceValue)
    }

    if FileExist(targetPath) {
        LogInformationConclusion("Skipped", logValuesForConclusion)
    } else {
        fileTimeCreated := AssignFileTimeAsLocalIso(filePath, "Created")
        FileCopy(filePath, targetPath)

        try {
            if !FileExist(targetPath) {
                throw Error("Copy did not produce target.")
            }
        } catch as fileNotCopiedCorrectlyError {
            LogInformationConclusion("Failed", logValuesForConclusion, fileNotCopiedCorrectlyError)
        }

        SetFileTimeFromLocalIsoDateTime(targetPath, fileTimeCreated, "Created")

        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

DeleteFile(filePath) {
    static methodName := RegisterMethod("DeleteFile(filePath As String [Constraint: Absolute Path])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Delete File (" . ExtractFilename(filePath) . ")", methodName, [filePath])

    try {
        FileDelete(filePath)
    } catch as fileDeleteFailedError {
        fileDeleteFailedError.Message := "File delete failed: " . filePath
        LogInformationConclusion("Failed", logValuesForConclusion, fileDeleteFailedError)
    }

    try {
        if FileExist(filePath) {
            throw Error("File still exists after deletion attempt: " . filePath)
        }
    } catch as fileStillExistsError {
        LogInformationConclusion("Failed", logValuesForConclusion, fileStillExistsError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

EnsureDirectoryExists(directoryPath) {
    static methodName := RegisterMethod("EnsureDirectoryExists(directoryPath As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Ensure Directory Exists (" . directoryPath . ")", methodName, [directoryPath])

    if !DirExist(directoryPath) {
        try {
            DirCreate(directoryPath)

            if !DirExist(directoryPath) {
                throw Error("Failed to create directory: " . directoryPath)
            }
        } catch as directoryError {
            LogInformationConclusion("Failed", logValuesForConclusion, directoryError)
        }

        LogInformationConclusion("Completed", logValuesForConclusion)
    } else {
        LogInformationConclusion("Skipped", logValuesForConclusion)
    }
}

FileExistsInDirectory(filename, directoryPath, fileExtension := "") {
    static methodName := RegisterMethod("FileExistsInDirectory(filename As String [Constraint: Locator], directoryPath As String [Constraint: Directory], fileExtension As String [Optional])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("File Exists in Directory (" . filename . ")", methodName, [filename, directoryPath, fileExtension])

    filesInDirectory := GetFilesFromDirectory(directoryPath, true)

    if filesInDirectory.Length = 0 {
        LogInformationConclusion("Completed", logValuesForConclusion)
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
        LogInformationConclusion("Completed", logValuesForConclusion)
        return ""
    } else if filesInDirectory.Length = 1  {
        LogInformationConclusion("Completed", logValuesForConclusion)
        return filesInDirectory[1]
    } else {
        try {
            throw Error("Too many files match the filename (" . filename . ") in the directory: " . directoryPath)
        } catch as tooManyMatchesError {
            LogInformationConclusion("Failed", logValuesForConclusion, tooManyMatchesError)
        }
    }
}

MoveFileToDirectory(filePath, directoryPath, overwrite := false) {
    static methodName := RegisterMethod("MoveFileToDirectory(filePath As String [Constraint: Absolute Path], directoryPath As String [Constraint: Directory], overwrite As Boolean [Optional: false])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Move File to Directory (" . ExtractFilename(filePath) . ")", methodName, [filePath, directoryPath, overwrite])

    filename   := ExtractFilename(filePath)
    targetPath := directoryPath . filename

    try {
        if !DirExist(directoryPath) {
            Throw Error("Destination directory (" . directoryPath . ") does not exist.")
        }
    } catch as destinationDirectoryDoesNotExistError {
        LogInformationConclusion("Failed", logValuesForConclusion, destinationDirectoryDoesNotExistError)
    }

    if !overwrite {
        if !FileExist(targetPath) || InStr(FileExist(targetPath), "D") {
            try {
                FileMove(filePath, targetPath, overwrite)
            } catch as moveError {
                LogInformationConclusion("Failed", logValuesForConclusion, moveError)
            }

            LogInformationConclusion("Completed", logValuesForConclusion)
        } else {
            LogInformationConclusion("Skipped", logValuesForConclusion)
        }
    } else {
        if filePath = targetPath {
            LogInformationConclusion("Skipped", logValuesForConclusion)
        } else {
            try {
                FileMove(filePath, targetPath, overwrite)
            } catch as moveError {
                LogInformationConclusion("Failed", logValuesForConclusion, moveError)
            }

            LogInformationConclusion("Completed", logValuesForConclusion)
        }
    }
}

ReadFileOnHashMatch(filePath, expectedHash) {
    static methodName := RegisterMethod("ReadFileOnHashMatch(filePath As String [Constraint: Absolute Path], expectedHash As String [Constraint: SHA-256])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Read File on Hash Match (" . ExtractFilename(filePath) . ")", methodName, [filePath, expectedHash])

    try {
        if !FileExist(filePath) {
            throw Error("Missing file: " . filePath)
        }
    } catch as fileNotFoundError {
        LogInformationConclusion("Failed", logValuesForConclusion, fileNotFoundError)
    }

    try {
        if StrLen(expectedHash) != 64 {
            throw Error("Invalid SHA-256 hash length (" . StrLen(expectedHash) . " characters): " . filePath)
        }
    } catch as invalidHashError {
        LogInformationConclusion("Failed", logValuesForConclusion, invalidHashError)
    }

    try {
        fileHash := Hash.File("SHA256", filePath)
        if fileHash != expectedHash {
            throw Error("Hash mismatch in " . filePath . "`n`nExpected: " . expectedHash . "`nResults: " . fileHash)
        }
    } catch as hashMismatchError {
        LogInformationConclusion("Failed", logValuesForConclusion, hashMismatchError)
    }

    fileBuffer := FileRead(filePath, "RAW")
    totalSize  := fileBuffer.Size
    fileText   := ""
    
    try {
        if totalSize = 0 {
            throw Error("File is empty: " . filePath)
        }
    } catch as emptyFileError {
        LogInformationConclusion("Failed", logValuesForConclusion, emptyFileError)
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
        try {
            throw Error("UTF-32 encoded text for file " . filePath . " is not supported.")
        } catch as utf32EncodingError {
            LogInformationConclusion("Failed", logValuesForConclusion, utf32EncodingError)
        }
    } else {
        fileText := StrGet(fileBuffer.Ptr, totalSize, "UTF-8")
    }

    if SubStr(fileText, 1, 1) = Chr(0xFEFF) {
        fileText := SubStr(fileText, 2)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
    return fileText
}

WriteBase64IntoFileWithHash(base64Text, filePath, expectedHash) {
    static methodName := RegisterMethod("WriteBase64IntoFileWithHash(base64Text As String [Constraint: Base64], filePath As String [Constraint: Absolute Save Path], expectedHash As String [Constraint: SHA-256])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Write Base64 into File with Hash" . " (" . ExtractFilename(filePath) . ")", methodName, [base64Text, filePath, expectedHash])
    
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
        LogInformationConclusion("Skipped", logValuesForConclusion)
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
                LogInformationConclusion("Failed", logValuesForConclusion, base64DecodingError)
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

            LogInformationConclusion("Failed", logValuesForConclusion, fileWriteOrMoveError)
        }
    }
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

ExtractDirectory(filePath) {
    static methodName := RegisterMethod("ExtractDirectory(filePath As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [filePath])

    SplitPath(filePath, , &directoryPath)

    if directoryPath != "" && SubStr(directoryPath, -1) != "\" {
        directoryPath .= "\"
    }

    return directoryPath
}

ExtractFilename(filePath, removeFileExtension := false) {
    static methodName := RegisterMethod("ExtractFilename(filePath As String, removeFileExtension As Boolean [Optional: false])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [filePath, removeFileExtension])

    SplitPath(filePath, &filenameWithExtension, , , &filenameWithoutExtension)

    filename := filenameWithExtension
    if removeFileExtension {
        filename := filenameWithoutExtension
    }

    return filename
}

ExtractParentDirectory(filePath) {
    static methodName := RegisterMethod("ExtractParentDirectory(filePath As String)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [filePath])

    SplitPath(filePath, , &directoryPath)
    SplitPath(directoryPath, , &parentFolderPath)

    if parentFolderPath != "" && SubStr(parentFolderPath, -1) != "\" {
        parentFolderPath .= "\"
    }

    return parentFolderPath
}

GetFilesFromDirectory(directoryPath, emptyDirectoryAllowed := false) {
    static methodName := RegisterMethod("GetFilesFromDirectory(directoryPath As String [Constraint: Directory], emptyDirectoryAllowed As Boolean [Optional: false])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [directoryPath, emptyDirectoryAllowed])

    files := []
    pattern := RTrim(directoryPath, "\/") . "\*"

    loop files, pattern, "F" {
        files.Push(A_LoopFileFullPath)
    }

    if !emptyDirectoryAllowed && files.Length = 0 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Directory exists but contains no files: " directoryPath)
    }

    return files
}

GetFoldersFromDirectory(directoryPath, emptyDirectoryAllowed := false) {
    static methodName := RegisterMethod("GetFoldersFromDirectory(directoryPath As String [Constraint: Directory], emptyDirectoryAllowed As Boolean [Optional: false])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [directoryPath, emptyDirectoryAllowed])

    folders := []
    pattern := RTrim(directoryPath, "\/") . "\*"

    loop files, pattern, "D" {
        folders.Push(A_LoopFileFullPath . "\")
    }

    if !emptyDirectoryAllowed && folders.Length = 0 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Directory exists but contains no folders: " . directoryPath)
    }

    return folders
}