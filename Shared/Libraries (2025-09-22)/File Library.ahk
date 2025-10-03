#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Logging Library.ahk

CleanOfficeLocksInFolder(directoryPath) {
    static methodName := RegisterMethod("CleanOfficeLocksInFolder(directoryPath As String [Type: Directory])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Clean Office Locks in Folder (" . directoryPath . ")", methodName, [directoryPath])

    deletedCount     := 0
    filesInDirectory := GetFileListFromDirectory(directoryPath, true)

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
    static methodName := RegisterMethod("ConvertCsvToObject(filePath As String [Type: Absolute Path], delimiter As String [Optional: |])", A_LineFile, A_LineNumber + 1)
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

    headerNames := StrSplit(Trim(allLines[1], " `t"), delimiter)

    rowsAsMaps := []
    Loop allLines.Length - 1 {
        currentLine := Trim(allLines[1 + A_Index], " `t")

        try { 
            if currentLine = "" {
                throw Error("Found an empty line on line #" . A_Index + 1 . ".")
            }
        } catch as emptyLineError {
            LogInformationConclusion("Failed", logValuesForConclusion, emptyLineError)
        }

        fieldValues := StrSplit(currentLine, delimiter)
        rowMap := Map()

        Loop headerNames.Length {
            headerName := Trim(headerNames[A_Index], " `t")
            valueText := (A_Index <= fieldValues.Length) ? Trim(fieldValues[A_Index], " `t") : ""
            rowMap[headerName] := valueText
        }

        rowsAsMaps.Push(rowMap)
    }
    
    LogInformationConclusion("Completed", logValuesForConclusion)
    return rowsAsMaps
}

CopyFileToTarget(filePath, targetDirectory, findValue := "", replaceValue := "") {
    static methodName := RegisterMethod("CopyFileToTarget(filePath As String [Type: Absolute Path], targetDirectory As String [Type: Directory], findValue As String [Optional], replaceValue As String [Optional])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Copy File to Target (" . ExtractFilename(filePath) . ")", methodName, [filePath, targetDirectory, findValue, replaceValue])

    try {
        if ((findValue = "" && replaceValue !== "") || (findValue !== "" && replaceValue = "")) {
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
    static methodName := RegisterMethod("DeleteFile(filePath As String [Type: Absolute Path])", A_LineFile, A_LineNumber + 1)
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
    static methodName := RegisterMethod("EnsureDirectoryExists(directoryPath As String [Type: Directory])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Ensure Directory Exists (" . directoryPath . ")", methodName, [directoryPath])

    try {
        if !DirExist(directoryPath) {
            DirCreate(directoryPath)
        }

        if !DirExist(directoryPath) {
            throw Error("Failed to create directory: " directoryPath)
        }
    } catch as directoryError {
        LogInformationConclusion("Failed", logValuesForConclusion, directoryError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

FileExistsInDirectory(filename, directoryPath, fileExtension := "") {
    static methodName := RegisterMethod("FileExistsInDirectory(filename As String [Type: Search], directoryPath As String [Type: Directory], fileExtension As String [Optional])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("File Exists in Directory (" . filename . ")", methodName, [filename, directoryPath, fileExtension])

    filesInDirectory := GetFileListFromDirectory(directoryPath, true)

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

GetFileListFromDirectory(directoryPath, emptyDirectoryAllowed := false) {
    static methodName := RegisterMethod("GetFileListFromDirectory(directoryPath As String [Type: Directory], emptyDirectoryAllowed As Boolean [Optional: false])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Get File List from Directory (" . directoryPath . ")", methodName, [directoryPath, emptyDirectoryAllowed])

    files := []
    pattern := RTrim(directoryPath, "\/") . "\*"

    Loop Files, pattern, "F"
    {
        files.Push(A_LoopFileFullPath)
    }

    try {
        if !emptyDirectoryAllowed && files.Length = 0 {
            throw Error("Directory exists but contains no files: " directoryPath)
        }
    } catch as emptyDirectoryError {
        LogInformationConclusion("Failed", logValuesForConclusion, emptyDirectoryError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
    return files
}

GetFolderListFromDirectory(directoryPath, emptyDirectoryAllowed := false) {
    static methodName := RegisterMethod("GetFolderListFromDirectory(directoryPath As String [Type: Directory], emptyDirectoryAllowed As Boolean [Optional: false])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Get Folder List from Directory (" . directoryPath . ")", methodName, [directoryPath, emptyDirectoryAllowed])

    folders := []
    pattern := RTrim(directoryPath, "\/") . "\*"

    Loop Files, pattern, "D"
    {
        folders.Push(A_LoopFileFullPath . "\")
    }

    try {
        if !emptyDirectoryAllowed && folders.Length = 0 {
            throw Error("Directory exists but contains no folders: " . directoryPath)
        }
    } catch as emptyDirectoryError {
        LogInformationConclusion("Failed", logValuesForConclusion, emptyDirectoryError)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
    return folders
}

MoveFileToDirectory(filePath, directoryPath, overwrite := false) {
    static methodName := RegisterMethod("MoveFileToDirectory(filePath As String [Type: Absolute Path], directoryPath As String [Type: Directory], overwrite As Boolean [Optional: false])", A_LineFile, A_LineNumber + 1)
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

    if overwrite = false {
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
    static methodName := RegisterMethod("ReadFileOnHashMatch(filePath As String [Type: Absolute Path], expectedHash As String [Type: SHA-256])", A_LineFile, A_LineNumber + 1)
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

        Loop beSize // 2 {
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

    if (SubStr(fileText, 1, 1) = Chr(0xFEFF)) {
        fileText := SubStr(fileText, 2)
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
    return fileText
}

WriteBase64IntoImageFileWithHash(base64Text, filePath, expectedHash) {
    static methodName := RegisterMethod("WriteBase64IntoImageFileWithHash(base64Text As String [Type: Base64], filePath As String [Type: Absolute Save Path], expectedHash As String [Type: SHA-256])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Write Base64 into Image File with Hash" . " (" . ExtractFilename(filePath) . ")", methodName, [base64Text, filePath, expectedHash])
    
    needsWrite := true

    if FileExist(filePath) {
        fileHash := Hash.File("SHA256", filePath)

        if !(StrUpper(fileHash) = StrUpper(expectedHash)) {
            FileDelete(filePath)
        } else {
            needsWrite := false
        }
    }

    if needsWrite = true {
        ; Decode Base64 via MSXML (single failure point)
        try {
            xmlDocument   := ComObject("MSXML2.DOMDocument.6.0")
            base64Element := xmlDocument.createElement("b64")
            base64Element.dataType := "bin.base64"
            base64Element.text := base64Text
            byteArray := base64Element.nodeTypedValue
            byteCount := byteArray.MaxIndex(1) + 1
            if byteCount <= 0 {
                throw Error("Decoded byte array is empty.")
            }
        } catch as base64DecodingError {
            LogInformationConclusion("Failed", logValuesForConclusion, base64DecodingError)
        }

        ; Atomic write (single failure point). No cleanup to avoid a second error.
        temporaryFilePath := filePath ".part"
        try {
            fileHandle := FileOpen(temporaryFilePath, "w")
            Loop byteCount {
                fileHandle.WriteUChar(byteArray[A_Index - 1])
            }
            fileHandle.Close()
            FileMove(temporaryFilePath, filePath, 1)
        } catch as fileWriteError {
            LogInformationConclusion("Failed", logValuesForConclusion, fileWriteError)

        }

        try {
            fileHash := Hash.File("SHA256", filePath)
            if fileHash != expectedHash {
                throw Error("Hash mismatch in " . filePath . "`n`nExpected: " . expectedHash . "`nResults: " . fileHash)
            }
        } catch as hashMismatchError {
            LogInformationConclusion("Failed", logValuesForConclusion, hashMismatchError)
        }

        LogInformationConclusion("Completed", logValuesForConclusion)
    } else {
        LogInformationConclusion("Skipped", logValuesForConclusion)
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
    if removeFileExtension = true {
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