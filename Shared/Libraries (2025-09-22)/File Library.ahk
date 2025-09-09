#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Logging Library.ahk

CleanOfficeLocksInFolder(directoryPath) {
    static methodName := RegisterMethod("CleanOfficeLocksInFolder(directoryPath As String [Type: Directory])" . LibraryTag(A_LineFile), A_LineNumber + 1)
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

CopyFileToTarget(filePath, targetDirectory, findValue := "", replaceValue := "") {
    static methodName := RegisterMethod("CopyFileToTarget(filePath As String [Type: Absolute Path], targetDirectory As String [Type: Directory], findValue As String [Optional], replaceValue As String [Optional])" . LibraryTag(A_LineFile), A_LineNumber + 1)
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
    static methodName := RegisterMethod("DeleteFile(filePath As String [Type: Absolute Path])" . LibraryTag(A_LineFile), A_LineNumber + 1)
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
    static methodName := RegisterMethod("EnsureDirectoryExists(directoryPath As String [Type: Directory])" . LibraryTag(A_LineFile), A_LineNumber + 1)
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
    static methodName := RegisterMethod("FileExistsInDirectory(filename As String [Type: Search], directoryPath As String [Type: Directory], fileExtension As String [Optional])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("File Exists in Directory (" . filename . ")", methodName, [filename, directoryPath, fileExtension])

    filesInDirectory := GetFileListFromDirectory(directoryPath, true)

    if filesInDirectory.Length = 0 {
        LogInformationConclusion("Completed", logValuesForConclusion)
        return ""
    }

    index := filesInDirectory.Length
    while (index >= 1) {
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
    static methodName := RegisterMethod("GetFileListFromDirectory(directoryPath As String [Type: Directory], emptyDirectoryAllowed As Boolean [Optional: false])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Get File List from Directory (" . directoryPath . ")", methodName, [directoryPath, emptyDirectoryAllowed])

    files := []
    pattern := RTrim(directoryPath, "\/") . "\*"

    Loop Files, pattern
    {
        if InStr(A_LoopFileAttrib, "D") { ; Skip Directories.
            continue
        }
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

MoveFileToDirectory(filePath, directoryPath, overwrite := false) {
    static methodName := RegisterMethod("MoveFileToDirectory(filePath As String [Type: Absolute Path], directoryPath As String [Type: Directory], overwrite As Boolean [Optional: false])" . LibraryTag(A_LineFile), A_LineNumber + 1)
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
    static methodName := RegisterMethod("ReadFileOnHashMatch(filePath As String [Type: Absolute Path], expectedHash As String [Type: SHA-256])" . LibraryTag(A_LineFile), A_LineNumber + 1)
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

    LogInformationConclusion("Completed", logValuesForConclusion)
    return FileRead(filePath)
}

WriteBase64IntoImageFileWithHash(base64Text, filePath, expectedHash) {
    static methodName := RegisterMethod("WriteBase64IntoImageFileWithHash(base64Text As String [Type: Base64], filePath As String [Type: Absolute Save Path], expectedHash As String [Type: SHA-256])" . LibraryTag(A_LineFile), A_LineNumber + 1)
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

; ******************** ;
; Helper Methods       ;
; ******************** ;

ExtractDirectory(filePath) {
    SplitPath(filePath, , &directoryPath)

    if directoryPath != "" && SubStr(directoryPath, -1) != "\" {
        directoryPath .= "\"
    }

    return directoryPath
}

ExtractFilename(filePath, removeFileExtension := false) {
    SplitPath(filePath, &filenameWithExtension, , , &filenameWithoutExtension)

    filename := filenameWithExtension
    if removeFileExtension = true {
        filename := filenameWithoutExtension
    }

    return filename
}