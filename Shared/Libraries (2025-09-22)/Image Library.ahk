#Requires AutoHotkey v2.0
#Include "..\jsongo_AHKv2 (2025-02-26)\jsongo.v2.ahk"
#Include Base Library.ahk
#Include File Library.ahk
#Include Logging Library.ahk

AssignSharedImages() {
    static methodName := RegisterMethod("AssignSharedImages()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Assign Shared Images", methodName)

    SplitPath(A_LineFile, , &libraryFileDirectory)
    SplitPath(libraryFileDirectory, , &parentDirectory)
    imageDirectory := parentDirectory . "\Images\"

    filesInDirectory := GetFileListFromDirectory(imageDirectory)

    position := filesInDirectory.Length
    while position >= 1
    {
        if RegExMatch(filesInDirectory[position], "i)\.ndjson$") {
            filesInDirectory.RemoveAt(position)
        }
        position -= 1
    }

    sharedImages := Map()

    for index, filePath in filesInDirectory {
        SplitPath(filePath, , , , &filenameNoExtension)
        sharedImages[filenameNoExtension] := filePath
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
    return sharedImages
}

CreateSharedImages(imageCatalogName) {
    static methodName := RegisterMethod("CreateSharedImages(imageCatalogName As String [Type: Search])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Create Shared Images", methodName, [imageCatalogName])

    SplitPath(A_LineFile, , &libraryFileDirectory)
    SplitPath(libraryFileDirectory, , &parentDirectory)
    imageDirectory := parentDirectory . "\Images\"

    if !DirExist(imageDirectory) {
        DirCreate(imageDirectory)
    }

    imageCatalogFilePath := imageDirectory . imageCatalogName . ".ndjson"

    try {
        if !FileExist(imageCatalogFilePath) {
            throw Error("Image Catalog not found: " . imageCatalogFilePath)
        }
    } catch as imageCatalogMissingError {
        LogInformationConclusion("Failed", logValuesForConclusion, imageCatalogMissingError)
    }

    static heroes := ConvertCsvToArrayOfMaps(ExtractParentDirectory(A_LineFile) . "Constants\Heroes (2025-09-20).csv")
    static variations := Map()
    for hero in heroes {
        heroName       := hero["Name"]
        firstCharacter := StrLower(SubStr(heroName, 1, 1))

        variations[firstCharacter] := heroName
    }

    sharedImages := []

    for line in StrSplit(FileRead(imageCatalogFilePath, "UTF-8"), "`n") {
        line := Trim(line, "`r`n ")
        if line = "" || SubStr(line, 1, 1) = "#" {
            continue
        }

        image := jsongo.Parse(line)
        image["Variation"] := variations[image["Variation"]]

        sharedImages.Push(image)
    }

    if imageCatalogName = "Image Library Catalog (2025-09-04)" {
        for index, image in sharedImages {
            if image["Name"] = "SMMS Query Successful" {
                image["Name"] := "SQL Server Management Studio Query executed successfully"
            }

            if image["Name"] = "Toad for Oracle Play" {
                image["Name"] := "Toad for Oracle Run selected apps"
            }
        }
    }

    filenameValues := []
    for index, image in sharedImages {
        filenameValues.Push(image["Name"] . " (" . image["Variation"] . ")." . image["Extension"])
    }

    BatchAppendSymbolLedger("F", filenameValues)

    hashValues := []
    for index, image in sharedImages {
        hashValues.Push(image["SHA-256"])
    }

    BatchAppendSymbolLedger("H", hashValues)

    for image in sharedImages {
        WriteBase64IntoImageFileWithHash(image["Base64"], imageDirectory . image["Name"] . " (" . image["Variation"] . ")." . image["Extension"], image["SHA-256"])
    }  

    LogInformationConclusion("Completed", logValuesForConclusion)
}

GetImageCoordinatesFromSegment(imageAlias, horizontalPercentRange, verticalPercentRange, secondsOfWaitingBeforeFailure := 60) {
    static methodName := RegisterMethod("GetImageCoordinatesFromSegment(imageAlias As String, horizontalPercentRange As String [Type: Percent Range], verticalPercentRange As String [Type: Percent Range], secondsOfWaitingBeforeFailure As Integer [Optional: 60])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Get Image Coordinates from Segment", methodName, [imageAlias, horizontalPercentRange, verticalPercentRange, secondsOfWaitingBeforeFailure])

    coordinatePair := ""

    horizontalPercentRange := StrReplace(horizontalPercentRange, ",", ".")
    horizontalParts        := StrSplit(horizontalPercentRange, "-")
    horizontalPercentStart := horizontalParts[1] + 0
    horizontalPercentEnd   := horizontalParts[2] + 0

    try {
        if horizontalPercentStart >= horizontalPercentEnd {
            throw Error("Wrong order horizontally. " . "End value of " . horizontalPercentEnd . " is smaller than or equal to than start value of " . horizontalPercentStart . ".")
        }
    } catch as horizontalWrongOrderError {
        LogInformationConclusion("Failed", logValuesForConclusion, horizontalWrongOrderError)
    }

    try {
        if horizontalPercentStart >= 100 {
            throw Error("Horizontal start value of " . horizontalPercentStart . " is invalid. Maximum allowed is 99.9.")
        }
    } catch as invalidHorizontalValueError {
        LogInformationConclusion("Failed", logValuesForConclusion, invalidHorizontalValueError)
    }

    verticalPercentRange := StrReplace(verticalPercentRange, ",", ".")
    verticalParts        := StrSplit(verticalPercentRange, "-")
    verticalPercentStart := verticalParts[1] + 0
    verticalPercentEnd   := verticalParts[2] + 0

    try {
        if verticalPercentStart >= verticalPercentEnd {
            throw Error("Wrong order vertically. " . "End value of " . verticalPercentEnd . " is smaller than or equal to than start value of " . verticalPercentStart . ".")
        }
    } catch as verticalWrongOrderError {
        LogInformationConclusion("Failed", logValuesForConclusion, verticalWrongOrderError)
    }

    try {
        if verticalPercentStart >= 100 {
            throw Error("Vertical start value of " . verticalPercentStart . " is invalid. Maximum allowed is 99.9.")
        }
    } catch as invalidVerticalValueError {
        LogInformationConclusion("Failed", logValuesForConclusion, invalidVerticalValueError)
    }

    CoordMode("Pixel", "Screen")
    screenWidth  := A_ScreenWidth
    screenHeight := A_ScreenHeight

    regionLeftPixel := Floor(screenWidth * horizontalPercentStart / 100)
    regionRightPixel := Ceil(screenWidth * horizontalPercentEnd / 100) - 1
    regionTopPixel := Floor(screenHeight * verticalPercentStart / 100)
    regionBottomPixel := Ceil(screenHeight * verticalPercentEnd / 100) - 1

    sharedImages := AssignSharedImages()
    chosenImages := Map()
    for imageNameNoExtension, imageFullPath in sharedImages {
        aliasStartPosition := InStr(imageNameNoExtension, " (")
        baseImageName := aliasStartPosition ? SubStr(imageNameNoExtension, 1, aliasStartPosition - 1) : imageNameNoExtension

        if baseImageName = imageAlias {
            chosenImages[imageNameNoExtension] := imageFullPath
        }
    }

    try {
        if chosenImages.Count = 0 {
            throw Error("No image variants found for alias: " imageAlias)
        }
    } catch as noImageVariantsFoundForAliasError {
        LogInformationConclusion("Failed", logValuesForConclusion, noImageVariantsFoundForAliasError)
    }

    startTime            := A_Now
    endTime              := DateAdd(startTime, secondsOfWaitingBeforeFailure, "Seconds")
    static headerBuffer  := Buffer(24, 0)
    imageDimensionsCache := Map()

    overlayVisibility := OverLayIsVisible()
    if overlayVisibility = True {
        OverlayChangeVisibility()
    }

    for aliasVariant, imagePath in chosenImages {
        fileHandle := FileOpen(imagePath, "r")

        try {
            if !fileHandle {
                Throw("Could not open file:`n" imagePath)
            }
        } catch as couldNotOpenFileError {
            LogInformationConclusion("Failed", logValuesForConclusion, couldNotOpenFileError)
        }

        try {
            if fileHandle.RawRead(headerBuffer, 24) < 24 {
                Throw("Could not read PNG header (need first 24 bytes).")
            }
        } catch as preloadError {
            LogInformationConclusion("Failed", logValuesForConclusion, preloadError)
        } finally {
            fileHandle.Close()
        }

        imageWidthPixels := (NumGet(headerBuffer, 16, "UChar") << 24)
                        | (NumGet(headerBuffer, 17, "UChar") << 16)
                        | (NumGet(headerBuffer, 18, "UChar") << 8)
                        |  NumGet(headerBuffer, 19, "UChar")
        imageHeightPixels := (NumGet(headerBuffer, 20, "UChar") << 24)
                        | (NumGet(headerBuffer, 21, "UChar") << 16)
                        | (NumGet(headerBuffer, 22, "UChar") << 8)
                        |  NumGet(headerBuffer, 23, "UChar")

        try {
            if imageWidthPixels <= 0 || imageHeightPixels <= 0 {
                throw Error("PNG dimensions look invalid. Use a proper PNG.")
            }
        } catch as preloadDimensionsError {
            LogInformationConclusion("Failed", logValuesForConclusion, preloadDimensionsError)
        }

        imageDimensionsCache[imagePath] := [imageWidthPixels, imageHeightPixels]
    }

    while A_Now < endTime {
        for aliasVariant, imagePath in chosenImages {
            imageDimensions   := imageDimensionsCache[imagePath]
            imageWidthPixels  := imageDimensions [1]
            imageHeightPixels := imageDimensions [2]

            foundLeftPixel := 0
            foundTopPixel  := 0

            if ImageSearch(&foundLeftPixel, &foundTopPixel, regionLeftPixel, regionTopPixel, regionRightPixel, regionBottomPixel, imagePath) {
                centerHorizontalCoordinate := foundLeftPixel + Floor(imageWidthPixels / 2)
                centerVerticalCoordinate   := foundTopPixel  + Floor(imageHeightPixels / 2)
                coordinatePair             := centerHorizontalCoordinate "x" centerVerticalCoordinate

                if overlayVisibility = True {
                    OverlayChangeVisibility()
                }

                LogInformationConclusion("Completed", logValuesForConclusion)
                return coordinatePair
            }
        }

        Sleep(800)
    }

    if overlayVisibility = True {
        OverlayChangeVisibility()
    }

    try {
        throw Error("Image (" . imageAlias . ")" . " not found inside coordinates " . regionLeftPixel . "x" . regionTopPixel . " to " . regionRightPixel . "x" . regionBottomPixel . ".")
    } catch as imageNotFoundError {
        LogInformationConclusion("Failed", logValuesForConclusion, imageNotFoundError)
    }
}