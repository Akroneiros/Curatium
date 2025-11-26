#Requires AutoHotkey v2.0
#Include ..\AHK_CNG (2021-11-03)\Class_CNG.ahk
#Include Application Library.ahk
#Include Base Library.ahk
#Include File Library.ahk
#Include Logging Library.ahk

global imageRegistry := Map()

CreateApplicationImages() {
    static methodName := RegisterMethod("CreateApplicationImages()", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Create Application Images", methodName)

    if !IsSet(applicationRegistry) {
        try {
            throw Error("Application Registry has not yet been initialized. Need to run RegisterApplications first.")
        } catch as applicationRegistryUninitializedError {
            LogInformationConclusion("Failed", logValuesForConclusion, applicationRegistryUninitializedError)
        }
    }

    imagesFileList := GetFilesFromDirectory(system["Images Directory"])
    applicationImageLibraryDataFiles := []
    for filePath in imagesFileList {
        if InStr(filePath, "Image Library Data") {
            filename := ExtractFilename(filePath, true)
            if RegExMatch(filename, "\((.*)\)\s*$", &capturedGroups) {
                applicationImageLibraryDataFiles.Push(capturedGroups[1])
            }
        }
    }

    installedApplicationsWithImageLibraryDataCount := 0
    for outerKey, innerMap in applicationRegistry {
        if innermap["Installed"] {
            for applicationName in applicationImageLibraryDataFiles {
                if outerKey = applicationName {
                    installedApplicationsWithImageLibraryDataCount := installedApplicationsWithImageLibraryDataCount + 1
                }
            }
        }
    }

    if installedApplicationsWithImageLibraryDataCount = 0 {
        LogInformationConclusion("Skipped", logValuesForConclusion)
    } else {
        switch system["Display Resolution"] {
            case "1920x1080":
                switch system["DPI Scale"] {
                    case "100%", "125%", "150%":
                        CreateImagesFromCatalog("Full High Definition")
                    default:
                        try {
                            throw Error("DPI scale unsupported: " . system["DPI Scale"] . ". For 1920x1080 the following scales are supported: 100%, 125%, 150%.")
                        } catch as dpiScaleUnsupportedError {
                            LogInformationConclusion("Failed", logValuesForConclusion, dpiScaleUnsupportedError)
                        }
                }
            case "2560x1440":
                switch system["DPI Scale"] {
                    case "100%", "125%", "150%":
                        CreateImagesFromCatalog("Quad High Definition")
                    default:
                        try {
                            throw Error("DPI scale unsupported: " . system["DPI Scale"] . ". For 2560x1440 the following scales are supported: 100%, 125%, 150%.")
                        } catch as dpiScaleUnsupportedError {
                            LogInformationConclusion("Failed", logValuesForConclusion, dpiScaleUnsupportedError)
                        }
                }
            case "3840x2160":
                switch system["DPI Scale"] {
                    case "100%", "125%", "150%", "175%":
                        CreateImagesFromCatalog("Ultra High Definition")
                    default:
                        try {
                            throw Error("DPI scale unsupported: " . system["DPI Scale"] . ". For 3840x2160 the following scales are supported: 100%, 125%, 150%, 175%.")
                        } catch as dpiScaleUnsupportedError {
                            LogInformationConclusion("Failed", logValuesForConclusion, dpiScaleUnsupportedError)
                        }
                }
            default:
                try {
                    throw Error("Display resolution unsupported: " . system["Display Resolution"] . ". The following display resolutions are supported: 1920x1080, 2560x1440, 3840x2160.")
                } catch as displayResolutionUnsupportedError {
                    LogInformationConclusion("Failed", logValuesForConclusion, displayResolutionUnsupportedError)
                }
        }

        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

CreateImagesFromCatalog(imageLibraryCatalogName) {
    static methodName := RegisterMethod("CreateImagesFromCatalog(imageCatalogName As String [Constraint: Locator])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Create Images from Catalog", methodName, [imageLibraryCatalogName])

    global imageRegistry

    projectImageCatalogFilePath  := system["Project Directory"] . "Image Library Catalog (" . imageLibraryCatalogName . ").csv"
    sharedImageCatalogFilePath   := system["Images Directory"] . "Image Library Catalog (" . imageLibraryCatalogName . ").csv"
    catalogDirectory             := unset
    imageLibraryCatalog          := unset

    if !IsSet(applicationRegistry) {
        try {
            throw Error("Application Registry has not yet been initialized. Need to run RegisterApplications first.")
        } catch as applicationRegistryUninitializedError {
            LogInformationConclusion("Failed", logValuesForConclusion, applicationRegistryUninitializedError)
        }
    }

    try {
        if !FileExist(projectImageCatalogFilePath) && !FileExist(sharedImageCatalogFilePath) {
            throw Error("Image Library Catalog not found: " . imageLibraryCatalogName)
        }
    } catch as imageCatalogMissingError {
        LogInformationConclusion("Failed", logValuesForConclusion, imageCatalogMissingError)
    }

    try {
        if FileExist(projectImageCatalogFilePath) && FileExist(sharedImageCatalogFilePath) {
            throw Error("Image Library Catalog with the same name found in both Images and Project: " . imageLibraryCatalogName)
        }
    } catch as imageCatalogDuplicateError {
        LogInformationConclusion("Failed", logValuesForConclusion, imageCatalogDuplicateError)
    }
    
    if FileExist(projectImageCatalogFilePath) && !FileExist(sharedImageCatalogFilePath) {
        catalogDirectory := system["Project Directory"]
        imageLibraryCatalog := ConvertCsvToArrayOfMaps(projectImageCatalogFilePath)
    } else if !FileExist(projectImageCatalogFilePath) && FileExist(sharedImageCatalogFilePath) {
        catalogDirectory := system["Images Directory"]
        imageLibraryCatalog := ConvertCsvToArrayOfMaps(sharedImageCatalogFilePath)
    }

    static applications      := unset
    static variants          := unset
    static fileSignatures    := unset
    static resolutions       := unset
    static scales            := unset
    static displayResolution := unset
    static dpiScale          := unset

    if !IsSet(applications) && !IsSet(variants) && !IsSet(fileSignatures) && !IsSet(resolutions) && !IsSet(scales) && !IsSet(displayResolution) && !IsSet(dpiScale) {
        applications := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Applications.csv")

        heroes := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Heroes (2025-09-20).csv")

        variants := Map()
        for hero in heroes {
            heroName       := hero["Name"]
            firstCharacter := StrLower(SubStr(heroName, 1, 1))

            variants[firstCharacter] := heroName
        }

        fileSignatures := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "File Signatures.csv")
        for index, rowMap in fileSignatures {
            rowMap["Maximum Base64 Signature"] := ConvertHexStringToBase64(rowMap["Maximum Hex Signature"])
        }

        resolutions := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Resolutions (2025-09-20).csv")
        for index, rowMap in resolutions {
            rowMap["Counter"] := index
        }

        scales := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Scales (2025-09-20).csv")
        for index, rowMap in scales {
            rowMap["Counter"] := index
        }

        displayResolution := ExtractRowFromArrayOfMapsOnHeaderCondition(resolutions, "Resolution", system["Display Resolution"])["Counter"] . ""
        dpiScale          := ExtractRowFromArrayOfMapsOnHeaderCondition(scales, "Scale", system["DPI Scale"])["Counter"] . ""
    }

    relevantImages       := []
    uniqueDataReferences := []

    for image in imageLibraryCatalog {
        if IsInteger(image["Image Library Data Reference"]) {
            image["Image Library Data Reference"] := ExtractRowFromArrayOfMapsOnHeaderCondition(applications, "Counter", image["Image Library Data Reference"])["Name"]

            if displayResolution = image["Display Resolution"] && dpiScale = image["DPI Scale"] && applicationRegistry[image["Image Library Data Reference"]]["Installed"] {
                relevantImages.Push(image)
                uniqueDataReferences.Push(image["Image Library Data Reference"])
            }
        } else {
            if displayResolution = image["Display Resolution"] && dpiScale = image["DPI Scale"] {
                relevantImages.Push(image)
                uniqueDataReferences.Push(image["Image Library Data Reference"])
            }
        }
    }

    if relevantImages.Length = 0 {
        LogInformationConclusion("Skipped", logValuesForConclusion)
    } else {
        pendingBase64ImageWriteQueue := []
        uniqueDataReferences         := RemoveDuplicatesFromArray(uniqueDataReferences)

        uniqueDataReferencesDirectories := uniqueDataReferences.Clone()
        for index, uniqueDataReferenceDirectory in uniqueDataReferencesDirectories {
            uniqueDataReferencesDirectories[index] := system["Images Directory"] . uniqueDataReferenceDirectory . "\"
        }

        BatchAppendSymbolLedger("D", uniqueDataReferencesDirectories)

        static screenWidth  := A_ScreenWidth
        static screenHeight := A_ScreenHeight
        for uniqueDataReference in uniqueDataReferences {
            libraryDataEntries := unset

            libraryDataEntries := ConvertCsvToArrayOfMaps(catalogDirectory . "Image Library Data (" . uniqueDataReference . ").csv")
            EnsureDirectoryExists(system["Images Directory"] . uniqueDataReference . "\")

            for image in relevantImages {
                for libraryData in libraryDataEntries {
                    if image["Counter Reference"] = libraryData["Counter"] && uniqueDataReference = image["Image Library Data Reference"] {
                        if !libraryData.Has("Directory") {
                            libraryData["Directory"] := image["Image Library Data Reference"]
                            libraryData["Filename"]  := libraryData["Name"] . " (" . variants[libraryData["Variant"]] . ")." . libraryData["Extension"]
                            libraryData["SHA-256"]   := DecodeBaseToSha256Hex(libraryData["SHA-256"], 86)
                            libraryData["Base64"]    := ExtractRowFromArrayOfMapsOnHeaderCondition(fileSignatures, "Extension", libraryData["Extension"])["Maximum Base64 Signature"] . libraryData["Base64"]
                            pendingBase64ImageWriteQueue.Push(libraryData)

                            if !imageRegistry.Has(libraryData["Directory"]) {
                                imageRegistry[libraryData["Directory"]] := Map()
                            }

                            if !imageRegistry[libraryData["Directory"]].Has(libraryData["Name"]) {
                                imageRegistry[libraryData["Directory"]][libraryData["Name"]] := []
                            }

                            path := system["Images Directory"] . libraryData["Directory"] . "\" . libraryData["Filename"]

                            horizontalRange      := StrReplace(image["Horizontal Range"], ",", ".")
                            horizontalParts      := StrSplit(horizontalRange, "-")
                            horizontalRangeStart := Floor(screenWidth * horizontalParts[1] / 100)
                            horizontalRangeEnd   := Ceil(screenWidth * horizontalParts[2] / 100) - 1

                            verticalRange        := StrReplace(image["Vertical Range"], ",", ".")
                            verticalParts        := StrSplit(verticalRange, "-")
                            verticalRangeStart   := Floor(screenHeight * verticalParts[1] / 100)
                            verticalRangeEnd     := Ceil(screenHeight * verticalParts[2] / 100) - 1

                            imageRegistry[libraryData["Directory"]][libraryData["Name"]].Push(Map(
                                "Path", path,
                                "Name", libraryData["Name"],
                                "Variant", StrLower(libraryData["Variant"]),
                                "Extension", libraryData["Extension"],
                                "Horizontal Range", horizontalRange,
                                "Horizontal Range Start", horizontalRangeStart,
                                "Horizontal Range End", horizontalRangeEnd,
                                "Vertical Range", verticalRange,
                                "Vertical Range Start", verticalRangeStart,
                                "Vertical Range End", verticalRangeEnd
                            ))

                            
                        }
                    }
                }
            }
        }

        filenameValues := []
        hashValues := []
        for image in pendingBase64ImageWriteQueue {
            filenameValues.Push(image["Filename"])
            hashValues.Push(image["SHA-256"])
        }

        BatchAppendSymbolLedger("F", filenameValues)
        BatchAppendSymbolLedger("H", hashValues)

        for image in pendingBase64ImageWriteQueue {
            filePath := system["Images Directory"] . image["Directory"] . "\" . image["Filename"]
            if FileExist(filePath) {
                if image["SHA-256"] != Hash.File("SHA256", filePath) {
                    DeleteFile(filePath)
                }
            }

            WriteBase64IntoFileWithHash(image["Base64"], filePath, image["SHA-256"])
        }

        for uniqueDataReference in uniqueDataReferences {
            for image in imageRegistry[uniqueDataReference] {
                directoryImages := imageRegistry[uniqueDataReference][image]

                for directoryImage in directoryImages {
                    imageDimensions := StrSplit(GetImageDimensions(directoryImage["Path"]), "x")
                    directoryImage["Width"]  := imageDimensions[1] + 0
                    directoryImage["Height"] := imageDimensions[2] + 0
                }
            }
        }

        LogInformationConclusion("Completed", logValuesForConclusion)
    }
}

SearchForDirectoryImage(directoryFolder, imageName, secondsToAttempt := 60, variant := "") {
    static methodName := RegisterMethod("SearchForDirectoryImage(directoryFolder As String, imageName As String, secondsToAttempt As Integer [Optional: 60], variant As String [Optional])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Search for Directory Image (" . directoryFolder . ", " . imageName . ")", methodName, [directoryFolder, imageName, secondsToAttempt, variant])

    try {
        if !imageRegistry.Has(directoryFolder) {
            throw Error("Directory folder for image not found: " . directoryFolder)
        }
    } catch as directoryFolderForImageNotFoundError {
        LogInformationConclusion("Failed", logValuesForConclusion, directoryFolderForImageNotFoundError)
    }

    try {
        if !imageRegistry[directoryFolder].Has(imageName) {
            throw Error("Image name not found: " . imageName)
        }
    } catch as imageNotFoundError {
        LogInformationConclusion("Failed", logValuesForConclusion, imageNotFoundError)
    }

    if secondsToAttempt = 0 {
        secondsToAttempt := 1
    }

    variant := StrLower(variant)
    if variant != "" {
        try {
            variantFound := false
            for image in imageRegistry[directoryFolder][imageName] {
                if variant = image["Variant"] {
                    variantFound := true
                }
            }

            if !variantFound {
                throw Error("Variant not found: " . variant)
            }
        } catch as variantNotFoundError {
            LogInformationConclusion("Failed", logValuesForConclusion, variantNotFoundError)
        }
    }

    overlayVisibility := OverLayIsVisible()
    if overlayVisibility {
        OverlayChangeVisibility()
    }

    CoordMode("Pixel", "Screen")
    static displayResolution := StrSplit(system["Display Resolution"], "x")
    static screenWidth  := displayResolution[1] + 0
    static screenHeight := displayResolution[2] + 0

    directoryImageVariants := unset
    if variant = "" {
        directoryImageVariants := imageRegistry[directoryFolder][imageName]
    } else {
        directoryImageVariants := []
        directoryImageVariantsLookup := imageRegistry[directoryFolder][imageName]
        for directoryImageVariantLookup in directoryImageVariantsLookup {
            if variant = directoryImageVariantLookup["Variant"] {
                directoryImageVariants.Push(directoryImageVariantLookup)
            }
        }
    }

    startTime := A_Now
    endTime   := DateAdd(startTime, secondsToAttempt, "Seconds")

    imageSearchResults := unset
    while A_Now < endTime && !IsSet(imageSearchResults) {
        for image in directoryImageVariants {
            horizontalCoordinate := 0
            verticalCoordinate := 0

            if ImageSearch(&horizontalCoordinate, &verticalCoordinate, image["Horizontal Range Start"], image["Vertical Range Start"], image["Horizontal Range End"], image["Vertical Range End"], image["Path"]) {
                horizontalCoordinate := horizontalCoordinate + Floor(image["Width"] / 2)
                verticalCoordinate   := verticalCoordinate + Floor(image["Height"] / 2)

                imageSearchResults := Map(
                    "Directory", directoryFolder,
                    "Name", imageName,
                    "Variant", image["Variant"],
                    "Coordinate Pair", horizontalCoordinate "x" verticalCoordinate,
                    "Seconds to Attempt", secondsToAttempt
                )
            }
        }

        Sleep(1000)
    }

    if !IsSet(imageSearchResults) {
        imageSearchResults := Map(
            "Directory", directoryFolder,
            "Name", imageName,
            "Variant", "",
            "Coordinate Pair", "",
            "Seconds to Attempt", secondsToAttempt
        )
    }

    if overlayVisibility {
        OverlayChangeVisibility()
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
    return imageSearchResults
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

ConvertImagesToBase64ImageLibrary(directoryPath) {
    static methodName := RegisterMethod("ConvertImagesToBase64ImageLibrary(directoryPath As String [Constraint: Directory])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [directoryPath])

    static applications   := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "Applications.csv")
    static fileSignatures := unset
    static resolutions    := unset
    static scales         := unset

    if !IsSet(fileSignatures) && !IsSet(resolutions) && !IsSet(scales) {
        fileSignatures := ConvertCsvToArrayOfMaps(system["Mappings Directory"] . "File Signatures.csv")
        for index, rowMap in fileSignatures {
            rowMap["Maximum Base64 Signature"] := ConvertHexStringToBase64(rowMap["Maximum Hex Signature"])
        }

        resolutions := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Resolutions (2025-09-20).csv")
        for index, rowMap in resolutions {
            rowMap["Counter"] := index
        }

        scales := ConvertCsvToArrayOfMaps(system["Constants Directory"] . "Scales (2025-09-20).csv")
        for index, rowMap in scales {
            rowMap["Counter"] := index
        }
    }

    static newLine       := "`r`n"
    static headerCatalog := "Image Library Data Reference|Counter Reference|Display Resolution|DPI Scale|Horizontal Range|Vertical Range" . newLine
    static headerData    := "Name|Variant|Counter|SHA-256|Extension|Base64" . newLine

    SplitPath(RTrim(directoryPath, "\/"), &referenceDirectoryName)
    referenceIsApplication := false
    for index, application in applications {
        if application["Name"] = referenceDirectoryName {
            referenceIsApplication := true
            break
        }
    }

    imageLibraryCatalogFilePath       := directoryPath . "Image Library Catalog (" . referenceDirectoryName . ")" . ".csv"
    imageLibraryDataReferenceFilePath := directoryPath . "Image Library Data (" . referenceDirectoryName . ")" . ".csv"
    imageLibraryDataReference         := unset
    if referenceIsApplication {
        imageLibraryDataReference := ExtractRowFromArrayOfMapsOnHeaderCondition(applications, "Name", referenceDirectoryName)["Counter"] + 0
    } else {
        imageLibraryDataReference := referenceDirectoryName
    }

    if FileExist(imageLibraryCatalogFilePath) {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Image Library Catalog (" . referenceDirectoryName . ") already exists, remove before proceeding.")
    }

    hashToCounter := Map()
    counter := 1
    if FileExist(imageLibraryDataReferenceFilePath) {
        imageLibraryDataReferenceFile := ConvertCsvToArrayOfMaps(imageLibraryDataReferenceFilePath)

        for index, rowMap in imageLibraryDataReferenceFile {
            if counter < rowMap["Counter"] {
                counter := rowMap["Counter"]
            }
        }

        counter := counter + 1
    }

    originalCounter := counter
    catalogEntries  := []
    dataEntries     := []

    actionImageDirectories := GetFoldersFromDirectory(directoryPath)
    for index, actionFolderPath in actionImageDirectories {
        SplitPath(RTrim(actionFolderPath, "\/"), &lastActionDirectoryName)

        if !RegExMatch(lastActionDirectoryName, "^\s*(.+?)\s*\(([a-p])\)\s*$", &matchResults) {
            LogHelperError(logValuesForConclusion, A_LineNumber, "Folder does not match format of Action Name (a...p): " . lastActionDirectoryName)
        }

        actionName   := Trim(matchResults[1])
        actionLetter := matchResults[2]

        loop files, actionFolderPath . "*", "F" {
            SplitPath(A_LoopFileName, , , , &filenameWithoutExtension)
            lastOpenParenthesisIndex := InStr(filenameWithoutExtension, "(", "On", -1)
            baseTextWithoutRanges    := RTrim(SubStr(filenameWithoutExtension, 1, lastOpenParenthesisIndex - 1))
            rangeContent             := SubStr(filenameWithoutExtension, lastOpenParenthesisIndex + 1, InStr(filenameWithoutExtension, ")", "On", -1) - lastOpenParenthesisIndex - 1)

            rangeParts := StrSplit(rangeContent, ",", " `t")

            validation := unset
            for rangePart in rangeParts {
                validation := ValidateDataUsingSpecification(rangePart, "String", "Percent Range")

                if validation != "" {
                    subfolder := StrSplit(RTrim(actionFolderPath, "\"), "\").Pop()
                    LogHelperError(logValuesForConclusion, A_LineNumber, "File " . Chr(34) . A_LoopFileName . Chr(34) . " in subfolder " . Chr(34) . subfolder . Chr(34) . " has invalid range value in parenthesis after percent. " . validation)
                }
            }

            horizontalPercentRange := rangeParts[1]
            verticalPercentRange   := rangeParts[2]

            parts      := StrSplit(baseTextWithoutRanges, "@", " `t")
            resolution := ExtractRowFromArrayOfMapsOnHeaderCondition(resolutions, "Resolution", parts[1])["Counter"]
            scale      := ExtractRowFromArrayOfMapsOnHeaderCondition(scales, "Scale", parts[2])["Counter"]

            fileHash    := Hash.File("SHA256", A_LoopFileFullPath)
            encodedHash := EncodeSha256HexToBase(fileHash, 86)
            base64Data  := GetBase64FromFile(A_LoopFileFullPath)

            extension := ""
            for index, rowMap in fileSignatures {
                maximumBase64Signature := rowMap["Maximum Base64 Signature"]
                if SubStr(base64Data, 1, StrLen(maximumBase64Signature)) = maximumBase64Signature {
                    extension  := rowMap["Extension"]
                    base64Data := SubStr(base64Data, StrLen(maximumBase64Signature) + 1)
                    break
                }
            }               

            if !hashToCounter.Has(fileHash) {
                hashToCounter[fileHash] := counter
                counter := counter + 1

                dataEntries.Push(actionName . "|" . actionLetter . "|" . hashToCounter[fileHash] . "|" . encodedHash . "|" . extension . "|" . base64Data)
            }
            
            catalogEntries.Push(imageLibraryDataReference . "|" . hashToCounter[fileHash] . "|" . resolution . "|" . scale . "|" . horizontalPercentRange . "|" . verticalPercentRange)
        }
    }

    csvDataString := unset
    if originalCounter = 1 {
        csvDataString := headerData . ConvertArrayIntoCsvString(dataEntries)
    } else {
        csvDataString := newLine . ConvertArrayIntoCsvString(dataEntries)
    }

    FileAppend(csvDataString, imageLibraryDataReferenceFilePath, "UTF-8")

    csvCatalogString := headerCatalog . ConvertArrayIntoCsvString(catalogEntries)
    FileAppend(csvCatalogString, imageLibraryCatalogFilePath, "UTF-8")
}

ExtractScreenCoordinates(imageSearchResults) {
    static methodName := RegisterMethod("ExtractScreenCoordinates(imageSearchResults As Object)", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [imageSearchResults])

    if imageSearchResults["Variant"] = "" && imageSearchResults["Coordinate Pair"] = "" {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Image (" . imageSearchResults["Name"] . ")" . " not found in directory (" . imageSearchResults["Directory"] . "). Tried for " . imageSearchResults["Seconds to Attempt"] . " seconds.")
    }

    coordinatePair := imageSearchResults["Coordinate Pair"]

    return coordinatePair
}

GetImageDimensions(imageAbsolutePath) {
    static methodName := RegisterMethod("GetImageDimensions(imageAbsolutePath As String [Constraint: Absolute Path])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [imageAbsolutePath])

    static gdiPlusStartupInputSize := (A_PtrSize = 8) ? 24 : 16
    static gdiPlusStartupInputBuffer := Buffer(gdiPlusStartupInputSize, 0)
    NumPut("UInt", 1, gdiPlusStartupInputBuffer, 0)
    NumPut("Ptr", 0, gdiPlusStartupInputBuffer, (A_PtrSize= 8 ) ? 8:4)
    NumPut("UInt", 0, gdiPlusStartupInputBuffer, (A_PtrSize= 8 ) ? 16:8)
    NumPut("UInt", 0, gdiPlusStartupInputBuffer, (A_PtrSize= 8 ) ? 20:12)

    static gdiPlusLoadLibrary := DllCall("LoadLibrary", "Str", "GdiPlus", "Ptr")

    gdiPlusStartupToken := 0
    gdiPlusStartupResult := DllCall("GdiPlus\GdiplusStartup", "Ptr*", &gdiPlusStartupToken, "Ptr", gdiPlusStartupInputBuffer.Ptr, "Ptr", 0, "UInt")
    if gdiPlusStartupResult != 0 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to initialize Windows GDI+. [GdiPlus\GdiplusStartup" . ", GDI+ Status Code: " . gdiPlusStartupResult . "]")
    }

    gdiPlusBitmapPointer := 0
    gdiPlusCreateBitmapStatus := DllCall("GdiPlus\GdipCreateBitmapFromFile", "WStr", imageAbsolutePath, "Ptr*", &gdiPlusBitmapPointer, "UInt")
    if gdiPlusCreateBitmapStatus != 0 || !gdiPlusBitmapPointer {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to create a bitmap object based on the image. [GdiPlus\GdipCreateBitmapFromFile" . ", GDI+ Status Code: " . gdiPlusCreateBitmapStatus . "]")
    }

    gdiPlusGetImageWidthStatus := DllCall("GdiPlus\GdipGetImageWidth", "Ptr", gdiPlusBitmapPointer, "UInt*", &imageWidthPixels := 0, "UInt")
    if gdiPlusGetImageWidthStatus != 0 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to get the width in pixels of the image. [GdiPlus\GdipGetImageWidth" . ", GDI+ Status Code: " . gdiPlusGetImageWidthStatus . "]")
    }

    gdiPlusGetImageHeightStatus := DllCall("GdiPlus\GdipGetImageHeight", "Ptr", gdiPlusBitmapPointer, "UInt*", &imageHeightPixels := 0, "UInt")
    if gdiPlusGetImageHeightStatus != 0 {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Failed to get the height in pixels of the image. [GdiPlus\GdipGetImageHeight" . ", GDI+ Status Code: " . gdiPlusGetImageHeightStatus . "]")
    }

    DllCall("Gdiplus\GdipDisposeImage", "Ptr", gdiPlusBitmapPointer, "UInt")
    DllCall("Gdiplus\GdiplusShutdown", "UPtr", gdiPlusStartupToken, "UInt")

    imageDimensions := imageWidthPixels . "x" . imageHeightPixels

    return imageDimensions
}

OverrideDirectoryImageVariant(directoryFolder, imageName, variant, horizontalRange, verticalRange) {
    static methodName := RegisterMethod("OverrideDirectoryImageVariant(directoryFolder As String, imageName As String, variant As String, horizontalRange As String [Constraint: Percent Range], verticalRange As String [Constraint: Percent Range])", A_LineFile, A_LineNumber + 1)
    logValuesForConclusion := LogHelperValidation(methodName, [directoryFolder, imageName, variant, horizontalRange, verticalRange])

    global imageRegistry

    if !imageRegistry.Has(directoryFolder) {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Directory folder for image not found: " . directoryFolder)
    }

    if !imageRegistry[directoryFolder].Has(imageName) {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Image name not found: " . imageName)
    }

    variant := StrLower(variant)
    variantFound := false
    for image in imageRegistry[directoryFolder][imageName] {
        if variant = image["Variant"] {
            variantFound := true
        }
    }

    if !variantFound {
        LogHelperError(logValuesForConclusion, A_LineNumber, "Variant not found: " . variant)
    }

    static displayResolution := StrSplit(system["Display Resolution"], "x")
    static screenWidth  := displayResolution[1] + 0
    static screenHeight := displayResolution[2] + 0

    horizontalRange      := StrReplace(horizontalRange, ",", ".")
    horizontalParts      := StrSplit(horizontalRange, "-")
    horizontalRangeStart := Floor(screenWidth * horizontalParts[1] / 100)
    horizontalRangeEnd   := Ceil(screenWidth * horizontalParts[2] / 100) - 1

    verticalRange        := StrReplace(verticalRange, ",", ".")
    verticalParts        := StrSplit(verticalRange, "-")
    verticalRangeStart   := Floor(screenHeight * verticalParts[1] / 100)
    verticalRangeEnd     := Ceil(screenHeight * verticalParts[2] / 100) - 1

    for image in imageRegistry[directoryFolder][imageName] {
        if variant = image["Variant"] {
            image["Horizontal Range"] := horizontalRange
            image["Horizontal Range Start"] := horizontalRangeStart
            image["Horizontal Range End"] := horizontalRangeEnd
            image["Vertical Range"] := verticalRange
            image["Vertical Range Start"] := verticalRangeStart
            image["Vertical Range End"] := verticalRangeEnd

            break
        }
    }
}