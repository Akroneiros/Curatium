#Requires AutoHotkey v2.0
#Include Application Library.ahk
#Include Base Library.ahk
#Include Chrono Library.ahk
#Include File Library.ahk
#Include Logging Library.ahk

ConvertImagesToBase64ImageLibrary(directoryPath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("directoryPath As String [Constraint: Directory]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [directoryPath], "Convert Images to Base64 Image Library (" . directoryPath . ")")

    static newLine       := "`r`n"
    static headerCatalog := "Image Library Data Reference|Counter Reference|Display Resolution|DPI Scale|Horizontal Range|Vertical Range" . newLine
    static headerData    := "Name|Variant|Counter|SHA-256|Extension|Base64" . newLine

    SplitPath(RTrim(directoryPath, "\/"), &referenceDirectoryName)
    referenceIsApplication := false
    for application in system["Mappings"]["Applications"] {
        if application["Name"] = referenceDirectoryName {
            referenceIsApplication := true
            break
        }
    }

    imageLibraryCatalogFilePath       := directoryPath . "Image Library Catalog (" . referenceDirectoryName . ")" . ".csv"
    imageLibraryDataReferenceFilePath := directoryPath . "Image Library Data (" . referenceDirectoryName . ")" . ".csv"
    imageLibraryDataReference         := unset
    if referenceIsApplication {
        imageLibraryDataReference := ExtractRowFromArrayOfMapsOnHeaderCondition(system["Mappings"]["Applications"], "Name", referenceDirectoryName)["Counter"] + 0
    } else {
        imageLibraryDataReference := referenceDirectoryName
    }

    if FileExist(imageLibraryCatalogFilePath) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Image Library Catalog (" . referenceDirectoryName . ") already exists, remove before proceeding.")
    }

    counter := 1
    if FileExist(imageLibraryDataReferenceFilePath) {
        imageLibraryDataReferenceHash    := GetFileHash(imageLibraryDataReferenceFilePath, "SHA-256")
        imageLibraryDataReferenceContent := ReadFileOnHashMatch(imageLibraryDataReferenceFilePath, imageLibraryDataReferenceHash)
        imageLibraryDataReferenceArray   := ParseDelimitedRowsToArrayOfMaps(imageLibraryDataReferenceContent)

        for rowMap in imageLibraryDataReferenceArray {
            if counter < rowMap["Counter"] {
                counter := rowMap["Counter"]
            }
        }

        counter := counter + 1
    }

    counterModified := false
    if counter != 1 {
        counterModified := true
    }

    catalogEntries  := []
    dataEntries     := []

    actionImageDirectories := GetFoldersFromDirectory(directoryPath)
    for actionFolderPath in actionImageDirectories {
        SplitPath(RTrim(actionFolderPath, "\/"), &actionDirectoryName)

        if !RegExMatch(actionDirectoryName, "^\s*(.+?)\s*\(([a-p])\)\s*$", &matchResults) {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Folder does not match format of Action Name (a...p): " . actionDirectoryName)
        }

        actionName   := Trim(matchResults[1])
        actionLetter := matchResults[2]

        Loop Files, actionFolderPath . "*", "F" {
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
                    LogConclusion("Failed", logConclusionData, A_LineNumber, 'File "' . A_LoopFileName . '" in subfolder "' . subfolder . '" has invalid range value in parenthesis after percent. ' . validation)
                }
            }

            horizontalPercentRange := rangeParts[1]
            verticalPercentRange   := rangeParts[2]

            parts      := StrSplit(baseTextWithoutRanges, "@", " `t")
            resolution := ExtractRowFromArrayOfMapsOnHeaderCondition(system["Constants"]["Resolutions"], "Resolution", parts[1])["Counter"]
            scale      := ExtractRowFromArrayOfMapsOnHeaderCondition(system["Constants"]["Scales"], "Scale", parts[2])["Counter"]

            fileHash    := GetFileHash(A_LoopFileFullPath, "SHA-256")
            encodedHash := EncodeSha256HexToBase(fileHash, 86)
            base64Data  := GetBase64FromFile(A_LoopFileFullPath)

            extension := ""
            for rowMap in system["Mappings"]["File Signatures"] {
                maximumBase64Signature := rowMap["Maximum Base64 Signature"]
                if SubStr(base64Data, 1, StrLen(maximumBase64Signature)) = maximumBase64Signature {
                    extension  := rowMap["Extension"]
                    base64Data := SubStr(base64Data, StrLen(maximumBase64Signature) + 1)
                    break
                }
            }               

            hashToCounter := Map()
            if !hashToCounter.Has(fileHash) {
                hashToCounter[fileHash] := counter
                counter := counter + 1

                dataEntries.Push(actionName . "|" . actionLetter . "|" . hashToCounter[fileHash] . "|" . encodedHash . "|" . extension . "|" . base64Data)
            }
            
            catalogEntries.Push(imageLibraryDataReference . "|" . hashToCounter[fileHash] . "|" . resolution . "|" . scale . "|" . horizontalPercentRange . "|" . verticalPercentRange)
        }
    }

    WriteTextToFile(headerCatalog . ConvertArrayToLineSeparatedString(catalogEntries), imageLibraryCatalogFilePath, "UTF-8-BOM")

    if counterModified {
        WriteTextToFile(ConvertArrayToLineSeparatedString(dataEntries), imageLibraryDataReferenceFilePath, "UTF-8-BOM", "Append Break")
    } else {
        WriteTextToFile(headerData . ConvertArrayToLineSeparatedString(dataEntries), imageLibraryDataReferenceFilePath, "UTF-8-BOM")
    }

    LogConclusion("Completed", logConclusionData)
}

CreateImagesFromCatalog(imageLibraryCatalogName) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("imageCatalogName As String", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [imageLibraryCatalogName], "Create Images from Catalog")

    global imageRegistry

    screenWidth  := A_ScreenWidth
    screenHeight := A_ScreenHeight

    projectImageCatalogFilePath  := system["Directories"]["Project"] . "Image Library Catalog (" . imageLibraryCatalogName . ").csv"
    sharedImageCatalogFilePath   := system["Directories"]["Images"] . "Image Library Catalog (" . imageLibraryCatalogName . ").csv"
    catalogDirectory             := unset
    imageLibraryCatalogArray     := unset

    if !FileExist(projectImageCatalogFilePath) && !FileExist(sharedImageCatalogFilePath) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Image Library Catalog not found: " . imageLibraryCatalogName)
    }

    if FileExist(projectImageCatalogFilePath) && FileExist(sharedImageCatalogFilePath) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Image Library Catalog with the same name found in both Images and Project: " . imageLibraryCatalogName)
    }
    
    if FileExist(projectImageCatalogFilePath) && !FileExist(sharedImageCatalogFilePath) {
        catalogDirectory := system["Directories"]["Project"]

        imageLibraryCatalogHash    := GetFileHash(projectImageCatalogFilePath, "SHA-256")
        imageLibraryCatalogContent := ReadFileOnHashMatch(projectImageCatalogFilePath, imageLibraryCatalogHash)
        imageLibraryCatalogArray   := ParseDelimitedRowsToArrayOfMaps(imageLibraryCatalogContent)
    } else if !FileExist(projectImageCatalogFilePath) && FileExist(sharedImageCatalogFilePath) {
        catalogDirectory := system["Directories"]["Images"]
        imageLibraryCatalogHash    := GetFileHash(sharedImageCatalogFilePath, "SHA-256")
        imageLibraryCatalogContent := ReadFileOnHashMatch(sharedImageCatalogFilePath, imageLibraryCatalogHash)
        imageLibraryCatalogArray   := ParseDelimitedRowsToArrayOfMaps(imageLibraryCatalogContent)
    }

    static variants          := unset
    static displayResolution := unset
    static dpiScale          := unset

    if !IsSet(variants) {
        configurationImageVariantPresetHash    := GetFileHash(system["Configuration"]["Settings"]["Image Variant Preset"], "SHA-256")
        configurationImageVariantPresetContent := ReadFileOnHashMatch(system["Configuration"]["Settings"]["Image Variant Preset"], configurationImageVariantPresetHash)
        configurationImageVariantPresetArray   := ParseDelimitedRowsToArrayOfMaps(configurationImageVariantPresetContent)

        variants := Map()
        for name in configurationImageVariantPresetArray {
            variantName    := name["Name"]
            firstCharacter := StrLower(SubStr(variantName, 1, 1))

            variants[firstCharacter] := variantName
        }

        displayResolution := ExtractRowFromArrayOfMapsOnHeaderCondition(system["Constants"]["Resolutions"], "Resolution", system["Environment"]["Display Resolution"])["Counter"] . ""
        dpiScale          := ExtractRowFromArrayOfMapsOnHeaderCondition(system["Constants"]["Scales"], "Scale", system["Environment"]["DPI Scale"])["Counter"] . ""
    }

    relevantImages       := []
    uniqueDataReferences := []

    for image in imageLibraryCatalogArray {
        if IsInteger(image["Image Library Data Reference"]) {
            image["Image Library Data Reference"] := ExtractRowFromArrayOfMapsOnHeaderCondition(system["Mappings"]["Applications"], "Counter", image["Image Library Data Reference"])["Name"]

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
        LogConclusion("Skipped", logConclusionData)
    } else {
        pendingBase64ImageWriteQueue := []
        uniqueDataReferences         := RemoveDuplicatesFromArray(uniqueDataReferences)

        uniqueDataReferencesDirectories := uniqueDataReferences.Clone()
        for index, uniqueDataReferenceDirectory in uniqueDataReferencesDirectories {
            uniqueDataReferencesDirectories[index] := system["Directories"]["Images"] . uniqueDataReferenceDirectory . "\"
        }

        BatchAppendSymbolLedger("Reference", uniqueDataReferencesDirectories)

        uniqueImageLibraryDataReferences := []
        for uniqueDataReference in uniqueDataReferences {
            uniqueImageLibraryDataReferences.Push(catalogDirectory . "Image Library Data (" . uniqueDataReference . ").csv")
        }

        for uniqueDataReference in uniqueDataReferences {
            uniqueImageLibraryDataReferences.Push(GetFileHash(catalogDirectory . "Image Library Data (" . uniqueDataReference . ").csv", "SHA-256"))
        }

        BatchAppendSymbolLedger("Reference", uniqueImageLibraryDataReferences)

        for uniqueDataReference in uniqueDataReferences {
            EnsureDirectoryExists(system["Directories"]["Images"] . uniqueDataReference . "\")

            libraryDataEntriesHash    := GetFileHash(catalogDirectory . "Image Library Data (" . uniqueDataReference . ").csv", "SHA-256")
            libraryDataEntriesContent := ReadFileOnHashMatch(catalogDirectory . "Image Library Data (" . uniqueDataReference . ").csv", libraryDataEntriesHash)
            libraryDataEntriesArray   := ParseDelimitedRowsToArrayOfMaps(libraryDataEntriesContent)

            for image in relevantImages {
                for libraryData in libraryDataEntriesArray {
                    if image["Counter Reference"] = libraryData["Counter"] && uniqueDataReference = image["Image Library Data Reference"] {
                        if !libraryData.Has("Directory") {
                            libraryData["Directory"] := image["Image Library Data Reference"]
                            libraryData["Filename"]  := libraryData["Name"] . " (" . variants[libraryData["Variant"]] . ")." . libraryData["Extension"]
                            libraryData["SHA-256"]   := DecodeBaseToSha256Hex(libraryData["SHA-256"], 86)
                            libraryData["Base64"]    := ExtractRowFromArrayOfMapsOnHeaderCondition(system["Mappings"]["File Signatures"], "Extension", libraryData["Extension"])["Maximum Base64 Signature"] . libraryData["Base64"]
                            pendingBase64ImageWriteQueue.Push(libraryData)

                            if !imageRegistry.Has(libraryData["Directory"]) {
                                imageRegistry[libraryData["Directory"]] := Map()
                            }

                            if !imageRegistry[libraryData["Directory"]].Has(libraryData["Name"]) {
                                imageRegistry[libraryData["Directory"]][libraryData["Name"]] := []
                            }

                            path := system["Directories"]["Images"] . libraryData["Directory"] . "\" . libraryData["Filename"]

                            horizontalRange      := StrReplace(image["Horizontal Range"], ",", ".")
                            horizontalParts      := StrSplit(horizontalRange, "-")
                            horizontalRangeStart := Floor(screenWidth * horizontalParts[1] / 100)
                            horizontalRangeEnd   := Ceil(screenWidth * horizontalParts[2] / 100) - 1

                            verticalRange        := StrReplace(image["Vertical Range"], ",", ".")
                            verticalParts        := StrSplit(verticalRange, "-")
                            verticalRangeStart   := Floor(screenHeight * verticalParts[1] / 100)
                            verticalRangeEnd     := Ceil(screenHeight * verticalParts[2] / 100) - 1

                            imageRegistry[libraryData["Directory"]][libraryData["Name"]].Push(Map(
                                "Path",                   path,
                                "Name",                   libraryData["Name"],
                                "Variant",                StrLower(libraryData["Variant"]),
                                "Extension",              libraryData["Extension"],
                                "Horizontal Range",       horizontalRange,
                                "Horizontal Range Start", horizontalRangeStart,
                                "Horizontal Range End",   horizontalRangeEnd,
                                "Vertical Range",         verticalRange,
                                "Vertical Range Start",   verticalRangeStart,
                                "Vertical Range End",     verticalRangeEnd
                            ))

                            
                        }
                    }
                }
            }
        }

        references := []
        for image in pendingBase64ImageWriteQueue {
            references.Push(system["Directories"]["Images"] . image["Directory"] . "\" . image["Filename"])
        }

        for image in pendingBase64ImageWriteQueue {
            references.Push(image["SHA-256"])
        }

        BatchAppendSymbolLedger("Reference", references)

        for image in pendingBase64ImageWriteQueue {
            filePath := system["Directories"]["Images"] . image["Directory"] . "\" . image["Filename"]
            if FileExist(filePath) {
                if image["SHA-256"] != GetFileHash(filePath, "SHA-256") {
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

        LogConclusion("Completed", logConclusionData)
    }
}

; **************************** ;
; Helper Methods               ;
; **************************** ;

ExtractImageCoordinates(imageSearchResults) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("imageSearchResults As Map", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [imageSearchResults])

    if imageSearchResults["Success"] = false {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Image (" . imageSearchResults["Name"] . ")" . " not found in directory (" . imageSearchResults["Directory"] . "). Failed after " . imageSearchResults["Times Attempted"] . " attempts with " . imageSearchResults["Medium Delay"] . " milliseconds delay between each attempt.")
    }

    coordinatePair := imageSearchResults["Coordinate Pair"]

    return coordinatePair
}

GetImageDimensions(imagePath) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("imagePath As String [Constraint: Path]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [imagePath])

    static gdiPlusStartupInputSize := (A_PtrSize = 8) ? 24 : 16
    static gdiPlusStartupInputBuffer := Buffer(gdiPlusStartupInputSize, 0)
    NumPut("UInt", 1, gdiPlusStartupInputBuffer, 0)
    NumPut("Ptr", 0, gdiPlusStartupInputBuffer, (A_PtrSize= 8 ) ? 8:4)
    NumPut("UInt", 0, gdiPlusStartupInputBuffer, (A_PtrSize= 8 ) ? 16:8)
    NumPut("UInt", 0, gdiPlusStartupInputBuffer, (A_PtrSize= 8 ) ? 20:12)

    static gdiPlusLoadLibrary := DllCall("LoadLibrary", "Str", "GdiPlus", "Ptr")

    gdiPlusStartupToken  := 0
    gdiPlusStartupResult := DllCall("GdiPlus\GdiplusStartup", "Ptr*", &gdiPlusStartupToken, "Ptr", gdiPlusStartupInputBuffer.Ptr, "Ptr", 0, "UInt")
    if gdiPlusStartupResult != 0 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to initialize Windows GDI+. [GdiPlus\GdiplusStartup" . ", GDI+ Status Code: " . gdiPlusStartupResult . "]")
    }

    gdiPlusBitmapPointer := 0
    gdiPlusCreateBitmapStatus := DllCall("GdiPlus\GdipCreateBitmapFromFile", "WStr", imagePath, "Ptr*", &gdiPlusBitmapPointer, "UInt")
    if gdiPlusCreateBitmapStatus != 0 || !gdiPlusBitmapPointer {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to create a bitmap object based on the image. [GdiPlus\GdipCreateBitmapFromFile" . ", GDI+ Status Code: " . gdiPlusCreateBitmapStatus . "]")
    }

    gdiPlusGetImageWidthStatus := DllCall("GdiPlus\GdipGetImageWidth", "Ptr", gdiPlusBitmapPointer, "UInt*", &imageWidthPixels := 0, "UInt")
    if gdiPlusGetImageWidthStatus != 0 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to get the width in pixels of the image. [GdiPlus\GdipGetImageWidth" . ", GDI+ Status Code: " . gdiPlusGetImageWidthStatus . "]")
    }

    gdiPlusGetImageHeightStatus := DllCall("GdiPlus\GdipGetImageHeight", "Ptr", gdiPlusBitmapPointer, "UInt*", &imageHeightPixels := 0, "UInt")
    if gdiPlusGetImageHeightStatus != 0 {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Failed to get the height in pixels of the image. [GdiPlus\GdipGetImageHeight" . ", GDI+ Status Code: " . gdiPlusGetImageHeightStatus . "]")
    }

    DllCall("Gdiplus\GdipDisposeImage", "Ptr", gdiPlusBitmapPointer, "UInt")
    DllCall("Gdiplus\GdiplusShutdown", "UPtr", gdiPlusStartupToken, "UInt")

    imageDimensions := imageWidthPixels . "x" . imageHeightPixels

    return imageDimensions
}

OverrideDirectoryImageVariant(directoryFolder, imageName, variant, horizontalRange, verticalRange) {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("directoryFolder As String, imageName As String, variant As String, horizontalRange As String [Constraint: Percent Range], verticalRange As String [Constraint: Percent Range]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [directoryFolder, imageName, variant, horizontalRange, verticalRange])

    global imageRegistry

    static displayResolution := StrSplit(system["Environment"]["Display Resolution"], "x")
    static screenWidth  := displayResolution[1] + 0
    static screenHeight := displayResolution[2] + 0

    if !imageRegistry.Has(directoryFolder) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Directory folder for image not found: " . directoryFolder)
    }

    if !imageRegistry[directoryFolder].Has(imageName) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Image name not found: " . imageName)
    }

    variant := StrLower(variant)
    variantFound := false
    for image in imageRegistry[directoryFolder][imageName] {
        if variant = image["Variant"] {
            variantFound := true
        }
    }

    if !variantFound {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Variant not found: " . variant)
    }

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
            image["Horizontal Range"]       := horizontalRange
            image["Horizontal Range Start"] := horizontalRangeStart
            image["Horizontal Range End"]   := horizontalRangeEnd
            image["Vertical Range"]         := verticalRange
            image["Vertical Range Start"]   := verticalRangeStart
            image["Vertical Range End"]     := verticalRangeEnd

            break
        }
    }
}

SearchForDirectoryImage(directoryFolder, imageName, timesToAttempt := 60, variant := "") {
    static qpcPreBuffer    := Buffer(8, 0)
    static timestampBuffer := Buffer(8, 0)
    static qpcPostBuffer   := Buffer(8, 0)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPreBuffer.Ptr, "Int")
    DllCall("Kernel32\GetSystemTimeAsFileTime", "Ptr", timestampBuffer.Ptr)
    DllCall("Kernel32\QueryPerformanceCounter", "Ptr", qpcPostBuffer.Ptr, "Int")

    static methodName := RegisterMethod("directoryFolder As String, imageName As String, timesToAttempt As Integer [Optional: 60], variant As String [Optional]", A_ThisFunc, A_LineFile, A_LineNumber + 1)
    logConclusionData := LogBeginning(methodName, NumGet(qpcPreBuffer, 0, "Int64"), NumGet(timestampBuffer, 0, "Int64"), NumGet(qpcPostBuffer, 0, "Int64"), [directoryFolder, imageName, timesToAttempt, variant])

    static defaultMethodSettingsSet := unset
    if !IsSet(defaultMethodSettingsSet) {
        ConfigureMethodSetting(methodName, "Medium Delay", 1000, 100, 10000)

        defaultMethodSettingsSet := true
    }

    settings := methodRegistry[methodName]["Settings"]
    
    mediumDelay := settings["Medium Delay"].Get("Value")

    if !imageRegistry.Has(directoryFolder) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Directory folder for image not found: " . directoryFolder)
    }

    if !imageRegistry[directoryFolder].Has(imageName) {
        LogConclusion("Failed", logConclusionData, A_LineNumber, "Image name not found: " . imageName)
    }

    if timesToAttempt = 0 {
        timesToAttempt := 1
    }

    variant := StrLower(variant)
    if variant != "" {
        variantFound := false
        for image in imageRegistry[directoryFolder][imageName] {
            if variant = image["Variant"] {
                variantFound := true
            }
        }

        if !variantFound {
            LogConclusion("Failed", logConclusionData, A_LineNumber, "Variant not found: " . variant)
        }
    }

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

    imageSearchResults := Map(
        "Directory",        directoryFolder,
        "Name",             imageName,
        "Times to Attempt", timesToAttempt,
        "Medium Delay",     mediumDelay,
        "Success",          false
    )

    overlayVisibility := OverLayIsVisible()
    if overlayVisibility {
        OverlayChangeVisibility()
    }

    horizontalCoordinate := 0
    verticalCoordinate   := 0

    CoordMode("Pixel", "Screen")
    Loop timesToAttempt {
        imageSearchResults["Times Attempted"] := A_Index

        for image in directoryImageVariants {
            if ImageSearch(&horizontalCoordinate, &verticalCoordinate, image["Horizontal Range Start"], image["Vertical Range Start"], image["Horizontal Range End"], image["Vertical Range End"], image["Path"]) {
                horizontalCoordinate := horizontalCoordinate + Floor(image["Width"] / 2)
                verticalCoordinate   := verticalCoordinate + Floor(image["Height"] / 2)

                imageSearchResults["Variant"]         := image["Variant"]
                imageSearchResults["Coordinate Pair"] := horizontalCoordinate "x" verticalCoordinate
                imageSearchResults["Success"]         := true

                break 2
            }
        }

        Sleep(mediumDelay)
    }

    if overlayVisibility {
        OverlayChangeVisibility()
    }

    return imageSearchResults
}