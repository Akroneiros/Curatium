#Requires AutoHotkey v2.0
#Include Base Library.ahk
#Include File Library.ahk
#Include Logging Library.ahk

AssignSharedImages() {
    static methodName := RegisterMethod("AssignSharedImages()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Assign Shared Images", methodName)

    SplitPath(A_LineFile, , &libraryFileDirectory)
    SplitPath(libraryFileDirectory, , &parentDirectory)
    imageDirectory := parentDirectory . "\Images\"

    filesInDirectory := GetFileListFromDirectory(imageDirectory)
    sharedImages := Map()

    for index, filePath in filesInDirectory {
        SplitPath(filePath, , , , &filenameNoExtension)
        sharedImages[filenameNoExtension] := filePath
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
    return sharedImages
}

CreateSharedImages() {
    static methodName := RegisterMethod("CreateSharedImages()" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Create Shared Images", methodName)

    SplitPath(A_LineFile, , &libraryFileDirectory)
    SplitPath(libraryFileDirectory, , &parentDirectory)
    imageDirectory := parentDirectory . "\Images\"

    if !DirExist(imageDirectory) {
        DirCreate(imageDirectory)
    }

    sharedImages := []
    variations := AssignHeroAliases()

    sharedImages.Push(["SMMS Query Successful", variations["a"], "5a48ec471fc18a17f25f5b8c99f9429f3269b74624f620dc39dfc94f35dffe41", "
    (
    iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAALHRFWHRDcmVhdGlvbiBUaW1lAFR1ZSAyMiBKdWwgMjAyNSAxMToxODoxNiArMDEwMJf/wuMAAAAHdElN
    RQfpBxYJEjEcNn/fAAAACXBIWXMAAAsSAAALEgHS3X78AAAABGdBTUEAALGPC/xhBQAAATZJREFUeNpj/PCshwEGmNk0GNnUmVik4SL//jz9/+vm31834CIsEIqRiZeF
    2/PCm8err089/PgwXNpW1jZUM9RAJOzP1+3//30GqQTaAFTNyhvedrxn+93tDNgAUE+eccbvzyuBekA2AM2edHYGLtVAsPr6aiCZaxj1+/MqJqC7wS5ZjanOV833TNqZ
    zZGbIXqAyoCKmYC+xKoaCNKM04DkrLOz4PaAggQYJsi+RDZeklfy+efnm29thogAlQEVM+FyN5rx6MEKAUDnAk1tPNAIZKMZDwcoNkDMA5qNy3iQBmBcAmMHasOtzUBT
    gWZjNR6oDKiYCRjzwHhBswSr8UBlQMWgmGblDWs/OQtPxEFUQyMOyAGmkyrLEmR7MFUDkwZQGQMkLTGgJL7V2BKfLErig0sTk7wB/hu0pZBeGB4AAAAASUVORK5CYII=
    )"])

    sharedImages.Push(["Toad for Oracle Play", variations["a"], "df3b39143bd11435f4937dcf2835512ffc8753d353bf4d9d2de2d6f41f9bacd1", "
    (
    iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAK3RFWHRDcmVhdGlvbiBUaW1lAFNhdCA5IEF1ZyAyMDI1IDIxOjAwOjQyICswMTAwmrqrOQAAAAd0SU1F
    B+kICRMCAwW4yrAAAAAJcEhZcwAACxIAAAsSAdLdfvwAAAAEZ0FNQQAAsY8L/GEFAAAAvUlEQVR42pVRwRHDIAyzM0HWgNnYIDzJBsyUFbJGNqAipq5D0h7Vg+OMZMmG
    DoOUEhHhPL6DLA9nKaT3R+WEhxij8CxQQZ1uYIhyzrgty+V5XatJCKFXiLu01CRqeE/FmrvacR1pnueuYg0m9ECVufUDW3y0iMplBk31OAPmdoV267Ntmx1A16qpIHDn
    va3Vey+p+vXxZx8ia9l+/7Q6qA93S+iAru6dbT89JxqDsPGPQwJl13/4iz0ksGzgBRovzTfB8AnSAAAAAElFTkSuQmCC
    )"])
    
    sharedImages.Push(["Toad for Oracle Play", variations["b"], "071ce63fd7a4f0862d527aca77762f0ffb776891b02b832d789090cd31fe34ae", "
    (
    iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAK3RFWHRDcmVhdGlvbiBUaW1lAFNhdCA5IEF1ZyAyMDI1IDIyOjI4OjI5ICswMTAwqKNjvwAAAAd0SU1F
    B+kICRQdFqBwNkAAAAAJcEhZcwAACxIAAAsSAdLdfvwAAAAEZ0FNQQAAsY8L/GEFAAABwUlEQVR42mP88OEDAwz8/HTmxulWIENeM15AKoABG2AC4reP15/eEfD13fGr
    x1vs7O3t7BweXF0C5AIFgVJYNNy+uMDF2fTro9a//xkYGBkYft1iZhf+9qjZxUETKIWmgQWIf//5/+XlIRVFEVVNVYY/Hxh+v7DRY/jPyPflzbnff9jRbfjwbAMvN9uF
    ewqM/38wfL98/8a2+w8f/2dkZeTQOHGFk5uTCegxZA2MB9c6Ah0NdAbQ4Ou3v0mZbgeKvrtopigvy8Bj+///v8OHDus7r0dxEgPjfwYOVQYWUQaGkyjW/7wtNfUKkL7h
    jOQHDdPq/XtbgN7VVXooxA8yGyjKxfn75at7t54LA9kzqhs1Wu1vlB2E+oGdz+TRs1/GOozMLP+ZGRl4uX6Li7Py8nEBuRoyN4AqArgCQHq67BHB+vvXf1aGe8dOc206
    rMcv+A9o/Kyl3EAuJ/tHiCJkPSANumZJc1aKq5p1vXrxkpWV5fu3P9++/QNygYJwp8P1MCInjXPHVp49OBXIcAvtlFexBDKAKl60XoTIcl/j5l2qwoIcKkZW4UCENQlB
    VAO9zsRABICrhvqBeNVEaUBWDQQABnPEiZDQDGsAAAAASUVORK5CYII=
    )"])

    sharedImages.Push(["Toad for Oracle Play", variations["c"], "f6e71e64fd851f514fff4c045aee86786ae1aee87383522b1ff84c955d5536af", "
    (
    iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAALHRFWHRDcmVhdGlvbiBUaW1lAHRpci4gNSBhdWcgMjAyNSAyMTowNzowOSArMDEwMEXL+RcAAAAHdElN
    RQfpCAUTBynpokibAAAACXBIWXMAAAsSAAALEgHS3X78AAAABGdBTUEAALGPC/xhBQAAAcFJREFUeNpj/PDhAwMM/Px05sbpViBDXjNeQCqAARtgAuK3j9ef3hHw9d3x
    q8db7Ozt7ewcHlxdAuQCBYFSWDTcvrjAxdn066PWv/8ZGBgZGH7dYmYX/vao2cVBEyiFpoEFiH//+f/l5SEVRRFVTVWGPx8Yfr+w0WP4z8j35c2533/Y0W348GwDLzfb
    hXsKjP9/MHy/fP/GtvsPH/9nZGXk0DhxhZObkwnoMWQNjAfXOgIdDXQG0ODrt79JmW4Hir67aKYoL8vAY/v//7/Dhw7rO69HcRID438GDlUGFlEGhpMo1v+8LTX1CpC+
    4YzkBw3T6v17W4De1VV6KMQPMhsoysX5++Wre7eeCwPZM6obNVrtb5QdhPqBnc/k0bNfxjqMzCz/mRkZeLl+i4uz8vJxAbkaMjeAKgK4AkB6uuwRwfr7139WhnvHTnNt
    OqzHL/gPaPyspdxALif7R4giZD0gDbpmSXNWiquadb168ZKVleX7tz/fvv0DcoGCcKfD9TAiJ41zx1aePTgVyHAL7ZRXsQQygCpetF6EyHJf4+ZdqsKCHCpGVuFAhDUJ
    QVQDvc7EQASAq4b6gXjVRGlAVg0EAAZzxImQ0AxrAAAAAElFTkSuQmCC
    )"])

    sharedImages.Push(["Toad for Oracle Search", variations["a"], "004f04d4d0f035d1dd14e579fc4f6fbc88184309f6a0bbebf2cf116f0d5efb81", "
    (
    iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAK3RFWHRDcmVhdGlvbiBUaW1lAFNhdCA5IEF1ZyAyMDI1IDIxOjAwOjMyICswMTAwkH+iIAAAAAd0SU1F
    B+kICRMCHIiwx0UAAAAJcEhZcwAACxIAAAsSAdLdfvwAAAAEZ0FNQQAAsY8L/GEFAAABK0lEQVR42mM8fuc7AwzsXdFbU1PT0tLiHFHMgAMwJ+bU7F/ZGx3grCjGCVT9
    /z+Dvf0+IBsiIq9lBRRBRkx///2DqEMDQBGgOFAWDTHuu/b12Lp+oIrq6hpkDa2tLUDSKqgQ3SSgi+FGQthAEm4hkL37yhdkBFUKQUCw4+InNBFko0GyBG04duzY////
    gSTCOS1ggBYaEDm4argeBggFVwExG24Jsur/YCGmZ8yaQGlGRvTAgIgcP34cjWRceewdXNHtfdMgMa3qlBVuJQQ0kBFsCcwIRpDRy468xYz/KBvh/2AHMSJZDVSt5JDJ
    8vcfuupYOxTViw8hTAQqZvmLqiPBURSiuhWsesH+12gKgDZgSUZA1cB0MnfvK0xZdCfN2v0K4hIgA9O1QAAAinT9pDUl/fgAAAAASUVORK5CYII=
    )"])

    sharedImages.Push(["Toad for Oracle Search", variations["b"], "e347e6201e7234a73255373f54e08782faa2d898300d095a92b5090942ae4767", "
    (
    iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAAK3RFWHRDcmVhdGlvbiBUaW1lAFNhdCA5IEF1ZyAyMDI1IDIyOjI5OjUwICswMTAw2DxESwAAAAd0SU1F
    B+kICRQeF/xaVRUAAAAJcEhZcwAACxIAAAsSAdLdfvwAAAAEZ0FNQQAAsY8L/GEFAAACSklEQVR42mM8fuc7AwwosR29cboVyJDXjH/OFs6ADTD9/ccg9n3FlzOeckwH
    rh5vsbO3t7NzeHB1CZALFARKARUgI6CGf7cvLnBxNv36qPXvfwYGRgaGX7eY2YW/PWp2cdAESgEVICMWoKbff/5/eXlIRVFEVVOV4c8Hht8vbPQY/jPyfXlz7vcfdrZ/
    qE5SYljNy8124Z4C4/8fDN8v37+x7f7Dx/8ZWRk5NE5c4eTmZNLmOf733384Yjy41hHoaKAzgAZfv/3tkdhOoDFqH0wU5WUZeGz///93+NDh7yobETaACMb/DByqDJy6
    QCbEZ1Dw8zbDfya4IAQxvny0Bxg4QO/qKj0Eyn37zgokuTh/A8lbz03//XqnoB1z6XsgwoZjL00ePftlrMPIzPKfmZGBl+u3uDgrLx8XkMuo0Hfoa2veQuFLl28Ks36F
    +AEYrP9///rPynDv2GmuTYf1+AX/AY2ftZT7BtORWRuviMpIlme7a2pIz9lyTZT9G1jD3/+6ZklzVoqrmnW9evGSlZXl+7c/377923vmoa2Vurqx7BsWJjlNKS8H1X1n
    HwMVg2y4xxok7r79wldLc4fEnmncUxfw+EQ2n7z2wsZYVoTvn7A44/3v/0SVJE5dfQEK1sWH3mJNM7du3jDSUzIxlfjE8P/lJ8Z7F148uXtfWVUd6KR/WJGridzqrecv
    nH3C+vXPi+vPDh84bfVrEVCcce7eVww4gKrQz92nH56+9sJUW5L98CRvzrMqLmqMs3bj1IAGuOZ5iGsZAwBn21E5n+dUaQAAAABJRU5ErkJggg==
    )"])

    sharedImages.Push(["Toad for Oracle Search", variations["c"], "8ac48d9c7f84a1ba3265bf253ea52254b2c402b862ceaa57dae32226617e965d", "
    (
    iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAIAAACQkWg2AAAALHRFWHRDcmVhdGlvbiBUaW1lAFdlZCAzMCBKdWwgMjAyNSAxOTo1NDo1NiArMDEwMFf0WbIAAAAHdElN
    RQfpBx4ROg25hXfVAAAACXBIWXMAAAsSAAALEgHS3X78AAAABGdBTUEAALGPC/xhBQAAAiFJREFUeNpj/P//PwMMvHq898bpViBDXjMeiBiwARYgvndlwb2rC7XMym+e
    7nR0dmD4z3jwwBJWDvFrpzqVtOOVdBKQNTAB8e2LC1ycTb8+av0LtIyRgeHXLWZ24W+Pml0cNIFSWGz4/ef/l5eHVBRFVDVVGf58YPj9wkYPaA3flzfnfv9hR9PA9PD6
    Ql5utgv3FBj//2D4fvn+jW33Hz7+z8jKyKFx4gonNycT0GMoNgA12Ds4AJ3B8Jvh+u1vmi43gaL3D6kryv92cXH4///f4UOtYrLOKH5gYPzPwKHKwKmLHiI/bzP8Z0IT
    Y3z5aM/V4y1A7+oqPQTyv31nBZJcnL+B5K3npv9+vVPQjkEOYkZgPCycYO3nyv73132Gv4wg1bxcIJ3fvt5gPrL//NMz11+ZaYo7GslY6UpCnfT7139WhnvHTnNtOqzH
    L/gPaPyspdw3mI7M2nhFVEayPNtdU0N6zpZrRy8/h2rQNUuas1Jc1azr1YuXrKws37/9+fbt394zD22t1NWNZd+wMMlpSnk5qO47+xjkpv9IYN+2md3lekB08cx2z6K1
    197/ufT375G/f1e9+HvgxU+fko1ANYzIaQkZ1M05ZqSnZGIq8Ynh/8tPjPcuvHhy9359kiUTAw7gaiK3euv5C2efsH798+L6s8MHTlv9WgQNJVx6Dl94svv0w9PXXphq
    S7IfnuTNeVbFRQ2fBjSwNMpYXMsYANyNBTDVex/0AAAAAElFTkSuQmCC
    )"])

    filenameValues := []
    for index, image in sharedImages {
        filenameValues.Push(image[1] . " (" . image[2] . ").png")
    }

    SymbolLedgerBatchAppend("F", filenameValues)

    for image in sharedImages {
        WriteBase64IntoImageFileWithHash(image[4], imageDirectory . image[1] . " (" . image[2] . ").png", image[3])
    }

    LogInformationConclusion("Completed", logValuesForConclusion)
}

RetrieveImageCoordinatesFromSegment(imageAlias, horizontalPercentRange, verticalPercentRange, secondsOfWaitingBeforeFailure := 60) {
    static methodName := RegisterMethod("RetrieveImageCoordinatesFromSegment(imageAlias As String, horizontalPercentRange As String [Type: Percent Range], verticalPercentRange As String [Type: Percent Range], secondsOfWaitingBeforeFailure As Integer [Optional: 60])" . LibraryTag(A_LineFile), A_LineNumber + 1)
    logValuesForConclusion := LogInformationBeginning("Retrieve Image Coordinates from Segment", methodName, [imageAlias, horizontalPercentRange, verticalPercentRange, secondsOfWaitingBeforeFailure])

    coordinatePair := ""

    horizontalParts := StrSplit(Trim(horizontalPercentRange), "-")
    horizontalPercentStart := horizontalParts[1] + 0
    horizontalPercentEnd   := horizontalParts[2] + 0

    try {
        if horizontalPercentStart > horizontalPercentEnd {
            throw Error("Wrong order horizontally. " . "End value of " . horizontalPercentEnd . " is smaller than start value of " . horizontalPercentStart . ".")
        }
    } catch as horizontalWrongOrderError {
        LogInformationConclusion("Failed", logValuesForConclusion, horizontalWrongOrderError)
    }

    try {
        if horizontalPercentStart >= 100 {
            throw Error("Horizontal start value of " . horizontalPercentStart . " is invalid. Maximum allowed is 99.")
        }
    } catch as invalidHorizontalValueError {
        LogInformationConclusion("Failed", logValuesForConclusion, invalidHorizontalValueError)
    }

    verticalParts := StrSplit(Trim(verticalPercentRange), "-")
    verticalPercentStart := verticalParts[1] + 0
    verticalPercentEnd   := verticalParts[2] + 0

    try {
        if verticalPercentStart > verticalPercentEnd {
            throw Error("Wrong order vertically. " . "End value of " . verticalPercentEnd . " is smaller than start value of " . verticalPercentStart . ".")
        }
    } catch as verticalWrongOrderError {
        LogInformationConclusion("Failed", logValuesForConclusion, verticalWrongOrderError)
    }

    try {
        if verticalPercentStart >= 100 {
            throw Error("Vertical start value of " . verticalPercentStart . " is invalid. Maximum allowed is 99.")
        }
    } catch as invalidVerticalValueError {
        LogInformationConclusion("Failed", logValuesForConclusion, invalidVerticalValueError)
    }

    CoordMode("Pixel", "Screen")
    screenWidth  := A_ScreenWidth
    screenHeight := A_ScreenHeight

    if horizontalPercentStart = 0 {
        regionLeftPixel := 0
    } else {
        regionLeftPixel := Round(screenWidth  * horizontalPercentStart / 100)
    }

    if horizontalPercentEnd = 100 {
        regionRightPixel := screenWidth - 1
    } else {
       regionRightPixel := Round(screenWidth  * horizontalPercentEnd / 100) - 1
    }
    
    if verticalPercentStart = 0 {
        regionTopPixel := 0
    } else {
        regionTopPixel := Round(screenHeight * verticalPercentStart / 100)
    }
    
    if verticalPercentEnd = 100 {
        regionBottomPixel := screenHeight - 1
    } else {
        regionBottomPixel := Round(screenHeight * verticalPercentEnd / 100) - 1
    }

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

    startTime    := A_Now
    endTime      := DateAdd(startTime, secondsOfWaitingBeforeFailure, "Seconds")
    headerBuffer := Buffer(24, 0)
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

    while (A_Now < endTime) {
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

        Sleep 800
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