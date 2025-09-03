Public Const frameworkDemonstrationVersion As String = "Framework Demonstration (v1, 2025-09-02)"

Sub Startup()
    Call CreateAbout()
End Sub

Sub Master()
    Dim originalWorkbook As String: originalWorkbook = mainWorkbook.FullName
    checkpointType = "Foundation" ' Foundation, Augmentation. '

    Dim defaultIntermissionStates As String: defaultIntermissionStates = "Reset View|Covert Mode|Save Workbook|Quit Excel"
    ' Intermission States: Break Script, Covert Mode, Duplicate Workbook, Open Workbook, Quit Excel, Reset View, Save Workbook, Testing Mode. '

    Dim reportDetails As String
    Dim reportVision As String
    Dim dependenciesList As String
    Dim retrievedDate As String

    Dim checkpointsArray() As String
    checkpointsArray = Split("Demonstration", "|")

    Dim index As Byte
    For index = 0 To UBound(checkpointsArray)
        Call CreateAbout()

        If CheckpointIsNew(checkpointsArray(index)) Then

            If checkpointsArray(index) = "Demonstration" Then
                reportDetails = frameworkDemonstrationVersion
                reportVision = "Demonstration of Spreadsheet Operations Template used in conjunction with AutoHotkey v2."
                dependenciesList = ""
                retrievedDate = ""

                Call ConfigureAbout(reportDetails, reportVision, dependenciesList, retrievedDate)
                Call SaveWorkbook(reportDetails & " " & Format(Now, "YYYY-MM-DD"), "C:\Export")
            End If

            Call ConfigureStyles()
            Call ConfigureLogging()

            Call LogCheckpoint(checkpointType, checkpointsArray(index), "Beginning")

            Call Demonstration(checkpointsArray(index))

            Call LogCheckpoint(checkpointType, checkpointsArray(index), "Conclusion")

            If checkpointsArray(index) = "Demonstration" Then
                Call Intermission(defaultIntermissionStates, checkpointsArray(index))
            End If
        End If
    Next index
End Sub

Sub Demonstration(checkpointValidation As String)
If checkpointValidation <> "Demonstration" Then Exit Sub

    ' Code '

End Sub