' Alteration  '
' Background  '
' Conjuration '
' Destruction '
' Elementals  '
' Formatting  '
' Logging     '
' Repetition  '
' Sequencing  '
' Validation  '

Option Compare Text
Option Explicit

Public baseTickCount As Double
Public depth As Long
Public helperColumn As String
Public operationSequenceNumber As Long
Public cellStyles As Object
Public constants As Object
Public environment As Object
Public international As Object
Public mappings As Object
Public methodRegistry As Object
Public report As Object
Public telemetry As Object
Public mainWorkbook As Workbook
Public aboutWorksheet As Worksheet
Public logWorksheet As Worksheet
Public runStatusWorksheet As Worksheet

Private Declare PtrSafe Function CloseHandle Lib "Kernel32" (ByVal objectHandle As LongPtr) As Long
Private Declare PtrSafe Function CreateFileW Lib "Kernel32" (ByVal filenamePointer As LongPtr, ByVal desiredAccess As Long, ByVal shareMode As Long, _
    ByVal securityAttributesPointer As LongPtr, ByVal creationDisposition As Long, ByVal flagsAndAttributes As Long, ByVal templateFileHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFileSizeEx Lib "Kernel32" (ByVal fileHandle As LongPtr, ByRef fileSize As Currency) As Long
Private Declare PtrSafe Sub GetSystemTime Lib "Kernel32" (ByRef systemTime As SystemTimeStructure)
Private Declare PtrSafe Function GetTickCount64 Lib "Kernel32" () As Currency
Private Declare PtrSafe Function QueryPerformanceCounter Lib "Kernel32" (ByRef queryPerformanceCounterValue As Currency) As Long
Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal millisecondsToPause As Long)

Private Type LogEntry
    operationSequenceNumber As Long
    methodName As String
    arguments As String
	tickCount As Double
End Type

Private Type SystemTimeStructure
    year As Integer
    month As Integer
    dayOfWeek As Integer
    day As Integer
    hour As Integer
    minute As Integer
    second As Integer
    milliseconds As Integer
End Type