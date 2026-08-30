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

Public baseQpc As Double
Public baseTickCount As Double
Public baseUtcTimestamp As String
Public helperColumn As String
Public operationSequenceNumber As Double
Public cellStyles As Object
Public environment As Object
Public international As Object
Public methodRegistry As Object
Public report As Object
Public telemetry As Object
Public importWorkbook As Workbook
Public mainWorkbook As Workbook
Public aboutWorksheet As Worksheet
Public logWorksheet As Worksheet
Public runStatusWorksheet As Worksheet

Private Declare PtrSafe Sub GetSystemTime Lib "Kernel32" (ByRef systemTime As SystemTimeStructure)
Private Declare PtrSafe Function GetTickCount64 Lib "Kernel32" () As Currency
Private Declare PtrSafe Function QueryPerformanceCounter Lib "Kernel32" (ByRef queryPerformanceCounterValue As Currency) As Long
Private Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal millisecondsToPause As Long)

Private Type LogEntry
    operationSequenceNumber As Long
    methodName As String
    arguments As String
    categoryName As String
	tickCount As Double
    outcome As String
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