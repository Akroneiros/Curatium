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

Public importWorkbook As Workbook
Public mainWorkbook As Workbook
Public cellStyles As Object
Public environment As Object
Public international As Object
Public methodRegistry As Object
Public report As Object
Public telemetry As Object

Private Declare PtrSafe Sub GetSystemTime Lib "Kernel32" (ByRef systemTime As SystemTimeStructure)
Private Declare PtrSafe Function GetTickCount64 Lib "Kernel32" () As Currency
Private Declare PtrSafe Function QueryPerformanceCounter Lib "Kernel32" (ByRef queryPerformanceCounterValue As Currency) As Long

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