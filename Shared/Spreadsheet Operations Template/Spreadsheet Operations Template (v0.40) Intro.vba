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
Public methodRegistry As Object
Public report As Object

Private Declare PtrSafe Function GetTickCount64 Lib "Kernel32" () As Currency
Private Declare PtrSafe Function QueryPerformanceCounter Lib "Kernel32" (ByRef queryPerformanceCounterValue As Currency) As Long

Private Type LogEntry
    operationSequenceNumber As Long
    methodName As String
    arguments As String
	tickCount As Double
End Type