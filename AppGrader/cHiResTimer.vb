'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'END
'Attribute VB_Name = "cHiResTimer"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False

Option Explicit On

Public Class cHIResTimer

    Private Structure LARGE_INTEGER
        Dim lowpart As Long
        Dim highpart As Long
    End Structure

    Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" _
         (ByRef lpPerformanceCount As Decimal) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" _
          (ByRef lpFrequency As Decimal) As Long

    Private period As Decimal
    Private startTime As Decimal
    Private timerFrequency As Decimal
    Private bhasHiResCounter As Boolean

    Public Sub StartTimer()
        Dim lR As Long
        lR = QueryPerformanceCounter(startTime)
    End Sub

    Public Sub StopTimer()
        Dim endTime As Decimal    ' currency
        Dim lR As Long
        lR = QueryPerformanceCounter(endTime)
        period = endTime - startTime
    End Sub

    Property ElapsedTime As Double
    Property HasHiResCounter As Boolean
    Property Frequency As Decimal

    'Public Property Get ElapsedTime() As Double
    '   ElapsedTime = period / (timerFrequency * 1#)
    'End Property

    'Public Property Get HasHiResCounter() As Boolean
    '   HasHiResCounter = bhasHiResCounter
    'End Property

    'Public Property Get Frequency() As Currency
    '   Frequency = timerFrequency
    'End Property

    Private Sub New()
        ' If the installed hardware supports a high-resolution performance counter,
        ' the return value is nonzero.
        ' If the function fails, the return value is zero. To get extended error
        ' information, call GetLastError. For example, if the installed hardware
        ' does not support a high-resolution performance counter, the function fails.
        Dim r As Long
        r = QueryPerformanceFrequency(timerFrequency)
        If (r <> 0) Then
            bhasHiResCounter = True
        End If
    End Sub
End Class
