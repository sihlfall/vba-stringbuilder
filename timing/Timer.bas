Attribute VB_Name = "Timer"
Private Declare Function getFrequency Lib "kernel32" _
Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
Private Declare Function getTickCount Lib "kernel32" _
Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

' Source: https://www.soa.org/news-and-publications/newsletters/compact/2014/january/com-2014-iss50/finding-excel-vba-bottlenecks/

Function MicroTimer() As Double

' Returns seconds.
Dim cyTicks1 As Currency
Static cyFrequency As Currency
' Initialize MicroTimer
MicroTimer = 0
' Get frequency.
If cyFrequency = 0 Then getFrequency cyFrequency
' Get ticks.
getTickCount cyTicks1
' Seconds = Ticks (or counts) divided by Frequency
If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency

End Function

