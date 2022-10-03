Attribute VB_Name = "CODE1"
Public Declare Function QueryPerformanceCounter Lib "Kernel32" _
                                 (X As Currency) As Boolean
Public Declare Function QueryPerformanceFrequency Lib "Kernel32" _
                                 (X As Currency) As Boolean
Public Declare Function QueryPerformanceFrequencyAny Lib "Kernel32" Alias _
    "QueryPerformanceFrequency" (lpFrequency As Any) As Long

Public Declare Function QueryPerformanceCounterAny Lib "Kernel32" Alias _
    "QueryPerformanceCounter" (lpPerformanceCount As Any) As Long


