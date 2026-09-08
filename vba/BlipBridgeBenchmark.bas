Attribute VB_Name = "BlipBridgeBenchmark"
Option Explicit
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long

' Run on disposable shapes. Timings include the wrapper but exclude name lookup.
Public Sub BenchmarkVBA(ByVal target As Shape, ByVal donor As Shape, ByVal path As String, Optional ByVal count As Long = 1000)
    Dim a As Currency, b As Currency, frequency As Currency, i As Long, handle As Long
    QueryPerformanceFrequency frequency
    handle = BB_RegisterTextureShape(donor)
    QueryPerformanceCounter a
    For i = 1 To count
        target.Fill.UserPicture path
    Next
    QueryPerformanceCounter b
    Debug.Print "VBA UserPicture total ms", (b - a) / frequency * 1000
    QueryPerformanceCounter a
    For i = 1 To count
        BB_ApplyTexture target, handle
    Next
    QueryPerformanceCounter b
    Debug.Print "VBA fallback wrapper total ms", (b - a) / frequency * 1000
    BB_ReleaseTexture handle
End Sub
