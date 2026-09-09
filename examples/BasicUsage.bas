Attribute VB_Name = "BasicUsage"
'==============================================================================
' BlipBridge - smallest useful example.
'
' Import BlipBridge.bas and this module, put BlipBridge.dll next to the
' presentation, then run Demo.
'==============================================================================
Option Explicit

''' Reads a file into a Byte array. Any source of bytes works - a file, a
''' resource, something generated at run time - because LoadTexture only ever
''' sees bytes.
Public Function LoadFileBytes(ByVal path As String) As Byte()
    Dim handle As Integer
    handle = FreeFile
    Open path For Binary Access Read As #handle
    Dim bytes() As Byte
    ReDim bytes(0 To LOF(handle) - 1)
    Get #handle, , bytes
    Close #handle
    LoadFileBytes = bytes
End Function

''' Fills every AutoShape on slide 1 from a single decode.
Public Sub Demo()
    If Not BlipBridge.IsAvailable Then
        MsgBox "BlipBridge is not available here." & vbCrLf & BlipBridge.Version
        Exit Sub
    End If

    Dim bytes() As Byte
    bytes = LoadFileBytes(ActivePresentation.path & "\texture.png")

    Dim tex As LongLong
    tex = BlipBridge.LoadTexture(bytes)          ' decode once

    On Error GoTo Cleanup
    Dim shp As Shape
    For Each shp In ActivePresentation.Slides(1).Shapes
        If shp.Type = msoAutoShape Or shp.Type = msoFreeform Then
            BlipBridge.ApplyTexture shp, tex     ' cheap, repeatable
        End If
    Next shp

Cleanup:
    BlipBridge.ReleaseTexture tex
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub

''' Two textures alternating over many Shapes - the animation-shaped workload.
''' Note the textures are loaded once, outside the loop.
Public Sub AnimateTwoFrames(Optional ByVal frames As Long = 60)
    BlipBridge.Initialize

    Dim a As LongLong, b As LongLong
    a = BlipBridge.LoadTexture(LoadFileBytes(ActivePresentation.path & "\frame1.png"))
    b = BlipBridge.LoadTexture(LoadFileBytes(ActivePresentation.path & "\frame2.png"))

    On Error GoTo Cleanup
    Dim targets As Object
    Set targets = ActivePresentation.Slides(1).Shapes

    Dim frame As Long, shp As Shape
    For frame = 1 To frames
        For Each shp In targets
            If shp.Type = msoAutoShape Then
                BlipBridge.ApplyTexture shp, IIf(frame Mod 2 = 0, a, b)
            End If
        Next shp
        DoEvents
    Next frame

Cleanup:
    BlipBridge.ReleaseTexture a
    BlipBridge.ReleaseTexture b
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub

''' Batch form. One call instead of N - tidier, but measured to be the same
''' speed, so choose it for clarity rather than performance.
Public Sub BatchDemo()
    BlipBridge.Initialize

    Dim bytes() As Byte
    bytes = LoadFileBytes(ActivePresentation.path & "\texture.png")
    Dim tex As LongLong
    tex = BlipBridge.LoadTexture(bytes)

    On Error GoTo Cleanup
    Dim source As Object
    Set source = ActivePresentation.Slides(1).Shapes

    Dim targets() As Object
    ReDim targets(0 To source.Count - 1)
    Dim count As Long, index As Long
    For index = 1 To source.Count
        If source(index).Type = msoAutoShape Then
            Set targets(count) = source(index)
            count = count + 1
        End If
    Next index
    If count = 0 Then GoTo Cleanup
    ReDim Preserve targets(0 To count - 1)

    Debug.Print "filled"; BlipBridge.ApplyTextureToAll(targets, tex)

Cleanup:
    BlipBridge.ReleaseTexture tex
    If Err.Number <> 0 Then MsgBox Err.Description
End Sub
