Attribute VB_Name = "BlipBridge"
'==============================================================================
' BlipBridge - fast reusable image textures for PowerPoint Shapes.
'
' Drop BlipBridge.dll next to your .pptm, import this module, and call:
'
'     BlipBridge.Initialize
'     tex = BlipBridge.LoadTexture(bytes)
'     BlipBridge.ApplyTexture shp, tex
'     BlipBridge.ReleaseTexture tex
'     BlipBridge.Shutdown
'
' No regsvr32, no ProgID, no CreateObject, no COM add-in. The DLL is loaded
' explicitly with LoadLibraryW so it can live beside the presentation.
'
' What this module hides from callers:
'   * finding and loading the DLL, and unloading it on Shutdown
'   * the architecture difference between 32-bit and 64-bit PowerPoint
'   * a clear message when the DLL beside you is the wrong architecture
'   * turning Byte() into a pointer and a length
'   * turning a Shape into the pointer the native side expects
'   * native error codes, raised as ordinary VBA errors with real messages
'
' ARCHITECTURE
'
' One module serves both 32-bit and 64-bit PowerPoint. The public API is the
' same on each; only the Declares below differ.
'
' Texture handles are opaque 64-bit tokens in the C ABI and stay 64-bit on both
' architectures - the ABI does not shrink to suit a 32-bit host. VBA makes that
' awkward, because LongLong exists only in 64-bit Office, so this module hands
' handles to callers as Variant:
'
'     Dim tex                       ' works on both
'     tex = BlipBridge.LoadTexture(bytes)
'     BlipBridge.ApplyTexture shp, tex
'
' On 64-bit the Variant carries a LongLong. On 32-bit it carries a Double, which
' represents every integer below 2^53 exactly - handles start near 16.7 million
' and increment by one, so the representation is exact for any handle this
' library will ever issue, and HandleParts verifies that rather than assuming it.
'
' Requirements: PowerPoint for Windows and VBA7. The accelerated backend is
' validated per Office build and per architecture and refuses anything else -
' Initialize will say so. On 32-bit there is no accelerated backend yet at all;
' see docs/windows_x86.md.
'==============================================================================
Option Explicit
Option Private Module

#If VBA7 Then
#Else
    ' VBA6 has no PtrSafe and no LongPtr, so it cannot declare this ABI at all.
    #Const BB_UNSUPPORTED_VBA = True
#End If

#If Win64 Then
#Else
''' A 64-bit texture handle as 32-bit VBA can express it: two Longs, low first,
''' which is the little-endian layout of the ABI's uint64_t. Passing this ByRef
''' is exactly passing a uint64_t*.
Private Type BBHandle
    Low As Long
    High As Long
End Type
#End If

'--- Win32 loader -------------------------------------------------------------
Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" _
    (ByVal lpLibFileName As LongPtr) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr

'--- BlipBridge C ABI ---------------------------------------------------------
' These look like ordinary static Declares, and the LoadLibraryW call above may
' look redundant. It is not, and the combination is the whole trick behind
' registration-free deployment:
'
'   * A bare Declare makes VBA ask the OS loader for "BlipBridge.dll", and the
'     OS search path does NOT include the folder holding the presentation, so
'     beside-the-pptm deployment would fail.
'   * EnsureLoaded calls LoadLibraryW with the full path first. Once a module is
'     loaded, the loader resolves the bare name to that already-loaded module
'     instead of searching disk again.
'
' So Initialize must run before any Declare below is used, which is why every
' public entry point calls it. Exported names are undecorated on both
' architectures - the 32-bit build strips the stdcall @N suffix - so the same
' names appear in both blocks and match the C header exactly.
'
' The blocks differ only in how a 64-bit handle is carried:
'
'   64-bit   LongLong, which is exactly the ABI's uint64_t.
'   32-bit   two ByVal Long for a by-value handle, and a two-Long UDT ByRef for
'            an out-parameter. That is not a trick: a stdcall frame passes a
'            64-bit value as two consecutive 32-bit slots, low word first, which
'            is precisely what two ByVal Longs produce. LongLong does not exist
'            in 32-bit VBA, so it cannot be spelled any other way.
#If Win64 Then
Private Declare PtrSafe Function BB_Init Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_Shutdown Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_LoadTexture Lib "BlipBridge.dll" _
    (ByVal bytesPtr As LongPtr, ByVal length As Long, ByRef outHandle As LongLong) As Long
Private Declare PtrSafe Function BB_LoadTexturePixels Lib "BlipBridge.dll" _
    (ByVal pixelsPtr As LongPtr, ByVal width As Long, ByVal height As Long, _
     ByVal stride As Long, ByRef outHandle As LongLong) As Long
Private Declare PtrSafe Function BB_GetAbiVersion Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_ApplyTexture Lib "BlipBridge.dll" _
    (ByVal shapePtr As LongPtr, ByVal texture As LongLong) As Long
Private Declare PtrSafe Function BB_ApplyTextureBatch Lib "BlipBridge.dll" _
    (ByVal shapesPtr As LongPtr, ByVal texturesPtr As LongPtr, _
     ByVal count As Long, ByRef applied As Long) As Long
Private Declare PtrSafe Function BB_ReleaseTexture Lib "BlipBridge.dll" _
    (ByVal texture As LongLong) As Long
Private Declare PtrSafe Function BB_ClearTextures Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_ApplyPicture Lib "BlipBridge.dll" _
    (ByVal shape As LongPtr, ByVal path As LongPtr) As Long
Private Declare PtrSafe Function BB_InvalidateShape Lib "BlipBridge.dll" _
    (ByVal shape As LongPtr) As Long
Private Declare PtrSafe Function BB_ClearPictureCache Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_GetPictureCacheStats Lib "BlipBridge.dll" _
    (ByRef textures As Long, ByRef shapes As Long, ByRef skipped As LongLong) As Long
Private Declare PtrSafe Function BB_GetTextureCount Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_GetCapabilities Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_GetLastError Lib "BlipBridge.dll" _
    (ByVal buffer As LongPtr, ByVal capacity As Long) As Long
Private Declare PtrSafe Function BB_GetVersionString Lib "BlipBridge.dll" _
    (ByVal buffer As LongPtr, ByVal capacity As Long) As Long
#Else
Private Declare PtrSafe Function BB_Init Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_Shutdown Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_LoadTexture Lib "BlipBridge.dll" _
    (ByVal bytesPtr As LongPtr, ByVal length As Long, ByRef outHandle As BBHandle) As Long
Private Declare PtrSafe Function BB_LoadTexturePixels Lib "BlipBridge.dll" _
    (ByVal pixelsPtr As LongPtr, ByVal width As Long, ByVal height As Long, _
     ByVal stride As Long, ByRef outHandle As BBHandle) As Long
Private Declare PtrSafe Function BB_GetAbiVersion Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_ApplyTexture Lib "BlipBridge.dll" _
    (ByVal shapePtr As LongPtr, ByVal textureLow As Long, ByVal textureHigh As Long) As Long
Private Declare PtrSafe Function BB_ApplyTextureBatch Lib "BlipBridge.dll" _
    (ByVal shapesPtr As LongPtr, ByVal texturesPtr As LongPtr, _
     ByVal count As Long, ByRef applied As Long) As Long
Private Declare PtrSafe Function BB_ReleaseTexture Lib "BlipBridge.dll" _
    (ByVal textureLow As Long, ByVal textureHigh As Long) As Long
Private Declare PtrSafe Function BB_ClearTextures Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_ApplyPicture Lib "BlipBridge.dll" _
    (ByVal shape As LongPtr, ByVal path As LongPtr) As Long
Private Declare PtrSafe Function BB_InvalidateShape Lib "BlipBridge.dll" _
    (ByVal shape As LongPtr) As Long
Private Declare PtrSafe Function BB_ClearPictureCache Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_GetPictureCacheStats Lib "BlipBridge.dll" _
    (ByRef textures As Long, ByRef shapes As Long, ByRef skipped As BBHandle) As Long
Private Declare PtrSafe Function BB_GetTextureCount Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_GetCapabilities Lib "BlipBridge.dll" () As Long
Private Declare PtrSafe Function BB_GetLastError Lib "BlipBridge.dll" _
    (ByVal buffer As LongPtr, ByVal capacity As Long) As Long
Private Declare PtrSafe Function BB_GetVersionString Lib "BlipBridge.dll" _
    (ByVal buffer As LongPtr, ByVal capacity As Long) As Long
#End If

'--- Error codes, mirroring blipbridge.h --------------------------------------
Public Enum BlipBridgeResult
    BB_OK = 0
    BB_E_INVALID_ARG = -1
    BB_E_NOT_INITIALIZED = -2
    BB_E_WRONG_THREAD = -3
    BB_E_UNSUPPORTED_HOST = -4
    BB_E_UNSUPPORTED_BUILD = -5
    BB_E_INVALID_HANDLE = -6
    BB_E_INVALID_SHAPE = -7
    BB_E_DECODE_FAILED = -8
    BB_E_APPLY_FAILED = -9
    BB_E_OUT_OF_MEMORY = -10
    BB_E_UNSUPPORTED_SHAPE = -11
    BB_E_FILE_NOT_FOUND = -12
    BB_E_FALLBACK_FAILED = -13
    BB_E_INTERNAL = -99
End Enum

'--- Capability bits ----------------------------------------------------------
Public Const BB_CAP_NATIVE_BACKEND As Long = &H1
Public Const BB_CAP_MEMORY_IMAGE As Long = &H2
Public Const BB_CAP_CACHED_TEXTURE As Long = &H4
Public Const BB_CAP_BATCH_APPLY As Long = &H8
Public Const BB_CAP_PICKUP_FALLBACK As Long = &H10
Public Const BB_CAP_RAW_PIXELS As Long = &H20
Public Const BB_CAP_APPLY_PICTURE As Long = &H40

''' ABI version this module was written against. The DLL reports its own with
''' BB_GetAbiVersion; a mismatch means an old .bas is paired with a newer DLL (or
''' the reverse), which must fail loudly rather than call something whose shape
''' this module has wrong.
Private Const BB_EXPECTED_ABI As Long = 2

Private Const BB_ERROR_BASE As Long = vbObjectError + 0.5E3
Private mModule As LongPtr
Private mReady As Boolean

'==============================================================================
' Public API
'==============================================================================

''' Loads the DLL and prepares the library. Safe to call more than once.
''' Raises a descriptive error when the host or Office build is unsupported.
Public Sub Initialize()
    If mReady Then Exit Sub
    EnsureLoaded
    Dim abi As Long
    abi = BB_GetAbiVersion()
    If abi <> BB_EXPECTED_ABI Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", _
                  "BlipBridge.dll reports ABI version " & abi & " but this " & _
                  "BlipBridge.bas expects " & BB_EXPECTED_ABI & _
                  ". Update whichever is older; they are not compatible."
    End If
    CheckResult BB_Init(), "Initialize"
    mReady = True
End Sub

''' True when Initialize would succeed, without raising. Use this to decide
''' whether to take the accelerated path at all.
Public Function IsAvailable() As Boolean
    On Error GoTo Unavailable
    Initialize
    IsAvailable = (Capabilities And BB_CAP_NATIVE_BACKEND) <> 0
    Exit Function
Unavailable:
    IsAvailable = False
End Function

''' Releases every texture and unloads the DLL. Safe without a matching
''' Initialize, and safe to call twice.
Public Sub Shutdown()
    If mModule <> 0 Then
        If mReady Then BB_Shutdown
        FreeLibrary mModule
        mModule = 0
    End If
    mReady = False
End Sub

''' Decodes image bytes once into a reusable texture and returns its handle.
''' The bytes are copied; the array is yours again as soon as this returns.
Public Function LoadTexture(ByRef bytes() As Byte) As Variant
    Initialize
    If LBound(bytes) > UBound(bytes) Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", "LoadTexture was given an empty array"
    End If
#If Win64 Then
    Dim handle As LongLong
    CheckResult BB_LoadTexture(VarPtr(bytes(LBound(bytes))), _
                               UBound(bytes) - LBound(bytes) + 1, handle), "LoadTexture"
    LoadTexture = handle
#Else
    Dim parts As BBHandle
    CheckResult BB_LoadTexture(VarPtr(bytes(LBound(bytes))), _
                               UBound(bytes) - LBound(bytes) + 1, parts), "LoadTexture"
    LoadTexture = HandleToVariant(parts)
#End If
End Function

''' Builds a texture from raw 32-bit BGRA pixels, skipping image decoding.
'''
''' Pixels are blue, green, red, alpha per pixel - the ordinary Windows in-memory
''' layout. Only that layout is accepted; convert others before calling.
''' `stride` is bytes per row and must be at least width*4.
'''
''' This is not faster than LoadTexture for an image you already have encoded -
''' measured, it is about 15-18% slower for the same dimensions. It exists so a
''' caller holding pixels does not have to encode a PNG first, which would cost
''' far more than either path.
Public Function LoadTexturePixels(ByRef pixels() As Byte, ByVal width As Long, _
                                  ByVal height As Long, _
                                  Optional ByVal stride As Long = 0) As Variant
    Initialize
    If stride <= 0 Then stride = width * 4
    If LBound(pixels) > UBound(pixels) Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", "LoadTexturePixels was given an empty array"
    End If
    Dim needed As Long
    needed = stride * height
    If (UBound(pixels) - LBound(pixels) + 1) < needed Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", _
                  "LoadTexturePixels needs at least stride*height (" & needed & ") bytes"
    End If
#If Win64 Then
    Dim handle As LongLong
    CheckResult BB_LoadTexturePixels(VarPtr(pixels(LBound(pixels))), width, height, _
                                     stride, handle), "LoadTexturePixels"
    LoadTexturePixels = handle
#Else
    Dim parts As BBHandle
    CheckResult BB_LoadTexturePixels(VarPtr(pixels(LBound(pixels))), width, height, _
                                     stride, parts), "LoadTexturePixels"
    LoadTexturePixels = HandleToVariant(parts)
#End If
End Function

''' Fills one Shape with a previously loaded texture.
'''
''' The Shape pointer is borrowed for the duration of the call only - the native
''' side never stores it - and `shp` being a live argument here is what keeps the
''' object alive for that call. Never cache the value of ObjPtr yourself.
Public Sub ApplyTexture(ByVal shp As Object, ByVal texture As Variant)
    Initialize
    If shp Is Nothing Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", "ApplyTexture was given Nothing"
    End If
#If Win64 Then
    CheckResult BB_ApplyTexture(ObjPtr(shp), CLngLng(texture)), "ApplyTexture"
#Else
    Dim parts As BBHandle
    parts = HandleParts(texture)
    CheckResult BB_ApplyTexture(ObjPtr(shp), parts.Low, parts.High), "ApplyTexture"
#End If
End Sub

''' Fills a Shape from an image file - the drop-in for Fill.UserPicture.
'''
''' This is the entry point most code should use. It needs to know nothing about
''' Shape classes, texture handles or lifetimes:
'''
'''     BlipBridge.UserPicture2 shp, "C:	exturesrick.png"
'''
''' Internally it takes the accelerated path for Shape classes that have a
''' validated one, and Office's own Fill.UserPicture for classes that do not.
''' The file is read and decoded once however many Shapes it is applied to, and
''' applying the same image to the same Shape twice in a row does no work at all.
'''
''' It does **not** hide failures. An unsupported Office build, a deleted Shape
''' or an unreadable file each raise with their own reason; only a Shape whose
''' *class* has no native path is quietly routed to the slower route.
'''
''' If you change a Shape's fill by other means, call InvalidateShape on it so
''' the next call here does real work rather than assuming its own last result
''' still holds.
Public Sub UserPicture2(ByVal shp As Object, ByVal path As String)
    Initialize
    If shp Is Nothing Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", "UserPicture2 was given Nothing"
    End If
    If Len(path) = 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", "UserPicture2 needs an image path"
    End If
    ' StrPtr hands over VBA's own UTF-16 buffer, which is exactly what the ABI
    ' takes - no conversion, so nothing can be lost for a non-ASCII path.
    CheckResult BB_ApplyPicture(ObjPtr(shp), StrPtr(path)), "UserPicture2"
End Sub

''' Forgets what UserPicture2 last applied to one Shape.
'''
''' Call it after changing that Shape's fill by any other means. Forgetting a
''' Shape that was never cached is not an error.
Public Sub InvalidateShape(ByVal shp As Object)
    Initialize
    If shp Is Nothing Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", "InvalidateShape was given Nothing"
    End If
    CheckResult BB_InvalidateShape(ObjPtr(shp)), "InvalidateShape"
End Sub

''' Empties the picture cache: every file-keyed texture and every remembered Shape.
'''
''' Worth calling after closing or reloading a presentation, or when a batch of
''' source images has been rewritten.
Public Sub ClearPictureCache()
    Initialize
    CheckResult BB_ClearPictureCache(), "ClearPictureCache"
End Sub

''' Cache statistics, as "textures=N shapes=N skipped=N".
'''
''' `skipped` is how many applies UserPicture2 avoided entirely because the Shape
''' already carried that image.
Public Function PictureCacheStats() As String
    Initialize
    Dim textures As Long, shapes As Long
#If Win64 Then
    Dim skipped As LongLong
    CheckResult BB_GetPictureCacheStats(textures, shapes, skipped), "PictureCacheStats"
#Else
    Dim skippedParts As BBHandle
    CheckResult BB_GetPictureCacheStats(textures, shapes, skippedParts), "PictureCacheStats"
    Dim skipped As Variant
    skipped = HandleToVariant(skippedParts)
#End If
    PictureCacheStats = "textures=" & textures & " shapes=" & shapes & _
                        " skipped=" & skipped
End Function

''' Fills many Shapes in one crossing into native code.
'''
''' `shapes` is a 0-based array of Shape objects and `textures` a matching array
''' of handles; repeat a handle to give several Shapes the same texture. Every
''' Shape stays referenced by the array for the whole call.
Public Function ApplyTextureBatch(ByRef shapes() As Object, _
                                  ByRef textures() As Variant) As Long
    Initialize
    Dim count As Long
    count = UBound(shapes) - LBound(shapes) + 1
    If count <= 0 Then Exit Function
    If (UBound(textures) - LBound(textures) + 1) <> count Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", _
                  "ApplyTextureBatch needs one texture handle per Shape"
    End If

    ' Build the pointer array the ABI expects. Holding `shapes` for the whole
    ' call is what keeps every one of these pointers valid.
    Dim pointers() As LongPtr
    ReDim pointers(0 To count - 1)
    Dim index As Long
    For index = 0 To count - 1
        If shapes(LBound(shapes) + index) Is Nothing Then
            Err.Raise BB_ERROR_BASE, "BlipBridge", _
                      "ApplyTextureBatch was given Nothing at index " & index
        End If
        pointers(index) = ObjPtr(shapes(LBound(shapes) + index))
    Next index

    ' The ABI wants a contiguous array of 64-bit handles. A Variant array is not
    ' that on either architecture, so one is built here in the exact layout the
    ' ABI expects: LongLong elements on 64-bit, low/high Long pairs on 32-bit.
    Dim applied As Long
    Dim status As Long
#If Win64 Then
    Dim raw() As LongLong
    ReDim raw(0 To count - 1)
    For index = 0 To count - 1
        raw(index) = CLngLng(textures(LBound(textures) + index))
    Next index
    status = BB_ApplyTextureBatch(VarPtr(pointers(0)), VarPtr(raw(0)), count, applied)
#Else
    Dim raw() As Long
    ReDim raw(0 To count * 2 - 1)
    Dim parts As BBHandle
    For index = 0 To count - 1
        parts = HandleParts(textures(LBound(textures) + index))
        raw(index * 2) = parts.Low
        raw(index * 2 + 1) = parts.High
    Next index
    status = BB_ApplyTextureBatch(VarPtr(pointers(0)), VarPtr(raw(0)), count, applied)
#End If
    ApplyTextureBatch = applied
    CheckResult status, "ApplyTextureBatch"
End Function

''' Convenience: one texture across many Shapes, the common renderer case.
Public Function ApplyTextureToAll(ByRef shapes() As Object, _
                                  ByVal texture As Variant) As Long
    Dim count As Long
    count = UBound(shapes) - LBound(shapes) + 1
    If count <= 0 Then Exit Function
    Dim textures() As Variant
    ReDim textures(0 To count - 1)
    Dim index As Long
    For index = 0 To count - 1
        textures(index) = texture
    Next index
    ApplyTextureToAll = ApplyTextureBatch(shapes, textures)
End Function

''' Releases one texture. Its handle is never valid again.
Public Sub ReleaseTexture(ByVal texture As Variant)
    If Not mReady Then Exit Sub
#If Win64 Then
    CheckResult BB_ReleaseTexture(CLngLng(texture)), "ReleaseTexture"
#Else
    Dim parts As BBHandle
    parts = HandleParts(texture)
    CheckResult BB_ReleaseTexture(parts.Low, parts.High), "ReleaseTexture"
#End If
End Sub

''' Releases every texture but keeps the library loaded.
Public Sub ClearTextures()
    If Not mReady Then Exit Sub
    CheckResult BB_ClearTextures(), "ClearTextures"
End Sub

''' How many textures are currently loaded.
Public Function TextureCount() As Long
    If Not mReady Then Exit Function
    TextureCount = BB_GetTextureCount()
End Function

''' Capability bits for this process; see the BB_CAP_* constants.
Public Function Capabilities() As Long
    Initialize
    Capabilities = BB_GetCapabilities()
End Function

''' Version and backend description, for example "0.2.0 (windows-x64, ...)".
Public Function Version() As String
    EnsureLoaded
    Version = ReadNativeString(AddressOf_GetVersionString)
End Function

'==============================================================================
' Internals
'==============================================================================

#If Win64 Then
#Else
''' Splits a caller's handle into the two Longs the 32-bit ABI takes.
'''
''' The Variant carries a Double on 32-bit, which is exact for every integer
''' below 2^53. Handles begin near 16.7 million and increment by one, so that
''' holds for any handle this library issues - but it is checked rather than
''' assumed, because a silently truncated handle would release or apply the
''' wrong texture, and that is not a failure anyone would trace back to here.
Private Function HandleParts(ByVal handle As Variant) As BBHandle
    Dim value As Double
    value = CDbl(handle)
    If value < 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", "Texture handle cannot be negative"
    End If
    If value <> Fix(value) Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", _
                  "Texture handle is not a whole number; it did not come from " & _
                  "LoadTexture, or it has been modified."
    End If
    ' 2^53 is where a Double stops representing consecutive integers exactly.
    If value > 9007199254740992# Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", _
                  "Texture handle exceeds what 32-bit VBA can carry exactly " & _
                  "(2^53). Report this: the handle space was expected to stay " & _
                  "far below it."
    End If

    Dim high As Double
    high = Fix(value / 4294967296#)
    Dim low As Double
    low = value - high * 4294967296#
    ' Longs are signed, so the top half of the 32-bit range wraps to negative.
    ' That is the correct bit pattern; the ABI reads it as unsigned.
    If low >= 2147483648# Then
        HandleParts.Low = CLng(low - 4294967296#)
    Else
        HandleParts.Low = CLng(low)
    End If
    HandleParts.High = CLng(high)
End Function

''' The inverse: two Longs from the ABI back into a caller-visible handle.
Private Function HandleToVariant(ByRef parts As BBHandle) As Variant
    Dim low As Double
    low = parts.Low
    If low < 0 Then low = low + 4294967296#      ' undo the signed wrap
    Dim high As Double
    high = parts.High
    If high < 0 Then high = high + 4294967296#
    HandleToVariant = high * 4294967296# + low
End Function
#End If

Private Sub EnsureLoaded()
    If mModule <> 0 Then Exit Sub

    Dim candidate As String
    Dim found As String
    Dim path As Variant
    Dim lastError As Long
    For Each path In SearchPaths()
        candidate = CStr(path)
        If Len(Dir$(candidate)) > 0 Then
            found = candidate
            mModule = LoadLibraryW(StrPtr(candidate))
            If mModule <> 0 Then Exit For
            lastError = Err.LastDllError
        End If
    Next path

    If mModule <> 0 Then Exit Sub

    If Len(found) = 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", _
                  "BlipBridge.dll was not found. Put it next to this presentation." & _
                  vbCrLf & "Looked in: " & Join(SearchPaths(), vbCrLf)
    End If

    ' The file is there but would not load. By far the most common reason is the
    ' wrong package for this Office, and ERROR_BAD_EXE_FORMAT says so exactly -
    ' so say it in words rather than leaving the user with error 193 or a "Bad
    ' DLL calling convention" further along.
    If lastError = 193 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge", _
                  "Architecture mismatch: " & found & " is the " & OtherArchitecture() & _
                  " build of BlipBridge, but this PowerPoint is " & HostArchitecture() & _
                  "." & vbCrLf & _
                  "Download the blipbridge-...-windows-" & HostArchitectureTag() & _
                  ".zip package instead." & vbCrLf & _
                  "To confirm your Office architecture: File > Account > About PowerPoint."
    End If

    Err.Raise BB_ERROR_BASE, "BlipBridge", _
              "BlipBridge.dll was found at " & found & " but could not be loaded " & _
              "(Windows error " & lastError & ")."
End Sub

''' "64-bit" or "32-bit" for this PowerPoint process.
Private Function HostArchitecture() As String
#If Win64 Then
    HostArchitecture = "64-bit"
#Else
    HostArchitecture = "32-bit"
#End If
End Function

''' The one this process is not, for naming the DLL that failed to load.
Private Function OtherArchitecture() As String
#If Win64 Then
    OtherArchitecture = "32-bit"
#Else
    OtherArchitecture = "64-bit"
#End If
End Function

''' The tag used in release archive names.
Private Function HostArchitectureTag() As String
#If Win64 Then
    HostArchitectureTag = "x64"
#Else
    HostArchitectureTag = "x86"
#End If
End Function

''' Where the DLL may live, most specific first. Beside the presentation is the
''' intended deployment; the others exist so a shared copy also works.
Private Function SearchPaths() As Variant
    Dim here As String
    On Error Resume Next
    here = ActivePresentation.path
    On Error GoTo 0

    Dim paths() As String
    ReDim paths(0 To 2)
    If Len(here) > 0 Then paths(0) = here & "\BlipBridge.dll"
    paths(1) = CurDir$() & "\BlipBridge.dll"
    paths(2) = Environ$("BLIPBRIDGE_DLL")

    Dim cleaned() As String
    Dim count As Long
    ReDim cleaned(0 To UBound(paths))
    Dim index As Long
    For index = 0 To UBound(paths)
        If Len(paths(index)) > 0 Then
            cleaned(count) = paths(index)
            count = count + 1
        End If
    Next index
    ReDim Preserve cleaned(0 To IIf(count = 0, 0, count - 1))
    SearchPaths = cleaned
End Function

''' Turns a native failure into a VBA error carrying the native message.
Private Sub CheckResult(ByVal status As Long, ByVal operation As String)
    If status = BB_OK Then Exit Sub
    Dim detail As String
    detail = ReadNativeString(AddressOf_GetLastError)
    If Len(detail) = 0 Then detail = "native error " & status
    Err.Raise BB_ERROR_BASE + Abs(status), "BlipBridge", _
              operation & " failed: " & detail & " (code " & status & ")"
End Sub

' The two native string getters share a shape: ask for the size, then read.
Private Enum NativeStringSource
    AddressOf_GetLastError = 1
    AddressOf_GetVersionString = 2
End Enum

Private Function ReadNativeString(ByVal source As NativeStringSource) As String
    Dim needed As Long
    If source = AddressOf_GetLastError Then
        needed = BB_GetLastError(0, 0)
    Else
        needed = BB_GetVersionString(0, 0)
    End If
    If needed <= 1 Then Exit Function

    Dim buffer() As Byte
    ReDim buffer(0 To needed - 1)
    If source = AddressOf_GetLastError Then
        BB_GetLastError VarPtr(buffer(0)), needed
    Else
        BB_GetVersionString VarPtr(buffer(0)), needed
    End If

    ' The ABI emits UTF-8; these messages are ASCII in practice, but decode
    ' properly so a non-ASCII path in a message survives.
    Dim result As String
    Dim index As Long
    For index = 0 To needed - 1
        If buffer(index) = 0 Then Exit For
        result = result & Chr$(buffer(index))
    Next index
    ReadNativeString = result
End Function
