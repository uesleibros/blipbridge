Attribute VB_Name = "BlipBridge"
'/**
' * BlipBridge - High-Performance Native Picture Fills for PowerPoint
' * @description Registration-free VBA wrapper for the BlipBridge C ABI. The module
' * automatically selects BlipBridge-x64.dll on 64-bit PowerPoint and
' * BlipBridge-x86.dll on 32-bit PowerPoint, while exposing one architecture-neutral
' * VBA API for cached textures, raw BGRA images, semantic Shape-safe picture fills,
' * batching, cache invalidation, capability discovery, and lifecycle management.
' * @author UesleiDev
' * @version ABI 2
' * @remarks The architecture is selected from the PowerPoint process, not Windows.
' * Both native DLLs may safely live beside the same presentation.
' */

Option Explicit
Option Private Module

#If VBA7 Then
#Else
    #Const BB_UNSUPPORTED_VBA = True
#End If

'/**
' * @enum BlipBridgeResult
' * @brief Native result codes returned by the BlipBridge C ABI.
' * @value BB_OK Operation completed successfully.
' * @value BB_E_INVALID_ARG One or more arguments are invalid.
' * @value BB_E_NOT_INITIALIZED BlipBridge has not been initialized.
' * @value BB_E_WRONG_THREAD The call was made from an unsupported thread.
' * @value BB_E_UNSUPPORTED_HOST The current process is not a supported PowerPoint host.
' * @value BB_E_UNSUPPORTED_BUILD The current Office build has not been validated.
' * @value BB_E_INVALID_HANDLE The texture handle is stale, zero, or otherwise invalid.
' * @value BB_E_INVALID_SHAPE The supplied Shape object is invalid or no longer usable.
' * @value BB_E_DECODE_FAILED The encoded image could not be decoded.
' * @value BB_E_APPLY_FAILED Office rejected or failed the native picture-fill transaction.
' * @value BB_E_OUT_OF_MEMORY A required native allocation failed.
' * @value BB_E_UNSUPPORTED_SHAPE The Shape is valid but not semantically eligible.
' * @value BB_E_FILE_NOT_FOUND The requested picture file does not exist.
' * @value BB_E_FALLBACK_FAILED The Office public fallback path failed.
' * @value BB_E_INTERNAL An unexpected internal failure occurred.
' */
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

'/** @description Capability flag indicating an active validated native Office backend. */
Public Const BB_CAP_NATIVE_BACKEND As Long = &H1

'/** @description Capability flag indicating memory-backed encoded-image loading. */
Public Const BB_CAP_MEMORY_IMAGE As Long = &H2

'/** @description Capability flag indicating reusable cached texture handles. */
Public Const BB_CAP_CACHED_TEXTURE As Long = &H4

'/** @description Capability flag indicating batch apply support. */
Public Const BB_CAP_BATCH_APPLY As Long = &H8

'/** @description Capability flag indicating availability of the public Office fallback. */
Public Const BB_CAP_PICKUP_FALLBACK As Long = &H10

'/** @description Capability flag indicating native BGRA32 raw-pixel loading. */
Public Const BB_CAP_RAW_PIXELS As Long = &H20

'/** @description Capability flag indicating the high-level UserPicture2 path. */
Public Const BB_CAP_APPLY_PICTURE As Long = &H40

'/** @description Public ABI version required by this VBA wrapper. */
Private Const BB_EXPECTED_ABI As Long = 2

'/** @description Base VBA error number used when translating native failures. */
Private Const BB_ERROR_BASE As Long = vbObjectError + 500

'/** @description UTF-8 code page used by native diagnostic strings. */
Private Const BB_CP_UTF8 As Long = 65001

'/** @description Native string selector for the current thread's last BlipBridge error. */
Private Const BB_NATIVE_STRING_LAST_ERROR As Long = 1

'/** @description Native string selector for the BlipBridge version string. */
Private Const BB_NATIVE_STRING_VERSION As Long = 2

#If Win64 Then
    '/** @description Native library selected for 64-bit PowerPoint. */
    Private Const BB_DLL_NAME As String = "BlipBridge-x64.dll"

    '/** @description Opposite-architecture library name used only for diagnostics. */
    Private Const BB_OTHER_DLL_NAME As String = "BlipBridge-x86.dll"
#Else
    '/** @description Native library selected for 32-bit PowerPoint. */
    Private Const BB_DLL_NAME As String = "BlipBridge-x86.dll"

    '/** @description Opposite-architecture library name used only for diagnostics. */
    Private Const BB_OTHER_DLL_NAME As String = "BlipBridge-x64.dll"

    '/**
    ' * @type BBHandle
    ' * @brief Binary representation of one opaque 64-bit texture handle in 32-bit VBA.
    ' * @remarks Low and High map directly to the little-endian uint64_t ABI layout.
    ' */
    Private Type BBHandle
        Low As Long
        High As Long
    End Type
#End If

'/** @description Native loader reference held by this wrapper. */
Private mModule As LongPtr

'/** @description True after BB_Init succeeds for the currently loaded native library. */
Private mReady As Boolean

#If VBA7 Then

'/** @description Loads a native module from an explicit Unicode file-system path. */
Private Declare PtrSafe Function LoadLibraryW Lib "kernel32" ( _
    ByVal lpLibFileName As LongPtr _
) As LongPtr

'/** @description Releases the loader reference acquired by LoadLibraryW. */
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" ( _
    ByVal hLibModule As LongPtr _
) As Long

'/** @description Converts UTF-8 diagnostic bytes returned by the native ABI to UTF-16 VBA text. */
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
) As Long

#If Win64 Then

Private Declare PtrSafe Function BB_Init Lib "BlipBridge-x64.dll" () As Long
Private Declare PtrSafe Function BB_Shutdown Lib "BlipBridge-x64.dll" () As Long
Private Declare PtrSafe Function BB_GetAbiVersion Lib "BlipBridge-x64.dll" () As Long

Private Declare PtrSafe Function BB_LoadTexture Lib "BlipBridge-x64.dll" ( _
    ByVal bytesPtr As LongPtr, _
    ByVal length As Long, _
    ByRef outHandle As LongLong _
) As Long

Private Declare PtrSafe Function BB_LoadTexturePixels Lib "BlipBridge-x64.dll" ( _
    ByVal pixelsPtr As LongPtr, _
    ByVal width As Long, _
    ByVal height As Long, _
    ByVal stride As Long, _
    ByRef outHandle As LongLong _
) As Long

Private Declare PtrSafe Function BB_ApplyTexture Lib "BlipBridge-x64.dll" ( _
    ByVal shapePtr As LongPtr, _
    ByVal texture As LongLong _
) As Long

Private Declare PtrSafe Function BB_ApplyTextureBatch Lib "BlipBridge-x64.dll" ( _
    ByVal shapesPtr As LongPtr, _
    ByVal texturesPtr As LongPtr, _
    ByVal count As Long, _
    ByRef applied As Long _
) As Long

Private Declare PtrSafe Function BB_ReleaseTexture Lib "BlipBridge-x64.dll" ( _
    ByVal texture As LongLong _
) As Long

Private Declare PtrSafe Function BB_ClearTextures Lib "BlipBridge-x64.dll" () As Long

Private Declare PtrSafe Function BB_ApplyPicture Lib "BlipBridge-x64.dll" ( _
    ByVal shapePtr As LongPtr, _
    ByVal pathPtr As LongPtr _
) As Long

Private Declare PtrSafe Function BB_InvalidateShape Lib "BlipBridge-x64.dll" ( _
    ByVal shapePtr As LongPtr _
) As Long

Private Declare PtrSafe Function BB_ClearPictureCache Lib "BlipBridge-x64.dll" () As Long

Private Declare PtrSafe Function BB_GetPictureCacheStats Lib "BlipBridge-x64.dll" ( _
    ByRef textures As Long, _
    ByRef shapes As Long, _
    ByRef skipped As LongLong _
) As Long

Private Declare PtrSafe Function BB_GetTextureCount Lib "BlipBridge-x64.dll" () As Long
Private Declare PtrSafe Function BB_GetCapabilities Lib "BlipBridge-x64.dll" () As Long

Private Declare PtrSafe Function BB_GetLastError Lib "BlipBridge-x64.dll" ( _
    ByVal buffer As LongPtr, _
    ByVal capacity As Long _
) As Long

Private Declare PtrSafe Function BB_GetVersionString Lib "BlipBridge-x64.dll" ( _
    ByVal buffer As LongPtr, _
    ByVal capacity As Long _
) As Long

#Else

Private Declare PtrSafe Function BB_Init Lib "BlipBridge-x86.dll" () As Long
Private Declare PtrSafe Function BB_Shutdown Lib "BlipBridge-x86.dll" () As Long
Private Declare PtrSafe Function BB_GetAbiVersion Lib "BlipBridge-x86.dll" () As Long

Private Declare PtrSafe Function BB_LoadTexture Lib "BlipBridge-x86.dll" ( _
    ByVal bytesPtr As LongPtr, _
    ByVal length As Long, _
    ByRef outHandle As BBHandle _
) As Long

Private Declare PtrSafe Function BB_LoadTexturePixels Lib "BlipBridge-x86.dll" ( _
    ByVal pixelsPtr As LongPtr, _
    ByVal width As Long, _
    ByVal height As Long, _
    ByVal stride As Long, _
    ByRef outHandle As BBHandle _
) As Long

Private Declare PtrSafe Function BB_ApplyTexture Lib "BlipBridge-x86.dll" ( _
    ByVal shapePtr As LongPtr, _
    ByVal textureLow As Long, _
    ByVal textureHigh As Long _
) As Long

Private Declare PtrSafe Function BB_ApplyTextureBatch Lib "BlipBridge-x86.dll" ( _
    ByVal shapesPtr As LongPtr, _
    ByVal texturesPtr As LongPtr, _
    ByVal count As Long, _
    ByRef applied As Long _
) As Long

Private Declare PtrSafe Function BB_ReleaseTexture Lib "BlipBridge-x86.dll" ( _
    ByVal textureLow As Long, _
    ByVal textureHigh As Long _
) As Long

Private Declare PtrSafe Function BB_ClearTextures Lib "BlipBridge-x86.dll" () As Long

Private Declare PtrSafe Function BB_ApplyPicture Lib "BlipBridge-x86.dll" ( _
    ByVal shapePtr As LongPtr, _
    ByVal pathPtr As LongPtr _
) As Long

Private Declare PtrSafe Function BB_InvalidateShape Lib "BlipBridge-x86.dll" ( _
    ByVal shapePtr As LongPtr _
) As Long

Private Declare PtrSafe Function BB_ClearPictureCache Lib "BlipBridge-x86.dll" () As Long

Private Declare PtrSafe Function BB_GetPictureCacheStats Lib "BlipBridge-x86.dll" ( _
    ByRef textures As Long, _
    ByRef shapes As Long, _
    ByRef skipped As BBHandle _
) As Long

Private Declare PtrSafe Function BB_GetTextureCount Lib "BlipBridge-x86.dll" () As Long
Private Declare PtrSafe Function BB_GetCapabilities Lib "BlipBridge-x86.dll" () As Long

Private Declare PtrSafe Function BB_GetLastError Lib "BlipBridge-x86.dll" ( _
    ByVal buffer As LongPtr, _
    ByVal capacity As Long _
) As Long

Private Declare PtrSafe Function BB_GetVersionString Lib "BlipBridge-x86.dll" ( _
    ByVal buffer As LongPtr, _
    ByVal capacity As Long _
) As Long

#End If
#End If

'/**
' * @function Initialize
' * @brief Loads the architecture-correct BlipBridge DLL and initializes its native backend.
' * @description Uses the PowerPoint process architecture to select BlipBridge-x64.dll or
' * BlipBridge-x86.dll. The DLL is loaded by full path first so registration-free
' * deployment works when the native files are stored beside the presentation.
' * @remarks Safe to call more than once. ABI mismatches fail before BB_Init is invoked.
' */
Public Sub Initialize()
#If VBA7 Then
    If mReady Then Exit Sub

    EnsureLoaded

    Dim abi As Long
    abi = BB_GetAbiVersion()

    If abi <> BB_EXPECTED_ABI Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.Initialize", _
                  "ABI mismatch: " & BB_DLL_NAME & " reports ABI " & CStr(abi) & _
                  ", but this BlipBridge.bas requires ABI " & CStr(BB_EXPECTED_ABI) & "."
    End If

    CheckResult BB_Init(), "Initialize"
    mReady = True
#Else
    Err.Raise BB_ERROR_BASE, "BlipBridge.Initialize", _
              "BlipBridge requires VBA7 or newer."
#End If
End Sub

'/**
' * @function IsAvailable
' * @brief Reports whether the accelerated native backend can initialize in this PowerPoint process.
' * @return True when initialization succeeds and the native-backend capability is active.
' * @remarks This function intentionally converts initialization failure to False.
' */
Public Function IsAvailable() As Boolean
    On Error GoTo Unavailable

    Initialize
    IsAvailable = (Capabilities And BB_CAP_NATIVE_BACKEND) <> 0
    Exit Function

Unavailable:
    IsAvailable = False
End Function

'/**
' * @function Shutdown
' * @brief Shuts down BlipBridge and releases this wrapper's native loader reference.
' * @description Releases native texture state through BB_Shutdown before dropping the
' * LoadLibraryW reference. A shutdown failure is reported after local wrapper state is reset.
' * @remarks Safe to call without a matching Initialize and safe to call repeatedly.
' */
Public Sub Shutdown()
#If VBA7 Then
    Dim status As Long
    Dim detail As String

    If mModule = 0 Then
        mReady = False
        Exit Sub
    End If

    If mReady Then
        status = BB_Shutdown()

        If status <> BB_OK Then
            detail = ReadNativeString(BB_NATIVE_STRING_LAST_ERROR)
        End If
    End If

    mReady = False

    Dim moduleToRelease As LongPtr
    moduleToRelease = mModule
    mModule = 0

    If moduleToRelease <> 0 Then
        FreeLibrary moduleToRelease
    End If

    If status <> BB_OK Then
        RaiseCapturedResult status, "Shutdown", detail
    End If
#Else
    mReady = False
#End If
End Sub

'/**
' * @function LoadTexture
' * @brief Decodes an encoded image buffer once and returns a reusable opaque texture handle.
' * @param bytes Encoded PNG, JPEG, BMP, or another format supported by the validated native decoder.
' * @return Architecture-neutral Variant carrying the opaque 64-bit BlipBridge texture token.
' * @remarks The native side copies or retains everything it needs before returning; the VBA array
' * may be reused immediately. A zero-length or unallocated array is rejected before the ABI call.
' */
Public Function LoadTexture(ByRef bytes() As Byte) As Variant
    Initialize

    Dim length As Long
    length = ByteArrayLength(bytes)

    If length <= 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.LoadTexture", _
                  "LoadTexture requires a non-empty byte array."
    End If

#If Win64 Then
    Dim handle As LongLong
    CheckResult BB_LoadTexture(VarPtr(bytes(LBound(bytes))), length, handle), "LoadTexture"
    LoadTexture = handle
#Else
    Dim parts As BBHandle
    CheckResult BB_LoadTexture(VarPtr(bytes(LBound(bytes))), length, parts), "LoadTexture"
    LoadTexture = HandleToVariant(parts)
#End If
End Function

'/**
' * @function LoadTexturePixels
' * @brief Creates a texture directly from a 32-bit BGRA pixel buffer.
' * @param pixels Contiguous BGRA32 pixel bytes ordered blue, green, red, alpha.
' * @param width Pixel width of the source surface.
' * @param height Pixel height of the source surface.
' * @param stride Number of bytes between consecutive rows. Zero selects width * 4.
' * @return Architecture-neutral Variant carrying the opaque 64-bit texture token.
' * @remarks Raw pixels avoid requiring callers that already own BGRA data to encode an image first.
' */
Public Function LoadTexturePixels( _
    ByRef pixels() As Byte, _
    ByVal width As Long, _
    ByVal height As Long, _
    Optional ByVal stride As Long = 0 _
) As Variant
    Initialize

    If width <= 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.LoadTexturePixels", _
                  "width must be greater than zero."
    End If

    If height <= 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.LoadTexturePixels", _
                  "height must be greater than zero."
    End If

    If stride = 0 Then
        If CDbl(width) * 4# > 2147483647# Then
            Err.Raise BB_ERROR_BASE, "BlipBridge.LoadTexturePixels", _
                      "width is too large for a VBA Long stride."
        End If

        stride = width * 4
    End If

    If stride < width * 4 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.LoadTexturePixels", _
                  "stride must be at least width * 4 for BGRA32 pixels."
    End If

    Dim available As Long
    available = ByteArrayLength(pixels)

    If available <= 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.LoadTexturePixels", _
                  "LoadTexturePixels requires a non-empty pixel array."
    End If

    Dim required As Double
    required = CDbl(stride) * CDbl(height)

    If required > CDbl(available) Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.LoadTexturePixels", _
                  "The pixel array is smaller than stride * height."
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

'/**
' * @function ApplyTexture
' * @brief Applies a previously loaded texture to one semantically supported PowerPoint Shape.
' * @param shp Live PowerPoint Shape-compatible object whose pointer is borrowed for this call only.
' * @param texture Opaque handle returned by LoadTexture or LoadTexturePixels.
' * @remarks The Shape pointer is never retained by BlipBridge. Native structural and semantic
' * eligibility validation still runs before the private Office transaction is allowed.
' */
Public Sub ApplyTexture(ByVal shp As Object, ByVal texture As Variant)
    Initialize

    If shp Is Nothing Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.ApplyTexture", _
                  "ApplyTexture requires a live Shape."
    End If

#If Win64 Then
    CheckResult BB_ApplyTexture(ObjPtr(shp), CLngLng(texture)), "ApplyTexture"
#Else
    Dim parts As BBHandle
    parts = HandleParts(texture)
    CheckResult BB_ApplyTexture(ObjPtr(shp), parts.Low, parts.High), "ApplyTexture"
#End If
End Sub

'/**
' * @function UserPicture2
' * @brief High-level cached replacement for PowerPoint Fill.UserPicture.
' * @param shp Destination PowerPoint Shape.
' * @param imagePath Unicode path to the image file.
' * @description Native-supported Shape classes use the accelerated cached path, known fallback
' * classes use Office's public Fill.UserPicture path, and semantically unsupported classes fail
' * explicitly. Unexpected native failures are never converted into silent fallback.
' * @remarks The native picture cache also skips a redundant apply when it can prove that the same
' * image is already the last BlipBridge-managed picture on the same cache-safe Shape.
' */
Public Sub UserPicture2(ByVal shp As Object, ByVal imagePath As String)
    Initialize

    If shp Is Nothing Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.UserPicture2", _
                  "UserPicture2 requires a live Shape."
    End If

    If LenB(imagePath) = 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.UserPicture2", _
                  "UserPicture2 requires an image path."
    End If

    CheckResult BB_ApplyPicture(ObjPtr(shp), StrPtr(imagePath)), "UserPicture2"
End Sub

'/**
' * @function InvalidateShape
' * @brief Invalidates UserPicture2's remembered last-picture state for one Shape.
' * @param shp Shape whose BlipBridge picture-cache state should be forgotten.
' * @description Call after changing the Shape's fill through another API so a later UserPicture2
' * invocation cannot incorrectly skip a required native transaction.
' */
Public Sub InvalidateShape(ByVal shp As Object)
    Initialize

    If shp Is Nothing Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.InvalidateShape", _
                  "InvalidateShape requires a live Shape."
    End If

    CheckResult BB_InvalidateShape(ObjPtr(shp)), "InvalidateShape"
End Sub

'/**
' * @function ClearPictureCache
' * @brief Clears all file-keyed UserPicture2 textures and remembered Shape picture states.
' * @remarks Useful after presentations are reloaded or source image assets are rewritten.
' */
Public Sub ClearPictureCache()
    Initialize
    CheckResult BB_ClearPictureCache(), "ClearPictureCache"
End Sub

'/**
' * @function PictureCacheStats
' * @brief Returns diagnostic counters for the UserPicture2 cache.
' * @return String formatted as "textures=N shapes=N skipped=N".
' */
Public Function PictureCacheStats() As String
    Initialize

    Dim textureCountValue As Long
    Dim shapeCountValue As Long

#If Win64 Then
    Dim skipped As LongLong
    CheckResult BB_GetPictureCacheStats(textureCountValue, shapeCountValue, skipped), _
                "PictureCacheStats"
#Else
    Dim skippedParts As BBHandle
    CheckResult BB_GetPictureCacheStats(textureCountValue, shapeCountValue, skippedParts), _
                "PictureCacheStats"

    Dim skipped As Variant
    skipped = HandleToVariant(skippedParts)
#End If

    PictureCacheStats = "textures=" & CStr(textureCountValue) & _
                        " shapes=" & CStr(shapeCountValue) & _
                        " skipped=" & CStr(skipped)
End Function

'/**
' * @function ApplyTextureBatch
' * @brief Applies one texture handle per Shape in a single VBA-to-native ABI transition.
' * @param shapes Array of live Shape objects.
' * @param textures Matching array of opaque texture handles.
' * @return Number of Shape entries successfully applied before the native call returned.
' * @remarks Batch is primarily a convenience API; current measurements do not show a material
' * speed advantage over individual applies because the Office transaction dominates each Shape.
' */
Public Function ApplyTextureBatch( _
    ByRef shapes() As Object, _
    ByRef textures() As Variant _
) As Long
    Initialize

    Dim shapeLower As Long
    Dim textureLower As Long
    Dim count As Long
    Dim textureCountValue As Long

    count = ObjectArrayCount(shapes, shapeLower)
    textureCountValue = VariantArrayCount(textures, textureLower)

    If count <= 0 Then Exit Function

    If textureCountValue <> count Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.ApplyTextureBatch", _
                  "ApplyTextureBatch requires one texture handle per Shape."
    End If

    Dim pointers() As LongPtr
    ReDim pointers(0 To count - 1)

    Dim index As Long

    For index = 0 To count - 1
        If shapes(shapeLower + index) Is Nothing Then
            Err.Raise BB_ERROR_BASE, "BlipBridge.ApplyTextureBatch", _
                      "Shape at batch index " & CStr(index) & " is Nothing."
        End If

        pointers(index) = ObjPtr(shapes(shapeLower + index))
    Next index

    Dim applied As Long
    Dim status As Long

#If Win64 Then
    Dim rawHandles() As LongLong
    ReDim rawHandles(0 To count - 1)

    For index = 0 To count - 1
        rawHandles(index) = CLngLng(textures(textureLower + index))
    Next index

    status = BB_ApplyTextureBatch(VarPtr(pointers(0)), VarPtr(rawHandles(0)), count, applied)
#Else
    Dim rawHandles() As Long
    ReDim rawHandles(0 To count * 2 - 1)

    Dim parts As BBHandle

    For index = 0 To count - 1
        parts = HandleParts(textures(textureLower + index))
        rawHandles(index * 2) = parts.Low
        rawHandles(index * 2 + 1) = parts.High
    Next index

    status = BB_ApplyTextureBatch(VarPtr(pointers(0)), VarPtr(rawHandles(0)), count, applied)
#End If

    ApplyTextureBatch = applied
    CheckResult status, "ApplyTextureBatch"
End Function

'/**
' * @function ApplyTextureToAll
' * @brief Applies one texture handle to every Shape in an array.
' * @param shapes Array of live Shape objects.
' * @param texture Opaque texture handle applied to every element.
' * @return Number of Shape entries successfully applied.
' */
Public Function ApplyTextureToAll( _
    ByRef shapes() As Object, _
    ByVal texture As Variant _
) As Long
    Dim lower As Long
    Dim count As Long

    count = ObjectArrayCount(shapes, lower)

    If count <= 0 Then Exit Function

    Dim textures() As Variant
    ReDim textures(0 To count - 1)

    Dim index As Long

    For index = 0 To count - 1
        textures(index) = texture
    Next index

    ApplyTextureToAll = ApplyTextureBatch(shapes, textures)
End Function

'/**
' * @function ReleaseTexture
' * @brief Releases the caller-owned reference represented by one texture handle.
' * @param texture Opaque texture handle previously returned by a load function.
' * @remarks The released handle must never be used again. Office may independently retain its own
' * references for Shape state, document ownership, Undo/Redo, or rendering.
' */
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

'/**
' * @function ClearTextures
' * @brief Releases every caller-owned reusable texture handle while keeping BlipBridge initialized.
' */
Public Sub ClearTextures()
    If Not mReady Then Exit Sub
    CheckResult BB_ClearTextures(), "ClearTextures"
End Sub

'/**
' * @function TextureCount
' * @brief Reports the number of reusable texture handles currently owned by BlipBridge.
' * @return Active texture handle count, or zero while the wrapper is not initialized.
' */
Public Function TextureCount() As Long
    If Not mReady Then Exit Function
    TextureCount = BB_GetTextureCount()
End Function

'/**
' * @function Capabilities
' * @brief Returns the native runtime capability bitmask for the current PowerPoint process.
' * @return Combination of BB_CAP_* flags.
' */
Public Function Capabilities() As Long
    Initialize
    Capabilities = BB_GetCapabilities()
End Function

'/**
' * @function Version
' * @brief Returns the native BlipBridge version/backend description without requiring BB_Init.
' * @return Human-readable native version string.
' */
Public Function Version() As String
#If VBA7 Then
    EnsureLoaded
    Version = ReadNativeString(BB_NATIVE_STRING_VERSION)
#Else
    Version = vbNullString
#End If
End Function

'/**
' * @function AbiVersion
' * @brief Returns the ABI version exported by the selected native library.
' * @return Native public C ABI version.
' */
Public Function AbiVersion() As Long
#If VBA7 Then
    EnsureLoaded
    AbiVersion = BB_GetAbiVersion()
#Else
    AbiVersion = 0
#End If
End Function

'/**
' * @function Architecture
' * @brief Reports the architecture of the current PowerPoint process.
' * @return "x64" for 64-bit PowerPoint or "x86" for 32-bit PowerPoint.
' */
Public Function Architecture() As String
#If Win64 Then
    Architecture = "x64"
#Else
    Architecture = "x86"
#End If
End Function

'/**
' * @function NativeLibraryName
' * @brief Reports the native DLL filename selected for this PowerPoint architecture.
' * @return BlipBridge-x64.dll or BlipBridge-x86.dll.
' */
Public Function NativeLibraryName() As String
    NativeLibraryName = BB_DLL_NAME
End Function

'/**
' * @function EnsureLoaded
' * @brief Locates and loads the architecture-correct native library by full Unicode path.
' * @description Searches beside the active presentation, the current working directory, then the
' * BLIPBRIDGE_DLL environment override. Loading by full path before using static VBA Declares lets
' * the OS resolve those declarations to the already-loaded architecture-specific module.
' */
Private Sub EnsureLoaded()
#If VBA7 Then
    If mModule <> 0 Then Exit Sub

    Dim candidates As Variant
    candidates = SearchPaths()

    Dim candidate As Variant
    Dim candidatePath As String
    Dim foundButFailed As String
    Dim lastError As Long

    For Each candidate In candidates
        candidatePath = CStr(candidate)

        If LenB(candidatePath) <> 0 Then
            If FileExists(candidatePath) Then
                foundButFailed = candidatePath
                mModule = LoadLibraryW(StrPtr(candidatePath))

                If mModule <> 0 Then Exit Sub

                lastError = Err.LastDllError
            End If
        End If
    Next candidate

    If mModule <> 0 Then Exit Sub

    If LenB(foundButFailed) <> 0 Then
        If lastError = 193 Then
            Err.Raise BB_ERROR_BASE, "BlipBridge.EnsureLoaded", _
                      "Architecture mismatch while loading " & foundButFailed & "." & vbCrLf & _
                      "This PowerPoint process requires " & BB_DLL_NAME & "."
        End If

        Err.Raise BB_ERROR_BASE, "BlipBridge.EnsureLoaded", _
                  "BlipBridge found " & foundButFailed & " but Windows could not load it " & _
                  "(error " & CStr(lastError) & ")."
    End If

    Dim wrongArchitecturePath As String
    wrongArchitecturePath = FindOtherArchitectureDll()

    If LenB(wrongArchitecturePath) <> 0 Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.EnsureLoaded", _
                  "Found " & wrongArchitecturePath & ", but this PowerPoint process is " & _
                  Architecture() & " and requires " & BB_DLL_NAME & "." & vbCrLf & _
                  "Both BlipBridge-x64.dll and BlipBridge-x86.dll may be placed beside the presentation."
    End If

    Err.Raise BB_ERROR_BASE, "BlipBridge.EnsureLoaded", _
              BB_DLL_NAME & " was not found." & vbCrLf & _
              "Place both architecture DLLs beside the presentation, or set BLIPBRIDGE_DLL " & _
              "to the full path of the correct native library." & vbCrLf & _
              "Searched:" & vbCrLf & Join(candidates, vbCrLf)
#Else
    Err.Raise BB_ERROR_BASE, "BlipBridge.EnsureLoaded", _
              "BlipBridge requires VBA7 or newer."
#End If
End Sub

'/**
' * @function SearchPaths
' * @brief Builds the ordered list of candidate native library paths.
' * @return Variant containing a compact String array of candidate full paths.
' */
Private Function SearchPaths() As Variant
    Dim paths(0 To 2) As String

    Dim presentationFolder As String
    presentationFolder = ActivePresentationFolder()

    If LenB(presentationFolder) <> 0 Then
        paths(0) = CombinePath(presentationFolder, BB_DLL_NAME)
    End If

    paths(1) = CombinePath(CurDir$(), BB_DLL_NAME)
    paths(2) = Environ$("BLIPBRIDGE_DLL")

    SearchPaths = CompactPaths(paths)
End Function

'/**
' * @function CompactPaths
' * @brief Removes empty entries from a fixed candidate-path array.
' * @param source Three-element String array produced by SearchPaths.
' * @return Variant containing a dense String array suitable for For Each and Join.
' */
Private Function CompactPaths(ByRef source() As String) As Variant
    Dim result() As String
    ReDim result(0 To UBound(source) - LBound(source))

    Dim count As Long
    Dim index As Long

    For index = LBound(source) To UBound(source)
        If LenB(source(index)) <> 0 Then
            result(count) = source(index)
            count = count + 1
        End If
    Next index

    If count = 0 Then
        ReDim result(0 To 0)
        result(0) = vbNullString
    Else
        ReDim Preserve result(0 To count - 1)
    End If

    CompactPaths = result
End Function

'/**
' * @function FindOtherArchitectureDll
' * @brief Finds the opposite-architecture BlipBridge DLL for a clearer missing-binary diagnostic.
' * @return Full path of a sibling opposite-architecture DLL, or an empty string when absent.
' */
Private Function FindOtherArchitectureDll() As String
    Dim presentationFolder As String
    presentationFolder = ActivePresentationFolder()

    If LenB(presentationFolder) <> 0 Then
        Dim candidate As String
        candidate = CombinePath(presentationFolder, BB_OTHER_DLL_NAME)

        If FileExists(candidate) Then
            FindOtherArchitectureDll = candidate
            Exit Function
        End If
    End If

    candidate = CombinePath(CurDir$(), BB_OTHER_DLL_NAME)

    If FileExists(candidate) Then
        FindOtherArchitectureDll = candidate
    End If
End Function

'/**
' * @function ActivePresentationFolder
' * @brief Returns the active presentation directory without failing when no saved path exists.
' * @return Presentation folder or an empty string for unsaved/unavailable presentations.
' */
Private Function ActivePresentationFolder() As String
    On Error GoTo Unavailable
    ActivePresentationFolder = ActivePresentation.Path
    Exit Function

Unavailable:
    ActivePresentationFolder = vbNullString
End Function

'/**
' * @function CombinePath
' * @brief Joins one folder and file name without creating a duplicate path separator.
' * @param folder Parent directory.
' * @param fileName Child file name.
' * @return Combined Windows path.
' */
Private Function CombinePath(ByVal folder As String, ByVal fileName As String) As String
    If LenB(folder) = 0 Then
        CombinePath = fileName
        Exit Function
    End If

    If Right$(folder, 1) = "\" Or Right$(folder, 1) = "/" Then
        CombinePath = folder & fileName
    Else
        CombinePath = folder & "\" & fileName
    End If
End Function

'/**
' * @function FileExists
' * @brief Tests whether one candidate native/image path resolves to a file.
' * @param filePath File-system path to test.
' * @return True when Dir$ can resolve a non-directory file entry.
' */
Private Function FileExists(ByVal filePath As String) As Boolean
    On Error GoTo Missing
    FileExists = LenB(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) <> 0
    Exit Function

Missing:
    FileExists = False
End Function

'/**
' * @function ByteArrayLength
' * @brief Returns the logical number of bytes in a dynamic Byte array.
' * @param bytes Byte array whose allocation state and bounds should be measured.
' * @return Element count, or zero when the dynamic array is unallocated.
' */
Private Function ByteArrayLength(ByRef bytes() As Byte) As Long
    On Error GoTo EmptyArray
    ByteArrayLength = UBound(bytes) - LBound(bytes) + 1
    Exit Function

EmptyArray:
    ByteArrayLength = 0
End Function

'/**
' * @function ObjectArrayCount
' * @brief Safely measures a dynamic Object array and returns its lower bound.
' * @param values Object array to inspect.
' * @param lower Receives the array's lower bound when allocated.
' * @return Element count, or zero when the array is unallocated.
' */
Private Function ObjectArrayCount(ByRef values() As Object, ByRef lower As Long) As Long
    On Error GoTo EmptyArray
    lower = LBound(values)
    ObjectArrayCount = UBound(values) - lower + 1
    Exit Function

EmptyArray:
    lower = 0
    ObjectArrayCount = 0
End Function

'/**
' * @function VariantArrayCount
' * @brief Safely measures a dynamic Variant array and returns its lower bound.
' * @param values Variant array to inspect.
' * @param lower Receives the array's lower bound when allocated.
' * @return Element count, or zero when the array is unallocated.
' */
Private Function VariantArrayCount(ByRef values() As Variant, ByRef lower As Long) As Long
    On Error GoTo EmptyArray
    lower = LBound(values)
    VariantArrayCount = UBound(values) - lower + 1
    Exit Function

EmptyArray:
    lower = 0
    VariantArrayCount = 0
End Function

'/**
' * @function CheckResult
' * @brief Converts one native BlipBridge result code into a descriptive VBA runtime error.
' * @param status Native BB_* result value.
' * @param operation Human-readable operation name included in the raised error.
' */
Private Sub CheckResult(ByVal status As Long, ByVal operation As String)
    If status = BB_OK Then Exit Sub

    Dim detail As String
    detail = ReadNativeString(BB_NATIVE_STRING_LAST_ERROR)

    RaiseCapturedResult status, operation, detail
End Sub

'/**
' * @function RaiseCapturedResult
' * @brief Raises a VBA error from a native status and an already-captured native message.
' * @param status Native BB_* result value.
' * @param operation Human-readable operation name.
' * @param detail Native diagnostic message captured while the DLL was callable.
' */
Private Sub RaiseCapturedResult( _
    ByVal status As Long, _
    ByVal operation As String, _
    ByVal detail As String _
)
    If LenB(detail) = 0 Then
        detail = "native error " & CStr(status)
    End If

    Err.Raise BB_ERROR_BASE + Abs(status), "BlipBridge." & operation, _
              operation & " failed: " & detail & " (code " & CStr(status) & ")"
End Sub

'/**
' * @function ReadNativeString
' * @brief Reads one length-prefixed-by-query UTF-8 string from the BlipBridge C ABI.
' * @param source BB_NATIVE_STRING_* selector.
' * @return Correctly decoded VBA Unicode string.
' */
Private Function ReadNativeString(ByVal source As Long) As String
    Dim needed As Long

    Select Case source
        Case BB_NATIVE_STRING_LAST_ERROR
            needed = BB_GetLastError(0, 0)

        Case BB_NATIVE_STRING_VERSION
            needed = BB_GetVersionString(0, 0)

        Case Else
            Err.Raise BB_ERROR_BASE, "BlipBridge.ReadNativeString", _
                      "Unknown native string source."
    End Select

    If needed <= 1 Then Exit Function

    Dim buffer() As Byte
    ReDim buffer(0 To needed - 1)

    Select Case source
        Case BB_NATIVE_STRING_LAST_ERROR
            BB_GetLastError VarPtr(buffer(0)), needed

        Case BB_NATIVE_STRING_VERSION
            BB_GetVersionString VarPtr(buffer(0)), needed
    End Select

    Dim payloadLength As Long
    payloadLength = NullTerminatedByteLength(buffer)

    If payloadLength <= 0 Then Exit Function

    ReadNativeString = Utf8ToString(buffer, payloadLength)
End Function

'/**
' * @function NullTerminatedByteLength
' * @brief Measures UTF-8 payload bytes before the first NUL terminator.
' * @param bytes Native byte buffer.
' * @return Number of non-NUL payload bytes.
' */
Private Function NullTerminatedByteLength(ByRef bytes() As Byte) As Long
    Dim index As Long

    For index = LBound(bytes) To UBound(bytes)
        If bytes(index) = 0 Then
            NullTerminatedByteLength = index - LBound(bytes)
            Exit Function
        End If
    Next index

    NullTerminatedByteLength = UBound(bytes) - LBound(bytes) + 1
End Function

'/**
' * @function Utf8ToString
' * @brief Converts a contiguous UTF-8 byte prefix into a native VBA Unicode String.
' * @param bytes Source byte array.
' * @param byteCount Number of UTF-8 bytes to decode.
' * @return Decoded UTF-16 VBA string.
' */
Private Function Utf8ToString(ByRef bytes() As Byte, ByVal byteCount As Long) As String
#If VBA7 Then
    If byteCount <= 0 Then Exit Function

    Dim charCount As Long
    charCount = MultiByteToWideChar(BB_CP_UTF8, 0, VarPtr(bytes(LBound(bytes))), _
                                    byteCount, 0, 0)

    If charCount <= 0 Then Exit Function

    Dim result As String
    result = String$(charCount, vbNullChar)

    If MultiByteToWideChar(BB_CP_UTF8, 0, VarPtr(bytes(LBound(bytes))), _
                           byteCount, StrPtr(result), charCount) <= 0 Then
        Exit Function
    End If

    Utf8ToString = result
#End If
End Function

#If Win64 Then
#Else

'/**
' * @function HandleParts
' * @brief Converts the 32-bit VBA Variant representation into exact low/high uint64_t words.
' * @param handle Caller-visible texture handle represented as an exact integral Double.
' * @return BBHandle containing the native low and high 32-bit words.
' * @remarks Values above 2^53 are rejected because 32-bit VBA Double can no longer represent every
' * consecutive integer exactly beyond that point.
' */
Private Function HandleParts(ByVal handle As Variant) As BBHandle
    Dim value As Double
    value = CDbl(handle)

    If value < 0# Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.HandleParts", _
                  "Texture handle cannot be negative."
    End If

    If value <> Fix(value) Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.HandleParts", _
                  "Texture handle must be an exact integer returned by BlipBridge."
    End If

    If value > 9007199254740992# Then
        Err.Raise BB_ERROR_BASE, "BlipBridge.HandleParts", _
                  "Texture handle exceeds the exact integer range supported by 32-bit VBA."
    End If

    Dim highValue As Double
    Dim lowValue As Double

    highValue = Fix(value / 4294967296#)
    lowValue = value - highValue * 4294967296#

    If lowValue >= 2147483648# Then
        HandleParts.Low = CLng(lowValue - 4294967296#)
    Else
        HandleParts.Low = CLng(lowValue)
    End If

    HandleParts.High = CLng(highValue)
End Function

'/**
' * @function HandleToVariant
' * @brief Converts native low/high uint64_t words into the exact 32-bit VBA Variant representation.
' * @param parts Native 64-bit handle words.
' * @return Variant containing an integral Double suitable for the public architecture-neutral API.
' */
Private Function HandleToVariant(ByRef parts As BBHandle) As Variant
    Dim lowValue As Double
    Dim highValue As Double

    lowValue = parts.Low
    highValue = parts.High

    If lowValue < 0# Then
        lowValue = lowValue + 4294967296#
    End If

    If highValue < 0# Then
        highValue = highValue + 4294967296#
    End If

    HandleToVariant = highValue * 4294967296# + lowValue
End Function

#End If
