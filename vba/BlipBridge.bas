Attribute VB_Name = "BlipBridge"
Option Explicit
Private Engine As Object

Public Sub BB_Init()
    Set Engine = CreateObject("BlipBridge.Engine")
End Sub

' Explicit fallback: donor must already be a normal Shape with a picture fill.
' Apply copies other formatting too, including line weight.
Public Function BB_RegisterTextureShape(ByVal donor As Object) As Long
    If Engine Is Nothing Then BB_Init
    BB_RegisterTextureShape = Engine.RegisterTextureShape(donor)
End Function

Public Sub BB_ApplyTexture(ByVal target As Object, ByVal handle As Long)
    Engine.ApplyTexture target, handle
End Sub

Public Sub BB_ReleaseTexture(ByVal handle As Long)
    Engine.ReleaseTexture handle
End Sub

Public Sub BB_ClearTextures()
    If Not Engine Is Nothing Then Engine.ClearTextures
End Sub

Public Function BB_GetCapabilities() As String
    If Engine Is Nothing Then BB_Init
    BB_GetCapabilities = Engine.GetCapabilities()
End Function
