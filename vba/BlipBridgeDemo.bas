Attribute VB_Name = "BlipBridgeDemo"
Option Explicit

Public Sub DemoPreloadedFill()
    Dim slide As Slide, donor As Shape, target As Shape, handle As Long

    Set slide = ActivePresentation.Slides(1)
    Set donor = slide.Shapes("texture_donor")
    Set target = slide.Shapes("poly_17")

    BB_Init

    Debug.Print BB_GetCapabilities()
    handle = BB_RegisterTextureShape(donor)

    BB_ApplyTexture target, handle
    BB_ReleaseTexture handle
End Sub
