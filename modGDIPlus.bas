Attribute VB_Name = "modGDIPlus"
' GDI+ module by 330k
' Copyright (C) 2010 330k, All rights reserved.
Option Explicit

Private Declare Function GdipPrivateAddMemoryFont Lib "gdiplus" (ByVal fontCollection As Long, ByVal memory As Long, ByVal length As Long) As Long
Private Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (fontCollection As Long) As Long
Private Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal family As Long, ByVal Name As Long, ByVal language As Integer) As Long
Private Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal fontCollection As Long, numFound As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal fontCollection As Long, ByVal numSought As Long, gpfamilies As Long, numFound As Long) As Long
Private Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (fontCollection As Long) As Long
Private Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal fontCollection As Long, ByVal filename As Long) As Long
Private Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (fontCollection As Long) As Long
Private Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal pen As Long, ByVal brush As Long) As Long
Private Declare Function GdipCreateTexture Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As Long, texture As Long) As Long
Private Declare Function GdipTranslateTextureTransform Lib "gdiplus" (ByVal brush As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As Long) As Long
Private Declare Function GdipResetTextureTransform Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipGetTextureImage Lib "gdiplus" (ByVal brush As Long, Image As Long) As Long
Private Declare Function GdipGetCompositingMode Lib "gdiplus" (ByVal graphics As Long, CompositingMd As Long) As Long
Private Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal graphics As Long, ByVal CompositingMd As Long) As Long
Private Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal graphics As Long, ByVal region As Long, ByVal CombineMd As Long) As Long
Private Declare Function GdipCreateRegion Lib "gdiplus" (region As Long) As Long
Private Declare Function GdipSetEmpty Lib "gdiplus" (ByVal region As Long) As Long
Private Declare Function GdipSetInfinite Lib "gdiplus" (ByVal region As Long) As Long
Private Declare Function GdipResetClip Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipCreateRegionRectI Lib "gdiplus" (Rect As Rect, region As Long) As Long
Private Declare Function GdipSetPenMode Lib "gdiplus" (ByVal pen As Long, ByVal penMode As Long) As Long
Private Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal region As Long, Rect As Rect, ByVal CombineMd As Long) As Long
Private Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal region As Long, ByVal region2 As Long, ByVal CombineMd As Long) As Long
Private Declare Function GdipGetRegionHRgn Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, hRgn As Long) As Long
Private Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, result As Long) As Long
Private Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal region As Long) As Long
Private Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal region As Long, ByVal path As Long, ByVal CombineMd As Long) As Long
Private Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal path As Long, region As Long) As Long
Private Declare Function GdipCreateMatrix Lib "gdiplus" (matrix As Long) As Long
Private Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal matrix As Long, cloneMatrix As Long) As Long
Private Declare Function GdipMultiplyMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal matrix2 As Long, ByVal order As Long) As Long
Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal order As Long) As Long
Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal Angle As Single, ByVal order As Long) As Long
Private Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal scaleX As Single, ByVal scaleY As Single, ByVal order As Long) As Long
Private Declare Function GdipShearMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal shearX As Single, ByVal shearY As Single, ByVal order As Long) As Long
Private Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal matrix As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal path As Long) As Long
Private Declare Function GdipClonePath Lib "gdiplus" (ByVal path As Long, clonePath As Long) As Long
Private Declare Function GdipCloneRegion Lib "gdiplus" (ByVal region As Long, cloneRegion As Long) As Long
Private Declare Function GdipTransformPath Lib "gdiplus" (ByVal path As Long, ByVal matrix As Long) As Long
Private Declare Function GdipTransformRegion Lib "gdiplus" (ByVal region As Long, ByVal matrix As Long) As Long
Private Declare Function GdipTransformMatrixPointsI Lib "gdiplus" (ByVal matrix As Long, pts As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal path As Long) As Long
Private Declare Function GdipFillRegion Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal region As Long) As Long
Private Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal region As Long, ByVal X As Single, ByVal Y As Single, ByVal graphics As Long, result As Long) As Long
Private Declare Function GdipIsVisibleRegionPoint Lib "gdiplus" (ByVal region As Long, ByVal X As Single, ByVal Y As Single, ByVal graphics As Long, result As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As Long, path As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipAddPathPath Lib "gdiplus" (ByVal path As Long, ByVal addingPath As Long, ByVal bConnect As Long) As Long
Private Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal path As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GdipAddPathPolygonI Lib "gdiplus" (ByVal path As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipAddPathEllipseI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipAddPathArcI Lib "gdiplus" (ByVal path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, Points As POINTAPI, ByVal count As Long, ByVal FillMd As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipDrawBeziersI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipAddPathBeziersI Lib "gdiplus" (ByVal path As Long, Points As POINTAPI, ByVal count As Long) As Long
Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As Long
Private Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal graphics As Long, ByVal lColor As Long) As Long
Private Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal Image As Long, Format As GUID) As Long
Private Declare Function GdipRecordMetafileI Lib "gdiplus" (ByVal referenceHdc As Long, ByVal etype As Long, frameRect As Rect, ByVal frameUnit As Long, ByVal Description As Long, metafile As Long) As Long
Private Declare Function GdipGetDC Lib "gdiplus" (ByVal graphics As Long, hdc As Long) As Long
Private Declare Function GdipReleaseDC Lib "gdiplus" (ByVal graphics As Long, ByVal hdc As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal PixelFormat As Long, scan0 As Any, bitmap As Long) As Long
Private Declare Function GdipSetImagePalette Lib "gdiplus" (ByVal Image As Long, palette As ColorPalette) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (Rect As Rect, ByVal color1 As Long, ByVal color2 As Long, ByVal mode As Long, ByVal WrapMd As Long, lineGradient As Long) As Long
Private Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal style As Long, ByVal forecolr As Long, ByVal backcolr As Long, brush As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipBitmapSetResolution Lib "gdiplus" (ByVal bitmap As Long, ByVal xdpi As Single, ByVal ydpi As Single) As Long
Private Declare Function GdipGetPathWorldBoundsI Lib "gdiplus" (ByVal path As Long, bounds As Rect, ByVal matrix As Long, ByVal pen As Long) As Long
Private Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, Rect As Rect) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, graphics As Long) As Long
Private Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal graphics As Long, bitmap As Long) As Long
Private Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, cloneImage As Long) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, PixelFormat As Long) As Long
Private Declare Function GdipGetPropertySize Lib "gdiplus" (ByVal Image As Long, totalBufferSize As Long, numProperties As Long) As Long
Private Declare Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal Image As Long, ByVal totalBufferSize As Long, ByVal numProperties As Long, allItems As Any) As Long
Private Declare Function GdipRemovePropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long) As Long
Private Declare Function GdipAlloc Lib "gdiplus.dll" (ByVal Size As Long) As Long
Private Declare Function GdipFree Lib "gdiplus.dll" (ByVal Ptr As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal stream As Object, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, Id As GUID) As Long
Private Declare Function StringFromCLSID Lib "ole32.dll" (pCLSID As GUID, lpszProgID As Long) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal filename As Long, ByRef bitmap As Long) As Long
Private Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, _
                                                            ByVal propSize As Long, ByRef Buffer As Any) As Long
Private Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal propId As Long, _
                                                                ByRef Size As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, LInput As GdiplusStartupInput, Optional ByVal lOutPut As Long = 0) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal bitmap As Long, ByRef hbmReturn As Long, _
                                                                    ByVal Background As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, bitmap As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipImageSelectActiveFrame Lib "gdiplus" _
                                                    (ByVal Image As Long, ByRef dimensionID As GUID, _
                                                     ByVal frameIndex As Long) As Long
Private Declare Function GdipImageGetFrameCount Lib "gdiplus" _
                                                (ByVal Image As Long, ByRef dimensionID As GUID, _
                                                 ByRef count As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" _
                                               (ByVal Image As Long, ByRef Width As Single, _
                                                ByRef Height As Single) As Long
Private Declare Function GdipSetPropertyItem Lib "gdiplus" (ByVal nImage As Long, item As PropertyItem) As Long
Private Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal Image As Long, resolution As Single) As Long
Private Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal Image As Long, resolution As Single) As Long
Private Declare Function GdipGetPropertyCount Lib "gdiplus" (ByVal Image As Long, numOfProperty As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal bitmap As Long, Rect As Rect, ByVal flags As Long, ByVal PixelFormat As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal bitmap As Long, lockedBitmapData As BitmapData) As Long
Private Declare Function GdipResetWorldTransform Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipGetWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal matrix As Long) As Long
Private Declare Function GdipSetWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal matrix As Long) As Long
Private Declare Function GdipScaleWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal sx As Single, ByVal sy As Single, ByVal order As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dx As Single, ByVal dy As Single, ByVal order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal order As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As Long) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal align As Long) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal graphics As Long, ByVal str As Long, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, boundingBox As RECTF, codepointsFitted As Long, linesFilled As Long) As Long
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal graphics As Long, ByVal mode As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal graphics As Long, ByVal str As Long, ByVal length As Long, ByVal thefont As Long, layoutRect As RECTF, ByVal StringFormat As Long, ByVal brush As Long) As Long
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal fontFamily As Long, ByVal emSize As Single, ByVal style As Long, ByVal unit As Long, createdfont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (nativeFamily As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipSetImageAttributesColorKeys Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As Long, ByVal enableFlag As Long, ByVal colorLow As Long, ByVal colorHigh As Long) As Long
Private Declare Function GdipSetImageAttributesRemapTable Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As Long, ByVal enableFlag As Long, ByVal mapSize As Long, map As ColorMap) As Long
Private Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal imageattr As Long, ByVal wrap As Long, ByVal argb As Long, ByVal bClamp As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal interpolation As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, SmoothingMd As Long) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As Long
Private Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal pen As Long, ByVal startCap As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal pen As Long, ByVal endCap As Long) As Long
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal pen As Long, ByVal LnJoin As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, graphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As Long, pen As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As Long
Private Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal bitmap As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ClrAdjType As Long, ByVal enableFlag As Long, colourMatrix As ColorMatrix, grayMatrix As Any, ByVal flags As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal dstx As Long, _
                                                               ByVal dsty As Long, ByVal dstwidth As Long, ByVal dstheight As Long, _
                                                               ByVal srcx As Long, ByVal srcy As Long, ByVal srcwidth As Long, ByVal srcheight As Long, _
                                                               ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, _
                                                               Optional ByVal CallBack As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipGetHemfFromMetafile Lib "gdiplus" (ByVal metafile As Long, hemf As Long) As Long
Private Declare Function GdipGetRegionBounds Lib "gdiplus" (ByVal region As Long, ByVal graphics As Long, Rect As RECTF) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal length As Long)

' Pour region auto
Private Type TRgnPoints
    o As Boolean
    T As Boolean
End Type
' Information icone
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
' Info scrollbar
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * 260
   szTypeName As String * 80
End Type
Private Type BitmapData
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As Long
    scan0 As Long
    Reserved As Long
End Type
Private Type ColorMap
    oldColor As Long
    newColor As Long
End Type
Private Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type
Private Type PicBmp
    Size As Long
    tType As Long
    hBmp As Long
    hpal As Long
    Reserved As Long
End Type
' Rectangle pour API
Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' Point pour API
Private Type POINTAPI
    X As Long
    Y As Long
End Type
' Rectangle pour API (single)
Private Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type
Private Type bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type PropertyItem
    Id As Long
    length As Long
Type As Integer
    Value As Long
End Type
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Type EncoderParameter
    GUID As GUID
    NumberOfValues As Long
    Type As Long
    Value As Long
End Type
Private Type EncoderParameters
    count As Long
    Parameter(0 To 15) As EncoderParameter
End Type
Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Type argb
    blue As Byte    ' Bleu
    green As Byte    ' Vert
    red As Byte    ' Rouge
    Alpha As Byte    ' Luminosite
End Type
Private Type BitmapInfoHeader
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
' Pour lire les fichiers integres type EMF (enhanced metafile)
Private Type SIZEL
    cx As Long
    cy As Long
End Type
Private Type ENHMETAHEADER
    iType As Long
    nSize As Long
    rclBounds As Rect
    rclFrame As Rect
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As SIZEL
    szlMillimeters As SIZEL
    cbPixelFormat As Long
    offPixelFormat As Long
    bOpenGL As Long
End Type
Private Type ColorPalette
    flags As Long
    count As Long
    Entries(0 To 255) As Long
End Type

Public Type RGBA
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbAlpha As Byte
End Type



Private Const OBJ_BITMAP As Long = 7
Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B&
Private Const HIMETRIC_INCH = 2540          ' Pour conversion Pouce<->Himetric
Private Const MM_HIMETRIC = 3
Private Const STAP_ALLOW_CONTROLS = &H2
Private Const STAP_ALLOW_NONCLIENT = &H1
Private Const STAP_ALLOW_WEBCONTENT = &H4
Private Const GMEM_MOVEABLE = &H2&
Private Const SRCCOPY = &HCC0020
Private Const SRCAND = &H8800C6
Private Const SRCPAINT = &HEE0086
Private Const CF_ENHMETAFILE = 14
Private Const CF_BITMAP = 40
Private Const LOGPIXELSY = 90
Private Const LOGPIXELSX = 88
Private Const PropertyTagFrameDelay As Long = &H5100&
Private Const PropertyTagTypeByte = 1
Private Const PropertyTagTypeASCII = 2
Private Const PropertyTagTypeShort = 3
Private Const PropertyTagTypeLong = 4
Private Const PropertyTagTypeRational = 5
Private Const PropertyTagTypeUndefined = 7
Private Const PropertyTagTypeSLong = 9
Private Const PropertyTagTypeSRational = 10
Private Const PixelFormat1bppIndexed = &H30101
Private Const PixelFormat4bppIndexed = &H30402
Private Const PixelFormat8bppIndexed = &H30803
Private Const PixelFormat16bppGreyScale = &H101004
Private Const PixelFormat16bppRGB555 = &H21005
Private Const PixelFormat16bppRGB565 = &H21006
Private Const PixelFormat16bppARGB1555 = &H61007
Private Const PixelFormat24bppRGB = &H21808
Private Const PixelFormat32bppRGB = &H22009
Private Const PixelFormat32bppARGB = &H26200A
Private Const PixelFormat32bppPARGB = &HE200B
Private Const PixelFormat48bppRGB = &H10300C
Private Const PixelFormat64bppARGB = &H34400D
Private Const PixelFormat64bppPARGB = &H1C400E
Private Const WS_EX_COMPOSITED = &H2000000
Private Const GWL_EXSTYLE = &HFFEC
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const MM_TEXT = 1
Private Const SHGFI_ICON = &H100

Private Const QUALITY_PARAMS As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const ENCODER_BMP    As String = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
Private Const ENCODER_JPG    As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
Private Const ENCODER_GIF    As String = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
Private Const ENCODER_TIF    As String = "{557CF405-1A04-11D3-9A73-0000F81EF32E}"
Private Const ENCODER_PNG    As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"

Public Enum FileFormat
    FileFormat_Bitmap = 0
    FileFormat_JPEG = 1
    FileFormat_GIF = 2
    FileFormat_TIFF = 3
    FileFormat_PNG = 4
End Enum

Public Function LoadPictureGDIP(strFileName As String) As RGBA()
    Dim gsiInput As GdiplusStartupInput
    Dim lngGDIPlusToken As Long
    Dim lngResult As Long
    Dim lngBitmap As Long
    Dim lngWidth As Single
    Dim lngHeight As Single
    Dim rctRect As Rect
    Dim bdBitmapData As BitmapData
    Dim rgbaData() As RGBA
    Dim i As Long
    Dim j As Long
    
    With gsiInput
        .GdiplusVersion = 1
        .DebugEventCallback = 0
        .SuppressBackgroundThread = 0
        .SuppressExternalCodecs = 0
    End With
    
    lngResult = GdiplusStartup(lngGDIPlusToken, gsiInput, 0)
    If lngResult = 0 Then
    
        lngResult = GdipLoadImageFromFile(StrPtr(strFileName), lngBitmap)
        If lngResult = 0 Then
            lngResult = GdipGetImageDimension(lngBitmap, lngWidth, lngHeight)
            
            rctRect.Right = lngWidth
            rctRect.Bottom = lngHeight
            lngResult = GdipBitmapLockBits(lngBitmap, rctRect, &H1, PixelFormat32bppARGB, bdBitmapData)

            ReDim rgbaData(0 To Abs(bdBitmapData.stride) \ 4 - 1, 0 To bdBitmapData.Height - 1) As RGBA
            
            RtlMoveMemory rgbaData(0, 0), ByVal bdBitmapData.scan0, Abs(bdBitmapData.stride) * bdBitmapData.Height
            
            LoadPictureGDIP = rgbaData
            
            lngResult = GdipBitmapUnlockBits(lngBitmap, bdBitmapData)
            lngResult = GdipDisposeImage(lngBitmap)
        End If
    
        lngResult = GdiplusShutdown(lngGDIPlusToken)
    End If
    
End Function

Public Function SavePictureGDIP(strFileName As String, rgbaData() As RGBA, Optional ffFileFormat As FileFormat = FileFormat_Bitmap) As Long
    Dim gsiInput As GdiplusStartupInput
    Dim lngGDIPlusToken As Long
    Dim lngResult As Long
    Dim lngBitmap As Long
    Dim lngWidth As Single
    Dim lngHeight As Single
    Dim rctRect As Rect
    Dim bdBitmapData As BitmapData
    Dim strCLSID(0 To 4) As String
    Dim guidID As GUID
    Dim strFileFormat(0 To 4) As String

    With gsiInput
        .GdiplusVersion = 1
        .DebugEventCallback = 0
        .SuppressBackgroundThread = 0
        .SuppressExternalCodecs = 0
    End With
    
    lngResult = GdiplusStartup(lngGDIPlusToken, gsiInput, 0)
    
    If lngResult = 0 Then
        lngWidth = UBound(rgbaData, 1) + 1
        lngHeight = UBound(rgbaData, 2) + 1
        
        lngResult = GdipCreateBitmapFromScan0(lngWidth, lngHeight, 4 * lngWidth, PixelFormat32bppARGB, rgbaData(0, 0), lngBitmap)
        If lngResult = 0 Then
            strFileFormat(0) = ENCODER_BMP
            strFileFormat(1) = ENCODER_JPG
            strFileFormat(2) = ENCODER_GIF
            strFileFormat(3) = ENCODER_TIF
            strFileFormat(4) = ENCODER_PNG
            
            lngResult = CLSIDFromString(StrPtr(strFileFormat(ffFileFormat)), guidID)
            lngResult = GdipSaveImageToFile(lngBitmap, StrPtr(strFileName), guidID, ByVal 0)
            
            lngResult = GdipDisposeImage(lngBitmap)
        End If
    
        lngResult = GdiplusShutdown(lngGDIPlusToken)
    End If
    
    SavePictureGDIP = lngResult
End Function
