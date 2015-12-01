Attribute VB_Name = "modBitmap"
' Bitmap module by 330k
' Copyright (C) 2010 330k, All rights reserved.
Option Explicit

Public Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPFILEHEADER
    bfType As String * 2
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bjOffBits As Long
End Type
Private Type BitmapInfoHeader
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImaze As Long
    biXPixPerMeter As Long
    biYPixPerMeter As Long
    biClrUsed As Long
    biClrImporant As Long
End Type

' Read a bitmap file (1, 4, 8, 24 and 32 bit) and return image as 2-dimension array of RGBTriple
Public Function ReadBitmap(strFileName As String) As RGBTRIPLE()
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim intFileNumber As Integer
    Dim bjHeader As BITMAPFILEHEADER
    Dim biHeader As BitmapInfoHeader
    Dim lngColors As Long
    Dim rgbData() As RGBTRIPLE
    Dim rgbTemp As RGBTRIPLE
    Dim rgbTable3() As RGBTRIPLE
    Dim rgbTable4() As RGBQUAD
    Dim bytTemp As Byte
    
    intFileNumber = FreeFile()
    Open strFileName For Binary As intFileNumber
        Get intFileNumber, , bjHeader
        Get intFileNumber, , biHeader
        
        ReDim rgbData(0 To biHeader.biHeight - 1, 0 To biHeader.biWidth - 1) As RGBTRIPLE
        n = (4 - (-(Int(-biHeader.biWidth * (biHeader.biBitCount / 8))) Mod 4)) Mod 4
        lngColors = IIf(biHeader.biClrUsed = 0, 2 ^ biHeader.biBitCount, biHeader.biClrUsed)
        
        Select Case biHeader.biBitCount
        Case 1
            ReDim rgbTable4(0 To lngColors - 1) As RGBQUAD
            Get intFileNumber, , rgbTable4
            rgbTable3 = ConvertRGBQuadToRGBTriple(rgbTable4)
            
            For i = UBound(rgbData, 1) To 0 Step -1
                For j = 0 To UBound(rgbData, 2) Step 8
                    Get intFileNumber, , bytTemp
                    
                    rgbData(i, j) = rgbTable3(bytTemp \ 128)
                    If j + 1 <= UBound(rgbData, 2) Then rgbData(i, j + 1) = rgbTable3(bytTemp \ 64 And 1)
                    If j + 2 <= UBound(rgbData, 2) Then rgbData(i, j + 2) = rgbTable3(bytTemp \ 32 And 1)
                    If j + 3 <= UBound(rgbData, 2) Then rgbData(i, j + 3) = rgbTable3(bytTemp \ 16 And 1)
                    If j + 4 <= UBound(rgbData, 2) Then rgbData(i, j + 4) = rgbTable3(bytTemp \ 8 And 1)
                    If j + 5 <= UBound(rgbData, 2) Then rgbData(i, j + 5) = rgbTable3(bytTemp \ 4 And 1)
                    If j + 6 <= UBound(rgbData, 2) Then rgbData(i, j + 6) = rgbTable3(bytTemp \ 2 And 1)
                    If j + 7 <= UBound(rgbData, 2) Then rgbData(i, j + 7) = rgbTable3(bytTemp And 1)
                    
                Next
                For j = 1 To n
                    Get intFileNumber, , bytTemp
                Next
            Next
            
        Case 4
            ReDim rgbTable4(0 To lngColors - 1) As RGBQUAD
            Get intFileNumber, , rgbTable4
            rgbTable3 = ConvertRGBQuadToRGBTriple(rgbTable4)
            
            For i = UBound(rgbData, 1) To 0 Step -1
                For j = 0 To UBound(rgbData, 2) Step 2
                    Get intFileNumber, , bytTemp
                    
                    rgbData(i, j) = rgbTable3(bytTemp \ 16)
                    If j + 1 <= UBound(rgbData, 2) Then rgbData(i, j + 1) = rgbTable3(bytTemp And 15)
                Next
                For j = 1 To n
                    Get intFileNumber, , bytTemp
                Next
            Next
            
        Case 8
            ReDim rgbTable4(0 To lngColors - 1) As RGBQUAD
            Get intFileNumber, , rgbTable4
            rgbTable3 = ConvertRGBQuadToRGBTriple(rgbTable4)
            
            For i = UBound(rgbData, 1) To 0 Step -1
                For j = 0 To UBound(rgbData, 2)
                    Get intFileNumber, , bytTemp
                    rgbData(i, j) = rgbTable3(bytTemp)
                Next
                For j = 1 To n
                    Get intFileNumber, , bytTemp
                Next
            Next
            
        Case 24
            For i = UBound(rgbData, 1) To 0 Step -1
                For j = 0 To UBound(rgbData, 2)
                    Get intFileNumber, , rgbData(i, j)
                Next
                For j = 1 To n
                    Get intFileNumber, , bytTemp
                Next
            Next
            
        Case 32
            ReDim rgbTable4(0 To biHeader.biHeight - 1, 0 To biHeader.biWidth - 1) As RGBQUAD
            
            Get intFileNumber, , rgbTable4
            rgbData = ConvertRGBQuadToRGBTriple2(rgbTable4)
            
        End Select
    Close
    
    ReadBitmap = rgbData
End Function

' Write a bitmap file (24-bit only) from 2-dimension array of RGBTriple
Public Sub WriteBitmap24(strFileName As String, rgbData() As RGBTRIPLE)
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim intFileNumber As Integer
    Dim bjHeader As BITMAPFILEHEADER
    Dim biHeader As BitmapInfoHeader
    Dim bytTemp As Byte
    
    lngHeight = UBound(rgbData, 1) + 1
    lngWidth = UBound(rgbData, 2) + 1
    n = (4 - (lngWidth * 3 Mod 4)) Mod 4
    
    With bjHeader
        .bfType = "BM"
        .bfSize = Len(bjHeader) + Len(biHeader) + 3 * lngHeight * lngWidth
        .bjOffBits = Len(bjHeader) + Len(biHeader)
    End With
    With biHeader
        .biSize = 40
        .biWidth = lngWidth
        .biHeight = lngHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = 0
        .biSizeImaze = 3 * lngHeight * lngWidth
        .biXPixPerMeter = 3780
        .biYPixPerMeter = 3780
        .biClrUsed = 0
        .biClrImporant = 0
    End With
    
    If Len(Dir(strFileName)) Then
        Kill strFileName
    End If
    
    intFileNumber = FreeFile()
    Open strFileName For Binary As intFileNumber
        Put intFileNumber, , bjHeader
        Put intFileNumber, , biHeader
        
        For i = lngHeight - 1 To 0 Step -1
            For j = 0 To lngWidth - 1
                Put intFileNumber, , rgbData(i, j)
            Next
            For j = 1 To n
                Put intFileNumber, , bytTemp
            Next
        Next
    Close
End Sub

' Read GIF or JPEG files as RGBTRIPLE()
Public Function ReadGIFJPEG(strFileName As String) As RGBTRIPLE()
    Dim objPicture As IPictureDisp
    Dim strTempFile As String
    
    Set objPicture = LoadPicture(strFileName)
    strTempFile = GetTempName()
    
    SavePicture objPicture, strTempFile
    
    ReadGIFJPEG = ReadBitmap(strTempFile)
    
    Kill strTempFile
End Function

' Private Functions
Private Function ConvertRGBQuadToRGBTriple(rgbSource() As RGBQUAD) As RGBTRIPLE()
    Dim i As Long
    Dim j As Long
    Dim rgbResult() As RGBTRIPLE
    
    ReDim rgbResult(0 To UBound(rgbSource, 1)) As RGBTRIPLE
    
    For i = 0 To UBound(rgbSource, 1)
        rgbResult(i).rgbBlue = rgbSource(i).rgbBlue
        rgbResult(i).rgbGreen = rgbSource(i).rgbGreen
        rgbResult(i).rgbRed = rgbSource(i).rgbRed
    Next
    
    ConvertRGBQuadToRGBTriple = rgbResult
End Function

Private Function ConvertRGBQuadToRGBTriple2(rgbSource() As RGBQUAD) As RGBTRIPLE()
    Dim i As Long
    Dim j As Long
    Dim rgbResult() As RGBTRIPLE
    
    ReDim rgbResult(0 To UBound(rgbSource, 1), 0 To UBound(rgbSource, 2)) As RGBTRIPLE
    
    For i = 0 To UBound(rgbSource, 1)
        For j = 0 To UBound(rgbSource, 2)
            rgbResult(i, j).rgbBlue = rgbSource(i, j).rgbBlue
            rgbResult(i, j).rgbGreen = rgbSource(i, j).rgbGreen
            rgbResult(i, j).rgbRed = rgbSource(i, j).rgbRed
        Next
    Next
    
    ConvertRGBQuadToRGBTriple2 = rgbResult
End Function
