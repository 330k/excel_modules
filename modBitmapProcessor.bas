Attribute VB_Name = "modBitmapProcessor"
' Bitmap Process module by 330k
' Copyright (C) 2010 330k, All rights reserved.
Option Explicit

Public Function ImageMeanFilter(rgbaData() As RGBA, lngRadius As Long) As RGBA()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim n As Long
    
    Dim sngR As Single
    Dim sngG As Single
    Dim sngB As Single
    Dim sngA As Single
    
    Dim rgbaResult() As RGBA
    
    ReDim rgbaResult(0 To UBound(rgbaData, 1), 0 To UBound(rgbaData, 2)) As RGBA
    
    For j = 0 To UBound(rgbaData, 2)
        For i = 0 To UBound(rgbaData, 1)
            sngR = 0
            sngG = 0
            sngB = 0
            sngA = 0
            n = 0
            
            For l = IIf(j > lngRadius, j - lngRadius, 0) To IIf(j < UBound(rgbaData, 2) - lngRadius, j + lngRadius, UBound(rgbaData, 2))
                For k = IIf(i > lngRadius, i - lngRadius, 0) To IIf(i < UBound(rgbaData, 1) - lngRadius, i + lngRadius, UBound(rgbaData, 1))
                    If (l - j) * (l - j) + (k - i) * (k - i) < lngRadius * lngRadius Then
                        sngR = sngR + rgbaData(k, l).rgbRed
                        sngG = sngG + rgbaData(k, l).rgbGreen
                        sngB = sngB + rgbaData(k, l).rgbBlue
                        sngA = sngA + rgbaData(k, l).rgbAlpha
                        n = n + 1
                    End If
                Next
            Next
            
            rgbaResult(i, j).rgbRed = sngR / n
            rgbaResult(i, j).rgbGreen = sngG / n
            rgbaResult(i, j).rgbBlue = sngB / n
            rgbaResult(i, j).rgbAlpha = sngA / n
            
            
        Next
    Next
    
    ImageMeanFilter = rgbaResult
End Function

Public Function ImageCrop(rgbaData() As RGBA, lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long) As RGBA()
    Dim i As Long
    Dim j As Long
    
    Dim rgbaResult() As RGBA
    
    ReDim rgbaResult(0 To lngRight - lngLeft - 1, 0 To lngBottom - lngTop - 1) As RGBA
    
    For j = 0 To UBound(rgbaResult, 2)
        For i = 0 To UBound(rgbaResult, 1)
            rgbaResult(i, j) = rgbaData(i + lngLeft, j + UBound(rgbaData, 2) - lngBottom)
        Next
    Next
    
    ImageCrop = rgbaResult
End Function

Public Function ImageReduce(rgbaData() As RGBA, lngReductionRatio As Long) As RGBA()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long
    
    Dim rgbaResult() As RGBA
    Dim lngR As Long
    Dim lngG As Long
    Dim lngB As Long
    Dim lngA As Long
    
    ReDim rgbaResult(0 To (UBound(rgbaData, 1) + 1) \ lngReductionRatio - 1, 0 To (UBound(rgbaData, 2) + 1) \ lngReductionRatio - 1) As RGBA
    
    For j = 0 To UBound(rgbaResult, 2)
        For i = 0 To UBound(rgbaResult, 1)
            lngR = 0
            lngG = 0
            lngB = 0
            lngA = 0
            
            For l = j * lngReductionRatio To (j + 1) * lngReductionRatio - 1
                For k = i * lngReductionRatio To (i + 1) * lngReductionRatio - 1
                    lngR = lngR + rgbaData(k, l).rgbRed
                    lngG = lngG + rgbaData(k, l).rgbGreen
                    lngB = lngB + rgbaData(k, l).rgbBlue
                    lngA = lngA + rgbaData(k, l).rgbAlpha
                Next
            Next
            
            rgbaResult(i, j).rgbRed = lngR \ (lngReductionRatio * lngReductionRatio)
            rgbaResult(i, j).rgbGreen = lngG \ (lngReductionRatio * lngReductionRatio)
            rgbaResult(i, j).rgbBlue = lngB \ (lngReductionRatio * lngReductionRatio)
            rgbaResult(i, j).rgbAlpha = lngA \ (lngReductionRatio * lngReductionRatio)
        Next
    Next
    
    ImageReduce = rgbaResult
End Function
