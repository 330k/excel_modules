Attribute VB_Name = "modXMLHTTPRequest"
' MSXML2.XMLHTTPRequest module by 330k
' Copyright (C) 2010 330k, All rights reserved.
Option Explicit

Public Function GetHTTPResponseAsString(strURI As String) As String
    Dim objHTTP  As Object
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    
    objHTTP.Open "GET", strURI, False
    objHTTP.Send
    
    GetHTTPResponseAsString = objHTTP.responseText
    
    Set objHTTP = Nothing
End Function

Public Function GetHTTPResponseAsBinary(strURI As String) As Byte()
    Dim objHTTP  As Object
    Dim bytResponse() As Byte
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    
    objHTTP.Open "GET", strURI, False
    objHTTP.Send
    
    ReDim bytResponse(0 To Len(objHTTP.responseBody) - 1) As Byte
    
    bytResponse = objHTTP.responseBody
    GetHTTPResponseAsBinary = bytResponse
    
    Set objHTTP = Nothing
End Function

