Attribute VB_Name = "modStringOperation"
' String operation module by 330k
' Copyright (C) 2010 330k, All rights reserved.
Option Explicit

' Return whether the pattern matches the string or not
Public Function StringMatchQ(strExpression As String, strPattern As String, Optional bGlobal As Boolean = False, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False) As Boolean
    Dim objReg          As Object
    
    Set objReg = CreateObject("VBScript.RegExp")
    objReg.Pattern = strPattern
    objReg.Global = bGlobal
    objReg.IgnoreCase = bIgnoreCase
    objReg.MultiLine = bMultiLine
    
    StringMatchQ = objReg.test(strExpression)
    
    Set objReg = Nothing
End Function

' Return matches as 2-dimension array of String
Public Function StringCases(strExpression As String, strPattern As String, Optional bGlobal As Boolean = False, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False) As String()
    Dim objReg          As Object
    Dim objMatches      As Object
    Dim objMatch        As Object
    Dim strMatches()    As String
    Dim i               As Long
    Dim j               As Long

    i = 0

    Set objReg = CreateObject("VBScript.RegExp")
    objReg.Pattern = strPattern
    objReg.Global = bGlobal
    objReg.IgnoreCase = bIgnoreCase
    objReg.MultiLine = bMultiLine

    Set objMatches = objReg.Execute(strExpression)
    If objMatches.count > 0 Then
        For Each objMatch In objMatches
            If objMatch.SubMatches.count > 0 Then
                strMatches(i, 0) = objMatch.SubMatches(0)
            Else
                ReDim Preserve strMatches(objMatches.count - 1, objMatch.SubMatches.count - 1)
                For j = 0 To objMatch.SubMatches.count - 1
                    strMatches(i, j) = objMatch.SubMatches(j)
                Next
            End If
            i = i + 1
        Next
    End If
    
    StringCases = strMatches
    
    Set objReg = Nothing
End Function

' Replace string with regular expression
Public Function StringReplace(strExpression As String, strPattern As String, strReplace As String, Optional bGlobal As Boolean = False, Optional bIgnoreCase As Boolean = False, Optional bMultiLine As Boolean = False) As String
    Dim objReg          As Object
    
    Set objReg = CreateObject("VBScript.RegExp")
    objReg.Pattern = strPattern
    objReg.Global = bGlobal
    objReg.IgnoreCase = bIgnoreCase
    objReg.MultiLine = bMultiLine
    
    StringReplace = objReg.Replace(strExpression, strReplace)
    
    Set objReg = Nothing
End Function
