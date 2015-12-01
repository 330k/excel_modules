Attribute VB_Name = "modArchive"
' Archive file (*.zip) operation module by 330k
' Copyright (C) 2010 330k, All rights reserved.
Option Explicit

' Extract zip file and return contained file pathes
Public Function ExtractZip(strZipFileName As String, varTargetFolder As Variant) As String()
    Dim objShell As Object
    Dim objIE As Object
    Dim objZip As Object
    Dim objTargetFolder As Object
    
    Set objShell = CreateObject("Shell.Application")
    Set objIE = GetObject("new:{C08AFD90-F2A1-11D1-8455-00A0C91F3880}")
    
    objIE.Navigate strZipFileName
    Set objZip = objIE.Document.Folder
    Set objTargetFolder = objShell.Namespace(varTargetFolder)
    
    objTargetFolder.CopyHere objZip.Items
    
    objIE.Quit
    
    Set objTargetFolder = Nothing
    Set objZip = Nothing
    Set objIE = Nothing
    Set objShell = Nothing
End Function
