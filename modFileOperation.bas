Attribute VB_Name = "modFileOperation"
' File operation module by 330k
' Copyright (C) 2010 330k, All rights reserved.
Option Explicit


Public Function DriveExists(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    DriveExists = objFSO.DriveExists(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function FileExists(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    FileExists = objFSO.FileExists(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function FolderExists(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    FolderExists = objFSO.FolderExists(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function GetAbsolutePathName(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    GetAbsolutePathName = objFSO.GetAbsolutePathName(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function GetBaseName(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    GetBaseName = objFSO.GetBaseName(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function GetExtensionName(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    GetExtensionName = objFSO.GetExtensionName(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function GetDriveName(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    GetDriveName = objFSO.GetDriveName(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function GetFileName(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    GetFileName = objFSO.GetFileName(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function GetParentFolderName(strFileName As String) As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    GetParentFolderName = objFSO.GetParentFolderName(strFileName)
    
    Set objFSO = Nothing
End Function

Public Function GetTempName() As String
    Dim objFSO As Object
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    GetTempName = objFSO.GetSpecialFolder(2) & "\" & objFSO.GetTempName()
    
    Set objFSO = Nothing
End Function

Public Function OpenFolderDialog(Optional strMessage As String = "") As String
    Dim oShell As Object
    Dim oFolder As Object
    
    Const ssfDRIVES = &H11
    Const BIF_RETURNONLYFSDIRS = &H1
    Const BIF_EDITBOX = &H10                ' IE5
    Const BIF_BROWSEINCLUDEFILES = &H4000   ' IE5

    Set oShell = CreateObject("Shell.Application")
    Set oFolder = oShell.BrowseForFolder(0, strMessage, 0, ssfDRIVES)
    Set oShell = Nothing
    
    If Not oFolder Is Nothing Then
        OpenFolderDialog = oFolder.Items.item.path
    End If
End Function
