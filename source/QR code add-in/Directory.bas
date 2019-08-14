Attribute VB_Name = "Directory"
Option Private Module
Option Explicit


Private Const TemporaryFolder = 2

Public Function FileExists(ByVal FileSpec As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(FileSpec)
End Function

Public Sub DeleteFile(ByVal FileSpec As String, Optional ByVal Force As Boolean = False)
    Call CreateObject("Scripting.FileSystemObject").DeleteFile(FileSpec, Force)
End Sub

Public Function GetParentFolderName(ByVal Path As String) As String
    GetParentFolderName = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Path)
End Function

Public Function GetBaseName(ByVal Path As String) As String
    GetBaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(Path)
End Function

Function GetFileName(ByVal Path As String) As String
    GetFileName = CreateObject("Scripting.FileSystemObject").GetFileName(Path)
End Function

Public Function GetTempFolderName() As String
    GetTempFolderName = CreateObject("Scripting.FileSystemObject").GetSpecialFolder(TemporaryFolder)
End Function

Public Function GetTempFileName() As String
    GetTempFileName = CreateObject("Scripting.FileSystemObject").GetTempName()
End Function

Public Function GetTempPath() As String
    GetTempPath = GetTempFolderName() & "\" & GetTempFileName()
End Function
