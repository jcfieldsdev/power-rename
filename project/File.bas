Attribute VB_Name = "File"
Option Explicit

Public Type FileToRename
    FullPath As String
    DirectoryName As String
    OldFileName As String
    NewFileName As String
    
    FileExists As Boolean
End Type

Public Function OpenFile(DirectoryName As String, FileName As String) As FileToRename
    Dim File As FileToRename
    
    With File
        Dim FullPath As String: FullPath = GetFullPath(DirectoryName, FileName)
        Dim FileExists As Boolean: FileExists = CheckIfFileExists(FullPath)
        
        If FileExists Then
            .FullPath = GetFullPath(DirectoryName, FileName)
            .DirectoryName = DirectoryName
            .OldFileName = FileName
            .NewFileName = FileName
        End If
        
        .FileExists = FileExists
    End With
    
    OpenFile = File
End Function

Public Sub RenameFile(File As FileToRename)
    Name File.OldFileName As File.NewFileName
    File.OldFileName = File.NewFileName
End Sub

Public Function GetFullPath(DirectoryName As String, FileName As String) As String
    GetFullPath = DirectoryName & "\" & FileName
End Function

Public Function CheckIfFileExists(FullPath As String) As Boolean
    On Error GoTo FILE_EXISTS_ERROR
    
    Dim Attributes As Integer: Attributes = GetAttr(FullPath)
    
    CheckIfFileExists = True
    Exit Function
    
FILE_EXISTS_ERROR:
    CheckIfFileExists = False
    Exit Function
End Function
