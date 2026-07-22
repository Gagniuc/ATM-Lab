Attribute VB_Name = "modDestroyFile"
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
Option Explicit

Private Const GENERIC_WRITE = &H40000000
Private Const TRUNCATE_EXISTING = 5
Private Const FILE_SHARE_WRITE = &H2


Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) _
As Long




Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100



Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" _
    (ByVal lpFileSpec As String) As Long
    
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" _
    (ByVal lpFileSpec As String, ByVal dwFileAttributes As Long) As Long
    
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long



Public Function DestroyFile(dFile As String) As Boolean

Dim fhnd As Long
Dim fSize As Long
     
     fhnd = CreateFile(dFile, GENERIC_WRITE, FILE_SHARE_WRITE, 0, TRUNCATE_EXISTING, FILE_FLAG_NO_BUFFERING Or FILE_FLAG_WRITE_THROUGH, 0)
   
   If fhnd = -1 Then
        Call CloseHandle(fhnd)
        DestroyFile = False
        Exit Function
     Else
        Call FlushFileBuffers(fhnd)
        Call CloseHandle(fhnd)
        fSize = FileLen(dFile)

            If fSize <> 0 Then
                DestroyFile = False
            Else
                DestroyFile = True
            End If
    End If
 
End Function


Public Function RenameAndKill(rFile As String) As Boolean
Dim TempName As String
Dim fDrive As String


      
        fDrive = Left$(rFile, 3)
        
        
        RenameAndKill = False
        
        TempName = fDrive & Format(Date, "ddmm") & Format(Time(), "mmss") & ".pag"
                  
                
           Name rFile As TempName
           Kill TempName
           RenameAndKill = True
        
         
End Function


Public Function ValidFileAttributes(aFile As String) As Boolean
 
 Dim fAttr As Long
 
 fAttr = GetFileAttributes(aFile)

    If fAttr And FILE_ATTRIBUTE_NORMAL Then
        ValidFileAttributes = True
    End If
    
    If fAttr And FILE_ATTRIBUTE_ARCHIVE Then
        ValidFileAttributes = True
    End If
    
    If fAttr And FILE_ATTRIBUTE_SYSTEM Then
        ValidFileAttributes = False
    End If
    
    If fAttr And FILE_ATTRIBUTE_HIDDEN Then
        ValidFileAttributes = False
    End If
    
    If fAttr And FILE_ATTRIBUTE_READONLY Then
        ValidFileAttributes = False
    End If
    
    If fAttr And FILE_ATTRIBUTE_TEMPORARY Then
        ValidFileAttributes = False
    End If

End Function

Public Sub NormalAttributes(aFile As String)
   SetFileAttributes aFile, FILE_ATTRIBUTE_NORMAL
End Sub


Public Function FileExist(aFile As String) As Boolean
 If Dir$(aFile) = "" Then
         If Dir$(aFile, vbHidden) = "" Then
                 FileExist = False
            Else
                 FileExist = True
         End If
    Else
        FileExist = True
End If
End Function
