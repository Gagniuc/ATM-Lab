Attribute VB_Name = "Module1"
Option Explicit
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0                                                                                     '
'==========================================================================================================='
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Type OSVERSIONINFO
        Dimensiune_ver_OS As Long
        Ver_Major As Long
        Ver_Minor As Long
        Nr_Compilat As Long
        Nr_Platforma As Long
End Type

Private VersiuneOS_Acum As OSVERSIONINFO
'==========================================================================================================='
Dim tmpBuffer As String * 255

Sub GetSysInfo()
'==========================================================================================================='
    GetComputerName tmpBuffer, 255
    Instalare.Nume_PC.Caption = Trim$(tmpBuffer)
'-----------------------------------------------------------------------------------------------------------'
    GetUserName tmpBuffer, 255
    Instalare.NumeUser.Caption = tmpBuffer
'-----------------------------------------------------------------------------------------------------------'
    VersiuneOS_Acum.Dimensiune_ver_OS = 148
    GetVersionEx VersiuneOS_Acum
'-----------------------------------------------------------------------------------------------------------'
    If VersiuneOS_Acum.Nr_Platforma = VER_PLATFORM_WIN32_WINDOWS Then
        
        If VersiuneOS_Acum.Ver_Minor = 0 Then
        Instalare.Platforma.Caption = "9x"
        Else
        Instalare.Platforma.Caption = "9x"
        End If
    
    ElseIf VersiuneOS_Acum.Nr_Platforma = VER_PLATFORM_WIN32_NT Then
        
        If VersiuneOS_Acum.Ver_Major = 4 Then
        Instalare.Platforma.Caption = "NT"
        Else
        Instalare.Platforma.Caption = "NT"
        End If
    
    End If

End Sub

