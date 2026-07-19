Attribute VB_Name = "IconInTaskBar"
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPIX) As Long
   
Public Type POINTAPIX
    x As Long
    y As Long
End Type
 
  Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
   End Type
  

Global t As NOTIFYICONDATA
Global Const NIM_ADD = &H0
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200

Private m_sPattern As String
Private m_lhFind As Long

Public Sub ElibereazaMem666()
On Error Resume Next
VorlonWorld69 = GetSetting("DeltaMemoryForce", "SectiuneBeta", "Monitor")
If VorlonWorld69 = "DA" Then Exit Sub

ReDim a(200) As String
Dim j, tt As Integer
Form1.Timer1 = False

For j = 0 To Form1.BaraValoareLim1.APaulText
a(j) = Space$(513000)
OptimMemorie.Cls
DoEvents
Form1.OptimMemorie.Line ((OptimMemorie.ScaleWidth / 20) * Int(j), 0)-((OptimMemorie.ScaleWidth / 20) * Int(j) + 5, 10), 10 * j, BF
Next j
OptimMemorie.Cls
Form1.Timer1 = True
End Sub


Public Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim k As Long
Dim sName As String
sName = Space$(128)
If k > 0 Then
sName = Left$(sName, k)
If lParam = 0 Then sName = UCase(sName)
If sName Like m_sPattern Then
m_lhFind = hwnd
EnumWinProc = 0
Exit Function
End If
End If
EnumWinProc = 1
End Function


