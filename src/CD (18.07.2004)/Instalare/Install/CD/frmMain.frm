VERSION 5.00
Begin VB.Form Instalare 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Instalare"
   ClientHeight    =   9735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14295
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   Picture         =   "frmMain.frx":2982
   ScaleHeight     =   649
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   953
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Fundal_instalare 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   3240
      Picture         =   "frmMain.frx":56EBB
      ScaleHeight     =   5865
      ScaleWidth      =   8385
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   8415
      Begin VB.PictureBox Bara_numar_fisiere 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00004040&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         Picture         =   "frmMain.frx":F74ED
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   423
         TabIndex        =   6
         Top             =   3120
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.PictureBox Bara_copiere 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00004040&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1080
         Picture         =   "frmMain.frx":FBFB7
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   423
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.Label Eu_la_suta_Total 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   3360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image Copiere 
         Height          =   465
         Left            =   2400
         Picture         =   "frmMain.frx":100A81
         Top             =   5040
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Image Deschide_director 
         Height          =   465
         Left            =   6720
         Picture         =   "frmMain.frx":10300F
         Top             =   4560
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Image Instalare_DMF 
         Height          =   465
         Left            =   2400
         Picture         =   "frmMain.frx":10559D
         Top             =   5040
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Image Iesire2 
         Height          =   465
         Left            =   4560
         Picture         =   "frmMain.frx":107B2B
         Top             =   5040
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Titlul 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Install Delta Memory Force V2.0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1560
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Label Nume_PC 
         BackStyle       =   0  'Transparent
         Caption         =   "PC"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Platforma 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistem"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label NumeUser 
         BackStyle       =   0  'Transparent
         Caption         =   "Utilizator"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label pe 
         BackStyle       =   0  'Transparent
         Caption         =   "on"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Eu_la_suta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   2520
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label CDir 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "c:\"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   4680
         Visible         =   0   'False
         Width           =   6255
      End
   End
   Begin VB.Timer Arata_La_suta 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   9960
   End
   Begin VB.TextBox Scurtatura_Executabil 
      Height          =   285
      Left            =   8040
      TabIndex        =   3
      Text            =   "H:\Documents and Settings\Paul Gagniuc\Desktop\g2.exe"
      Top             =   9960
      Width           =   5655
   End
   Begin VB.TextBox Titlul_scurtaturii 
      Height          =   285
      Left            =   8040
      TabIndex        =   2
      Text            =   "Delta Memory Force V2.0"
      Top             =   10320
      Width           =   5055
   End
   Begin VB.TextBox Argument 
      Height          =   285
      Left            =   8040
      TabIndex        =   1
      Top             =   10680
      Width           =   1815
   End
   Begin VB.Timer Inchide_secunda 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   9960
   End
   Begin VB.Shape Bordura 
      BorderColor     =   &H00FFFFFF&
      Height          =   5925
      Left            =   3225
      Top             =   1305
      Visible         =   0   'False
      Width           =   8445
   End
   Begin VB.Image iesire0 
      Height          =   465
      Left            =   3000
      Picture         =   "frmMain.frx":10A0B9
      Top             =   10560
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image instalare0 
      Height          =   465
      Left            =   1440
      Picture         =   "frmMain.frx":10C647
      Top             =   10560
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image cale0 
      Height          =   465
      Left            =   6120
      Picture         =   "frmMain.frx":10EBD5
      Top             =   10560
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image copy0 
      Height          =   465
      Left            =   4560
      Picture         =   "frmMain.frx":111163
      Top             =   10560
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Cale_g 
      Height          =   465
      Left            =   6120
      Picture         =   "frmMain.frx":1136F1
      Top             =   9960
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Copy_g 
      Height          =   465
      Left            =   4560
      Picture         =   "frmMain.frx":115C7F
      Top             =   9960
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Cg 
      Height          =   465
      Left            =   3000
      Picture         =   "frmMain.frx":11820D
      Top             =   9960
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Image Ig 
      Height          =   465
      Left            =   1440
      Picture         =   "frmMain.frx":11A79B
      Top             =   9960
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Shape Bordura_Q 
      BorderColor     =   &H00FFFFFF&
      Height          =   9735
      Left            =   0
      Top             =   0
      Width           =   14295
   End
   Begin VB.Image Meniu 
      Height          =   375
      Index           =   4
      Left            =   12720
      Top             =   9000
      Width           =   975
   End
   Begin VB.Image Vizionare 
      Height          =   450
      Index           =   4
      Left            =   12750
      Picture         =   "frmMain.frx":11CD29
      Top             =   8925
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Dimensiune 
      BackStyle       =   0  'Transparent
      Caption         =   "0kb"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Meniu 
      Height          =   375
      Index           =   3
      Left            =   2880
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Image Meniu 
      Height          =   375
      Index           =   2
      Left            =   1920
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Image Meniu 
      Height          =   375
      Index           =   1
      Left            =   960
      Top             =   7440
      Width           =   1335
   End
   Begin VB.Image Meniu 
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image Vizionare 
      Height          =   390
      Index           =   3
      Left            =   2895
      Picture         =   "frmMain.frx":11E373
      Top             =   9045
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Image Vizionare 
      Height          =   450
      Index           =   2
      Left            =   1920
      Picture         =   "frmMain.frx":120B85
      Top             =   8250
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Image Vizionare 
      Height          =   465
      Index           =   1
      Left            =   900
      Picture         =   "frmMain.frx":1227E7
      Top             =   7470
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Vizionare 
      Height          =   465
      Index           =   0
      Left            =   270
      Picture         =   "frmMain.frx":124B85
      Top             =   6690
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Instalare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()

'muzica fundal
'mk dir **************************************************************************
'shortcut

Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
    ByVal cchBuffer As Long) As Long
Private Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" _
    (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, _
    ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, _
    ByVal fPrivate As Long, ByVal sParent As String) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long


Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOREDRAW = &H8
Const SWP_NOACTIVATE = &H10
Const SWP_FRAMECHANGED = &H20
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOCOPYBITS = &H100
Const SWP_NOOWNERZORDER = &H200

Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SecurityAttributes
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Private Const OPTION_NON_VOLATILE = &H0
Private Const FILE_ATTRIBUTE_NORMAL = &H80
 
Dim mWindowsDir As String
Dim mDeskTopPath As String
Dim mStartMenuPath As String
Dim mProgramsPath As String
Dim mDeskTopPathAbsolute As String
Dim mStartMenuPathAbsolute As String
Dim mProgramsPathAbsolute As String
Dim mTitleRef As String
Dim mKeyHandle As Long
Dim mresult

Private Const Mouse_eu = (-12)
Private Incarca_mouse_Vechi As Long

Public Paul_FSize As Variant
Public Paul_BTest As Variant


Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Public sunet_este As Variant


Private Sub Arata_La_suta_Timer()
Eu_la_suta.Caption = (100 - Int(100 * Paul_BTest / Paul_FSize)) & " %"
End Sub

Private Sub Copiere_Click()
On Error Resume Next

Dim dd0 As String
Dim Nume_copiat0 As String

dd0 = App.Path & "\KIT\DMF_KIT.zip"

Nume_copiat0 = "DMF_KIT.zip"

Eu_la_suta.Caption = "0 %"

Dimensiune.Caption = Copiere_Fisier(dd0, CDir.Caption & Nume_copiat0)


Inchide_secunda.Enabled = True
End Sub

Private Sub Copiere_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Copiere.Picture = Copy_g.Picture

If sunet_este = 6 Then
Else
sunet_este = 6
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If

End Sub

Private Sub Deschide_director_Click()
Calea_directorului_de_instalare.Show
End Sub

Private Sub Deschide_director_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Deschide_director.Picture = Cale_g.Picture

If sunet_este = 8 Then
Else
sunet_este = 8
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Mouse_nou As Long

Dim x
Dim bula As String

Paul_BTest = 1
Paul_FSize = 1

If App.PrevInstance Then End

Mouse_nou = LoadCursorFromFile(App.Path & "\CD\Aferent\Cruce.ani")
Incarca_mouse_Vechi = SetClassLong(Instalare.hwnd, Mouse_eu, Mouse_nou)

GetSysInfo
bula = "Play " & App.Path & "\CD\Sunet\S_fundal.wav"
x = mciExecute(bula)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long

Vizionare(0).Visible = False
Vizionare(1).Visible = False
Vizionare(2).Visible = False
Vizionare(3).Visible = False
Vizionare(4).Visible = False

Copiere.Picture = copy0.Picture
Instalare_DMF.Picture = instalare0.Picture
Iesire2.Picture = iesire0.Picture
Deschide_director.Picture = cale0.Picture

sunet_este = 10 'adica alceva .. ca nu conteaza ...

If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Instalare.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Calea_directorului_de_instalare
End Sub

Private Sub Fundal_instalare_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Copiere.Picture = copy0.Picture
Instalare_DMF.Picture = instalare0.Picture
Iesire2.Picture = iesire0.Picture
Deschide_director.Picture = cale0.Picture

sunet_este = 10

End Sub

Private Sub Iesire2_Click()
Fundal_instalare.Visible = False
Bordura.Visible = False

CDir.Visible = False
Deschide_director.Visible = False
Instalare_DMF.Visible = False
Iesire2.Visible = False
Dimensiune.Visible = False
Eu_la_suta.Visible = False
Eu_la_suta_Total.Visible = False
Bara_copiere.Visible = False
Bara_numar_fisiere.Visible = False

NumeUser.Visible = False
Nume_PC.Visible = False
Platforma.Visible = False
pe.Visible = False
Titlul.Visible = False

Copiere.Visible = False
End Sub

Private Sub Iesire2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Iesire2.Picture = Cg.Picture

If sunet_este = 7 Then
Else
sunet_este = 7
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If

End Sub

Private Sub Inchide_secunda_Timer()
Fundal_instalare.Visible = False
Bordura.Visible = False

CDir.Visible = False
Deschide_director.Visible = False
Instalare_DMF.Visible = False
Iesire2.Visible = False
Dimensiune.Visible = False
Eu_la_suta.Visible = False
Eu_la_suta_Total.Visible = False
Bara_copiere.Visible = False
Bara_numar_fisiere.Visible = False

NumeUser.Visible = False
Nume_PC.Visible = False
Platforma.Visible = False
pe.Visible = False
Titlul.Visible = False

Copiere.Visible = False

Eu_la_suta.Caption = "0 %"
Eu_la_suta_Total.Caption = "0 %"

Arata_La_suta.Enabled = False

Bara_copiere.Cls
Bara_numar_fisiere.Cls

Inchide_secunda.Enabled = False
End Sub

Private Sub Instalare_DMF_Click()
On Error Resume Next
Dim dd1(1 To 8) As String
Dim Nume_copiat(1 To 8) As String

MkDir CDir.Caption & "DMF"

Nume_copiat(1) = "g2.exe"
Nume_copiat(2) = "DMF_2004.DLL"
Nume_copiat(3) = "DMF.doc"
Nume_copiat(4) = "mscomctl.ocx"
Nume_copiat(5) = "VB6STKIT.DLL"
Nume_copiat(6) = "msvcrt.dll"
Nume_copiat(7) = "Rsrc32.dll"
Nume_copiat(8) = "winmm.dll"

dd1(1) = App.Path & "\Instalare\g2.exe"
dd1(2) = App.Path & "\Instalare\DMF_2004.DLL"
dd1(3) = App.Path & "\Instalare\DMF.doc"
dd1(4) = App.Path & "\Instalare\mscomctl.ocx"
dd1(5) = App.Path & "\Instalare\VB6STKIT.DLL"
dd1(6) = App.Path & "\Instalare\msvcrt.dll"
dd1(7) = App.Path & "\Instalare\Rsrc32.dll"
dd1(8) = App.Path & "\Instalare\winmm.dll"

For xy = 1 To 8

Dimensiune.Caption = Copiere_Fisier(dd1(xy), CDir.Caption & "DMF\" & Nume_copiat(xy))

xx = (Bara_numar_fisiere.ScaleWidth / 8) * xy
Bara_numar_fisiere.Line (0, 0)-(xx, 20), &HFFFF&, BF

Eu_la_suta_Total.Caption = Int(xy * (100 / 8)) & " %"
DoEvents
Next xy

Inchide_secunda.Enabled = True


Scurtatura_Executabil.Text = CDir.Caption & "DMF\g2.exe"

If Fa_linkuri = True Then
MsgBox "Installation process complete !" & vbCrLf & "         www.NovusOrdo.ro"
End If

End Sub

Private Sub Instalare_DMF_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Instalare_DMF.Picture = Ig.Picture

If sunet_este = 9 Then
Else
sunet_este = 9
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If
End Sub

Private Sub Meniu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Index = 0 Then
Vizionare(0).Visible = True

If sunet_este = 0 Then
Else
sunet_este = 0
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If

Vizionare(1).Visible = False
Vizionare(2).Visible = False
Vizionare(3).Visible = False
Vizionare(4).Visible = False
End If

If Index = 1 Then
Vizionare(1).Visible = True

If sunet_este = 1 Then
Else
sunet_este = 1
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If

Vizionare(0).Visible = False
Vizionare(2).Visible = False
Vizionare(3).Visible = False
Vizionare(4).Visible = False
End If

If Index = 2 Then
Vizionare(2).Visible = True

If sunet_este = 2 Then
Else
sunet_este = 2
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If

Vizionare(0).Visible = False
Vizionare(1).Visible = False
Vizionare(3).Visible = False
Vizionare(4).Visible = False
End If

If Index = 3 Then
Vizionare(3).Visible = True

If sunet_este = 3 Then
Else
sunet_este = 3
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If

Vizionare(0).Visible = False
Vizionare(1).Visible = False
Vizionare(2).Visible = False
Vizionare(4).Visible = False
End If

If Index = 4 Then
Vizionare(4).Visible = True

If sunet_este = 4 Then
Else
sunet_este = 4
Sunet (App.Path & "\CD\Sunet\xx.wav")
End If

Vizionare(0).Visible = False
Vizionare(1).Visible = False
Vizionare(2).Visible = False
Vizionare(3).Visible = False
End If
End Sub

Private Sub Meniu_Click(Index As Integer)
On Error Resume Next

If Index = 0 Then
Shell App.Path & "\DMF\g2.exe"
End If

If Index = 1 Then
Titlul.Caption = "Install Delta Memory Force V2.0"

Fundal_instalare.Visible = True
Bordura.Visible = True

Bara_copiere.Visible = True
Bara_numar_fisiere.Visible = True

CDir.Visible = True
Deschide_director.Visible = True

Instalare_DMF.Visible = True
Iesire2.Visible = True

Eu_la_suta.Visible = True
Eu_la_suta_Total.Visible = True



NumeUser.Visible = True
Nume_PC.Visible = True
Platforma.Visible = True
pe.Visible = True
Titlul.Visible = True
Copiere.Visible = False
End If

If Index = 2 Then
Titlul.Caption = "Copy Delta Memory Force V2.0 kit"

Fundal_instalare.Visible = True
Bordura.Visible = True

CDir.Visible = True
Deschide_director.Visible = True

NumeUser.Visible = True
Nume_PC.Visible = True
Platforma.Visible = True
pe.Visible = True

Bara_copiere.Visible = True
Eu_la_suta.Visible = True

Titlul.Visible = True

Copiere.Visible = True
Iesire2.Visible = True
Bara_numar_fisiere.Visible = False
End If


If Index = 3 Then
ShellExecute hwnd, "open", App.Path & "\CD\Aferent\DMF.doc", vbNullString, vbNullString, conSwNormal
End If

If Index = 4 Then
Unload Calea_directorului_de_instalare
Unload Me
End If
End Sub


Function Copiere_Fisier(src As String, dst As String) As Single
On Error Resume Next
 Static Buf$
 Dim BTest!, FSize!
 Dim Chunk%, F1%, F2%

 Const BUFSIZE = 1024

'Un buffer mai mare e mai bun dar nu tre' sa depasesc 64k

   If Dir(src) = "" Then MsgBox "Fisierul nu a fost gasit": Exit Function
   If Len(Dir(dst)) Then
      If MsgBox(UCase(dst) & Chr(13) & Chr(10) & "Fisierul exista ! " & vbCrLf & vbCrLf & "Suprascriem fisierul?", 4) <> 6 Then Exit Function
      Kill dst
   End If
 
 Bara_copiere.Cls
 
 Arata_La_suta.Enabled = True
 
   On Error GoTo Eruare
   F1 = FreeFile
   Open src For Binary As F1
   F2 = FreeFile
   Open dst For Binary As F2
 
   FSize = LOF(F1)
   BTest = FSize - LOF(F2)
   Do
      If BTest < BUFSIZE Then
         Chunk = BTest
      Else
         Chunk = BUFSIZE
      End If
      Buf = String(Chunk, " ")
      Get F1, , Buf
      Put F2, , Buf
      BTest = FSize - LOF(F2)

zz = (Bara_copiere.ScaleWidth / 100) * (100 - Int(100 * BTest / FSize))
Bara_copiere.Line (0, 0)-(zz, 20), &HFFFF&, BF

Paul_BTest = BTest
Paul_FSize = FSize

DoEvents
   Loop Until BTest = 0
   Close F1
   Close F2

   Copiere_Fisier = FSize
   Exit Function

Eruare:
   MsgBox "Copierea fisierului nu a putut fi facuta !"
   Close F1
   Close F2
   
   Arata_La_suta.Enabled = False
   
   Exit Function
End Function


Private Function Fa_linkuri() As Boolean
    On Error GoTo 666
    Fa_linkuri = False
    Dim mExeFileSpec As String
    Dim mDestPath As String
    Dim mArgList As String
    Dim mTitleRef As String
    Dim mPrivate As Boolean
    Dim mParent As String

    mExeFileSpec = Trim(Scurtatura_Executabil.Text)
    mArgList = Trim(Argument.Text)
    mTitleRef = Trim(Titlul_scurtaturii.Text)
    mPrivate = True
    mParent = "$(Programs)"
    

mDestPath = mDeskTopPath
OSfCreateShellLink mDestPath, mTitleRef, mExeFileSpec, mArgList, mPrivate, mParent


mDestPath = mStartMenuPath
OSfCreateShellLink mDestPath, mTitleRef, mExeFileSpec, mArgList, mPrivate, mParent


mDestPath = mProgramsPath
OSfCreateShellLink mDestPath, mTitleRef, mExeFileSpec, mArgList, mPrivate, mParent

Fa_linkuri = True
Exit Function
666:
End Function

Private Sub Form_activate()
On Error Resume Next
    If Fisierul_Exista(App.Path & "\" & App.EXEName & ".exe") Then
         Titlul_scurtaturii.Text = "Delta Memory Force V2.0"
         Scurtatura_Executabil.Text = App.Path & "\" & App.EXEName & ".exe"
    End If
    mWindowsDir = Director_WINDOWS
    mDeskTopPath = "..\..\Desktop"
    mStartMenuPath = ".."
    mProgramsPath = mStartMenuPath & "\Programs"
    mDeskTopPathAbsolute = mWindowsDir & "\Desktop"
    mStartMenuPathAbsolute = mWindowsDir & "\Start Menu"
    mProgramsPathAbsolute = mStartMenuPathAbsolute & "\Programs"
    
 End Sub



Private Function Director_WINDOWS() As String
    On Error Resume Next
    Dim mBuffer As String * 256
    Dim mDir As String
    mresult = GetWindowsDirectory(mBuffer, 256)
    If Len(mresult) = 0 Then
         MsgBox "Directorul sistemului nu a fost gasit !"
         Exit Function
    End If
    mDir = Mid$(mBuffer, 1, mresult)
    Director_WINDOWS = mDir
End Function



Private Function Scurtatura(inSpec) As String
    Dim i
    Dim ShortSpec As String
    Dim mBuffer As String
    Dim mBufLen As Long
    mBufLen = 164
    mBuffer = String(mBufLen, 0)
    i = GetShortPathName(inSpec, mBuffer, mBufLen)
    Scurtatura = Left$(mBuffer, i)
End Function

Function Fisierul_Exista(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim mFile
    Dim i
    mFile = Scurtatura(inFileSpec)
    i = FreeFile
    Open mFile For Input As i
    If Err Then
        Fisierul_Exista = False
    Else
        Close i
        Fisierul_Exista = True
    End If
End Function
