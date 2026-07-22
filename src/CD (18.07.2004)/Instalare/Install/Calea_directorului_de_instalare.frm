VERSION 5.00
Begin VB.Form Calea_directorului_de_instalare 
   Caption         =   "Director"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   Icon            =   "Calea_directorului_de_instalare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   6090
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Cale_director1 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   4095
   End
   Begin VB.CommandButton Iesire 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
   End
   Begin VB.DirListBox Directorul_este 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFF00&
      Height          =   2340
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin VB.DriveListBox partitie 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Da 
      Caption         =   "Ok"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Image Fundal 
      Height          =   2160
      Index           =   5
      Left            =   4320
      Picture         =   "Calea_directorului_de_instalare.frx":0CCA
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Fundal 
      Height          =   2160
      Index           =   4
      Left            =   2160
      Picture         =   "Calea_directorului_de_instalare.frx":23A9
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Fundal 
      Height          =   2160
      Index           =   3
      Left            =   0
      Picture         =   "Calea_directorului_de_instalare.frx":3A88
      Top             =   2160
      Width           =   2160
   End
   Begin VB.Image Fundal 
      Height          =   2160
      Index           =   2
      Left            =   4320
      Picture         =   "Calea_directorului_de_instalare.frx":5167
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Fundal 
      Height          =   2160
      Index           =   1
      Left            =   2160
      Picture         =   "Calea_directorului_de_instalare.frx":6846
      Top             =   0
      Width           =   2160
   End
   Begin VB.Image Fundal 
      Height          =   2160
      Index           =   0
      Left            =   0
      Picture         =   "Calea_directorului_de_instalare.frx":7F25
      Top             =   0
      Width           =   2160
   End
End
Attribute VB_Name = "Calea_directorului_de_instalare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


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

Private Sub Da_Click()
On Error Resume Next
Cale_director1.Text = Directorul_este.Path

baubau = Right(Directorul_este.Path, 1)
If baubau = "\" Then
rtty = Directorul_este.Path
Else
rtty = Directorul_este.Path & "\"
End If

Instalare.CDir.Caption = rtty
Me.Hide
End Sub

Private Sub Directorul_este_Change()
Cale_director1.Text = Directorul_este.Path
End Sub

Private Sub partitie_Change()
On Error Resume Next
Directorul_este.Path = partitie.Drive
End Sub

Private Sub Iesire_Click()
Unload Me
End Sub
