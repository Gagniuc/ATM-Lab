VERSION 5.00
Begin VB.Form VVV 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Formular construire atasament !"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   Icon            =   "Formular.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Formular.frx":038A
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Iesire 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Atasament 
      Caption         =   "Buy Delta Memory Force - online !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox ID 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   3720
      Picture         =   "Formular.frx":2ACF
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "= Online =        "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      X1              =   40
      X2              =   288
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   480
      Picture         =   "Formular.frx":399A
      Top             =   480
      Width           =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   104
      X2              =   208
      Y1              =   136
      Y2              =   136
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy  product ID"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   120
      Picture         =   "Formular.frx":3D24
      Top             =   75
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   320
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "                     www.ramforces.com®"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "  admin@ramforces.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   720
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your product ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "VVV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Atasament_Click()
ShellExecute hwnd, "open", "http://www.ramforces.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Form_Load()
ss = GetSetting("DeltaMemoryForce", "Trial", "Key")
VVV.ID.Text = ss
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(VVV.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Iesire_Click()
Unload Me
End Sub

Private Sub Label4_Click()
Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub
