VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   LinkTopic       =   "Form3"
   ScaleHeight     =   2760
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "  Software && Copyright© by P.A.G."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Software && Copyright by Paul Aurelian Gagniuc"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   3360
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      BackStyle       =   0  'Transparent
      Caption         =   "Software && Copyright by Paul Aurelian Gagniuc"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " admin@ramforces.com"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "  Software && Copyright© by P.A.G."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuovus Ordo , Anno Domini 2002"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label PaulAurelianGagniuc 
      BackStyle       =   0  'Transparent
      Caption         =   "www.ramforces.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "DesprePaul2.frx":0000
      MousePointer    =   2  'Cross
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   1800
      MouseIcon       =   "DesprePaul2.frx":08CA
      MousePointer    =   2  'Cross
      Picture         =   "DesprePaul2.frx":1194
      Top             =   360
      Width           =   1320
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************

'                       Software & Copyright by Paul Gagniuc Aurelian

'*******************************************************************************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form3.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form3.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label5_Click()
Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Private Sub PaulAurelianGagniuc_Click()
ShellExecute hwnd, "open", "http://www.ramforces.com", vbNullString, vbNullString, conSwNormal
End Sub
