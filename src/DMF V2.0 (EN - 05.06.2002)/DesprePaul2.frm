VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   LinkTopic       =   "Form3"
   ScaleHeight     =   2910
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5160
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008000&
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Image Image2 
      Height          =   4620
      Left            =   -120
      Picture         =   "DesprePaul2.frx":0000
      Top             =   -120
      Width           =   1755
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " admin@ramforces.com"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2760
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuovus Ordo , Anno Domini 2003"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   2
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2400
      MouseIcon       =   "DesprePaul2.frx":4130
      MousePointer    =   2  'Cross
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   3000
      MouseIcon       =   "DesprePaul2.frx":49FA
      MousePointer    =   99  'Custom
      Picture         =   "DesprePaul2.frx":52C4
      ToolTipText     =   "Exit"
      Top             =   360
      Width           =   1320
   End
   Begin VB.Label Label1 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   3855
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

Private Const SRCCOPY = &HCC0020
Private Const RDW_ERASE = &H4

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Dim Moving As Boolean
Dim mX As Long
Dim mY As Long

Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form3.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
PaulAurelianGagniuc.ForeColor = &H80&
Label5.ForeColor = &H80&
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
PaulAurelianGagniuc.ForeColor = &H80&
Label5.ForeColor = &H80&
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form3.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
PaulAurelianGagniuc.ForeColor = &H80&
Label5.ForeColor = &H80&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form3.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub

Private Sub Label5_Click()
Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub


Private Sub PaulAurelianGagniuc_Click()
ShellExecute hwnd, "open", "http://www.ramforces.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label5.ForeColor = &HFF&
PaulAurelianGagniuc.ForeColor = &H80&
End Sub

Private Sub PaulAurelianGagniuc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
PaulAurelianGagniuc.ForeColor = &HFF&
Label5.ForeColor = &H80&
End Sub

Private Sub Image1_Click()
Dim CDC As Long, DTWND As Long, DTDC As Long
Dim CBMP As Long, frmWidth As Long, frmHeight As Long
Dim PrevX As Long, PrevY As Long
Dim TempDC As Long, TempBMP As Long
Dim TempRect As RECT


frmWidth = Form3.Width \ Screen.TwipsPerPixelX
frmHeight = Form3.Height \ Screen.TwipsPerPixelY


DTWND = GetDesktopWindow
DTDC = GetDC(DTWND)
CDC = CreateCompatibleDC(DTDC)
CBMP = CreateCompatibleBitmap(DTDC, frmWidth, frmHeight)
Call SelectObject(CDC, CBMP)
TempDC = CreateCompatibleDC(DTDC)
TempBMP = CreateCompatibleBitmap(DTDC, frmWidth, frmHeight)
SelectObject TempDC, TempBMP

ffff = Me.Top
uuuu = Me.Left
Pic.Top = 0
Pic.Left = 0
Pic.Width = Me.Width
Pic.Height = Me.Height

PrevX = Me.Left \ Screen.TwipsPerPixelX
PrevY = Me.Top \ Screen.TwipsPerPixelY

Me.Hide
TempRect.Left = PrevX
TempRect.Right = PrevX + frmWidth
TempRect.Top = PrevY
TempRect.Bottom = PrevY + frmHeight
RedrawWindow DTWND, TempRect, ByVal 0&, RDW_ERASE
DoEvents
BitBlt CDC, 0&, 0&, frmWidth, frmHeight, DTDC, PrevX, PrevY, SRCCOPY
Me.Show
Me.Top = ffff
Me.Left = uuuu
DoEvents
BitBlt TempDC, 0&, 0&, frmWidth, frmHeight, DTDC, PrevX, PrevY, SRCCOPY

Pic.ZOrder 0
Pic.Visible = True
BitBlt Pic.hdc, 0&, 0&, frmWidth, frmHeight, TempDC, 0&, 0&, SRCCOPY
Pic.Picture = Pic.Image

Pic.AutoRedraw = False
Dim I As Long, tX As Long, tY As Long, SqSize As Long
SqSize = Val(2)
Randomize Timer
For I = 1 To Val(100000)
 tX = Rnd(5) * frmWidth
 tY = Rnd(7) * frmHeight
 BitBlt Pic.hdc, tX, tY, SqSize, SqSize, CDC, tX, tY, SRCCOPY
Next I
DeleteDC CDC
DeleteDC TempDC
Unload Me
End Sub
