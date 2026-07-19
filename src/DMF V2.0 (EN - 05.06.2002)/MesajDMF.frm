VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3720
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Delta Memory Force 2.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   240
      Picture         =   "MesajDMF.frx":0000
      Top             =   240
      Width           =   420
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label MesajDMF 
      BackColor       =   &H00000000&
      Caption         =   "Deleted files !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
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

Private Sub Command1_Click()
Dim CDC As Long, DTWND As Long, DTDC As Long
Dim CBMP As Long, frmWidth As Long, frmHeight As Long
Dim PrevX As Long, PrevY As Long
Dim TempDC As Long, TempBMP As Long
Dim TempRect As RECT


frmWidth = Form5.Width \ Screen.TwipsPerPixelX
frmHeight = Form5.Height \ Screen.TwipsPerPixelY


DTWND = GetDesktopWindow
DTDC = GetDC(DTWND)
CDC = CreateCompatibleDC(DTDC)
CBMP = CreateCompatibleBitmap(DTDC, frmWidth, frmHeight)
Call SelectObject(CDC, CBMP)
TempDC = CreateCompatibleDC(DTDC)
TempBMP = CreateCompatibleBitmap(DTDC, frmWidth, frmHeight)
SelectObject TempDC, TempBMP

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

