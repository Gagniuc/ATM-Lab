VERSION 5.00
Begin VB.Form VVV 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Formular construire atasament !"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5775
   Icon            =   "formular.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
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
      Top             =   2520
      Width           =   1815
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
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   3135
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
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy product ID from the box :"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "= Online ="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "formular.frx":038A
      Top             =   75
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   960
      Picture         =   "formular.frx":0714
      Top             =   540
      Width           =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00008080&
      X1              =   40
      X2              =   328
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008080&
      X1              =   8
      X2              =   376
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C0C0&
      X1              =   0
      X2              =   384
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "                    www.ramforces.com®"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your product ID:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "VVV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Atasament_Click()
ShellExecute hWnd, "open", "http://www.ramforces.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Form_Load()
rtty80 = GetSetting("DeltaMemoryForce", "Trial", "FullProgram")
If rtty80 <> "" Then
Register.Show
Me.Hide
End If
ss = GetSetting("DeltaMemoryForce", "Trial", "Key")
VVV.ID.Text = ss
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(VVV.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub

Private Sub Iesire_Click()
Unload Me
End Sub

