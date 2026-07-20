VERSION 5.00
Begin VB.Form Register 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   Icon            =   "Register.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "www.ramforces.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Register to XXXXXXXXXXXXXXXX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   4320
      Picture         =   "Register.frx":038A
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Register to XXXXXXXXXXXXXXXX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1540
      Width           =   5175
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   120
      Picture         =   "Register.frx":1255
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Delta Memory Force"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Delta Memory Force"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C000&
      Height          =   1935
      Left            =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Detinaror = GetSetting("DeltaMemoryForce", "Trial", "Name")
Label3.Caption = Detinaror
Label4.Caption = Detinaror
End Sub

Private Sub Image3_Click()
Unload Me
Unload VVV
End Sub

Private Sub Label2_Click()
Unload Me
Unload VVV
End Sub

Private Sub Label3_Click()
Unload Me
Unload VVV
End Sub
