VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Delta Memory Force"
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8895
   Icon            =   "DeltaHelp.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4575
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar Vf 
      Height          =   4335
      Left            =   8640
      Max             =   32
      TabIndex        =   20
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4305
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   9120
         Left            =   0
         ScaleHeight     =   9090
         ScaleWidth      =   8505
         TabIndex        =   6
         Top             =   -8880
         Width           =   8535
         Begin VB.Timer Timer1 
            Interval        =   10
            Left            =   7800
            Top             =   7920
         End
         Begin VB.Timer Timer2 
            Interval        =   2000
            Left            =   7320
            Top             =   7920
         End
         Begin VB.Image Image5 
            Height          =   645
            Left            =   7320
            Picture         =   "DeltaHelp.frx":08CA
            Top             =   840
            Width           =   705
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "RAM :"
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
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label Label26 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Delta Memory Force"
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
            Left            =   2760
            TabIndex        =   32
            Top             =   120
            Width           =   2895
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00008000&
            X1              =   120
            X2              =   8280
            Y1              =   8760
            Y2              =   8760
         End
         Begin VB.Label Label18 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   " DMF can stop NetBus or Sub7 etc and kill the file. "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   8280
            Width           =   8175
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   360
            Picture         =   "DeltaHelp.frx":161F
            Top             =   120
            Width           =   240
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Up limit ^"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   2040
            TabIndex        =   22
            Top             =   5760
            Width           =   855
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Free RAM ^"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   2640
            TabIndex        =   21
            Top             =   6480
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            Caption         =   "Delta Memory Force"
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
            Left            =   2880
            TabIndex        =   19
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"DeltaHelp.frx":19A9
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   855
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   6855
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000C000&
            X1              =   120
            X2              =   8280
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label Label3 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"DeltaHelp.frx":1AAD
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   975
            Left            =   120
            TabIndex        =   17
            Top             =   2160
            Width           =   7695
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000C000&
            X1              =   120
            X2              =   8280
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"DeltaHelp.frx":1CA1
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Left            =   120
            TabIndex        =   16
            Top             =   3720
            Width           =   7455
         End
         Begin VB.Line Line3 
            BorderColor     =   &H0000C000&
            X1              =   120
            X2              =   8280
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Facilities :"
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
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "About ""Delta Memory Force"" :"
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
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   2655
         End
         Begin VB.Image Image1 
            Height          =   2265
            Left            =   120
            Picture         =   "DeltaHelp.frx":1D74
            Top             =   4680
            Width           =   4110
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Limita Superioara ^"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   2400
            TabIndex        =   13
            Top             =   5880
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Memoria Libera ^"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   1560
            TabIndex        =   12
            Top             =   6600
            Width           =   1575
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Down limit ^"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   3120
            TabIndex        =   11
            Top             =   6960
            Width           =   1095
         End
         Begin VB.Image Image2 
            Height          =   975
            Left            =   6000
            Picture         =   "DeltaHelp.frx":7CDA
            Top             =   4680
            Width           =   2220
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "CPU usage ^"
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   7080
            TabIndex        =   10
            Top             =   5760
            Width           =   975
         End
         Begin VB.Line Line4 
            BorderColor     =   &H0000C000&
            X1              =   120
            X2              =   8280
            Y1              =   7320
            Y2              =   7320
         End
         Begin VB.Label Label12 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"DeltaHelp.frx":AAA4
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   615
            Left            =   4440
            TabIndex        =   9
            Top             =   6240
            Width           =   3855
         End
         Begin VB.Label Label13 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"DeltaHelp.frx":AB2B
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   7680
            Width           =   8175
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Processes :"
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
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   7440
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3975
         Left            =   0
         ScaleHeight     =   3945
         ScaleWidth      =   8505
         TabIndex        =   1
         Top             =   240
         Width           =   8535
         Begin VB.CommandButton Command2 
            Caption         =   "Buy !"
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
            Left            =   6480
            TabIndex        =   23
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   4
            Text            =   "     admin@ramforces.com"
            Top             =   3240
            Width           =   2895
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Ok"
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
            Left            =   7200
            TabIndex        =   3
            Top             =   3600
            Width           =   975
         End
         Begin VB.CommandButton PlayHelpul 
            Caption         =   "Replay"
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
            Left            =   6120
            TabIndex        =   2
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Software && Copyright by Paul Gagniuc"
            ForeColor       =   &H00004000&
            Height          =   255
            Left            =   0
            TabIndex        =   31
            Top             =   3720
            Width           =   4575
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   7800
            Picture         =   "DeltaHelp.frx":ABE6
            Top             =   0
            Width           =   240
         End
         Begin VB.Line Line10 
            BorderColor     =   &H00008000&
            X1              =   0
            X2              =   8160
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line9 
            BorderColor     =   &H00008000&
            X1              =   0
            X2              =   8160
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00008000&
            X1              =   0
            X2              =   8160
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00008000&
            X1              =   0
            X2              =   8160
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label24 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Option :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   0
            Width           =   975
         End
         Begin VB.Label Label23 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"DeltaHelp.frx":AF70
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   495
            Left            =   0
            TabIndex        =   29
            Top             =   2520
            Width           =   8175
         End
         Begin VB.Label Label22 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "- The program can run immediately after the operating system is loaded."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   2040
            Width           =   8175
         End
         Begin VB.Label Label21 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "- Run under service mode; (the program is invisible for Ctrl-Alt-Del key combination); "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   1560
            Width           =   8175
         End
         Begin VB.Label Label20 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "- The program recovery after CPU is under ""n""% capacity;(the ""n""% value can be set by user); "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   0
            TabIndex        =   26
            Top             =   1080
            Width           =   8175
         End
         Begin VB.Label Label19 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   $"DeltaHelp.frx":B007
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   375
            Left            =   0
            TabIndex        =   25
            Top             =   480
            Width           =   8295
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            X1              =   2760
            X2              =   5880
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact by email at :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   1080
            TabIndex        =   5
            Top             =   3240
            Width           =   1815
         End
      End
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000C000&
      Height          =   4575
      Left            =   0
      Top             =   0
      Width           =   8895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
Dim paul As Boolean
Dim g As Variant
Dim tt As Variant
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
VVV.Show
End Sub

Private Sub Form_Load()
rtty80 = GetSetting("DeltaMemoryForce", "Trial", "FullProgram")
If rtty80 <> "" Then Command2.Visible = False
Timer1.Enabled = False
Vf.min = 0
tt = Picture2.Top
Vf.Max = Picture2.Top
Vf.Value = Vf.Max
paul = True
Picture2.Top = Vf.Max - Vf.Value
Picture3.Top = (Vf.Max - Vf.Value) + Picture2.ScaleHeight
Timer2.Enabled = True
Label25.Caption = "Software && Copyright© by " & Chr(80) & Chr(97) & Chr(117) & Chr(108) & Chr(32) & Chr(71) & Chr(97) & Chr(103) & Chr(110) & Chr(105) & Chr(117) & Chr(99)
End Sub

Private Sub PlayHelpul_Click()
Vf.min = 0
Vf.Max = tt
Vf.Value = Vf.Max
paul = True
End Sub

Private Sub Timer1_Timer()
If paul Then
If Vf.Value > -100 Then
paul = False
GoTo 4
End If
Vf.Value = Vf.Value + 10
End If
4:
Picture2.Top = Vf.Max - Vf.Value
Picture3.Top = (Vf.Max - Vf.Value) + Picture2.ScaleHeight
End Sub

Private Sub Timer2_Timer()
Timer1.Enabled = True
End Sub
