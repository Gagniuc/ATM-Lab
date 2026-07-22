VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Delta Memory Force 2.0 - Copyright by Paul Aurelian Gagniuc"
   ClientHeight    =   5775
   ClientLeft      =   450
   ClientTop       =   225
   ClientWidth     =   9135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "GraficOptim4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   609
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerNT 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   6960
   End
   Begin VB.PictureBox Picture28 
      BackColor       =   &H00004000&
      Height          =   15
      Left            =   4920
      ScaleHeight     =   15
      ScaleWidth      =   4095
      TabIndex        =   114
      Top             =   480
      Width           =   4095
   End
   Begin VB.PictureBox Picture27 
      BackColor       =   &H00004000&
      Height          =   15
      Left            =   4920
      ScaleHeight     =   15
      ScaleWidth      =   4095
      TabIndex        =   113
      Top             =   2460
      Width           =   4095
   End
   Begin VB.PictureBox Picture26 
      BackColor       =   &H00008000&
      Height          =   15
      Left            =   4920
      ScaleHeight     =   15
      ScaleWidth      =   4095
      TabIndex        =   112
      Top             =   2475
      Width           =   4095
   End
   Begin VB.PictureBox Picture25 
      BackColor       =   &H00004000&
      Height          =   1995
      Left            =   8985
      ScaleHeight     =   1995
      ScaleWidth      =   15
      TabIndex        =   111
      Top             =   480
      Width           =   15
   End
   Begin VB.PictureBox Picture24 
      BackColor       =   &H00008000&
      Height          =   1995
      Left            =   9000
      ScaleHeight     =   1995
      ScaleWidth      =   15
      TabIndex        =   110
      Top             =   480
      Width           =   15
   End
   Begin VB.PictureBox Picture23 
      BackColor       =   &H00004000&
      Height          =   1995
      Left            =   4920
      ScaleHeight     =   1995
      ScaleWidth      =   15
      TabIndex        =   109
      Top             =   480
      Width           =   15
   End
   Begin VB.PictureBox Picture22 
      BackColor       =   &H00004000&
      Height          =   15
      Left            =   2400
      ScaleHeight     =   15
      ScaleWidth      =   1800
      TabIndex        =   108
      Top             =   1575
      Width           =   1800
   End
   Begin VB.PictureBox Picture21 
      BackColor       =   &H00004000&
      Height          =   15
      Left            =   2400
      ScaleHeight     =   15
      ScaleWidth      =   1800
      TabIndex        =   107
      Top             =   585
      Width           =   1800
   End
   Begin VB.PictureBox Picture20 
      BackColor       =   &H00004000&
      Height          =   975
      Left            =   2385
      ScaleHeight     =   975
      ScaleWidth      =   15
      TabIndex        =   106
      Top             =   600
      Width           =   15
   End
   Begin VB.PictureBox Picture19 
      BackColor       =   &H00008000&
      Height          =   975
      Left            =   2400
      ScaleHeight     =   975
      ScaleWidth      =   15
      TabIndex        =   105
      Top             =   600
      Width           =   15
   End
   Begin VB.PictureBox Picture18 
      BackColor       =   &H00008000&
      Height          =   15
      Left            =   2400
      ScaleHeight     =   15
      ScaleWidth      =   1800
      TabIndex        =   104
      Top             =   1560
      Width           =   1800
   End
   Begin VB.PictureBox Picture17 
      BackColor       =   &H00008000&
      Height          =   15
      Left            =   2400
      ScaleHeight     =   15
      ScaleWidth      =   1800
      TabIndex        =   103
      Top             =   600
      Width           =   1800
   End
   Begin VB.PictureBox Picture16 
      BackColor       =   &H00004000&
      Height          =   975
      Left            =   4095
      ScaleHeight     =   975
      ScaleWidth      =   15
      TabIndex        =   102
      Top             =   600
      Width           =   15
   End
   Begin VB.PictureBox Picture15 
      BackColor       =   &H00008000&
      Height          =   975
      Left            =   4080
      ScaleHeight     =   975
      ScaleWidth      =   15
      TabIndex        =   101
      Top             =   600
      Width           =   15
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00008000&
      Height          =   975
      Left            =   4200
      ScaleHeight     =   975
      ScaleWidth      =   15
      TabIndex        =   100
      Top             =   600
      Width           =   15
   End
   Begin VB.PictureBox Picture13 
      BackColor       =   &H00008000&
      Height          =   15
      Left            =   480
      ScaleHeight     =   15
      ScaleWidth      =   3675
      TabIndex        =   99
      Top             =   2265
      Width           =   3675
   End
   Begin VB.PictureBox Picture12 
      BackColor       =   &H00008000&
      Height          =   2175
      Left            =   465
      ScaleHeight     =   2175
      ScaleWidth      =   15
      TabIndex        =   98
      Top             =   2280
      Width           =   15
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00004000&
      Height          =   15
      Left            =   480
      ScaleHeight     =   15
      ScaleWidth      =   3675
      TabIndex        =   97
      Top             =   4410
      Width           =   3675
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H00008000&
      Height          =   15
      Left            =   480
      ScaleHeight     =   15
      ScaleWidth      =   3675
      TabIndex        =   96
      Top             =   4425
      Width           =   3675
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00004000&
      Height          =   2175
      Left            =   4110
      ScaleHeight     =   2175
      ScaleWidth      =   15
      TabIndex        =   95
      Top             =   2280
      Width           =   15
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00004000&
      Height          =   15
      Left            =   480
      ScaleHeight     =   15
      ScaleWidth      =   3675
      TabIndex        =   94
      Top             =   4440
      Width           =   3675
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00008000&
      Height          =   2175
      Left            =   4125
      ScaleHeight     =   2175
      ScaleWidth      =   15
      TabIndex        =   93
      Top             =   2280
      Width           =   15
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00004000&
      Height          =   2175
      Left            =   4140
      ScaleHeight     =   2175
      ScaleWidth      =   15
      TabIndex        =   92
      Top             =   2280
      Width           =   15
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00004000&
      Height          =   15
      Left            =   480
      ScaleHeight     =   15
      ScaleWidth      =   3675
      TabIndex        =   91
      Top             =   2280
      Width           =   3675
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00004000&
      Height          =   2175
      Left            =   480
      ScaleHeight     =   2175
      ScaleWidth      =   15
      TabIndex        =   90
      Top             =   2280
      Width           =   15
   End
   Begin VB.PictureBox APeste 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      Picture         =   "GraficOptim4.frx":0CCA
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   89
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox PrioApply 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      Picture         =   "GraficOptim4.frx":1D91
      ScaleHeight     =   315
      ScaleWidth      =   1755
      TabIndex        =   88
      Top             =   7080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox PriorA 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      Picture         =   "GraficOptim4.frx":2E87
      ScaleHeight     =   345
      ScaleWidth      =   1830
      TabIndex        =   87
      Top             =   7800
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.PictureBox yyyySS 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8280
      Picture         =   "GraficOptim4.frx":3CF2
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   86
      Top             =   5880
      Width           =   375
   End
   Begin VB.TextBox Edit_ProcessLabel 
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
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   5040
      MousePointer    =   1  'Arrow
      TabIndex        =   83
      Text            =   "C:\g.exe"
      Top             =   2940
      Width           =   3855
   End
   Begin VB.PictureBox Prioritati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   3  'Vertical Line
      ForeColor       =   &H0000FF00&
      Height          =   1365
      Left            =   8640
      Picture         =   "GraficOptim4.frx":40EC
      ScaleHeight     =   4
      ScaleMode       =   0  'User
      ScaleWidth      =   210
      TabIndex        =   76
      Top             =   3720
      Width           =   240
   End
   Begin VB.PictureBox PicCeasProcesor 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   1125
      Index           =   1
      Left            =   4920
      Picture         =   "GraficOptim4.frx":44C5
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   74
      Top             =   3840
      Width           =   2205
      Begin VB.PictureBox LED50 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   810
         ScaleHeight     =   30
         ScaleWidth      =   120
         TabIndex        =   116
         Top             =   150
         Width           =   150
      End
      Begin VB.PictureBox Limita75 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   210
         ScaleHeight     =   30
         ScaleWidth      =   120
         TabIndex        =   115
         Top             =   480
         Width           =   150
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   90
         Left            =   0
         ScaleHeight     =   4
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   145
         TabIndex        =   75
         Top             =   1050
         Width           =   2205
      End
   End
   Begin VB.Timer TimDel 
      Interval        =   100
      Left            =   2520
      Top             =   6600
   End
   Begin VB.PictureBox PictV 
      Height          =   255
      Left            =   2280
      Picture         =   "GraficOptim4.frx":C71B
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   73
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PictN 
      Height          =   255
      Left            =   1920
      Picture         =   "GraficOptim4.frx":CBEB
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   72
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox KillProgram 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   71
      Top             =   3600
      Width           =   375
   End
   Begin VB.CheckBox DeleteProces 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Terminate && Delete"
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
      Left            =   5040
      TabIndex        =   70
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Timer TimerForma 
      Interval        =   10
      Left            =   2160
      Top             =   6600
   End
   Begin VB.Timer TimerScrie 
      Interval        =   50
      Left            =   1800
      Top             =   6600
   End
   Begin VB.PictureBox CompresProces2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4215
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   68
      Top             =   3480
      Width           =   540
   End
   Begin VB.PictureBox TerD 
      Height          =   255
      Left            =   4800
      Picture         =   "GraficOptim4.frx":D065
      ScaleHeight     =   195
      ScaleWidth      =   1875
      TabIndex        =   67
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox RefD 
      Height          =   255
      Left            =   2880
      Picture         =   "GraficOptim4.frx":E8A7
      ScaleHeight     =   195
      ScaleWidth      =   1875
      TabIndex        =   66
      Top             =   7440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox TerminaV 
      Height          =   255
      Left            =   120
      Picture         =   "GraficOptim4.frx":10129
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   65
      Top             =   7320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox RefreshV 
      Height          =   255
      Left            =   120
      Picture         =   "GraficOptim4.frx":1196B
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   64
      Top             =   7080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox TerminaG 
      Height          =   255
      Left            =   120
      Picture         =   "GraficOptim4.frx":131AD
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   63
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox RefreshG 
      Height          =   255
      Left            =   120
      Picture         =   "GraficOptim4.frx":149EF
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   62
      Top             =   6600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox ProcessG 
      Height          =   255
      Left            =   6720
      Picture         =   "GraficOptim4.frx":1554B
      ScaleHeight     =   195
      ScaleWidth      =   2235
      TabIndex        =   61
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox AscundeD 
      Height          =   375
      Left            =   4800
      Picture         =   "GraficOptim4.frx":1612A
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   60
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox OptimD 
      Height          =   375
      Left            =   2880
      Picture         =   "GraficOptim4.frx":169F6
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   59
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox iesireD 
      Height          =   255
      Left            =   3960
      Picture         =   "GraficOptim4.frx":178FC
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   58
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox DespreD 
      Height          =   255
      Left            =   4080
      Picture         =   "GraficOptim4.frx":1853E
      ScaleHeight     =   195
      ScaleWidth      =   795
      TabIndex        =   57
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox ajutorD 
      Height          =   255
      Left            =   3000
      Picture         =   "GraficOptim4.frx":19000
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   56
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox optiuniD 
      Height          =   255
      Left            =   2640
      Picture         =   "GraficOptim4.frx":19C42
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   55
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox ascundeV 
      Height          =   375
      Left            =   2880
      Picture         =   "GraficOptim4.frx":1AE84
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   54
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox ascundeG 
      Height          =   375
      Left            =   2880
      Picture         =   "GraficOptim4.frx":1B535
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   53
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox optimV 
      Height          =   375
      Left            =   4800
      Picture         =   "GraficOptim4.frx":1BD84
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   52
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox optimG 
      Height          =   375
      Left            =   4800
      Picture         =   "GraficOptim4.frx":1C834
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   51
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox ProcesV 
      Height          =   255
      Left            =   6720
      Picture         =   "GraficOptim4.frx":1D3B5
      ScaleHeight     =   195
      ScaleWidth      =   2235
      TabIndex        =   50
      Top             =   6600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox ProcesG 
      Height          =   255
      Left            =   6720
      Picture         =   "GraficOptim4.frx":1DDA0
      ScaleHeight     =   195
      ScaleWidth      =   2235
      TabIndex        =   49
      Top             =   6360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.PictureBox optiuniV 
      Height          =   255
      Left            =   4920
      Picture         =   "GraficOptim4.frx":1E727
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   48
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox optiuniG 
      Height          =   255
      Left            =   4920
      Picture         =   "GraficOptim4.frx":1F969
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   47
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox ajutorV 
      Height          =   255
      Left            =   6360
      Picture         =   "GraficOptim4.frx":20100
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   46
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox ajutorG 
      Height          =   255
      Left            =   6360
      Picture         =   "GraficOptim4.frx":20D42
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   45
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox DespreN 
      Height          =   255
      Left            =   7320
      Picture         =   "GraficOptim4.frx":2135C
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   44
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox DespreG 
      Height          =   255
      Left            =   7320
      Picture         =   "GraficOptim4.frx":21E1E
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   43
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox iesireV 
      Height          =   255
      Left            =   8040
      Picture         =   "GraficOptim4.frx":224BB
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   42
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox iesireG 
      Height          =   255
      Left            =   8040
      Picture         =   "GraficOptim4.frx":230FD
      ScaleHeight     =   195
      ScaleWidth      =   915
      TabIndex        =   41
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox LedAprins 
      Height          =   135
      Left            =   1920
      Picture         =   "GraficOptim4.frx":23709
      ScaleHeight     =   75
      ScaleWidth      =   315
      TabIndex        =   40
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox LedStins 
      Height          =   135
      Left            =   2280
      Picture         =   "GraficOptim4.frx":2398B
      ScaleHeight     =   75
      ScaleWidth      =   315
      TabIndex        =   39
      Top             =   5880
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   162
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":23C0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":23FA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":24345
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":246E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":24A7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":24E19
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":251B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":25551
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":258ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":25C89
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":26025
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":263C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2675D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":26AF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":26E95
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":27231
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":275CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":27969
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":27D05
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":280A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2843D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":287D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":28B75
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":28F11
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":292AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":29649
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":299E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":29D81
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2A11D
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2A4B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2A855
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2ABF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2AF95
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2B339
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2B6DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2BA81
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2BE25
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2C1C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2C56D
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2C911
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2CCB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2D059
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2D3FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2D7A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2DB45
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2DEE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2E28D
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2E631
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2E9D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2ED79
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2F11D
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2F4C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2F865
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2FC09
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":2FFAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":30351
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":306F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":30A99
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":30E3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":311E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":31585
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":31929
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":31CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":32071
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":32415
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":327B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":32B5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":32F01
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":332A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":33649
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":339ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":33D89
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":34125
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":344C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3485D
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":34BF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":34F95
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":35331
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":356CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":35A69
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":35E05
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":361A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3654D
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":368F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":36C95
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":37039
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":373DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":37781
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":37B25
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":37EC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3826D
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":38611
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":389B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":38D59
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":390FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":394A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":39845
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":39BE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":39F8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3A331
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3A6D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3AA79
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3AE1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3B1C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3B565
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3B909
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3BCAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3C051
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3C3F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3C799
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3CB3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3CEE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3D285
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3D629
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3D9CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3DD71
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3E115
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3E4B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3E85D
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3EC01
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3EFA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3F349
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3F6ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3FA91
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":3FE35
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":401D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":4057D
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":40921
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":40CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":41069
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":4140D
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":417A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":41B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":41EE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":4227D
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":42619
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":429B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":42D51
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":430ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":43489
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":43825
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":43BC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":43F5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":442F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":44695
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":44A31
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":44DCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":45169
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":45505
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":458A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":45C3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":45FD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":46375
            Key             =   ""
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":46711
            Key             =   ""
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":46AAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":46E49
            Key             =   ""
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":471E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":47581
            Key             =   ""
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":4791D
            Key             =   ""
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":47CB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":48055
            Key             =   ""
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GraficOptim4.frx":483F1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimLED 
      Interval        =   300
      Left            =   840
      Top             =   6120
   End
   Begin VB.PictureBox LED 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4320
      Picture         =   "GraficOptim4.frx":4878D
      ScaleHeight     =   105
      ScaleWidth      =   345
      TabIndex        =   38
      ToolTipText     =   "LED"
      Top             =   4560
      Width           =   375
   End
   Begin DMForce.BaraValoareLim BaraValoareLim2 
      Height          =   2175
      Left            =   240
      TabIndex        =   37
      Top             =   2280
      Width           =   135
      _extentx        =   238
      _extenty        =   3836
   End
   Begin DMForce.BaraValoareLim BaraValoareLim1 
      Height          =   2175
      Left            =   120
      TabIndex        =   36
      Top             =   2280
      Width           =   135
      _extentx        =   238
      _extenty        =   4048
   End
   Begin VB.Timer UnicTimer 
      Interval        =   100
      Left            =   480
      Top             =   6120
   End
   Begin VB.PictureBox LimitaSus 
      BackColor       =   &H0000FF00&
      Height          =   15
      Left            =   480
      ScaleHeight     =   15
      ScaleWidth      =   3735
      TabIndex        =   33
      Top             =   4560
      Width           =   3735
   End
   Begin VB.Timer Alunecos 
      Interval        =   10
      Left            =   120
      Top             =   6120
   End
   Begin VB.ListBox ProcessList 
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
      ForeColor       =   &H0000C000&
      Height          =   2010
      ItemData        =   "GraficOptim4.frx":48A0F
      Left            =   4920
      List            =   "GraficOptim4.frx":48A16
      TabIndex        =   31
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox SecundaText 
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
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   "00"
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox MinutText 
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
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2280
      TabIndex        =   18
      Text            =   "00"
      Top             =   4680
      Width           =   375
   End
   Begin VB.Timer OptimizareTimp 
      Interval        =   1000
      Left            =   4920
      Top             =   960
   End
   Begin VB.PictureBox PicturaSys 
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "GraficOptim4.frx":48A27
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   5895
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicturaSys 
      Height          =   240
      Index           =   1
      Left            =   360
      Picture         =   "GraficOptim4.frx":48DB1
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   30
      Top             =   5895
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicturaSys 
      Height          =   255
      Index           =   2
      Left            =   600
      Picture         =   "GraficOptim4.frx":4913B
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   29
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicturaSys 
      Height          =   240
      Index           =   3
      Left            =   840
      Picture         =   "GraficOptim4.frx":494C5
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   28
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicturaSys 
      Height          =   240
      Index           =   4
      Left            =   1080
      Picture         =   "GraficOptim4.frx":4984F
      ScaleHeight     =   180
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox LimitaJos 
      BackColor       =   &H000000FF&
      Height          =   15
      Left            =   480
      ScaleHeight     =   15
      ScaleWidth      =   3735
      TabIndex        =   25
      Top             =   4440
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   480
      ScaleHeight     =   2145
      ScaleWidth      =   3645
      TabIndex        =   24
      Top             =   2280
      Width           =   3675
   End
   Begin VB.PictureBox CompresProces 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4200
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   20
      Top             =   2760
      Width           =   540
   End
   Begin VB.PictureBox OptimMemorie 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   480
      Picture         =   "GraficOptim4.frx":49BD9
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   19
      Top             =   4560
      Width           =   3735
   End
   Begin VB.PictureBox HardLiber 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4920
      ScaleHeight     =   7
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   14
      Top             =   5520
      Width           =   2175
   End
   Begin VB.PictureBox PicCPU 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2400
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   118
      TabIndex        =   11
      Top             =   600
      Width           =   1800
   End
   Begin VB.Timer TimerCPU 
      Interval        =   100
      Left            =   5640
      Top             =   960
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   135
      Left            =   2400
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   6000
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Left            =   5280
      Top             =   960
   End
   Begin VB.Label Nume_PC 
      Caption         =   "PC"
      Height          =   255
      Left            =   3600
      TabIndex        =   119
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Platforma 
      Caption         =   "Sistem"
      Height          =   255
      Left            =   360
      TabIndex        =   118
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label NumeUser 
      Caption         =   "Utilizator"
      Height          =   255
      Left            =   3600
      TabIndex        =   117
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H00008000&
      Height          =   315
      Left            =   4920
      Top             =   2910
      Width           =   4095
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      Height          =   315
      Left            =   4920
      Top             =   2910
      Width           =   4095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      X1              =   320
      X2              =   320
      Y1              =   26
      Y2              =   384
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   320
      X2              =   320
      Y1              =   28
      Y2              =   380
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00008000&
      Height          =   1575
      Left            =   7200
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      X1              =   472
      X2              =   472
      Y1              =   216
      Y2              =   228
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H00008000&
      Height          =   285
      Left            =   4920
      Top             =   3210
      Width           =   4095
   End
   Begin VB.Image MasterPaulGagniuc 
      Height          =   360
      Left            =   7200
      Picture         =   "GraficOptim4.frx":4B073
      Top             =   5280
      Width           =   1845
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      Height          =   1575
      Left            =   7200
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      X1              =   472
      X2              =   472
      Y1              =   216
      Y2              =   230
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      Height          =   285
      Left            =   4920
      Top             =   3210
      Width           =   4095
   End
   Begin VB.Label Status_Threads 
      BackColor       =   &H00000000&
      Caption         =   "999"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   8280
      TabIndex        =   85
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Status_Priority 
      BackColor       =   &H00000000&
      Caption         =   "priotity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6120
      TabIndex        =   84
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Priority :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   5280
      TabIndex        =   82
      Top             =   3240
      Width           =   705
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Threads :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   7440
      TabIndex        =   81
      Top             =   3240
      Width           =   795
   End
   Begin VB.Image Image7 
      Height          =   210
      Left            =   8280
      Picture         =   "GraficOptim4.frx":4BEDE
      Top             =   4800
      Width           =   330
   End
   Begin VB.Image Image6 
      Height          =   210
      Left            =   8280
      Picture         =   "GraficOptim4.frx":4C2E4
      Top             =   4080
      Width           =   330
   End
   Begin VB.Image Image5 
      Height          =   210
      Left            =   8280
      Picture         =   "GraficOptim4.frx":4C6EA
      Top             =   4440
      Width           =   330
   End
   Begin VB.Image Image4 
      Height          =   210
      Left            =   8280
      Picture         =   "GraficOptim4.frx":4CAF0
      Top             =   3720
      Width           =   330
   End
   Begin VB.Label Edit_Normal 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
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
      Left            =   7320
      TabIndex        =   80
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Edit_Idle 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Idle"
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
      Left            =   7320
      TabIndex        =   79
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Edit_High 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "High"
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
      Left            =   7320
      TabIndex        =   78
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Edit_RealTime 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Real Time"
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
      Left            =   7320
      TabIndex        =   77
      Top             =   3720
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   4200
      Picture         =   "GraficOptim4.frx":4CEF6
      Top             =   1680
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   120
      Picture         =   "GraficOptim4.frx":4D843
      Top             =   75
      Width           =   240
   End
   Begin VB.Label ScrieDelta 
      Caption         =   "Delta Memory Force 2.0"
      Height          =   255
      Left            =   120
      TabIndex        =   69
      Top             =   7560
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image TerminaProces 
      Height          =   240
      Left            =   7080
      Picture         =   "GraficOptim4.frx":4DBCD
      Top             =   2580
      Width           =   1920
   End
   Begin VB.Image RefreshQ 
      Height          =   240
      Left            =   4920
      Picture         =   "GraficOptim4.frx":4DDF9
      Top             =   2580
      Width           =   1935
   End
   Begin VB.Image Procesee 
      Height          =   240
      Left            =   2400
      Picture         =   "GraficOptim4.frx":4E026
      Top             =   5040
      Width           =   2280
   End
   Begin VB.Image Iesire 
      Height          =   240
      Left            =   3720
      Picture         =   "GraficOptim4.frx":4E265
      Top             =   5400
      Width           =   960
   End
   Begin VB.Image Despre 
      Height          =   240
      Left            =   2760
      Picture         =   "GraficOptim4.frx":4E485
      Top             =   5400
      Width           =   840
   End
   Begin VB.Image Ajutor 
      Height          =   240
      Left            =   1680
      Picture         =   "GraficOptim4.frx":4E6A3
      Top             =   5400
      Width           =   960
   End
   Begin VB.Image Optiuni 
      Height          =   240
      Left            =   120
      Picture         =   "GraficOptim4.frx":4E8C3
      Top             =   5400
      Width           =   1440
   End
   Begin VB.Image Ascunde 
      Height          =   360
      Left            =   240
      Picture         =   "GraficOptim4.frx":4EAE9
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Image ElibereazaMem 
      Height          =   360
      Left            =   240
      Picture         =   "GraficOptim4.frx":4ED1F
      Top             =   600
      Width           =   1920
   End
   Begin VB.Label ValBaraJos 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   4500
      Width           =   375
   End
   Begin VB.Label ValBaraSus 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label NrProceseText 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Process:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   5385
      Left            =   4800
      Top             =   390
      Width           =   4335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   " :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   4680
      Width           =   135
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Optimize -  "
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00004000&
      BorderWidth     =   3
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label DeltaMemoryMov 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "              Delta Memory Force  2.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   0
      TabIndex        =   23
      Top             =   15
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   5385
      Left            =   0
      Top             =   390
      Width           =   4815
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total C:\"
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
      Left            =   4920
      TabIndex        =   17
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Free C:\"
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
      Left            =   4920
      TabIndex        =   16
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label DiskTot 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "439.567 Mb"
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
      Left            =   5760
      TabIndex        =   15
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label CPUDans 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "  -100%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4125
      TabIndex        =   13
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "CPU"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label DiskLiber 
      BackStyle       =   0  'Transparent
      Caption         =   "100.456 Mb"
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
      Left            =   5640
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "   Lim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4170
      TabIndex        =   8
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "TOT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0 Mb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Mem :"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Free  Mem :"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1275
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1275
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private lngDC As Long
Private blnLoop As Boolean

Private Declare Function GetSystemDirectory Lib "KERNEL32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "KERNEL32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function RegisterServiceProcess Lib "KERNEL32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetCurrentProcessId Lib "KERNEL32" () As Long
Private Declare Function GetCurrentProcess Lib "KERNEL32" () As Long

Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function QueryPerformanceCounter Lib "KERNEL32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "KERNEL32" (lpFrequency As LARGE_INTEGER) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Sub GlobalMemoryStatus Lib "KERNEL32" (lpBuffer As MEMORYSTATUS)
Private Declare Function pBGetFreeSystemResources Lib "rsrc32.dll" Alias "_MyGetFreeSystemResources32@4" (ByVal iResType As Integer) As Integer

Private Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetDiskFreeSpace Lib "KERNEL32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "KERNEL32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()


'*****************************************************************************
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'*****************************************************************************

Private Type infostructre
exe As String
threads As Long
priority As Long
processId As Long
End Type

Private Type PreferencesStruc
Language As String
SetOwnPriority As String
LoadWithWindows As String
Log As String
LogFile As String
End Type

Private Type CounterInfo 'Windows NT
    hCounter As Long
    strName As String
End Type

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
Private Type LARGE_INTEGER
lowpart As Long
highpart As Long
End Type

Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion      As Long
dwMinorVersion      As Long
dwBuildNumber       As Long
dwPlatformId        As Long
szCSDVersion        As String * 128
End Type

Private Type MEMORYSTATUS
dwLength        As Long
dwMemoryLoad    As Long
dwTotalPhys     As Long
dwAvailPhys     As Long
dwTotalPageFile As Long
dwAvailPageFile As Long
dwTotalVirtual  As Long
dwAvailVirtual  As Long
End Type

'*****************************************************************************
Const LB_SETHORIZONTALEXTENT = &H194
'*****************************************************************************

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP

Public WinCx, Spatial$, SpatialUnu$, CaleWin&, CaleSystem&, SysCx  As String

Private SetOwnPriority As Boolean
Private LogEvents As Boolean
Private LogFile As String

Private CurEdit As Long
Private Edit_Current As Integer

Private info(1 To 500) As infostructre, totalinf As Integer, Myname As String

Private CurrentProcess As infostructre
Private Preferences As PreferencesStruc
Private HotKeyHandle As Long
Private Pref_SetTo As Long

Private Const SideBar_Main = 0
Private Const SideBar_Edit = 1
Private Const SideBar_Prefereces = 2
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const VK_END = &H23

'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP


Private Const REG_DWORD = 4
Private Const HKEY_DYN_DATA = &H80000006

Private Const SR = 0
Private Const GDI = 1
Private Const USR = 2
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Const RSP_SIMPLE_SERVICE = 1
Const RSP_UNREGISTER_SERVICE = 0
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Private Const CPUWidth = 110 '2
Private Const CPUHeight = 65 '2
Private Const CPUSchelet = 10 '2
Private Const Jumate = CPUHeight \ 2 '2
Private Const pWidth = 240 '1
Private Const pHeight = 140 '1
Private Const pGrid = 20 '1
Private Const tInterval = 100 '1

Private Const SRCCOPY = &HCC0020
Private Const PS_SOLID = &H8000&

Private CPU_PE_NT As New CPU_la_suta


'**********************************************************
Dim sl As Integer
Dim R1 As Double
Dim r2 As Double
Const vbPI = 3.141592654
Const Deg2Rad = vbPI / 180
Dim cx As Integer, cy As Integer, r As Single
Dim s As Double
Dim CPUceasRau As Double
'**********************************************************

Dim CPUcont As Long '2
Dim CPUY As Long '2
Dim CPUdc As Long, LinB As Long, LinC1 As Long '2
Dim counter As Long '1
Dim oldY As Long '1
Dim hDCh As Long, hPenB As Long, hPenC As Long '1

Dim CPU As Integer '2
Dim I As Integer '1
Dim PaulTOT As Variant '1
Dim Vechi As Variant '1
Dim YY, rrr, rrrj, FXCounter As Integer
Dim It As Integer
Dim min, sec, TimpMin As Variant
Dim Amorsat, jos, josA, josD, josI, josO, josPrior, josAA, josPP, josRR, josTT As Boolean
Dim FormaMare, paulForm, ggpp, EraInformationala As Boolean
Dim PaulOptim As Variant
Dim RauRau As Variant
Dim VariabilaMem As Variant
Dim nn, PaulTransmisie, CPULaSuta As Integer
Dim DeleteJ As String


Private Sub Ajutor_Click()
Form4.Show
End Sub

Private Sub Ajutor_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josA = True
Ajutor.Picture = ajutorG.Picture
End Sub

Private Sub Ajutor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josA Then
Exit Sub
Else
Ajutor.Picture = ajutorD.Picture
End If
Iesire.Picture = iesireV.Picture
Despre.Picture = DespreN.Picture
Optiuni.Picture = optiuniV.Picture
Procesee.Picture = ProcesV.Picture
End Sub

Private Sub Ajutor_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josA = False
Ajutor.Picture = ajutorV.Picture
End Sub

Private Sub Alunecos_Timer()
rrrj = rrrj + 5
DeltaMemoryMov.Left = rrrj
If rrrj >= Int(305 / 2) Then
Alunecos.Enabled = False
rrrj = 0
End If
End Sub





'*********************************************
Private Sub DeleteProces_Click()
If DeleteProces.Value = 1 Then
KillProgram.Picture = PictV.Picture
Else
KillProgram.Picture = PictN.Picture
End If
End Sub

Private Sub DeltaMemoryMov_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form1.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub

Private Sub Despre_Click()
Form3.Show
End Sub

Private Sub Despre_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josD = True
Despre.Picture = DespreG.Picture
End Sub

Private Sub Despre_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josD Then
Exit Sub
Else
Despre.Picture = DespreD.Picture
End If
Iesire.Picture = iesireV.Picture
Ajutor.Picture = ajutorV.Picture
Optiuni.Picture = optiuniV.Picture
Procesee.Picture = ProcesV.Picture
End Sub

Private Sub Despre_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josD = False
Despre.Picture = DespreN.Picture
End Sub

Private Sub Edit_ProcessLabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
TerminaProces.Picture = TerminaV.Picture
RefreshQ.Picture = RefreshV.Picture
End Sub

Private Sub ElibereazaMem_Click()
On Error Resume Next
VorlonWorld69 = GetSetting("DeltaMemoryForce", "SectiuneBeta", "Monitor")
If VorlonWorld69 = "DA" Then Exit Sub

ReDim a(200) As String
Dim j, tt As Integer
Image3.Visible = True
Timer1 = False
TimLED.Enabled = True
For j = 0 To BaraValoareLim1.APaulText
a(j) = Space$(513000)
OptimMemorie.Cls
DoEvents
OptimMemorie.Line ((OptimMemorie.ScaleWidth / BaraValoareLim1.APaulText) * Int(j), 0)-((OptimMemorie.ScaleWidth / BaraValoareLim1.APaulText) * Int(j) + 5, 10), 10 * j, BF
Next j
OptimMemorie.Cls
TimLED.Enabled = False
Timer1 = True
Dim keys(255) As Byte
keys(&H91) = 2
tmp = SetKeyboardState(keys(0))
LED.Picture = LedStins.Picture
Image3.Visible = False
End Sub

Private Sub ElibereazaMem_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josO = True
ElibereazaMem.Picture = optimG.Picture
End Sub

Private Sub ElibereazaMem_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josO Then
Exit Sub
Else
ElibereazaMem.Picture = OptimD.Picture
End If
Ascunde.Picture = ascundeV.Picture
End Sub

Private Sub ElibereazaMem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josO = False
ElibereazaMem.Picture = optimV.Picture
End Sub

Private Sub Form_Load()
Dim MS As MEMORYSTATUS
Dim PaulTOT, PaulLiber As Variant
Dim hBmp As Long
Dim BMPCPU As Long
EraInformationala = True
Dim I As Integer
On Error Resume Next
TimDel.Enabled = False

ggpp = True
Amorsat = True
CPULaSuta = 50
nn = 0
sec = 0

On Error Resume Next

Platforma_este

Spatial$ = Space(260)
CaleWin& = GetWindowsDirectory(Spatial$, 260)
WinCx = Left(Spatial$, CaleWin&)

SpatialUnu$ = Space(260)
CaleSystem& = GetSystemDirectory(SpatialUnu$, 260)
SysCx = Left(SpatialUnu$, CaleSystem&)


Strategicyy = GetSetting("DeltaMemoryForce", "SectiuneBeta", "StrategicY")
Strategicxx = GetSetting("DeltaMemoryForce", "SectiuneBeta", "StrategicX")

If Strategicxx <> "" Then
Form1.Left = Strategicxx
End If

If Strategicyy <> "" Then
Form1.Top = Strategicyy
End If

DeltaRaspuns = GetSetting("DeltaMemoryForce", "SectiuneBeta", "OptimizareInterval")
CPUAcumLim = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "CPULim")
GCPUVer = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "CPUVer")
DMFCamuflat = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "Service")

If DMFCamuflat = "1" Then
service
Else
UnService
End If

If GCPUVer = "1" Then
CPULaSuta = CPUAcumLim
Else
CPULaSuta = 50
End If

If DeltaRaspuns = "Da" Then
TimpMin = GetSetting("DeltaMemoryForce", "SectiuneBeta", "Minute")
min = TimpMin
Amorsat = False
OptimizareTimp.Enabled = True
Label12.Visible = True
MinutText.Visible = True
Label13.Visible = True
SecundaText.Visible = True
Else
mim = 0
Label12.Visible = False
MinutText.Visible = False
Label13.Visible = False
SecundaText.Visible = False
OptimizareTimp.Enabled = False
End If
 Alunecos.Enabled = False
    
Picture1.ScaleMode = 3
Form1.ScaleMode = 3
Picture1.Height = 145
Picture1.Width = 245
counter = 0


CompresProces.Left = Picture1.Left + Picture1.Width + 3
CompresProces2.Left = Picture1.Left + Picture1.Width + 3


LimitaSus.Width = Picture1.Width - 2
LimitaSus.Left = Picture1.Left + 1

LimitaJos.Width = Picture1.Width - 2
LimitaJos.Left = Picture1.Left + 1

PicCPU.ScaleMode = 3
Form1.ScaleMode = 3
PicCPU.Height = 65
PicCPU.Width = 115

CPUcont = 0
CompresProces.Top = Label8.Top + 12

FormaMare = True

LimitaSus.Visible = False
LimitaJos.Visible = False

Timer2.Enabled = True
UnicTimer.Enabled = False

RefreshProcessList

Form1.Width = 4830
Shape3.Width = 321
Shape2.Width = 321
DeltaMemoryMov.Left = 0

LimitaInfLoad = GetSetting("DeltaMemoryForce", "SectiuneBeta", "LimitaInferioara")
If LimitaInfLoad = "" Then
BaraValoareLim2.AAAStart = 1920
GoTo 2
End If
BaraValoareLim2.AAAStart = LimitaInfLoad
2:

LimitaSupLoad = GetSetting("DeltaMemoryForce", "SectiuneBeta", "LimitaSuperioara")
If LimitaSupLoad = "" Then
BaraValoareLim1.AAAStart = 996
GoTo 3
End If
BaraValoareLim1.AAAStart = LimitaSupLoad
3:

ElibereazaMem.Picture = optimV.Picture
Ascunde.Picture = ascundeV.Picture

Optiuni.Picture = optiuniV.Picture
Ajutor.Picture = ajutorV.Picture
Despre.Picture = DespreN.Picture
Iesire.Picture = iesireV.Picture

Procesee.Picture = ProcesV.Picture

RefreshQ.Picture = RefreshV.Picture
TerminaProces.Picture = TerminaV.Picture


If LimitaInfLoad < LimitaSupLoad Then
BaraValoareLim2.AAAStart = 1920
BaraValoareLim1.AAAStart = 996
End If
DeltaMemoryMov.Caption = ""
KillProgram.Picture = PictN.Picture
Me.Show
CPUdc = CreateCompatibleDC(PicCPU.hdc)
BMPCPU = CreateCompatibleBitmap(PicCPU.hdc, CPUWidth, CPUHeight)
Call SelectObject(CPUdc, BMPCPU)
LinB = CreatePen(PS_SOLID, 0, &H8000&)
LinC1 = CreatePen(PS_SOLID, 0, &HFFFF&)
Call SelectObject(CPUdc, LinB)

hDCh = CreateCompatibleDC(Picture1.hdc)
hBmp = CreateCompatibleBitmap(Picture1.hdc, pWidth, pHeight)
Call SelectObject(hDCh, hBmp)
hPenB = CreatePen(PS_SOLID, 0, &H8000&)
hPenC = CreatePen(PS_SOLID, 0, &HFFFF&)
Call SelectObject(hDCh, hPenB)

For I = pGrid To pHeight - 1 Step pGrid '1
Picture1.Line (0, I)-(pWidth, I), &H8000&
Next

For I = pGrid - (counter Mod pGrid) To pWidth - 1 Step pGrid '1
Picture1.Line (I, 0)-(I, pHeight), &H8000&
Next

For CPU = CPUSchelet To CPUHeight - 1 Step CPUSchelet '2
PicCPU.Line (0, CPU)-(CPUWidth, CPU), &H8000&
Next

For CPU = CPUSchelet - (CPUcont Mod CPUSchelet) To CPUWidth - 1 Step CPUSchelet
PicCPU.Line (CPU, 0)-(CPU, CPUHeight), &H8000&    '2
Next

MS.dwLength = Len(MS) '1
Call GlobalMemoryStatus(MS)
PaulTOT = Format$(MS.dwTotalPhys / 1024, "#,###")
P = InStr(PaulTOT, ",") - 1
PaulOptim = Mid(PaulTOT, 1, P)
Label8.Caption = PaulOptim & " Mb"


BaraValoareLim2.AAAwww = PaulOptim

BaraValoareLim1.AAAwww = PaulOptim


Call BitBlt(hDCh, 0, 0, pWidth, pHeight, Picture1.hdc, 0, 0, SRCCOPY) '1

Call BitBlt(CPUdc, 0, 0, CPUWidth, CPUHeight, PicCPU.hdc, 0, 0, SRCCOPY)
TimerCPU.Interval = 10 '2
TimerCPU.Enabled = True
CPUY = Jumate

PaulLiber = Format$(MS.dwAvailPhys / 1024, "#,###")
rr = InStr(PaulLiber, ",") - 1
rr2 = InStr(PaulTOT, ",") - 1
ert = Mid(PaulTOT, 1, rr2)
er = Mid(PaulLiber, 1, rr)
I = 150 - ((150 / ert) * er)
oldY = I

CompresProces2.Height = Picture1.Height - I - 16
If CompresProces2.Height < 5 Then CompresProces2.Visible = False

Timer1.Interval = 10
Timer1.Enabled = True
GetBytes

'--------- aici era NT

 It = 0
t.cbSize = Len(t) 'Iconita jos
t.hwnd = PicturaSys(0).hwnd
t.uID = 1&
t.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
t.uCallbackMessage = WM_MOUSEMOVE
t.hIcon = PicturaSys(It).Picture
t.szTip = "DeltaMemoryForce 2.0" & Chr$(0)
Shell_NotifyIcon NIM_ADD, t
TimerScrie.Enabled = True
DeltaOptim = GetSetting("DeltaMemoryForce", "SectiuneDelta", "Optimizare")
DeltaStartAdevar = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "Start")
If DeltaStartAdevar = "Da" Then
If DeltaOptim = "Da" Then
UnicTimer.Enabled = True
End If
End If
TimLED.Interval = 300
TimLED.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "LimitaInferioara", BaraValoareLim2.AStop
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "LimitaSuperioara", BaraValoareLim1.AStop
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "StrategicY", Form1.Top
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "StrategicX", Form1.Left

'De la procese !
Call UnregisterHotKey(Me.hwnd, HotKeyHandle)

End Sub

Private Sub Iesire_Click()
  Dim intX As Integer, intY As Integer
    Dim intI As Integer, intJ As Integer
    Dim intWidth As Integer, intHeight As Integer
    
    Timer3.Enabled = False
    Timer1.Enabled = False
    
    intWidth = Screen.Width / Screen.TwipsPerPixelX
    intHeight = Screen.Height / Screen.TwipsPerPixelY
    

    Form1.Visible = vbTrue
    Form1.AutoRedraw = vbFalse '
    
    Randomize
    
    blnLoop = vbTrue
    Do While blnLoop = vbTrue
        intX = (intWidth - 128) * Rnd
        intY = (intHeight - 128) * Rnd
        
        intI = 2 * Rnd - 1
        intJ = 2 * Rnd - 1
        
        
        Call BitBlt(Form1.hdc, intX + intI, intY + intJ, 128, 128, Form1.hdc, intX, intY, vbSrcCopy)
        
        DoEvents
    ttrryy = ttrryy + 1
    If ttrryy = 20000 Then GoTo 232
    Loop
232:
    Set frmMain = Nothing
    Call ReleaseDC(0, lngDC)
    
Shell_NotifyIcon NIM_DELETE, t
Unload Me
Unload Form2
Unload Form3
Unload Form4
Unload OptiuniDelta
Unload IconitaSistem
Unload VVV
End Sub

Private Sub Iesire_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josI = True
Iesire.Picture = iesireG.Picture
End Sub

Private Sub Iesire_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josI Then
Exit Sub
Else
Iesire.Picture = iesireD.Picture
End If
Procesee.Picture = ProcesV.Picture
Despre.Picture = DespreN.Picture
Ajutor.Picture = ajutorV.Picture
Optiuni.Picture = optiuniV.Picture
End Sub

Private Sub Iesire_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josI = False
Iesire.Picture = iesireV.Picture
End Sub

Private Sub Image2_Click()
If EraInformationala = True Then
EraInformationala = False
Form1.Height = 405
Me.Move 1, -10
Else
Form1.Height = 5775
EraInformationala = True
End If
End Sub

Private Sub KillProgram_Click()
If DeleteProces.Value = 1 Then
KillProgram.Picture = PictN.Picture
DeleteProces.Value = 0
Else
KillProgram.Picture = PictV.Picture
DeleteProces.Value = 1
End If
End Sub

Private Sub MasterPaulGagniuc_Click()
SetProcessPriority CurrentProcess.processId, Pref_SetTo
RefreshProcessList
End Sub

Private Sub MasterPaulGagniuc_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josPrior = True
MasterPaulGagniuc.Picture = PrioApply.Picture
End Sub

Private Sub MasterPaulGagniuc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josPrior Then
Exit Sub
Else
MasterPaulGagniuc.Picture = APeste.Picture
End If
End Sub

Private Sub MasterPaulGagniuc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josPrior = False
MasterPaulGagniuc.Picture = APeste.Picture
End Sub

Private Sub OptimizareTimp_Timer()
aici = GetSetting("DeltaMemoryForce", "SectiuneBeta", "Amorsare")
If aici = "Amors" Then Amorsat = True

If Amorsat Then
DeltaRaspuns = GetSetting("DeltaMemoryForce", "SectiuneBeta", "OptimizareInterval")
If DeltaRaspuns = "Da" Then
TimpMin = GetSetting("DeltaMemoryForce", "SectiuneBeta", "Minute")
min = TimpMin
MinutText.Text = min
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Amorsare", " "
Label12.Visible = True
MinutText.Visible = True
Label13.Visible = True
SecundaText.Visible = True
OptimizareTimp.Enabled = True
sec = 0
End If
Amorsat = False
End If
If sec = 0 Then
sec = 59
If min <= 0 Then
OptimizareTimp.Enabled = False
ElibereazaMem_Click
min = TimpMin
MinutText.Text = min
sec = 0
OptimizareTimp.Enabled = True
Else
min = min - 1
MinutText.Text = min
End If
Else
sec = sec - 1
SecundaText.Text = sec
End If
End Sub

Private Sub Optiuni_Click()
OptiuniDelta.Show
OptiuniDelta.Gagauta.Text = "Da"
OptiuniDelta.Timer3.Enabled = False
OptiuniDelta.Timer4.Enabled = False
OptiuniDelta.Timer3.Enabled = True
OptiuniDelta.Timer4.Enabled = True
OptiuniDelta.Label2.Caption = ""
'i = 46
End Sub

Private Sub Optiuni_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
jos = True
Optiuni.Picture = optiuniG.Picture
End Sub

Private Sub Optiuni_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If jos Then
Exit Sub
Else
Optiuni.Picture = optiuniD.Picture
End If
Iesire.Picture = iesireV.Picture
Despre.Picture = DespreN.Picture
Ajutor.Picture = ajutorV.Picture
Procesee.Picture = ProcesV.Picture
End Sub

Private Sub Optiuni_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
jos = False
Optiuni.Picture = optiuniV.Picture
End Sub

Private Sub Procesee_Click()

If Platforma.Caption = "NT" Then

'MsgBox "Delta Memory Force shows processes only for 9x and Me systems" & vbCrLf & "Press Alt-Ctrl-Del and go to Task Manager."
Form5.MesajDMF.Caption = "Delta Memory Force shows processes only for 9x and Me systems"
Form5.MesajDMF2.Caption = "Press Alt-Ctrl-Del and go to Task Manager."
Form5.Show
Else

If FormaMare Then
RefreshProcessList

Form1.Width = 9150
Shape3.Width = 609
Shape2.Width = 609
FormaMare = False
rrrj = 0
Alunecos.Enabled = True
Else
Form1.Width = 4830
Shape3.Width = 321
Shape2.Width = 321
DeltaMemoryMov.Left = 0
FormaMare = True
Alunecos.Enabled = False
rrrj = 0
End If
End If

End Sub

Private Sub Procesee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josPP = True
Procesee.Picture = ProcesG.Picture
End Sub

Private Sub Procesee_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josPP Then
Exit Sub
Else
Procesee.Picture = ProcessG.Picture
End If
Iesire.Picture = iesireV.Picture
Despre.Picture = DespreN.Picture
Ajutor.Picture = ajutorV.Picture
Optiuni.Picture = optiuniV.Picture
End Sub

Private Sub Procesee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josPP = False
Procesee.Picture = ProcesV.Picture
End Sub

Private Sub ProcessList_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TerminaProces_Click
End Sub

Private Sub ProcessList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
TerminaProces.Picture = TerminaV.Picture
RefreshQ.Picture = RefreshV.Picture
End Sub

Private Sub RefreshQ_Click()
Form1.ScaleMode = 1
AddHorizontalScrollbar ProcessList
RefreshProcessList
Form1.ScaleMode = 3
GetBytes
End Sub

Private Sub RefreshQ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josRR = True
RefreshQ.Picture = RefD.Picture
End Sub

Private Sub RefreshQ_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josRR Then
Exit Sub
Else
RefreshQ.Picture = RefreshG.Picture
End If
TerminaProces.Picture = TerminaV.Picture
End Sub

Private Sub RefreshQ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josRR = False
RefreshQ.Picture = RefreshV.Picture
End Sub

Private Sub TerminaProces_Click()
TB_KillTask
RefreshProcessList
End Sub

Private Sub TerminaProces_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josTT = True
TerminaProces.Picture = TerD.Picture
End Sub

Private Sub TerminaProces_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josTT Then
Exit Sub
Else
TerminaProces.Picture = TerminaG.Picture
End If
RefreshQ.Picture = RefreshV.Picture
End Sub

Private Sub TerminaProces_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josTT = False
TerminaProces.Picture = TerminaV.Picture
End Sub

Private Sub TimDel_Timer()
On Error Resume Next
If DeleteProces.Value = 1 Then


'UUUUUUUUUUUUUUUUUUUUUUU

z_Recheck:


If FileExist(DeleteJ) Then
If ValidFileAttributes(DeleteJ) Then
If Not DestroyFile(DeleteJ) Then
Form5.MesajDMF.Caption = "  = Delta Memory Force - couldn't destroy the file = "
Form5.Show
Exit Sub
Else

If RenameAndKill(DeleteJ) Then
Form5.MesajDMF.Caption = "  = The file " & "of this proces is deleted = "
Form5.Show
Else
Form5.MesajDMF.Caption = "  = Coulnd't rename and delete the file = "
Form5.Show
End If
End If
Else
NormalAttributes DeleteJ
GoTo z_Recheck
End If
Else
End If
'UUUUUUUUUUUUUUUUUUUUUUU
End If
TimDel.Enabled = False
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim MS As MEMORYSTATUS
Dim PaulTOT, PaulLiber As Variant

MS.dwLength = Len(MS)
Call GlobalMemoryStatus(MS)
PaulTOT = Format$(MS.dwTotalPhys / 1024, "#,###")
PaulLiber = Format$(MS.dwAvailPhys / 1024, "#,###")
rr = InStr(PaulLiber, ",") - 1
rr2 = InStr(PaulTOT, ",") - 1
ert = Mid(PaulTOT, 1, rr2)
er = Mid(PaulLiber, 1, rr)
I = Picture1.Height - ((Picture1.Height / ert) * er)

If InStr(PaulLiber, ",") > 0 Then
WWRRG = " Mb"
Else
WWRRG = " b"
End If
Label1.Caption = PaulLiber & WWRRG
Label2.Caption = PaulTOT & " Mb"
If er <> Vechi Then
t.szTip = "Free RAM - " & er & " Mb" & Chr$(0)
Label9.Visible = False
Label9.Top = Picture1.Top + (I - 5)
Label9.Caption = "- " & er & " Mb"
CompresProces2.Top = Label9.Top + 12
CompresProces2.Height = Picture1.Height - I - 16
Label9.Visible = True
PaulTransmisie = er
It = PaulTransmisie
If It = 0 Then It = 1
If It > 160 Then It = 161
t.hIcon = ImageList1.ListImages(It + 1).Picture
Shell_NotifyIcon NIM_MODIFY, t

If Label9.Top > (Label7.Top - 12) Then
CompresProces2.Visible = False
Label9.Top = Label7.Top - 12
Label9.Caption = "/  " & er & " Mb"
ggpp = False
Else
ggpp = True
End If
P = InStr(PaulTOT, ",") - 1
PaulOptim = Mid(PaulTOT, 1, P)
CompresProces.Cls
CompresProces2.Cls
For AI = 1 To I Step 4
pu = (20 * AI)
CompresProces.Line (0, AI)-(35, AI + 1), &HFF&, BF

If ggpp Then
CompresProces2.Visible = True
CompresProces2.Line (0, AI)-(35, AI + 1), &H4000&, BF
End If


CompresProces.Height = Label9.Top - Label8.Top - 12
Next AI
GetBytes
RefreshProcessList
End If


Vechi = er

Call BitBlt(hDCh, 0, 0, pWidth - 1, pHeight, hDCh, 1, 0, SRCCOPY)

If counter Mod pGrid = 0 Then
Call MoveToEx(hDCh, pWidth - 2, 0, 0)
Call LineTo(hDCh, pWidth - 2, pHeight)
End If


Call SelectObject(hDCh, hPenC)
Call MoveToEx(hDCh, pWidth - 3, oldY, 0)
Call LineTo(hDCh, pWidth - 2, I)
Call SelectObject(hDCh, hPenB)
Call BitBlt(Picture1.hdc, 0, 0, pWidth, pHeight, hDCh, 0, 0, SRCCOPY)
counter = counter + 1
oldY = I

LimitaJos.Top = Picture1.Top + (Picture1.Height - (Picture1.Height / ert) * BaraValoareLim2.APaulText) + 5
LimitaJos.Visible = True

LimitaSus.Top = Picture1.Top + (Picture1.Height - (Picture1.Height / ert) * BaraValoareLim1.APaulText) + 5
LimitaSus.Visible = True
pa = BaraValoareLim2.APaulText
ValBaraJos.Caption = pa


ps = BaraValoareLim1.APaulText
ValBaraSus.Caption = ps

If BaraValoareLim2.APaulText > er And RauRau <= CPULaSuta Then
ElibereazaMem_Click
End If
End Sub
  

Private Sub InitCPU()
Dim lData As Long, lType As Long, lSize As Long
Dim hKey As Long
Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
If Qry <> 0 Then
MsgBox " error "
End
End If
lType = REG_DWORD
lSize = 4
Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
Qry = RegCloseKey(hKey)
End Sub

Private Sub Timer3_Timer()
Dim Zeus As Long, lType As Long, lSize As Long
Dim hKey As Long

Zeus = CPU_PE_NT.CurrentCPUUsage

Statbar = Int(Zeus / 10)
RauRau = Int(Zeus)

'----------------------- ce urmeaza nu conteaza -------------------------------------

If RauRau > 35 Then
Label5.ForeColor = &HFFFF&
Else
Label5.ForeColor = &HFF00&
End If
If RauRau > 70 Then
Label5.ForeColor = &HFF&
End If

Label5.Caption = RauRau & " %"
Qry = RegCloseKey(hKey)

Picture2.Cls
xx = (Picture2.ScaleWidth / 100) * RauRau
If xx >= YY Then

For rr = xx To (xx + 10)
YY = rr
Picture2.Line (YY - 2, 0)-(YY, 10), &H4000&, BF
Picture2.Line (YY, 0)-(YY + 2, 10), &HC0C0&, BF
Next rr
Else
YY = YY - 4
If YY <= xx Then YY = xx + 2
Picture2.Line (YY, 0)-(YY + 2, 10), &HC0C0&, BF
End If
Picture2.Line (0, 0)-(xx, 10), &HC000&, BF


BulaComando = Zeus


CPU = PicCPU.Height - ((PicCPU.Height / 100) * BulaComando)
Call BitBlt(CPUdc, 0, 0, CPUWidth - 1, CPUHeight, CPUdc, 1, 0, SRCCOPY)
If CPUcont Mod CPUSchelet = 0 Then
Call MoveToEx(CPUdc, CPUWidth - 2, 0, 0)
Call LineTo(CPUdc, CPUWidth - 2, CPUHeight)
End If
Call SelectObject(CPUdc, LinC1)
Call MoveToEx(CPUdc, CPUWidth - 3, CPUY, 0)

Call LineTo(CPUdc, CPUWidth - 2, CPU)

If CPU <> Vechi1 Then
CPUDans.Top = PicCPU.Top + CPU - (CPUDans.Height / 2)
CPUDans.Caption = ""
CPUDans.Caption = "- " & RauRau & "%"
End If
CPUceasRau = RauRau
'****************************
If RauRau > 25 Then
Limita75.BackColor = &HC0C0&
Else
Limita75.BackColor = &H4040&
End If

If RauRau > 50 Then
LED50.BackColor = &HFF&
Else
LED50.BackColor = &H80&
End If

analogmeter 1, 2, 0, 100, 0, 30, 1, vbGreen, 0.1, CPUceasRau
Vechi1 = CPU

Call SelectObject(CPUdc, LinB)
Call BitBlt(PicCPU.hdc, 0, 0, CPUWidth, CPUHeight, CPUdc, 0, 0, SRCCOPY)
CPUcont = CPUcont + 1
CPUY = CPU
End Sub

Public Sub GetBytes()
Dim r As Long, BytesFreeToCalller As Currency, TotalBytes As Currency
Dim TotalFreeBytes As Currency, TotalBytesUsed As Currency
Const RootPathName = "C:\"
Call GetDiskFreeSpaceEx(RootPathName, BytesFreeToCalller, TotalBytes, TotalFreeBytes)
    
hardTotal = Format$(TotalBytes * 10000, "###,###,###,##0")
hard = Format$(BytesFreeToCalller * 10000, "###,###,###,##0")
HardOcupat = Format$((TotalBytes - TotalFreeBytes) * 10000, "###,###,###,##0")

DiskLiber.Caption = hard & " b"

DiskTot.Caption = hardTotal & " b"
For ro = 0 To ((HardLiber.ScaleWidth / Int(hardTotal)) * Int(HardOcupat))
pu = (&HC000& * ro)
HardLiber.Line (ro, 0)-(ro + 1, 10), pu, BF
rrr = rrr + 1
Next ro
End Sub


Private Sub Ascunde_Click()
It = 0
t.cbSize = Len(t)
t.hwnd = PicturaSys(0).hwnd
t.uID = 1&
t.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
t.uCallbackMessage = WM_MOUSEMOVE
t.hIcon = PicturaSys(It).Picture
t.szTip = "DeltaMemoryForce 2.0" & Chr$(0)
Shell_NotifyIcon NIM_ADD, t
Me.Hide
End Sub

Private Sub Ascunde_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josAA = True
Ascunde.Picture = ascundeG.Picture
End Sub

Private Sub Ascunde_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josAA Then
Exit Sub
Else
Ascunde.Picture = AscundeD.Picture
End If
ElibereazaMem.Picture = optimV.Picture
End Sub

Private Sub Ascunde_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josAA = False
Ascunde.Picture = ascundeV.Picture
End Sub

Private Sub PicturaSys_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Hex(x) = "1E3C" Then
IconitaSistem.PopupMenu IconitaSistem.Optiuni666
End If
End Sub



Private Sub EnumProcess(Optional ByVal sExeName As String = "")
Dim lSnapShot As Long
Dim lNextProcess As Long
Dim tPE As PROCESSENTRY32
lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
If lSnapShot <> -1 Then
lstProcess.Clear
tPE.dwSize = Len(tPE)
lNextProcess = Process32First(lSnapShot, tPE)
Do While lNextProcess
If sExeName = Left$(tPE.szExeFile, Len(sExeName)) And Len(sExeName) > 0 Then
Dim lProcess As Long
Dim lExitCode As Long
lProcess = OpenProcess(0, False, tPE.th32ProcessID)
TerminateProcess lProcess, lExitCode
CloseHandle lProcess
DeleteJ = sExeName
TimDel.Enabled = True
Else
lstProcess.AddItem tPE.szExeFile
End If
lNextProcess = Process32Next(lSnapShot, tPE)
NrProcese = NrProcese + 1
Loop
NrProceseText.Caption = " Total process : " & NrProcese
CloseHandle (lSnapShot)
Else
lstProcess.AddItem " Error !"
End If

AddHorizontalScrollbar ProcessList

End Sub


Private Sub TimerForma_Timer()
If PointerInForm(Me) Then
paulForm = False
Exit Sub
Else
If paulForm Then Exit Sub
MasterPaulGagniuc.Picture = PriorA.Picture
Iesire.Picture = iesireV.Picture
Ajutor.Picture = ajutorV.Picture
Despre.Picture = DespreN.Picture
Optiuni.Picture = optiuniV.Picture
Procesee.Picture = ProcesV.Picture
RefreshQ.Picture = RefreshV.Picture
TerminaProces.Picture = TerminaV.Picture
ElibereazaMem.Picture = optimV.Picture
Ascunde.Picture = ascundeV.Picture
paulForm = True
End If
End Sub

'Private Sub TimerNT_Timer()
'UpdateValues
'End Sub

Private Sub UnicTimer_Timer()
If RauRau < 50 Then
ElibereazaMem_Click
UnicTimer.Enabled = False
Else
UnicTimer.Enabled = True
End If
End Sub

Private Sub Timled_Timer()
Static n
Dim keys(255) As Byte
keys(&H91) = n
tmp = SetKeyboardState(keys(0))
n = n + 1
nn = nn + 1

If nn = 2 Then LED.Picture = LedAprins.Picture
If nn = 1 Then LED.Picture = LedStins.Picture
If nn = 2 Then nn = 0
If n = 2 Then n = 0
DoEvents
End Sub

Private Sub service()
Dim pid As Long
Dim reGserv As Long
pid = GetCurrentProcessId()
reGserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub

Private Sub UnService()
Dim pid As Long
Dim reGserv As Long
pid = GetCurrentProcessId()
reGserv = RegisterServiceProcess(pid, _
RSP_UNREGISTER_SERVICE)
End Sub

Private Sub InTop()
    Dim ira
    ira = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Private Sub NuTop()
    Dim ira
    ira = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Iesire.Picture = iesireV.Picture
Despre.Picture = DespreN.Picture
Ajutor.Picture = ajutorV.Picture
Optiuni.Picture = optiuniV.Picture
MasterPaulGagniuc.Picture = PriorA.Picture
Procesee.Picture = ProcesV.Picture

Ascunde.Picture = ascundeV.Picture
ElibereazaMem.Picture = optimV.Picture

RefreshQ.Picture = RefreshV.Picture
TerminaProces.Picture = TerminaV.Picture
End Sub


Private Sub TimerScrie_Timer()
 'PlaySound App.Path & "\Setup.wav"
FXCounter = FXCounter + 1
StrTemp = Left$(ScrieDelta.Caption, FXCounter)
TextLen = Len(ScrieDelta.Caption)
DeltaMemoryMov.Caption = "            " & StrTemp

If Len(StrTemp) = Len(ScrieDelta.Caption) Then
TimerScrie.Enabled = False
StrTemp = ""
FXCounter = 0
End If
'PlaySound App.Path & "\"
End Sub

Private Function PointerInForm(lForm As Object) As Boolean
    Dim iRight%, iBottom%, iLeft%, iTop%
    Dim pt As POINTAPIX
    iRight = (lForm.Left + lForm.Width) / Screen.TwipsPerPixelX
    iBottom = (lForm.Top + lForm.Height) / Screen.TwipsPerPixelY
    iLeft = lForm.Left / Screen.TwipsPerPixelX
    iTop = lForm.Top / Screen.TwipsPerPixelY
    GetCursorPos pt
    If (pt.x <= iRight) And (pt.x >= iLeft) _
    And (pt.y >= iTop) And (pt.y <= iBottom) Then
    PointerInForm = True
Else
    PointerInForm = False
End If
End Function

'*****************************************************************************
Private Sub AddHorizontalScrollbar(target As Object)
Dim MaxLen As Integer
For mxlc = 0 To target.ListCount - 1
CrLen = Len(target.List(mxlc))
If CrLen > MaxLen Then MaxLen = CrLen: myls$ = target.List(mxlc)
Next mxlc
tgw = target.Parent.TextWidth(myls$ & " ") / Screen.TwipsPerPixelX
x = SendMessageByNum(target.hwnd, LB_SETHORIZONTALEXTENT, tgw, 0)
End Sub
'*****************************************************************************

Private Sub analogmeter(Index As Integer, Mtype As Integer, Emin As Double, Emax As Double, _
Mmin As Double, Mmax As Double, Handw As Integer, Color As String, Handl As Double, _
Value As Double)
If Mtype = 1 Then
Mdeg = 0
cx = PicCeasProcesor(Index).Width / 2
cy = PicCeasProcesor(Index).Height / 2
End If
If Mtype = 2 Then
Mdeg = 270
cx = PicCeasProcesor(Index).Width / 2
cy = (PicCeasProcesor(Index).Height - 2)
End If
If Mtype = 3 Then
Mdeg = 0
cx = PicCeasProcesor(Index).Width / 2
cy = PicCeasProcesor(Index).Height / 2
End If
If Mtype = 4 Then
Mdeg = 0
cx = PicCeasProcesor(Index).Width / 2
cy = PicCeasProcesor(Index).Height / 2
End If
r = IIf(cx > cy, cy, cx) - 5
sl = r * Handl


R1 = Emax - Emin
r2 = Mmax - Mmin
s = ((r2 / R1) * Value) + Mmin
  

  sin_ = Sin((Mdeg - s * 6) * Deg2Rad) * (r - sl) + cx
  cos_ = Cos((Mdeg - s * 6) * Deg2Rad) * (r - sl) + cy
  If Mtype = 3 Then
  sin2_ = Sin((180 - s * 6) * Deg2Rad) * ((r - sl) / 1.5) + cx
  cos2_ = Cos((180 - s * 6) * Deg2Rad) * ((r - sl) / 1.5) + cy
  ar1 = Sin((45 - s * 6) * Deg2Rad) * ((r - sl) / 2) + cx
  ar2 = Cos((45 - s * 6) * Deg2Rad) * ((r - sl) / 2) + cy
  ar3 = Sin((315 - s * 6) * Deg2Rad) * ((r - sl) / 2) + cx
  ar4 = Cos((315 - s * 6) * Deg2Rad) * ((r - sl) / 2) + cy
  End If
  If Mtype = 4 Then
  sin2_ = Sin((180 - s * 6) * Deg2Rad) * (r - sl) + cx
  cos2_ = Cos((180 - s * 6) * Deg2Rad) * (r - sl) + cy
  End If
  oldcolor = vbBlack
  PicCeasProcesor(Index).ForeColor = Color
  PicCeasProcesor(Index).DrawWidth = Handw
  PicCeasProcesor(Index).Cls
  PicCeasProcesor(Index).Line (cx, cy)-(sin_, cos_)
  If Mtype = 3 Then
  PicCeasProcesor(Index).Line (cx, cy)-(sin2_, cos2_)
  PicCeasProcesor(Index).Line (sin_, cos_)-(ar1, ar2)
  PicCeasProcesor(Index).Line (sin_, cos_)-(ar3, ar4)
  PicCeasProcesor(Index).FillColor = Color
  PicCeasProcesor(Index).FillStyle = 0
  PicCeasProcesor(Index).Circle (sin2_, cos2_), 5
  End If
  If Mtype = 4 Then
  PicCeasProcesor(Index).Line (cx, cy)-(sin2_, cos2_)
  End If
  PicCeasProcesor(Index).ForeColor = oldcolor

End Sub


'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
Sub RefreshProcessList()
Dim hSnapshot As Long
Dim uProcess As PROCESSENTRY32
Dim r As Long
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If hSnapshot = 0 Then Exit Sub
uProcess.dwSize = Len(uProcess)
r = ProcessFirst(hSnapshot, uProcess)

ProcessList.Clear
totalinf = 1
tttt = 0
Do While r
        
x$ = uProcess.szExeFile
x$ = Trim$(Mid$(x$, 1, InStr(x$, Chr$(0))))
        
ProcessList.AddItem UCase$(x$)
        
info(totalinf).exe = x$
info(totalinf).priority = Trim$(Str$(Val(uProcess.pcPriClassBase)))
info(totalinf).threads = uProcess.cntThreads
info(totalinf).processId = uProcess.th32ProcessID
        
totalinf = totalinf + 1
r = ProcessNext(hSnapshot, uProcess)
tttt = tttt + 1
Loop
NrProceseText.Caption = "Total process : " & tttt
Call CloseHandle(hSnapshot)

Form1.ScaleMode = 1
AddHorizontalScrollbar ProcessList
Form1.ScaleMode = 3

ProcessList.ListIndex = 0
End Sub




Sub KillTask(Myname As String)
Const PROCESS_ALL_ACCESS = 0
Dim uProcess As PROCESSENTRY32
Dim rProcessFound As Long
Dim hSnapshot As Long
Dim szExename As String
Dim exitCode As Long
Dim myProcess As Long
Dim AppKill As Boolean
Dim appCount As Integer
Dim I As Integer
On Local Error GoTo Finish
appCount = 0
Const TH32CS_SNAPPROCESS As Long = 2&
    
uProcess.dwSize = Len(uProcess)
hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
Do While rProcessFound
I = InStr(1, uProcess.szExeFile, Chr(0))
szExename = LCase$(Left$(uProcess.szExeFile, I - 1))
        
If Right$(szExename, Len(Myname)) = LCase$(Myname) Then
KillApp = True
appCount = appCount + 1
myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
AppKill = TerminateProcess(myProcess, exitCode)
Call CloseHandle(myProcess)
End If

rProcessFound = ProcessNext(hSnapshot, uProcess)
Loop
DeleteJ = Myname
Call CloseHandle(hSnapshot)
TimDel.Enabled = True
Finish:
End Sub

'####################################################
'AAAAAAAAAAAAAAAAAAAAAAAAAAAici
Private Sub ProcessList_Click()
lx = ProcessList.ListIndex + 1
pb = info(lx).priority
Select Case pb
Case Is = 4
prefix$ = " Idle"
Status_Priority.ForeColor = &HFF0000
Case Is = 8
prefix$ = " Normal"
Status_Priority.ForeColor = &HFF00&
Case Is = 13
prefix$ = " High"
Status_Priority.ForeColor = &HFFFF&
Case Is = 24
prefix$ = " RealTime"
Status_Priority.ForeColor = &HFF&
End Select
CurrentProcess.exe = info(lx).exe
CurrentProcess.processId = info(lx).processId
CurrentProcess.threads = info(lx).threads
CurrentProcess.priority = info(lx).priority
Status_Priority.Caption = prefix$
Status_Threads.Caption = info(lx).threads

Edit_ProcessLabel.Text = LCase$(CurrentProcess.exe)
Select Case CurrentProcess.priority
Case Is = 4
UpDateGraph "idle"
Case Is = 8
UpDateGraph "normal"
Case Is = 13
UpDateGraph "high"
Case Is = 24
UpDateGraph "realtime"
End Select
End Sub


Private Sub Edit_High_Click()
'1111111111111111111111111111111
UpDateGraph "high"
End Sub

Private Sub Edit_Idle_Click()
UpDateGraph "idle"
End Sub

Private Sub Edit_Normal_Click()
UpDateGraph "normal"
End Sub

Private Sub Edit_RealTime_Click()
UpDateGraph "realtime"
End Sub

Private Sub Pref_Graph_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
AcV = Round(4 - y)
If Edit_Current <> AcV Then Else Exit Sub
Select Case AcV
Case Is = 0
UpDateGraph "idle"
Case Is = 1
UpDateGraph "normal"
Case Is = 2
UpDateGraph "high"
Case Is = 3
UpDateGraph "realtime"
End Select
End Sub



Private Sub TB_KillTask()

If InStr(ProcessList.Text, "kernel32.dll") <> 0 Or InStr(ProcessList.Text, "mdm.exe") <> 0 Or InStr(ProcessList.Text, "msgsrv32.exe") <> 0 Then
If WarnAboutSysfiles = 1 Then Else Exit Sub
End If
KillTask ProcessList.Text
RefreshProcessList
'TimDel.Enabled = True
End Sub

'2222222222222222222333333333333333333333333333333
Sub SetProcessPriority(WhatProcess As Long, Setto As Long)
handletoclose = OpenProcess(PROCESS_ALL_ACCESS, -1, WhatProcess)
SetPriorityClass handletoclose, Setto
CloseHandle handletoclose
End Sub

Sub UpDateGraph(RefValue As String)

Edit_RealTime.ForeColor = &H8000&
Edit_High.ForeColor = &H8000&
Edit_Normal.ForeColor = &H8000&
Edit_Idle.ForeColor = &H8000&

Select Case RefValue

Case Is = "idle"
Edit_Idle.ForeColor = &HFF0000
yyyySS.Top = Edit_Idle.Top
Edit_Current = 0
Pref_SetTo = IDLE_PRIORITY_CLASS

Case Is = "normal"
Edit_Normal.ForeColor = &HFF00&
yyyySS.Top = Edit_Normal.Top
Edit_Current = 1
Pref_SetTo = NORMAL_PRIORITY_CLASS

Case Is = "high"
Edit_High.ForeColor = &HFFFF&
yyyySS.Top = Edit_High.Top
Edit_Current = 2
Pref_SetTo = HIGH_PRIORITY_CLASS

Case Is = "realtime"
Edit_RealTime.ForeColor = &HFF&
yyyySS.Top = Edit_RealTime.Top
Edit_Current = 3
Pref_SetTo = REALTIME_PRIORITY_CLASS
End Select

End Sub






