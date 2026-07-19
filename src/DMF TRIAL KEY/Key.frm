VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Sistem Criptologic by Paul Gagniuc"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   Icon            =   "Key.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Raspuns pt Client "
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
      Left            =   360
      TabIndex        =   17
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prog Decrypt"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Key intoarcere"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generare Key && &Encript"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Text            =   "Key88888"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Line Line29 
      BorderColor     =   &H000000FF&
      X1              =   2280
      X2              =   2280
      Y1              =   1200
      Y2              =   1680
   End
   Begin VB.Line Line28 
      BorderColor     =   &H000000FF&
      X1              =   2160
      X2              =   2280
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line27 
      BorderColor     =   &H000000FF&
      X1              =   2400
      X2              =   2280
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line26 
      BorderColor     =   &H000000FF&
      X1              =   5040
      X2              =   5760
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line25 
      BorderColor     =   &H000000FF&
      X1              =   5040
      X2              =   5040
      Y1              =   3720
      Y2              =   4800
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   5280
      Picture         =   "Key.frx":038A
      Top             =   3840
      Width           =   1320
   End
   Begin VB.Line Line24 
      BorderColor     =   &H000000FF&
      X1              =   5160
      X2              =   5040
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line23 
      BorderColor     =   &H000000FF&
      X1              =   4920
      X2              =   5040
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line22 
      BorderColor     =   &H000000FF&
      X1              =   5040
      X2              =   5040
      Y1              =   2280
      Y2              =   3240
   End
   Begin VB.Line Line21 
      BorderColor     =   &H000000FF&
      X1              =   3960
      X2              =   3840
      Y1              =   2160
      Y2              =   2040
   End
   Begin VB.Line Line20 
      BorderColor     =   &H000000FF&
      X1              =   3960
      X2              =   3840
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Line Line19 
      BorderColor     =   &H000000FF&
      X1              =   5520
      X2              =   5400
      Y1              =   240
      Y2              =   360
   End
   Begin VB.Line Line18 
      BorderColor     =   &H000000FF&
      X1              =   5280
      X2              =   5400
      Y1              =   240
      Y2              =   360
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000000FF&
      X1              =   5400
      X2              =   5400
      Y1              =   360
      Y2              =   120
   End
   Begin VB.Line Line16 
      BorderColor     =   &H000000FF&
      X1              =   4080
      X2              =   5400
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000000FF&
      X1              =   4080
      X2              =   4080
      Y1              =   2040
      Y2              =   120
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000000FF&
      X1              =   3840
      X2              =   4080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      X1              =   3600
      X2              =   3480
      Y1              =   3360
      Y2              =   3240
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      X1              =   3600
      X2              =   3480
      Y1              =   3120
      Y2              =   3240
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      X1              =   5040
      X2              =   3480
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Key trial"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Decript cu Key din reg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   14
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000000FF&
      X1              =   5880
      X2              =   5760
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   5640
      X2              =   5760
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   1800
      X2              =   1680
      Y1              =   4560
      Y2              =   4440
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   1680
      X2              =   1560
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1680
      X2              =   1680
      Y1              =   4800
      Y2              =   4440
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   5040
      X2              =   1680
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   5760
      X2              =   5760
      Y1              =   2280
      Y2              =   3720
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Eu "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Trial la instalare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   4200
      X2              =   4320
      Y1              =   1320
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   4200
      X2              =   4320
      Y1              =   1080
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   3840
      X2              =   4320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Key Kript"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Key Kript"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Key clar"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'             Software & Copyright by Paul Gagniuc
Private Sub Command2_Click()
Dim g As String
4:
stru = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Text1.Text = ""
For q = 1 To 20
5:
Randomize
j = Int(34 * Rnd(34))
If j = 0 Then GoTo 5
g = g + Mid(stru, j, 1)
Next q
Text1.Text = g

ty = "DeltaMemoryForcePaul"
we = Text1.Text
For hh = 1 To Len(ty)
dmf = Mid(ty, hh, 1)
dd = Asc(dmf)
fg = (dd + (Asc(Mid(we, hh, 1))) - 5)
gforce = gforce & Chr(fg)
Next hh
Text3.Text = gforce
End Sub


Private Sub Command4_Click()
tyt = Text5.Text
we = "DeltaMemoryForcePaul"
For hh = 1 To Len(tyt)
dmf = Mid(tyt, hh, 1)
dd = Asc(dmf)
fg = (dd - (Asc(Mid(we, hh, 1))) + 5)
gforce69 = gforce69 & Chr(fg)
Next hh
Text4.Text = gforce69

ty = "DMFbyPaulGagniuc6669"
we = gforce69
For hh = 1 To Len(ty)
dmf = Mid(ty, hh, 1)
dd = Asc(dmf)
fg = (dd + (Asc(Mid(we, hh, 1))) - 5)
gforce = gforce & Chr(fg)
Next hh
Text6.Text = gforce
End Sub

Private Sub Command5_Click()

PaulGagniuc = Text9.Text
DeltaM = Text5.Text
tyt = DeltaM

we = "DeltaMemoryForcePaul"
For hh = 1 To Len(tyt)
dmf = Mid(tyt, hh, 1)
dd = Asc(dmf)
fg = (dd - (Asc(Mid(we, hh, 1))) + 5)
gforce69 = gforce69 & Chr(fg)
Next hh
we = gforce69
ty = PaulGagniuc
For hh = 1 To Len(ty)
dmf = Mid(ty, hh, 1)
dd = Asc(dmf)
fg = (dd - (Asc(Mid(we, hh, 1))) + 5)
gforce = gforce & Chr(fg)
Next hh

If gforce = "DMFbyPaulGagniuc6669" Then
Text8.Text = gforce
Else
MsgBox "Invalid Key - !"
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub
