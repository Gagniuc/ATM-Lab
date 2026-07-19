VERSION 5.00
Begin VB.UserControl BaraValoareLim 
   BackColor       =   &H00000000&
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   135
   ScaleHeight     =   2220
   ScaleWidth      =   135
   Begin VB.PictureBox BaraDeltaForce 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      Picture         =   "BaraDeltaForce.ctx":0000
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   1
      Top             =   720
      Width           =   145
   End
   Begin VB.PictureBox BaraSub 
      BackColor       =   &H000000FF&
      Height          =   2175
      Left            =   45
      ScaleHeight     =   2175
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.Timer Tim 
      Interval        =   10
      Left            =   360
      Top             =   2160
   End
End
Attribute VB_Name = "BaraValoareLim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
Dim StartX As Integer
Dim StartY As Integer
Dim oo, ww As Integer
Private Sub BaraDeltaForce_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StartX = X
StartY = Y
End Sub

Private Sub BaraDeltaForce_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 0 Then Exit Sub
If (BaraDeltaForce.Top - (StartY - Y)) <= BaraSub.Top Then
BaraDeltaForce.Top = BaraSub.Top
Exit Sub
End If
If ((BaraDeltaForce.Top - (StartY - Y))) >= (BaraSub.Top + BaraSub.Height) - BaraDeltaForce.Height Then
BaraDeltaForce.Top = (BaraSub.Top + BaraSub.Height) - BaraDeltaForce.Height
Exit Sub
End If
BaraDeltaForce.Top = BaraDeltaForce.Top - (StartY - Y)
End Sub

Private Sub Tim_Timer()
rr = (BaraSub.Height - ((BaraDeltaForce.Top) - BaraSub.Top))
oo = Int(rr * ww / BaraSub.Height)
End Sub

Public Property Get APaulText() As Integer
    APaulText = oo
End Property

Public Property Get AStop() As String
    AStop = BaraDeltaForce.Top
End Property

Public Property Get ALatime() As Integer
    ALatime = BaraDeltaForce.ScaleHeight / 2
End Property

Public Property Let AAAwww(ByVal www As String)
ww = www
End Property

Public Property Let AAAStart(ByVal Start As String)
BaraDeltaForce.Top = Start
End Property

Private Sub UserControl_Initialize()
Tim.Enabled = True
End Sub

