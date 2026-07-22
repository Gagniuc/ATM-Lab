VERSION 5.00
Begin VB.Form OptiuniDelta 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "OptiuniDelta"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5175
   Icon            =   "OptiuniDelta.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Gagauta 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3720
      TabIndex        =   28
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   4800
      TabIndex        =   27
      Top             =   30
      Width           =   255
   End
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   3240
      Top             =   1080
   End
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   3600
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4440
      TabIndex        =   26
      Text            =   "Options"
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox VORLON 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "OptiuniDelta.frx":0CCA
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   25
      Top             =   1080
      Width           =   375
   End
   Begin VB.CheckBox CaMonitor 
      BackColor       =   &H00000000&
      Caption         =   "Act like a monitor"
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
      Left            =   240
      TabIndex        =   24
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   4320
      Top             =   1080
   End
   Begin VB.PictureBox Deasupra 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      Picture         =   "OptiuniDelta.frx":1144
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox VOF 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "OptiuniDelta.frx":15BE
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   18
      Top             =   2880
      Width           =   375
   End
   Begin VB.PictureBox VRC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "OptiuniDelta.frx":1A38
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   17
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox VPM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "OptiuniDelta.frx":1EB2
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   16
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox VPSER 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "OptiuniDelta.frx":232C
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox VOP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "OptiuniDelta.frx":27A6
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   14
      Top             =   1800
      Width           =   375
   End
   Begin VB.PictureBox NeVazut 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      Picture         =   "OptiuniDelta.frx":2C20
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Vazut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      Picture         =   "OptiuniDelta.frx":309A
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox VPS 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "OptiuniDelta.frx":356A
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   11
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox CPULimita 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "50"
      Top             =   2520
      Width           =   375
   End
   Begin VB.CheckBox CPULaSuta 
      BackColor       =   &H00000000&
      Caption         =   "Recovery after CPU is under "
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
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CheckBox Pservice 
      BackColor       =   &H00000000&
      Caption         =   "Run under service mode"
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
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CheckBox PornireMinim 
      BackColor       =   &H00000000&
      Caption         =   "Run minimize mode"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3960
      Top             =   1080
   End
   Begin VB.PictureBox Py 
      BackColor       =   &H0000FF00&
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4695
      TabIndex        =   4
      Top             =   4800
      Width           =   4695
   End
   Begin VB.PictureBox Px 
      BackColor       =   &H0000FF00&
      Height          =   3255
      Left            =   5280
      ScaleHeight     =   3255
      ScaleWidth      =   15
      TabIndex        =   5
      Top             =   0
      Width           =   15
   End
   Begin VB.TextBox MinuteOptim 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "30"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CheckBox OptMin 
      BackColor       =   &H00000000&
      Caption         =   "Optimize the RAM  at every         minutes"
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
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   3975
   End
   Begin VB.CheckBox OptLaPornire 
      BackColor       =   &H00000000&
      Caption         =   "Optimize at start up"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox PornireCuSistem 
      BackColor       =   &H00000000&
      Caption         =   "Run at start up"
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
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1470
      Left            =   2400
      Picture         =   "OptiuniDelta.frx":39E4
      Top             =   600
      Width           =   2550
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   120
      Picture         =   "OptiuniDelta.frx":66D7
      Top             =   75
      Width           =   240
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004000&
      Height          =   255
      Left            =   3600
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004000&
      Height          =   255
      Left            =   1920
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   255
      Left            =   3600
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   255
      Left            =   1920
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00004000&
      Height          =   255
      Left            =   240
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      Height          =   255
      Left            =   240
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Aplay 
      BackColor       =   &H00000000&
      Caption         =   "     Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   3600
      TabIndex        =   23
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Cancel 
      BackColor       =   &H00000000&
      Caption         =   "     Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Ok 
      BackColor       =   &H00000000&
      Caption         =   "        Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      X1              =   0
      X2              =   344
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C000&
      Height          =   3975
      Left            =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "                        Options"
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
      TabIndex        =   20
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   2520
      Width           =   255
   End
End
Attribute VB_Name = "OptiuniDelta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPIX) As Long
   
Private Type POINTAPIX
    x As Long
    y As Long
End Type

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Const Citire = 1
Const HKEY_CURRENT_USER = &H80000001
Const CaleRegistru = "Software\Microsoft\Windows\CurrentVersion\Run"
Const Scrie = &H20006
Dim CaleCaii As Long
Dim paul, josG3, josG4, josG2, cc, pp, aa, www, SSAmors, okss, cancelss, afara As Boolean
Dim g, jjop As String
Dim gh As Boolean
Dim fgfg As Integer

Dim I, t As Integer
Dim paul69 As Boolean
Dim ff As String

Private Type Soricel
    x As Long
    y As Long
End Type

Private Sub Timer3_Timer()

If Gagauta.Text = "Da" Then
I = 46
t = 0
jjop = ""
Label2.Caption = ""
paul69 = True
Gagauta.Text = "Nu"
End If

If Not paul69 Then GoTo 2
1:
t = t + 1
paul69 = False
ff = Mid(Text1.Text, t, 1)
2:
I = I + 1
Text2.Text = Chr(I)
If Len(Label2.Caption) = Len(Text1.Text & "                   ") Then
Timer3.Enabled = False
I = 46
t = 0
Text2.Text = ""
paul69 = True
End If
If Chr(I) = ff Then
Text2.ForeColor = &HFF00&
jjop = jjop & ff
Label2.Caption = "                   " & jjop
I = 46
paul69 = True
Timer4.Enabled = True
GoTo 1
End If
End Sub

Private Sub Timer4_Timer()
Text2.ForeColor = &H8000&
Timer4.Enabled = False
End Sub



Private Sub Aplay_Click()
'***************************&
If CaMonitor.Value = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Monitor", "DA"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Mon", "1"
Else
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Mon", "NU"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Mon", "0"
End If

If OptMin.Value = 1 Then
dd = MinuteOptim.Text
tt = "Da"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Verificat", "DA"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "OptimizareInterval", tt
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Minute", dd
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Amorsare", "Amors"
Form1.OptimizareTimp.Enabled = True
Else
tt = "Nu"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Verificat", "Nu"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "OptimizareInterval", tt
Form1.Label12.Visible = False
Form1.MinutText.Visible = False
Form1.Label13.Visible = False
Form1.SecundaText.Visible = False
Form1.OptimizareTimp.Enabled = False
Form1.OptimizareTimp.Enabled = False
Form1.MinutText.Text = "00"
Form1.SecundaText.Text = "00"
End If

If PornireCuSistem.Value = 1 Then
If RegOpenKeyEx(HKEY_CURRENT_USER, CaleRegistru, 0, Scrie, CaleCaii) Then Exit Sub
RegSetValueEx CaleCaii, App.Title, 0, Citire, ByVal App.Path & "\" & App.EXEName & ".exe", Len(App.Path & "\" & App.EXEName & ".exe")
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Start", "Da"
Else
If RegOpenKeyEx(HKEY_CURRENT_USER, CaleRegistru, 0, Scrie, CaleCaii) Then Exit Sub
RegDeleteValue CaleCaii, App.Title
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Start", "Nu"
End If

If OptLaPornire.Value = 1 Then
dd = "Da"
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "Optimizare", dd
Else
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "Optimizare", "Nu"
End If
If PornireMinim.Value = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "Minim", "Da"
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "MinimX", "Da"
Else
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "Minim", "Nu"
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "MinimX", "Nu"
End If

If CPULaSuta = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "CPULim", CPULimita.Text
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "CPUVer", "1"
Else
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "CPULim", "50"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "CPUVer", "0"
End If

If Pservice.Value = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Service", "1"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Vaid", "1"
Else
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Service", "0"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Vaid", "0"
End If

If PornireMinim.Value = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Minimizat", "1"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "MinOk", "1"
Else
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Minimizat", "0"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "MinOk", "0"
End If
End Sub

Private Sub Aplay_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josG4 = True
Aplay.ForeColor = &HFF00&
Aplay.FontSize = 8
End Sub

Private Sub Aplay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josG4 Then
Exit Sub
Else
If Not aa Then
Aplay.ForeColor = &HFFFF&
Aplay.FontSize = 10
End If
aa = True
End If
pp = False
If cancelss Then
Cancel.ForeColor = &HC0C0&
Cancel.FontSize = 8
End If
cancelss = False
www = True
SSAmors = True
okss = True

afara = False
End Sub

Private Sub Aplay_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josG4 = False
End Sub


Private Sub Cancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josG3 = True
Cancel.ForeColor = &HFF00&
Cancel.FontSize = 8
End Sub

Private Sub Cancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josG3 Then
Exit Sub
Else
If Not pp Then
Cancel.ForeColor = &HFFFF&
Cancel.FontSize = 10
End If
pp = True
End If
cc = False
aa = False
If SSAmors Then
Ok.ForeColor = &HC0C0&
Ok.FontSize = 8
Aplay.ForeColor = &HC0C0&
Aplay.FontSize = 8
End If
SSAmors = False
www = True
okss = True
cancelss = True

afara = False
End Sub

Private Sub Cancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josG3 = False
End Sub

Private Sub CPULimita_KeyPress(KeyAscii As Integer)
If KeyAscii > 47 And KeyAscii <= 57 Then
Else
KeyAscii = 0
CPULimita.Text = ""
End If
End Sub

Private Sub Form_Load()
Amorsare = GetSetting("DeltaMemoryForce", "SectiuneBeta", "Verificat")
StartWindows = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "Start")
OptimizareInceput = GetSetting("DeltaMemoryForce", "SectiuneDelta", "Optimizare")
minmin = GetSetting("DeltaMemoryForce", "SectiuneDelta", "MinimX")
CPUL = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "CPUVer")
txtCPU = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "CPULim")
ValidService = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "Vaid")
VMinOk = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "MinOk")
VorlonWorld = GetSetting("DeltaMemoryForce", "SectiuneBeta", "Mon")
TimpulOptimizarii = GetSetting("DeltaMemoryForce", "SectiuneBeta", "Minute")
Timer3 = True
I = 46
Label2.Caption = ""
paul69 = True

If TimpulOptimizarii = "" Then TimpulOptimizarii = "30"
If txtCPU = "" Then txtCPU = "50"

If VorlonWorld = "1" Then
CaMonitor.Value = 1
VORLON.Picture = Vazut.Picture
Else
CaMonitor.Value = 0
VORLON.Picture = NeVazut.Picture
End If

If VMinOk = "1" Then
PornireMinim.Value = 1
VPM.Picture = Vazut.Picture
Else
PornireMinim.Value = 0
VPM.Picture = NeVazut.Picture
End If

If ValidService = "1" Then
Pservice.Value = 1
VPSER.Picture = Vazut.Picture
Else
Pservice.Value = 0
VPSER.Picture = NeVazut.Picture
End If

If TimpulOptimizarii <> 30 Then
MinuteOptim.Text = TimpulOptimizarii
End If


If txtCPU <> 50 Then
CPULimita.Text = txtCPU
End If

If CPUL = "1" Then
CPULaSuta.Value = 1
VRC.Picture = Vazut.Picture
Else
CPULaSuta.Value = 0
VRC.Picture = NeVazut.Picture
End If

If Amorsare = "DA" Then
OptMin.Value = 1
VOF.Picture = Vazut.Picture
Else
OptMin.Value = 0
VOF.Picture = NeVazut.Picture
End If

If StartWindows = "Da" Then
PornireCuSistem.Value = 1
VPS.Picture = Vazut.Picture
Else
PornireCuSistem.Value = 0
VPS.Picture = NeVazut.Picture
End If

If OptimizareInceput = "Da" Then
OptLaPornire.Value = 1
VOP.Picture = Vazut.Picture
Else
OptLaPornire.Value = 0
VOP.Picture = NeVazut.Picture
End If

If minmin = "Da" Then
PornireMinim.Value = 1
Else
PornireMinim.Value = 0
End If

End Sub



Private Sub Image3_Click()
rtty80 = GetSetting("DeltaMemoryForce", "Trial", "FullProgram")
If rtty80 = "" Then VVV.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(OptiuniDelta.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub


Private Sub MinuteOptim_KeyPress(KeyAscii As Integer)
If KeyAscii > 47 And KeyAscii <= 57 Then
Else
KeyAscii = 0
MinuteOptim.Text = ""
End If
End Sub

Private Sub Ok_Click()
'**************************&
If CaMonitor.Value = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Monitor", "DA"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Mon", "1"
Else
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Monitor", "NU"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Mon", "0"
End If

If OptMin.Value = 1 Then
dd = MinuteOptim.Text
tt = "Da"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Verificat", "DA"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "OptimizareInterval", tt
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Minute", dd
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Amorsare", "Amors"
Form1.OptimizareTimp.Enabled = True
Else
tt = "Nu"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "Verificat", "Nu"
SaveSetting "DeltaMemoryForce", "SectiuneBeta", "OptimizareInterval", tt
Form1.Label12.Visible = False
Form1.MinutText.Visible = False
Form1.Label13.Visible = False
Form1.SecundaText.Visible = False
Form1.OptimizareTimp.Enabled = False
Form1.OptimizareTimp.Enabled = False
Form1.MinutText.Text = "00"
Form1.SecundaText.Text = "00"
End If

If PornireCuSistem.Value = 1 Then
If RegOpenKeyEx(HKEY_CURRENT_USER, CaleRegistru, 0, Scrie, CaleCaii) Then Exit Sub
RegSetValueEx CaleCaii, App.Title, 0, Citire, ByVal App.Path & "\" & App.EXEName & ".exe", Len(App.Path & "\" & App.EXEName & ".exe")
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Start", "Da"
Else
If RegOpenKeyEx(HKEY_CURRENT_USER, CaleRegistru, 0, Scrie, CaleCaii) Then Exit Sub
RegDeleteValue CaleCaii, App.Title
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Start", "Nu"
End If

If OptLaPornire.Value = 1 Then
dd = "Da"
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "Optimizare", dd
Else
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "Optimizare", "Nu"
End If
If PornireMinim.Value = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "Minim", "Da"
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "MinimX", "Da"
Else
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "Minim", "Nu"
SaveSetting "DeltaMemoryForce", "SectiuneDelta", "MinimX", "Nu"
End If

If CPULaSuta = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "CPULim", CPULimita.Text
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "CPUVer", "1"
Else
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "CPULim", "50"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "CPUVer", "0"
End If

If Pservice.Value = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Service", "1"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Vaid", "1"
Else
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Service", "0"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Vaid", "0"
End If

If PornireMinim.Value = 1 Then
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Minimizat", "1"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "MinOk", "1"
Else
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Minimizat", "0"
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "MinOk", "0"
End If

Unload Me

End Sub

Private Sub Cancel_Click()
Timer3.Enabled = False
Timer4.Enabled = False
jjop = ""
Label2.Caption = ""
Unload Me
End Sub

Private Sub CPULaSuta_Click()
If CPULaSuta.Value = 1 Then
VRC.Picture = Vazut.Picture
Else
VRC.Picture = NeVazut.Picture
End If
End Sub

Private Sub Ok_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
josG2 = True
Ok.ForeColor = &HFF00&
Ok.FontSize = 8
End Sub

Private Sub OK_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If josG2 Then
Exit Sub
Else
If Not cc Then
Ok.ForeColor = &HFFFF&
Ok.FontSize = 10
End If
cc = True
End If
pp = False
If okss Then
Cancel.ForeColor = &HC0C0&
Cancel.FontSize = 8
End If
okss = False
www = True
SSAmors = True
cancelss = True

afara = False
End Sub

Private Sub Ok_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
josG2 = False
End Sub

Private Sub OptLaPornire_Click()
If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If
End Sub

Private Sub OptMin_Click()
If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If
End Sub

Private Sub PornireCuSistem_Click()
If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If
End Sub

Private Sub PornireMinim_Click()
If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If
End Sub

Private Sub CaMonitor_Click()
If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If
End Sub

Private Sub Pservice_Click()
If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If
End Sub



Private Sub VOF_Click()
If OptMin.Value = 1 Then
 VOF.Picture = NeVazut.Picture
OptMin.Value = 0
Else
 VOF.Picture = Vazut.Picture
OptMin.Value = 1
End If
End Sub

Private Sub VOP_Click()
If OptLaPornire.Value = 1 Then
VOP.Picture = NeVazut.Picture
OptLaPornire.Value = 0
Else
VOP.Picture = Vazut.Picture
OptLaPornire.Value = 1
End If
End Sub



Private Sub VORLON_Click()
'***
If CaMonitor.Value = 1 Then
VORLON.Picture = NeVazut.Picture
CaMonitor.Value = 0
Else
VORLON.Picture = Vazut.Picture
CaMonitor.Value = 1
End If
End Sub

Private Sub VORLON_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'***
If CaMonitor.Value = 1 Then
Exit Sub
Else
VORLON.Picture = Deasupra.Picture
End If

If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If

If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If

If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If

If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If

If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If
paul = False
End Sub

Private Sub VPM_Click()
If PornireMinim.Value = 1 Then
VPM.Picture = NeVazut.Picture
PornireMinim.Value = 0
Else
VPM.Picture = Vazut.Picture
PornireMinim.Value = 1
End If
End Sub

Private Sub VPS_Click()
If PornireCuSistem.Value = 1 Then
VPS.Picture = NeVazut.Picture
PornireCuSistem.Value = 0
Else
VPS.Picture = Vazut.Picture
PornireCuSistem.Value = 1
End If
End Sub

Private Sub VPSER_Click()
If Pservice.Value = 1 Then
VPSER.Picture = NeVazut.Picture
Pservice.Value = 0
Else
VPSER.Picture = Vazut.Picture
Pservice.Value = 1
End If
End Sub

Private Sub VRC_Click()
If CPULaSuta.Value = 1 Then
VRC.Picture = NeVazut.Picture
CPULaSuta.Value = 0
Else
VRC.Picture = Vazut.Picture
CPULaSuta.Value = 1
End If
End Sub

Private Sub VRC_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If CPULaSuta.Value = 1 Then
Exit Sub
Else
VRC.Picture = Deasupra.Picture
End If

If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If

If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If

If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If

If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If

If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If

If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If

paul = False
End Sub

Private Sub VPSER_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Pservice.Value = 1 Then
Exit Sub
Else
VPSER.Picture = Deasupra.Picture
End If

If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If

If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If

If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If

If CPULaSuta.Value = 1 Then
VRC.Picture = Vazut.Picture
Else
VRC.Picture = NeVazut.Picture
End If

If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If

If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If

paul = False
End Sub

Private Sub VPM_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If PornireMinim.Value = 1 Then
Exit Sub
Else
VPM.Picture = Deasupra.Picture
End If


If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If

If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If

If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If

If CPULaSuta.Value = 1 Then
VRC.Picture = Vazut.Picture
Else
VRC.Picture = NeVazut.Picture
End If

If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If

If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If

paul = False
End Sub

Private Sub VPS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If PornireCuSistem.Value = 1 Then
Exit Sub
Else
VPS.Picture = Deasupra.Picture
End If


If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If

If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If

If CPULaSuta.Value = 1 Then
VRC.Picture = Vazut.Picture
Else
VRC.Picture = NeVazut.Picture
End If

If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If

If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If

If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If

paul = False
End Sub

Private Sub VOF_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If OptMin.Value = 1 Then
Exit Sub
Else
VOF.Picture = Deasupra.Picture
End If

If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If

If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If

If CPULaSuta.Value = 1 Then
VRC.Picture = Vazut.Picture
Else
VRC.Picture = NeVazut.Picture
End If

If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If

If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If

If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If

paul = False
End Sub

Private Sub VOP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If OptLaPornire.Value = 1 Then
Exit Sub
Else
VOP.Picture = Deasupra.Picture
End If

If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If

If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If

If CPULaSuta.Value = 1 Then
VRC.Picture = Vazut.Picture
Else
VRC.Picture = NeVazut.Picture
End If

If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If

If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If

If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If

paul = False
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If www Then
Ok.ForeColor = &HC0C0&
Ok.FontSize = 8

Cancel.ForeColor = &HC0C0&
Cancel.FontSize = 8

Aplay.ForeColor = &HC0C0&
Aplay.FontSize = 8
End If
www = False
aa = False
pp = False
cc = False
If paul Then Exit Sub


If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If

If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If

If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If

If CPULaSuta.Value = 1 Then
VRC.Picture = Vazut.Picture
Else
VRC.Picture = NeVazut.Picture
End If

If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If

If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If

If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If
paul = True
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


Private Sub Timer2_Timer()
If PointerInForm(Me) Then
Exit Sub
Else
If afara Then GoTo 30
Cancel.ForeColor = &HC0C0&
Cancel.FontSize = 8
Aplay.ForeColor = &HC0C0&
Aplay.FontSize = 8
Ok.ForeColor = &HC0C0&
Ok.FontSize = 8

afara = True
30:
If paul Then Exit Sub

If CaMonitor.Value = 1 Then
VORLON.Picture = Vazut.Picture
Else
VORLON.Picture = NeVazut.Picture
End If

If OptLaPornire.Value = 1 Then
VOP.Picture = Vazut.Picture
Else
VOP.Picture = NeVazut.Picture
End If

If OptMin.Value = 1 Then
VOF.Picture = Vazut.Picture
Else
VOF.Picture = NeVazut.Picture
End If

If PornireCuSistem.Value = 1 Then
VPS.Picture = Vazut.Picture
Else
VPS.Picture = NeVazut.Picture
End If

If CPULaSuta.Value = 1 Then
VRC.Picture = Vazut.Picture
Else
VRC.Picture = NeVazut.Picture
End If

If Pservice.Value = 1 Then
VPSER.Picture = Vazut.Picture
Else
VPSER.Picture = NeVazut.Picture
End If

If PornireMinim.Value = 1 Then
VPM.Picture = Vazut.Picture
Else
VPM.Picture = NeVazut.Picture
End If
paul = True

End If
End Sub

        




