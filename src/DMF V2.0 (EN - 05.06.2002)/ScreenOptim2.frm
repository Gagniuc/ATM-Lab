VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1575
   ClientLeft      =   3855
   ClientTop       =   1005
   ClientWidth     =   8490
   Icon            =   "ScreenOptim2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "ScreenOptim2.frx":038A
   ScaleHeight     =   105
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   566
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox FrafStelar 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "NU"
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      ScaleHeight     =   225
      ScaleWidth      =   465
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Formular 
      Caption         =   "Buy online !"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   1125
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Name69 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
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
      Left            =   2400
      MaxLength       =   25
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox R1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
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
      Left            =   2400
      MaxLength       =   25
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Cancel 
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
      Left            =   4920
      TabIndex        =   2
      Top             =   1125
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Plata 
      Caption         =   "Register"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   1125
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   600
      Top             =   720
   End
   Begin VB.Image Image2 
      Height          =   1275
      Left            =   6375
      Picture         =   "ScreenOptim2.frx":781B
      Top             =   75
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   5400
      Picture         =   "ScreenOptim2.frx":9069
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Tranzit 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   4920
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Reg69 
      BorderColor     =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   1680
      Top             =   60
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label sec 
      BackColor       =   &H00000000&
      Caption         =   "         DMF Trial mode"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
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

Dim a As String
Dim b As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Const Citire = 1
Const HKEY_CURRENT_USER = &H80000001
Const CaleRegistru = "Software\Microsoft\Windows\CurrentVersion\Run"
Const Scrie = &H20006

Dim d As Integer

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance Then End
Timer2 = False

StartWindows = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "Start")
OptimizareInceput = GetSetting("DeltaMemoryForce", "SectiuneDelta", "Optimizare")
minmin = GetSetting("DeltaMemoryForce", "SectiuneDelta", "MinimX")
CPUL = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "CPUVer")
txtCPU = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "CPULim")
ValidService = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "Vaid")
VMinOk = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "MinOk")


If StartWindows = "" And OptimizareInceput = "" And minmin = "" And CPUL = "" And txtCPU = "" And ValidService = "" Then
If RegOpenKeyEx(HKEY_CURRENT_USER, CaleRegistru, 0, Scrie, CaleCaii) Then Exit Sub
RegSetValueEx CaleCaii, App.Title, 0, Citire, ByVal App.Path & "\" & App.EXEName & ".exe", Len(App.Path & "\" & App.EXEName & ".exe")
SaveSetting "DeltaMemoryForce", "SectiuneAlfa", "Start", "Da"
End If


Form2.Left = 3800
Form2.Top = Form2.Top - Form2.Height
Timer2 = False
Transforma

Cappopa = Chr(67) & Chr(104) & Chr(97) & Chr(114) & Chr(108) & Chr(105) & Chr(122) & Chr(101)
misloc = Chr(84) & Chr(104) & Chr(101) & Chr(114) & Chr(111) & Chr(110)
popa = Chr(75) & Chr(97) & Chr(105) & Chr(115) & Chr(101) & Chr(114)
rr = GetSetting(Cappopa, misloc, popa)
Cappopa = Empty
misloc = Empty
popa = Empty


If rr = "" Then End
Timer1.Enabled = False
BackColor = vbBlack
WindowState = vbNormal
AutoRedraw = True
Font.Name = "TimesNewRoman"
a = "www.ramforces.com"
End Sub

Private Sub Formular_Click()
VVV.Show
End Sub

Private Sub Plata_Click()
PaulGagniuc = R1.Text
DeltaM = GetSetting("DeltaMemoryForce", "Trial", "Key")
tyt = DeltaM

we = Chr(68) & Chr(101) & Chr(108) & Chr(116) & Chr(97) & Chr(77) & Chr(101) & Chr(109) & Chr(111) & Chr(114) & Chr(121) & Chr(70) & Chr(111) & Chr(114) & Chr(99) & Chr(101) & Chr(80) & Chr(97) & Chr(117) & Chr(108)
For hh = 1 To Len(tyt)
dmf = Mid(tyt, hh, 1)
dd = Asc(dmf)
fg = (dd - (Asc(Mid(we, hh, 1))) + 5)
gforce69 = gforce69 & Chr(fg)
Next hh
we = gforce69
tY = PaulGagniuc
For hh = 1 To Len(tY)
dmf = Mid(tY, hh, 1)
dd = Asc(dmf)
fg = (dd - (Asc(Mid(we, hh, 1))) + 5)
gforce = gforce & Chr(fg)
Next hh
rere = " Register to " & Name69.Text
If gforce = Chr(68) & Chr(77) & Chr(70) & Chr(98) & Chr(121) & Chr(80) & Chr(97) & Chr(117) & Chr(108) & Chr(71) & Chr(97) & Chr(103) & Chr(110) & Chr(105) & Chr(117) & Chr(99) & Chr(54) & Chr(54) & Chr(54) & Chr(57) Then
SaveSetting "DeltaMemoryForce", "Trial", "FullProgram", PaulGagniuc
SaveSetting "DeltaMemoryForce", "Trial", "Name", rere
MsgBox "Valid Key - Restart  Delta Memory Force!"
Unload Me
Else
MsgBox " Invalid Key !"
Unload Me
End If
End Sub

Private Sub Timer1_Timer()
Timer1 = False
If Tranzit.Caption <> "Peste" Then
Timer2 = True
Else
Timer2 = False
End If
End Sub

Private Sub Timer2_Timer()
DeltaMin = GetSetting("DeltaMemoryForce", "SectiuneAlfa", "Minimizat")

If DeltaMin = "1" Then
Form1.Hide
Else
Form1.Show
End If

Unload Me
End Sub

Private Sub Start()
If FrafStelar.Text = "DA" Then Exit Sub
Do
For b = 1 To 64 Step 4
Sleep 1
Cls
ForeColor = RGB(80 + b, 0, 0)
Font.Size = b
CurrentX = (ScaleWidth - TextWidth(a)) \ 2
CurrentY = (ScaleHeight - TextHeight(a)) \ 2
Print a
DoEvents
Next
For b = 64 To 1 Step -4
Sleep 1
Cls
ForeColor = RGB(155 + b, 0, 0)
Font.Size = b
CurrentX = (ScaleWidth - TextWidth(a)) \ 2
CurrentY = (ScaleHeight - TextHeight(a)) \ 2
Print a
DoEvents
Next
DoEvents
If a = sec.Caption Then GoTo 455
If a = "Delta Memory Force 2.0 " Then a = sec.Caption
If a = "www.ramforces.com" Then a = "Delta Memory Force 2.0 "

Loop
455:
Cls
Timer1.Enabled = True
Image2.Visible = True
Exit Sub
End Sub

Private Sub Form_Activate()
Call Start
End Sub



