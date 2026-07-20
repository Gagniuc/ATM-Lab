VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3270
   ClientLeft      =   3855
   ClientTop       =   1125
   ClientWidth     =   5310
   Icon            =   "ScreenOptim2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "ScreenOptim2.frx":038A
   ScaleHeight     =   3270
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "admin@ramforces.com"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
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
   Begin VB.Label sec 
      BackColor       =   &H00000000&
      Caption         =   " Software && Copyright by P.A.G."
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
      TabIndex        =   3
      Top             =   2880
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "www.ramforces.com"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nuovus Ordo , Anno Domini 2002"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   3255
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   " Delta Memory Force"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
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

Timer1 = True
End Sub

Private Sub Timer1_Timer()
Timer1 = False
Timer2 = True
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
