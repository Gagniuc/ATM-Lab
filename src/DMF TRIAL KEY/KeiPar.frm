VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copyright by Paul Gagniuc"
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "KeiPar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "$"
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'pws: PaulGagniucNOS

Private Sub Command1_Click()
    If Text1.Text = Chr(80) & Chr(97) & Chr(117) & Chr(108) & Chr(71) & Chr(97) & Chr(103) & Chr(110) & Chr(105) & Chr(117) & Chr(99) & Chr(78) & Chr(79) & Chr(83) Then
Form1.Show
Else
    MsgBox " Nu ! & Pa Pa !"
End If
Unload Me
End Sub
