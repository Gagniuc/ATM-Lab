VERSION 5.00
Begin VB.Form IconitaSistem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Optiuni"
   ClientHeight    =   1965
   ClientLeft      =   9555
   ClientTop       =   810
   ClientWidth     =   1755
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   1755
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Image imgExit 
      Height          =   315
      Left            =   120
      Picture         =   "IconitaSistem.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgCaution 
      Height          =   315
      Left            =   120
      Picture         =   "IconitaSistem.frx":0B9D
      Top             =   1560
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgDelete 
      Height          =   315
      Left            =   120
      Picture         =   "IconitaSistem.frx":1464
      Top             =   480
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgStop 
      Height          =   315
      Left            =   120
      Picture         =   "IconitaSistem.frx":1DCD
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Image imgYield 
      Height          =   315
      Left            =   120
      Picture         =   "IconitaSistem.frx":26B4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Menu Optiuni666 
      Caption         =   ""
      Begin VB.Menu OptimizeazaTree 
         Caption         =   ""
      End
      Begin VB.Menu Arata666 
         Caption         =   ""
      End
      Begin VB.Menu Ascunde666 
         Caption         =   ""
      End
      Begin VB.Menu Optiuni 
         Caption         =   ""
      End
      Begin VB.Menu Inchide 
         Caption         =   ""
      End
   End
End
Attribute VB_Name = "IconitaSistem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************

'                     Software & Copyright by Paul Gagniuc Aurelian

'***********************************************************************************
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Const MF_BITMAP = &H4&
Private Const MFT_BITMAP = MF_BITMAP
Private Const MIIM_TYPE = &H10

Dim pga As Variant

Private Sub Arata666_Click()
Form1.Show
End Sub

Private Sub Ascunde666_Click()
Form1.Hide
End Sub


Private Sub OptimizeazaTree_Click()
ElibereazaMem666
End Sub

Private Sub Inchide_Click()
 Shell_NotifyIcon NIM_DELETE, t
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload OptiuniDelta
Unload VVV
Unload Me
End Sub

Private Sub Optiuni_Click()
OptiuniDelta.Show
End Sub

Private Sub Form_Load()
    SetMenuBitmap Me, Array(0, 0), imgExit.Picture
    SetMenuBitmap Me, Array(0, 1), imgDelete.Picture
    SetMenuBitmap Me, Array(0, 2), imgStop.Picture
    SetMenuBitmap Me, Array(0, 3), imgYield.Picture
    SetMenuBitmap Me, Array(0, 4), imgCaution.Picture
End Sub

Public Sub SetMenuBitmap(ByVal frm As Form, ByVal item_numbers As Variant, ByVal pic As Picture)
Dim menu_handle As Long
Dim i As Integer
Dim menu_info As MENUITEMINFO

    menu_handle = GetMenu(frm.hwnd)
    For i = LBound(item_numbers) To UBound(item_numbers) - 1
        menu_handle = GetSubMenu(menu_handle, item_numbers(i))
    Next i

    With menu_info
        .cbSize = Len(menu_info)
        .fMask = MIIM_TYPE
        .fType = MFT_BITMAP
        .dwTypeData = pic
    End With

    SetMenuItemInfo menu_handle, _
        item_numbers(UBound(item_numbers)), _
        True, menu_info
End Sub
