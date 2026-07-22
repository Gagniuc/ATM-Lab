VERSION 5.00
Begin VB.Form Form6 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form6"
   ClientHeight    =   3225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   LinkTopic       =   "Form6"
   ScaleHeight     =   3225
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private lngDC As Long ' hDC of the screen, available to every sub/function, wich allows us to call ReleaseDC(0, lngDC) in cExit
Private blnLoop As Boolean


Private Sub Form_Load()
 Dim intX As Integer, intY As Integer
    Dim intI As Integer, intJ As Integer
    Dim intWidth As Integer, intHeight As Integer
    
    Form6.Height = Form1.Height
    Form6.Width = Form1.Width
    Form6.Top = Form1.Top
    Form6.Left = Form1.Left
    Form6.Visible = False
        
    intWidth = Screen.Width / Screen.TwipsPerPixelX 'Screenwidth
    intHeight = Screen.Height / Screen.TwipsPerPixelY 'Screenheight
    

    Form6.Visible = vbTrue ' Make form visible
    Form6.AutoRedraw = vbFalse ' Set autoredraw to 0 (or your graphics-card might cause a reboot)
    
    Randomize
    
    blnLoop = vbTrue
    Do While blnLoop = vbTrue
        intX = (intWidth - 128) * Rnd
        intY = (intHeight - 128) * Rnd
        
        intI = 2 * Rnd - 1
         intJ = 2 * Rnd - 2

        
        ' Move a part of the screen 1 pixel in a semi-random direction, to get the "melting" effect
        Call BitBlt(Form6.hdc, intX - intI, intY - intJ, 128, 128, Form6.hdc, intX, intY, vbSrcCopy)
        
        DoEvents
    ttrryy = ttrryy + 1
    If ttrryy = 20000 Then GoTo 232
    Loop
232:
    Set frmMain = Nothing ' Remove form from memory
    Call ReleaseDC(0, lngDC) ' Release the screen-hDC

Shell_NotifyIcon NIM_DELETE, t
Unload Form1
Unload Form2
Unload Form3
Unload Form4
Unload OptiuniDelta
Unload IconitaSistem
Unload VVV
Unload Me
End Sub
