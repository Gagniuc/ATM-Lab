Attribute VB_Name = "Securitate"
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
"GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Dim Inteligenta, Gagniuc, Dream As Boolean
Dim WinDir, decpitat  As String
Dim msBuffer As String * 255

Const BUFFERSIZE = 255

Public Sub Transforma()

fsiz = ShowFileSize(App.Path & "\" & App.EXEName & ".exe")
If fsiz > 627100 Then End

ssddp = GetSetting("DeltaMemoryForce", "Trial", "FullProgram")
ssoop = GetSetting("DeltaMemoryForce", "Trial", "Key")

If ssddp = "" Then GoTo 2

PaulGagniuc = ssddp
tyt = ssoop

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

If gforce = Chr(68) & Chr(77) & Chr(70) & Chr(98) & Chr(121) & Chr(80) & Chr(97) & Chr(117) & Chr(108) & Chr(71) & Chr(97) & Chr(103) & Chr(110) & Chr(105) & Chr(117) & Chr(99) & Chr(54) & Chr(54) & Chr(54) & Chr(57) Then
Detinaror = GetSetting("DeltaMemoryForce", "Trial", "Name")
Form2.sec.Caption = Detinaror
Exit Sub
Else
End If


2:
Gagniuc = True
Inteligenta = False
Dream = False
Cappopa9 = Chr(119) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115) & Chr(51) & Chr(50) & Chr(73) & Chr(66) & Chr(77)
misloc9 = Chr(73) & Chr(66) & Chr(77)
popa9 = Chr(101) & Chr(114) & Chr(114) & Chr(111) & Chr(114)
fostdanu = GetSetting(Cappopa9, misloc9, popa9)
Cappopa9 = Empty
misloc9 = Empty
popa9 = Empty
Cappopa69 = Chr(82) & Chr(101) & Chr(103) & Chr(51) & Chr(50) & Chr(71) & Chr(80) & Chr(83)
misloc69 = Chr(87) & Chr(105) & Chr(110) & Chr(71) & Chr(80) & Chr(83)
popa69 = Chr(83) & Chr(97) & Chr(116)
OglindaOglinjoara = GetSetting(Cappopa69, misloc69, popa69)
Cappopa69 = Empty
misloc69 = Empty
popa69 = Empty

Dim lBytes As Long
lBytes = GetWindowsDirectory(msBuffer, BUFFERSIZE)
WinDir = Left$(msBuffer, lBytes)
uiui = WinDir & Chr(92) & Chr(83) & Chr(121) & Chr(115) & Chr(116) & Chr(101) & Chr(109) & Chr(51) & Chr(50) & Chr(46) & Chr(112) & Chr(103) & Chr(97)

On Error GoTo 4
Open uiui For Input As #1
Input #1, mm
Close #1
GoTo 5
4:
If fostdanu <> "" And OglindaOglinjoara <> "" Then
Open uiui For Output As #1
Print #1, Chr(83) & Chr(121) & Chr(115) & Chr(116) & Chr(101) & Chr(109) & Chr(32) & Chr(114) & Chr(117) & Chr(110)
Close #1
End If
mm = ""
5:
If mm <> "" And fostdanu = "" And OglindaOglinjoara = "" Then End
If fostdanu = "" And OglindaOglinjoara = "" And mm <> "" Or fostdanu <> "" And OglindaOglinjoara = "" Or fostdanu = "" And OglindaOglinjoara <> "" Then

Cappopa9 = Chr(119) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115) & Chr(51) & Chr(50) & Chr(73) & Chr(66) & Chr(77)
misloc9 = Chr(73) & Chr(66) & Chr(77)
popa9 = Chr(101) & Chr(114) & Chr(114) & Chr(111) & Chr(114)
SaveSetting Cappopa9, misloc9, popa9, "6"
Cappopa9 = Empty
misloc9 = Empty
popa9 = Empty

Cappopa69 = Chr(82) & Chr(101) & Chr(103) & Chr(51) & Chr(50) & Chr(71) & Chr(80) & Chr(83)
misloc69 = Chr(87) & Chr(105) & Chr(110) & Chr(71) & Chr(80) & Chr(83)
popa69 = Chr(83) & Chr(97) & Chr(116)
SaveSetting Cappopa69, misloc69, popa69, "6"
Cappopa69 = Empty
misloc69 = Empty
popa69 = Empty

Inteligenta = True
GoTo 34
End If
Cappopa8 = Chr(77) & Chr(79) & Chr(86) & Chr(65) & Chr(76) & Chr(79)
misloc8 = Chr(109) & Chr(111) & Chr(100) & Chr(101) & Chr(109)
popa8 = Chr(108) & Chr(111) & Chr(97) & Chr(100)
siguranta = GetSetting(Cappopa8, misloc8, popa8)
Cappopa8 = Empty
misloc8 = Empty
popa8 = Empty

Cappopa = Chr(67) & Chr(104) & Chr(97) & Chr(114) & Chr(108) & Chr(105) & Chr(122) & Chr(101)
misloc = Chr(84) & Chr(104) & Chr(101) & Chr(114) & Chr(111) & Chr(110)
popa = Chr(75) & Chr(97) & Chr(105) & Chr(115) & Chr(101) & Chr(114)
rr = GetSetting(Cappopa, misloc, popa)
Cappopa = Empty
misloc = Empty
popa = Empty

If Inteligenta And rr <> "" Or Inteligenta And siguranta <> "" Then
End
PaulDMF
End
End If

QDelta = GetSetting("DeltaMemoryForce", "Trial", "Key")

If fostdanu <> "6" Then
If fostdanu = "" And OglindaOglinjoara = "" And mm = "" And rr = "" And siguranta = "" And QDelta = "" Then

Dim g As String
stru = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
For q = 1 To 20
7:
Randomize
j = Int(34 * Rnd(34))
If j = 0 Then GoTo 7
g = g + Mid(stru, j, 1)
Next q

ty = "DeltaMemoryForcePaul"
we = g
For hh = 1 To Len(ty)
dmf = Mid(ty, hh, 1)
dd = Asc(dmf)
fg = (dd + (Asc(Mid(we, hh, 1))) - 5)
gforce = gforce & Chr(fg)
Next hh
SaveSetting "DeltaMemoryForce", "Trial", "Key", gforce

rr = Empty
siguranta = Empty

Open uiui For Output As #1
Print #1, Chr(83) & Chr(121) & Chr(115) & Chr(116) & Chr(101) & Chr(109) & Chr(32) & Chr(114) & Chr(117) & Chr(110)
Close #1

Cappopa9 = Chr(119) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115) & Chr(51) & Chr(50) & Chr(73) & Chr(66) & Chr(77)
misloc9 = Chr(73) & Chr(66) & Chr(77)
popa9 = Chr(101) & Chr(114) & Chr(114) & Chr(111) & Chr(114)
SaveSetting Cappopa9, misloc9, popa9, "6"
Cappopa9 = Empty
misloc9 = Empty
popa9 = Empty

Cappopa69 = Chr(82) & Chr(101) & Chr(103) & Chr(51) & Chr(50) & Chr(71) & Chr(80) & Chr(83)
misloc69 = Chr(87) & Chr(105) & Chr(110) & Chr(71) & Chr(80) & Chr(83)
popa69 = Chr(83) & Chr(97) & Chr(116)
SaveSetting Cappopa69, misloc69, popa69, "6"
Cappopa69 = Empty
misloc69 = Empty
popa69 = Empty

Cappopa = Chr(67) & Chr(104) & Chr(97) & Chr(114) & Chr(108) & Chr(105) & Chr(122) & Chr(101)
misloc = Chr(84) & Chr(104) & Chr(101) & Chr(114) & Chr(111) & Chr(110)
popa = Chr(75) & Chr(97) & Chr(105) & Chr(115) & Chr(101) & Chr(114)
SaveSetting Cappopa, misloc, popa, "0"
Cappopa = Empty
misloc = Empty
popa = Empty

Cappopa2 = Chr(73) & Chr(110) & Chr(116) & Chr(101) & Chr(108) & Chr(51) & Chr(50) & Chr(83) & Chr(121) & Chr(115)
misloc2 = Chr(109) & Chr(105) & Chr(99) & Chr(114) & Chr(111) & Chr(115) & Chr(111) & Chr(102) & Chr(116)
popa2 = Chr(119) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115)
SaveSetting Cappopa2, misloc2, popa2, "0"
Cappopa2 = Empty
misloc2 = Empty
popa2 = Empty

Cappopa8 = Chr(77) & Chr(79) & Chr(86) & Chr(65) & Chr(76) & Chr(79)
misloc8 = Chr(109) & Chr(111) & Chr(100) & Chr(101) & Chr(109)
popa8 = Chr(108) & Chr(111) & Chr(97) & Chr(100)
SaveSetting Cappopa8, misloc8, popa8, "0"
Cappopa8 = Empty
misloc8 = Empty
popa8 = Empty
End If

Else
Gagniuc = False
If fostdanu <> "6" Then
End
PaulDMF
End If
End If

If fostdanu <> OglindaOglinjoara Then
End
PaulDMF
Else

mm = Empty
Cappopa = Chr(67) & Chr(104) & Chr(97) & Chr(114) & Chr(108) & Chr(105) & Chr(122) & Chr(101)
misloc = Chr(84) & Chr(104) & Chr(101) & Chr(114) & Chr(111) & Chr(110)
popa = Chr(75) & Chr(97) & Chr(105) & Chr(115) & Chr(101) & Chr(114)
rr = GetSetting(Cappopa, misloc, popa)
Cappopa = Empty
misloc = Empty
popa = Empty

Cappopa2 = Chr(73) & Chr(110) & Chr(116) & Chr(101) & Chr(108) & Chr(51) & Chr(50) & Chr(83) & Chr(121) & Chr(115)
misloc2 = Chr(109) & Chr(105) & Chr(99) & Chr(114) & Chr(111) & Chr(115) & Chr(111) & Chr(102) & Chr(116)
popa2 = Chr(119) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115)
danu = GetSetting(Cappopa2, misloc2, popa2)
Cappopa2 = Empty
misloc2 = Empty
popa2 = Empty

Cappopa8 = Chr(77) & Chr(79) & Chr(86) & Chr(65) & Chr(76) & Chr(79)
misloc8 = Chr(109) & Chr(111) & Chr(100) & Chr(101) & Chr(109)
popa8 = Chr(108) & Chr(111) & Chr(97) & Chr(100)
siguranta = GetSetting(Cappopa8, misloc8, popa8)
Cappopa8 = Empty
misloc8 = Empty
popa8 = Empty

If rr = danu Then
GoTo 9
Else
End
PaulDMF
End If

End If
9:
decpitat = rr
If Not Gagniuc And rr = "" Then
End
PaulDMF
End If
If rr <> decpitat Then
End
PaulDMF
End If

If rr <> siguranta Then
If danu = siguranta Then
End
PaulDMF
End If
End If

If danu <> siguranta Then
If rr = siguranta Then
End
PaulDMF
End If
End If

If siguranta <> danu Then
If rr = danu Then
End
PaulDMF
End If
End If
danu = Empty
siguranta = Empty
ttyy = 30 - rr
Form2.sec.Caption = "      Start up " & ttyy & " of 30"
rr = rr + 1
If rr > 30 Then GoTo 34
Cappopa = Chr(67) & Chr(104) & Chr(97) & Chr(114) & Chr(108) & Chr(105) & Chr(122) & Chr(101)
misloc = Chr(84) & Chr(104) & Chr(101) & Chr(114) & Chr(111) & Chr(110)
popa = Chr(75) & Chr(97) & Chr(105) & Chr(115) & Chr(101) & Chr(114)
SaveSetting Cappopa, misloc, popa, rr
Cappopa = Empty
misloc = Empty
popa = Empty

Cappopa2 = Chr(73) & Chr(110) & Chr(116) & Chr(101) & Chr(108) & Chr(51) & Chr(50) & Chr(83) & Chr(121) & Chr(115)
misloc2 = Chr(109) & Chr(105) & Chr(99) & Chr(114) & Chr(111) & Chr(115) & Chr(111) & Chr(102) & Chr(116)
popa2 = Chr(119) & Chr(105) & Chr(110) & Chr(100) & Chr(111) & Chr(119) & Chr(115)
SaveSetting Cappopa2, misloc2, popa2, rr
Cappopa2 = Empty
misloc2 = Empty
popa2 = Empty

Cappopa8 = Chr(77) & Chr(79) & Chr(86) & Chr(65) & Chr(76) & Chr(79)
misloc8 = Chr(109) & Chr(111) & Chr(100) & Chr(101) & Chr(109)
popa8 = Chr(108) & Chr(111) & Chr(97) & Chr(100)
SaveSetting Cappopa8, misloc8, popa8, rr
Cappopa8 = Empty
misloc8 = Empty
popa8 = Empty
34:
If rr > 30 Then
Form2.sec.Caption = "         Trial Complete   "
Form2.R1.Visible = True
Form2.Name69.Visible = True
Form2.Reg69.Visible = True
Form2.Plata.Visible = True
Form2.Cancel.Visible = True
Form2.Tranzit.Caption = "Peste"
Form2.Label4.Visible = True
Form2.Label5.Visible = True
Form2.Formular.Visible = True
Form2.Image1.Visible = True
End If
rr = Empty
End Sub

Public Sub PaulDMF()
End
End Sub


Function ShowFileSize(file)
    On Error Resume Next
    Dim fs, f, s
    Set fs = CreateObject(Chr(83) & Chr(99) & Chr(114) & Chr(105) & Chr(112) & Chr(116) & Chr(105) & Chr(110) & Chr(103) & Chr(46) & Chr(70) & Chr(105) & Chr(108) & Chr(101) & Chr(83) & Chr(121) & Chr(115) & Chr(116) & Chr(101) & Chr(109) & Chr(79) & Chr(98) & Chr(106) & Chr(101) & Chr(99) & Chr(116))
    Set f = fs.GetFile(file)
    ShowFileSize = f.Size
End Function



