Attribute VB_Name = "FileSystem"
Public Type Pattern
totalmes As Integer
tempo As Integer
OnOff(1 To 5) As Integer
Reverb(1 To 5) As Integer
Detune(1 To 5) As Integer
Chaos(1 To 5) As Integer
delay(1 To 5) As Integer
idis(1 To 5) As Integer
irev(1 To 5) As Integer
Pitch(1 To 5) As Long
Volume(1 To 5) As Long
delayint(1 To 5) As Integer
ReverbInt(1 To 5) As Integer
End Type

Public Type SEQu
Sequence(5, 16) As String * 16
End Type

Public Type ln
Length(1 To 5) As Long
End Type

Public W1BYTE() As Byte, W2BYTE() As Byte, W3BYTE() As Byte, W4BYTE() As Byte, W5BYTE() As Byte, P As Pattern, L As ln, s As SEQu, lI As Integer, cm As String

Sub FileOUT(Frm As Form, Filename As String): Dim ff As Integer: ff = FreeFile
On Error Resume Next

'Fill Header & Pattern Structures
P.totalmes = tmes: P.tempo = Frm.Slider4.Value
For c = 1 To 5: For lI = 1 To 16
s.Sequence(c, lI) = SEQ(c, lI)
Next
P.OnOff(c) = CInt(cON(c))
P.Reverb(c) = CInt(Frm.Check1(c).Value)
P.Detune(c) = CInt(Frm.Check2(c).Value)
P.Chaos(c) = CInt(Frm.Check3(c).Value)
P.delay(c) = CInt(Frm.Check4(c).Value)
P.idis(c) = CInt(Frm.Check5(c).Value)
P.irev(c) = CInt(Frm.Check6(c).Value)
P.Pitch(c) = CLng(Frm.Slider1(c).Value)
P.ReverbInt(c) = CInt(Frm.Slider2(c).Value)
P.Volume(c) = CLng(Frm.Slider3(c).Value)
P.delayint(c) = CInt(Frm.Slider5(c).Value)
Next

'Fill Wave Byte Arrays
ReDim W1BYTE(FileLen(WavePaths(1)))
Open WavePaths(1) For Binary As #ff
Get #ff, , W1BYTE()
Close #ff
ReDim W2BYTE(FileLen(WavePaths(2)))
Open WavePaths(2) For Binary As #ff
Get #ff, , W2BYTE()
Close #ff
ReDim W3BYTE(FileLen(WavePaths(3)))
Open WavePaths(3) For Binary As #ff
Get #ff, , W3BYTE()
Close #ff
ReDim W4BYTE(FileLen(WavePaths(4)))
Open WavePaths(4) For Binary As #ff
Get #ff, , W4BYTE()
Close #ff
ReDim W5BYTE(FileLen(WavePaths(5)))
Open WavePaths(5) For Binary As #ff
Get #ff, , W5BYTE()
Close #ff

For I = 1 To 5
L.Length(I) = FileLen(WavePaths(I))
Next

Kill Filename
Open Filename For Binary Access Write As #ff
Put #ff, , P
Put #ff, LenB(P), s
Put #ff, LenB(P) + LenB(s), L

Put #ff, LenB(P) + LenB(L) + LenB(s), W1BYTE()
Put #ff, LenB(P) + LenB(L) + LenB(s) + UBound(W1BYTE), W2BYTE()
Put #ff, LenB(P) + LenB(L) + LenB(s) + UBound(W1BYTE) + UBound(W2BYTE), W3BYTE()
Put #ff, LenB(P) + LenB(L) + LenB(s) + UBound(W1BYTE) + UBound(W2BYTE) + UBound(W3BYTE), W4BYTE()
Put #ff, LenB(P) + LenB(L) + LenB(s) + UBound(W1BYTE) + UBound(W2BYTE) + UBound(W3BYTE) + UBound(W4BYTE), W5BYTE()

Close #ff

End Sub

Sub FileIN(Frm As Form, Filename As String)
Dim cc As Integer, ff As Integer: ff = FreeFile
'On Error Resume Next

'Fill Structures with FileData
Open Filename For Binary Access Read As #ff
Get #ff, , P
Get #ff, LenB(P), s
Get #ff, LenB(P) + LenB(s), L

ReDim W1BYTE(L.Length(1))
ReDim W2BYTE(L.Length(2))
ReDim W3BYTE(L.Length(3))
ReDim W4BYTE(L.Length(4))
ReDim W5BYTE(L.Length(5))

Get #ff, LenB(P) + LenB(L) + LenB(s), W1BYTE()
Get #ff, LenB(P) + LenB(L) + LenB(s) + UBound(W1BYTE), W2BYTE()
Get #ff, LenB(P) + LenB(L) + LenB(s) + UBound(W1BYTE) + UBound(W2BYTE), W3BYTE()
Get #ff, LenB(P) + LenB(L) + LenB(s) + UBound(W1BYTE) + UBound(W2BYTE) + UBound(W3BYTE), W4BYTE()
Get #ff, LenB(P) + LenB(L) + LenB(s) + UBound(W1BYTE) + UBound(W2BYTE) + UBound(W3BYTE) + UBound(W4BYTE), W5BYTE()

Close #ff

'Set Tempo & Total Meassures
Frm.Slider4.Value = P.tempo: tmes = P.totalmes: Frm.Label9 = tmes

'Set other Switches & Sliders
For cc = 1 To 5
cON(cc) = CBool(P.OnOff(cc))
Frm.Check1(cc).Value = P.Reverb(cc)
Frm.Check2(cc).Value = P.Detune(cc)
Frm.Check3(cc).Value = P.Chaos(cc)
Frm.Check4(cc).Value = P.delay(cc)
Frm.Check5(cc).Value = P.idis(cc)
Frm.Check6(cc).Value = P.irev(cc)
Frm.Slider1(cc).Value = P.Pitch(cc)
Frm.Slider2(cc).Value = P.ReverbInt(cc)
Frm.Slider3(cc).Value = P.Volume(cc)
Frm.Slider5(cc).Value = P.delayint(cc)
Next

'Display Beat Pattern
For cc = 1 To 5: For lI = 1 To 16
SEQ(cc, lI) = s.Sequence(cc, lI)
Next
Select Case P.OnOff(cc)
Case False
Frm.Command3(cc).Caption = "Off"
Case True
Frm.Command3(cc).Caption = "On"
End Select
DS.SetVolume cc, P.Volume(cc)
Next
Frm.beatdisplay1.inDis SEQ(1, 1)
Frm.beatdisplay2.inDis SEQ(2, 1)
Frm.beatdisplay3.inDis SEQ(3, 1)
Frm.beatdisplay4.inDis SEQ(4, 1)
Frm.beatdisplay5.inDis SEQ(5, 1)

WavePaths(1) = App.Path & "\temp1.wav"
WavePaths(2) = App.Path & "\temp2.wav"
WavePaths(3) = App.Path & "\temp3.wav"
WavePaths(4) = App.Path & "\temp4.wav"
WavePaths(5) = App.Path & "\temp5.wav"

ReDim WBYTE(0), EWBYTE(0), DRB(0)

For cc = 1 To 5
Frm.Picture5(cc).BackColor = &H40&
'Kill App.Path & "\temp1.wav"
Open WavePaths(cc) For Binary As #ff
Select Case cc
Case 1
Put #ff, , W1BYTE()
Case 2
Put #ff, , W2BYTE()
Case 3
Put #ff, , W3BYTE()
Case 4
Put #ff, , W4BYTE()
Case 5
Put #ff, , W5BYTE()
End Select
Close #ff
Distort cc
Reverse cc
DisRev cc
Frm.Picture5(cc).BackColor = &H8000&
doPlot Frm.Picture5(cc), WavePaths(cc), (Rnd * 1) + 2
Next

DS.LoadWavToChannel 1, App.Path & "\temp1.wav"
DS.LoadWavToChannel 2, App.Path & "\temp2.wav"
DS.LoadWavToChannel 3, App.Path & "\temp3.wav"
DS.LoadWavToChannel 4, App.Path & "\temp4.wav"
DS.LoadWavToChannel 5, App.Path & "\temp5.wav"

DS.LoadWavToChannel 6, App.Path & "\tempdis1.wav"
DS.LoadWavToChannel 7, App.Path & "\tempdis2.wav"
DS.LoadWavToChannel 8, App.Path & "\tempdis3.wav"
DS.LoadWavToChannel 9, App.Path & "\tempdis4.wav"
DS.LoadWavToChannel 10, App.Path & "\tempdis5.wav"

DS.LoadWavToChannel 11, App.Path & "\temprev1.wav"
DS.LoadWavToChannel 12, App.Path & "\temprev2.wav"
DS.LoadWavToChannel 13, App.Path & "\temprev3.wav"
DS.LoadWavToChannel 14, App.Path & "\temprev4.wav"
DS.LoadWavToChannel 15, App.Path & "\temprev5.wav"

DS.LoadWavToChannel 16, App.Path & "\tdisrev1.wav"
DS.LoadWavToChannel 17, App.Path & "\tdisrev2.wav"
DS.LoadWavToChannel 18, App.Path & "\tdisrev3.wav"
DS.LoadWavToChannel 19, App.Path & "\tdisrev4.wav"
DS.LoadWavToChannel 20, App.Path & "\tdisrev5.wav"

End Sub
