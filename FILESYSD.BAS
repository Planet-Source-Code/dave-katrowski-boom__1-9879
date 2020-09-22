Attribute VB_Name = "modFilesystem"
Public Type Head
totalmes As Integer
tempo As Integer
End Type
Public Type Pattern
Pitch As Long
Volume As Long
Reverb As Integer
ReverbInt As Integer
Detune As Integer
OnOff As Integer
Sequence(1 To 48) As String * 16
Length As Long
End Type

Public W1BYTE() As Byte, W2BYTE() As Byte, W3BYTE() As Byte, W4BYTE() As Byte, W5BYTE() As Byte, h As Head, p(1 To 5) As Pattern

Sub FileOUT(Frm As Form, Filename As String): Static c As Integer, n(1 To 5) As Long, ff As Integer: ff = FreeFile
On Error Resume Next

'Fill Header & Pattern Structures
h.tempo = Frm.Slider4.Value: h.totalmes = tmes

For c = 1 To 5

    For I = 1 To 48
        p(c).Sequence(I) = SEQ(c, I)
    Next
    
p(c).Detune = Frm.Check2(c).Value
p(c).OnOff = CInt(cON(c))
p(c).Reverb = Frm.Check1(c).Value
p(c).Pitch = Frm.Slider1(c).Value
p(c).ReverbInt = Frm.Slider2(c).Value
p(c).Volume = Frm.Slider3(c).Value
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
p(I).Length = FileLen(WavePaths(I))
Next

Kill Filename
Open Filename For Binary Access Write As #ff
Put #ff, , h
For I = 1 To 5
Put #ff, , p(I)
Next
'Put #ff, Len(p), W1BYTE()
'Put #ff, Len(p) + UBound(W1BYTE), W2BYTE()
'Put #ff, Len(p) + UBound(W1BYTE) + UBound(W2BYTE), W3BYTE()
'Put #ff, Len(p) + UBound(W1BYTE) + UBound(W2BYTE) + UBound(W3BYTE), W4BYTE()
'Put #ff, Len(p) + UBound(W1BYTE) + UBound(W2BYTE) + UBound(W3BYTE) + UBound(W4BYTE), W5BYTE()
Close #ff

End Sub

Sub FileIN(Frm As Form, Filename As String): Static c As Integer, ff As Integer: ff = FreeFile
On Error Resume Next

'Fill Structures with FileData
Open Filename For Binary Access Read As #ff
Get #ff, , h
Get #ff, Len(h), p(1)
Get #ff, Len(h) + Len(p(1)), p(2)
Get #ff, Len(h) + Len(p(1)) + Len(p(2)), p(3)
Get #ff, Len(h) + Len(p(1)) + Len(p(2)) + Len(p(3)), p(4)
Get #ff, Len(h) + Len(p(1)) + Len(p(2)) + Len(p(3)) + Len(p(4)), p(5)
'ReDim W1BYTE(p.Length(1))
'ReDim W2BYTE(p.Length(2))
'ReDim W3BYTE(p.Length(3))
'ReDim W4BYTE(p.Length(4))
'ReDim W5BYTE(p.Length(5))

'Get #ff, Len(p), W1BYTE()
'Get #ff, Len(p) + UBound(W1BYTE), W2BYTE()
'Get #ff, Len(p) + UBound(W1BYTE) + UBound(W2BYTE), W3BYTE()
'Get #ff, Len(p) + UBound(W1BYTE) + UBound(W2BYTE) + UBound(W3BYTE), W4BYTE()
'Get #ff, Len(p) + UBound(W1BYTE) + UBound(W2BYTE) + UBound(W3BYTE) + UBound(W4BYTE), W5BYTE()
Close #ff

'Set Tempo & Total Meassures
Frm.Slider4.Value = h.tempo: tmes = h.totalmes: Frm.Label9 = tmes

'Set other Switches & Sliders
For c = 1 To 5
Frm.Check2(c).Value = p(c).Detune
cON(c) = p(c).OnOff
Frm.Check1(c).Value = p(c).Reverb
Frm.Slider1(c).Value = p(c).Pitch
Frm.Slider2(c).Value = p(c).ReverbInt
Frm.Slider3(c).Value = p(c).Volume
Next

'Create TempWaves from FileData/
'Load them to DS/ then kill them.
'Kill App.Path & "\temp.wav"
'Open App.Path & "\temp.wav" For Binary As #ff
'Put #ff, , W1BYTE()
'Close #ff
'doPlot Frm.Picture5(1), App.Path & "\temp.wav"
'DS.LoadWavToChannel 1, App.Path & "\temp.wav"
'Kill App.Path & "\temp.wav"
'Open App.Path & "\temp.wav" For Binary As #ff
'Put #ff, , W2BYTE()
'Close #ff
'doPlot Frm.Picture5(2), App.Path & "\temp.wav"
'DS.LoadWavToChannel 2, App.Path & "\temp.wav"
'Kill App.Path & "\temp.wav"
'Open App.Path & "\temp.wav" For Binary As #ff
'Put #ff, , W3BYTE()
'Close #ff
'doPlot Frm.Picture5(3), App.Path & "\temp.wav"
'DS.LoadWavToChannel 3, App.Path & "\temp.wav"
'Kill App.Path & "\temp.wav"
'Open App.Path & "\temp.wav" For Binary As #ff
'Put #ff, , W4BYTE()
'Close #ff
'doPlot Frm.Picture5(4), App.Path & "\temp.wav"
'DS.LoadWavToChannel 4, App.Path & "\temp.wav"
'Kill App.Path & "\temp.wav"
'Open App.Path & "\temp.wav" For Binary As #ff
'Put #ff, , W5BYTE()
'Close #ff
'doPlot Frm.Picture5(5), App.Path & "\temp.wav"
'DS.LoadWavToChannel 5, App.Path & "\temp.wav"
'Kill App.Path & "\temp.wav"

'Display Beat Pattern
For c = 1 To 5: For I = 1 To 48
SEQ(c, I) = p(c).Sequence(I)
Next: Next
Frm.beatdisplay1.inDis SEQ(1, 1)
Frm.beatdisplay2.inDis SEQ(2, 1)
Frm.beatdisplay3.inDis SEQ(3, 1)
Frm.beatdisplay4.inDis SEQ(4, 1)
Frm.beatdisplay5.inDis SEQ(5, 1)
End Sub
