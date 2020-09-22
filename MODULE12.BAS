Attribute VB_Name = "modWavEffects"
Type WAVEHDR
    lpData As String
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type
Type WAVEFORMAT
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
End Type

Public wf As WAVEFORMAT, wh As WAVEHDR
Public WBYTE() As Byte, EWBYTE() As Byte, DRB() As Byte, lng As Long



Sub Distort(ChannelNum As Integer): On Error Resume Next
Dim File1 As String, File2 As String, cb As Integer, iB As Long
Dim ff1 As Integer: ff1 = FreeFile
File1 = App.Path & "\tempdis" & ChannelNum & ".wav"
File2 = App.Path & "\temp" & ChannelNum & ".wav"
Kill File1
Open File2 For Binary Access Read As #ff1
Dim ff2 As Integer: ff2 = FreeFile
Open File1 For Binary Access Write As #ff2

Get #ff1, , wh
Get #ff1, , wf

lng = LOF(ff1)
ReDim WBYTE(lng), EWBYTE(lng)

Get #ff1, Len(wh) + Len(wf), WBYTE()

For iB = 0 To UBound(WBYTE())
cb = WBYTE(iB)
cb = cb + (0.3 * (cb + 100))
If cb < 255 Then
EWBYTE(iB) = cb
Else
EWBYTE(iB) = WBYTE(iB)
End If
Next

Put #ff2, , wh
Put #ff2, , wf
Put #ff2, Len(wh) + Len(wf), EWBYTE()
Close #ff1
Close #ff2
Exit Sub
End Sub

Sub Reverse(ChannelNum As Integer): On Error Resume Next
Dim File1 As String, File2 As String, ub As Long, iR As Long, iB As Long
Dim ff1 As Integer: ff1 = FreeFile
File1 = App.Path & "\temprev" & ChannelNum & ".wav"
File2 = App.Path & "\temp" & ChannelNum & ".wav"
Kill File1
Open File2 For Binary Access Read As #ff1
Dim ff2 As Integer: ff2 = FreeFile
Open File1 For Binary Access Write As #ff2

Get #ff1, , wh
Get #ff1, , wf

lng = LOF(ff1)
ReDim WBYTE(lng), EWBYTE(lng)

Get #ff1, , WBYTE() 'Len(wh) + Len(wf), WBYTE()

iR = 0
For iB = UBound(WBYTE()) To 0 Step -1
EWBYTE(iR) = WBYTE(iB)
iR = iR + 1
Next

Put #ff2, , wh
Put #ff2, , wf
Put #ff2, Len(wh) + Len(wf), EWBYTE()
Close #ff1
Close #ff2
End Sub

Sub DisRev(ChannelNum As Integer): On Error Resume Next
Dim File1 As String, File2 As String, ub As Long, iR As Long, cb As Integer, iB As Long
Dim ff1 As Integer: ff1 = FreeFile
File1 = App.Path & "\tdisrev" & ChannelNum & ".wav"
File2 = App.Path & "\temp" & ChannelNum & ".wav"
Kill File1
Open File2 For Binary Access Read As #ff1
Dim ff2 As Integer: ff2 = FreeFile
Open File1 For Binary Access Write As #ff2

Get #ff1, , wh
Get #ff1, , wf

lng = LOF(ff1)
ReDim WBYTE(lng), EWBYTE(lng), DRB(lng)

Get #ff1, Len(wh) + Len(wf), WBYTE()

For iB = 0 To UBound(WBYTE())
cb = WBYTE(iB)
cb = cb + (0.3 * (cb + 100))
If cb < 255 Then
EWBYTE(iB) = cb
Else
EWBYTE(iB) = WBYTE(iB)
End If
Next
iR = 0
For iB = UBound(WBYTE()) To 0 Step -1
DRB(iR) = EWBYTE(iB)
iR = iR + 1
Next

Put #ff2, , wh
Put #ff2, , wf
Put #ff2, Len(wh) + Len(wf), DRB()
Close #ff1
Close #ff2
End Sub

