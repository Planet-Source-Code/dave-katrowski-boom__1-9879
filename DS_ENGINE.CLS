VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DS_Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------=
'DirectSound Class for 2D Games=
'===============================
'Author: Dave Katrowski        =
'------------------------------=
Private WithEvents Echo1 As ccrpTimer
Attribute Echo1.VB_VarHelpID = -1
Private WithEvents Echo2 As ccrpTimer
Attribute Echo2.VB_VarHelpID = -1
Private WithEvents Echo3 As ccrpTimer
Attribute Echo3.VB_VarHelpID = -1
Private WithEvents Echo4 As ccrpTimer
Attribute Echo4.VB_VarHelpID = -1
Private WithEvents Echo5 As ccrpTimer
Attribute Echo5.VB_VarHelpID = -1

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const MAX_Channels As Integer = 20
Private Const Max_EchoSteps As Integer = 2

Private DX7 As New DirectX7, DS7 As DirectSound
Private BufDesc As DSBUFFERDESC, PCM As WAVEFORMATEX

Private DSB(1 To MAX_Channels) As DirectSoundBuffer
Private Echo_DSB(1 To MAX_Channels, 1 To Max_EchoSteps) As DirectSoundBuffer, dstep(1 To MAX_Channels) As Integer, sTeps As Integer, cn As Integer

Private BufferState(1 To MAX_Channels) As Boolean, st As Long
Public Sub LoadWavToChannel(ChannelNumber As Integer, Filename As String)
On Error Resume Next
Set DSB(ChannelNumber) = Nothing
On Local Error GoTo BadWaveFile_Exit
Set DSB(ChannelNumber) = DS7.CreateSoundBufferFromFile(Filename, BufDesc, PCM)
Let BufferState(ChannelNumber) = True: Exit Sub
BadWaveFile_Exit: Set DSB(ChannelNumber) = Nothing
End Sub
Public Sub SetFrequency(ChannelNumber As Integer, Value As Long): On Error Resume Next
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
Call DSB(ChannelNumber).SetFrequency(Value)
End Sub
Public Function GetFrequency(ChannelNumber As Integer) As Long: On Error Resume Next
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Function
Let GetFrequency = DSB(ChannelNumber).GetFrequency
End Function
Public Sub SetVolume(ChannelNumber As Integer, Value As Long): On Error Resume Next
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
 DSB(ChannelNumber).SetVolume Value
End Sub
Public Function GetVolume(ChannelNumber As Integer) As Long
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Function
Let GetVolume = DSB(ChannelNumber).GetVolume
End Function
Public Sub SetPan(ChannelNumber As Integer, Value As Long)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
DSB(ChannelNumber).SetPan Value
End Sub
Public Function GetPan(ChannelNumber As Integer) As Long
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Function
Let GetPan = DSB(ChannelNumber).GetPan
End Function
Public Sub PlaySound(ChannelNumber As Integer, Optional Looped As Boolean)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
On Error Resume Next
If Looped Then GoTo lloop
    DSB(ChannelNumber).SetCurrentPosition 0: DSB(ChannelNumber).PlaY DSBPLAY_DEFAULT: Exit Sub
lloop:
    DSB(ChannelNumber).SetCurrentPosition 0: DSB(ChannelNumber).PlaY DSBPLAY_LOOPING
End Sub
Public Sub PlayEcho(ChannelNumber As Integer, step As Integer, Interval As Integer)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
If step > Max_EchoSteps Then Let step = Max_EchoSteps
If Interval = 0 Then Let Interval = 50

dstep(ChannelNumber) = 0: sTeps = step: cn = ChannelNumber

Select Case cn
Case 1, 6, 11, 16
Echo1.Interval = Interval
Echo1.Enabled = True
Case 2, 7, 12, 17
Echo2.Interval = Interval
Echo2.Enabled = True
Case 3, 8, 13, 18
Echo3.Interval = Interval
Echo3.Enabled = True
Case 4, 9, 14, 19
Echo4.Interval = Interval
Echo4.Enabled = True
Case 5, 10, 15, 20
Echo5.Interval = Interval
Echo5.Enabled = True
End Select
End Sub
Private Sub Echo1_Timer(ByVal Milliseconds As Long)
dstep(cn) = dstep(cn) + 1
If dstep(cn) > sTeps Then Echo1.Enabled = False: Exit Sub
Set Echo_DSB(cn, dstep(cn)) = DS7.DuplicateSoundBuffer(DSB(cn))
Echo_DSB(cn, dstep(cn)).PlaY DSBPLAY_DEFAULT
Echo1.Interval = Echo1.Interval * dstep(cn)
End Sub
Private Sub Echo2_Timer(ByVal Milliseconds As Long)
dstep(cn) = dstep(cn) + 1
If dstep(cn) > sTeps Then Echo2.Enabled = False: Exit Sub
Set Echo_DSB(cn, dstep(cn)) = DS7.DuplicateSoundBuffer(DSB(cn))
Echo_DSB(cn, dstep(cn)).PlaY DSBPLAY_DEFAULT
Echo2.Interval = Echo2.Interval * dstep(cn)
End Sub
Private Sub Echo3_Timer(ByVal Milliseconds As Long)
dstep(cn) = dstep(cn) + 1
If dstep(cn) > sTeps Then Echo3.Enabled = False: Exit Sub
Set Echo_DSB(cn, dstep(cn)) = DS7.DuplicateSoundBuffer(DSB(cn))
Echo_DSB(cn, dstep(cn)).PlaY DSBPLAY_DEFAULT
Echo3.Interval = Echo3.Interval * dstep(cn)
End Sub
Private Sub Echo4_Timer(ByVal Milliseconds As Long)
dstep(cn) = dstep(cn) + 1
If dstep(cn) > sTeps Then Echo4.Enabled = False: Exit Sub
Set Echo_DSB(cn, dstep(cn)) = DS7.DuplicateSoundBuffer(DSB(cn))
Echo_DSB(cn, dstep(cn)).PlaY DSBPLAY_DEFAULT
Echo4.Interval = Echo4.Interval * dstep(cn)
End Sub
Private Sub Echo5_Timer(ByVal Milliseconds As Long)
dstep(cn) = dstep(cn) + 1
If dstep(cn) > sTeps Then Echo5.Enabled = False: Exit Sub
Set Echo_DSB(cn, dstep(cn)) = DS7.DuplicateSoundBuffer(DSB(cn))
Echo_DSB(cn, dstep(cn)).PlaY DSBPLAY_DEFAULT
Echo5.Interval = Echo5.Interval * dstep(cn)
End Sub
Public Sub StopSound(ChannelNumber As Integer): On Error Resume Next
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
DSB(ChannelNumber).Stop
End Sub
Public Sub ClearBuffer(ChannelNumber As Integer)
If ChannelNumber > MAX_Channels Or Not BufferState(ChannelNumber) Then Exit Sub
Set DSB(ChannelNumber) = Nothing
Let BufferState(ChannelNumber) = False
End Sub
Public Sub Initialize_Engine(Hwnd As Long)
Set DS7 = DX7.DirectSoundCreate("")
DS7.SetCooperativeLevel Hwnd, DSSCL_EXCLUSIVE

    With PCM
        .nSize = LenB(PCM)
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 2
        .lSamplesPerSec = 44100
        .nBitsPerSample = 16
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With

BufDesc.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_STATIC

Set Echo1 = New ccrpTimer: Set Echo2 = New ccrpTimer
Set Echo3 = New ccrpTimer: Set Echo4 = New ccrpTimer
Set Echo5 = New ccrpTimer
End Sub
Public Sub Terminate_Engine()

    Static I As Integer
    For I = 1 To MAX_Channels
        Set DSB(I) = Nothing
        For ii = 1 To Max_EchoSteps
            Set Echo_DSB(I, ii) = Nothing
        Next
    Next

Set DS7 = Nothing
Set DX7 = Nothing

Echo1.Enabled = False: Set Echo1 = Nothing
Echo2.Enabled = False: Set Echo2 = Nothing
Echo3.Enabled = False: Set Echo3 = Nothing
Echo4.Enabled = False: Set Echo4 = Nothing
Echo5.Enabled = False: Set Echo5 = Nothing
End Sub
Public Function GetReadableVolume(Volume As Long) As String

    Select Case Volume
        Case Is >= -10000: GetReadableVolume = "0"
        Case Is >= -3000: GetReadableVolume = "1"
        Case Is >= -2000: GetReadableVolume = "2"
        Case Is >= -1750: GetReadableVolume = "3"
        Case Is >= -1500: GetReadableVolume = "4"
        Case Is >= -1250: GetReadableVolume = "5"
        Case Is >= -1000: GetReadableVolume = "6"
        Case Is >= -750: GetReadableVolume = "7"
        Case Is >= -500: GetReadableVolume = "8"
        Case Is >= -250: GetReadableVolume = "9"
        Case 0: GetReadableVolume = "10"
    End Select

End Function

Public Function GetSpeakerConfig() As String

    Select Case DS7.GetSpeakerConfig
        Case DSSPEAKER_HEADPHONE
            GetSpeakerConfig = "Speakers are set to: [Headphone]"
        Case DSSPEAKER_MONO
            GetSpeakerConfig = "Speakers are set to: [Mono]"
        Case DSSPEAKER_QUAD
            GetSpeakerConfig = "Speakers are set to: [Quadraphonic]"
        Case DSSPEAKER_SURROUND
            GetSpeakerConfig = "Speakers are set to: [Surround]"
        Case DSSPEAKER_STEREO
            GetSpeakerConfig = "Speakers are set to: [Stereo]"
     End Select
    
End Function

Public Sub SetSpeakerConfig(Command As String)
    
    Select Case LCase(Command)
        Case Is = "headphone"
            DS7.SetSpeakerConfig DSSPEAKER_HEADPHONE
        Case "stereo"
            DS7.SetSpeakerConfig DSSPEAKER_STEREO
        Case "mono"
            DS7.SetSpeakerConfig DSSPEAKER_MONO
        Case "surround"
            DS7.SetSpeakerConfig DSSPEAKER_SURROUND
        Case "quad"
            DS7.SetSpeakerConfig DSSPEAKER_QUAD
        Case Else
            DS7.SetSpeakerConfig DSSPEAKER_STEREO
    End Select

End Sub

Public Sub Pause(ms As Long)
Sleep ms
End Sub

