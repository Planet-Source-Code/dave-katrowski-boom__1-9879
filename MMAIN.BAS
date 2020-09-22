Attribute VB_Name = "modMain"
Declare Function timeGetTime Lib "winmm.dll" () As Long

Public DS As New DS_Engine

Public cON(1 To 5) As Boolean

Public WavePaths(1 To 5) As String, SEQ(1 To 5, 1 To 16) As String, lastplot As String

Public curmes As Integer, tick As Integer, tmes As Integer, mes As Integer, c As Integer, m As Integer, I As Integer, st As Long, plotState As Integer
Sub Main()
Load SplashScreen: SplashScreen.Show: SplashScreen.Refresh
DS.Pause 100
Load fMain
SplashScreen.Hide: fMain.Show: fMain.Refresh
End Sub

Function BpmToInterval(Bpm As Integer)
If Not Bpm = 0 Then Let BpmToInterval = (1000 / (Bpm / 60)) / 4
End Function

Sub center(obj As Object)
obj.Top = (Screen.Height - obj.Height) / 2
obj.Left = (Screen.Width - obj.Width) / 2
End Sub

 
Public Function RndRange(ByVal Min As Long, ByVal Max As Long) As Long
Randomize
RndRange = (Rnd * (Max - Min + 1)) + Min
End Function
Public Function GetCounterBeat(nBN As Integer) As Integer

    Select Case nBN
        Case Is < 5: GetCounterBeat = 1
        Case 5, 6, 7, 8: GetCounterBeat = 2
        Case 9, 10, 11, 12: GetCounterBeat = 3
        Case Is >= 13: GetCounterBeat = 4
    End Select
    
End Function
