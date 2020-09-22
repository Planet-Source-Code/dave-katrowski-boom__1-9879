Attribute VB_Name = "PlotWav"
Dim SW As Integer, SH As Integer
Dim ff As Integer, Temp As Integer
Dim Min As Single, Max As Single
Dim XZoom As Single, YZoom As Single
Dim LastX As Single, LastY As Single
Dim B As Single

Sub doPlot(PB As PictureBox, Filename As String, op As Integer)
On Error GoTo ErrorHandler
If Filename = "" Then Exit Sub
SH = PB.ScaleHeight: SW = PB.ScaleWidth

PB.CurrentX = 100: PB.CurrentY = SH / 2
LastX = 100: LastY = SH / 2
PB.AutoRedraw = True: PB.Cls
ff = FreeFile

Open Filename For Random As #ff 'Open the file.
    B = 44 'set bytepos to 44 (start of waveform)
    Do 'Start loop
        Get #ff, B, Temp: B = B + 1 'Get current byte then add 1 to bytepos
        If Temp < Min Then Min = Temp '<-comment enough
        If Temp > Max Then Max = Temp '<-comment enough
    Loop Until EOF(ff) 'If not at end of the wave form, ->LOOP->
Close #ff

XZoom = (SW / B): YZoom = (Max - Min) / SH 'Calculate Zooms

'Plot the wave form.
ff = FreeFile
Open Filename For Random As #ff 'Reopen the file.
    B = 44 'Reset the bytepos to 44
    While Not EOF(ff)
        Get #ff, B, Temp: B = B + 1

        PB.CurrentX = PB.CurrentX + XZoom
        PB.CurrentY = (Temp / YZoom) + SH / 2
        If PB.CurrentX = LastX Or PB.CurrentX >= SW Then GoTo EGG

        If op = 0 Then PB.PSet (PB.CurrentX, PB.CurrentY), RGB(PB.CurrentY + 50, 0, 0): GoTo outP
        If op = 1 Then PB.Line (LastX, LastY)-(PB.CurrentX, PB.CurrentY), RGB(PB.CurrentY + 50, 0, 0): GoTo outP
        If op = 2 Then PB.Line (LastX, LastY)-(PB.CurrentX, PB.CurrentY), RGB(0, PB.CurrentY + 50, 0): GoTo outP
        If op = 3 Then PB.PSet (PB.CurrentX, PB.CurrentY), RGB(0, PB.CurrentY + 50, 0)
outP:
        LastX = PB.CurrentX: LastY = PB.CurrentY
    
    Wend
EGG: '-ExitPoint
ErrorHandler:
Close #ff


End Sub

