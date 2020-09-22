Attribute VB_Name = "modLED"
'Class: DmkWare GFX
'Title: LED Engine.
'Notes: Needs an Array of 10 PictureBoxes (LED()), and A PictureBox
'       for viewing (LEDScreen). See Example Form

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY As Long = &HCC0020

Public Sub DrawLED(DestPB As Object, Value As String, GFXPBARRAY As Object, Align As Integer)
Static CurChar As String, Ip As Integer, X As Integer, Y As Integer, ValLen As Integer
If Not DestPB.ScaleMode = 3 Then DestPB.ScaleMode = 3
If Not DestPB.AutoRedraw = True Then DestPB.AutoRedraw = True
DestPB.Cls: ValLen = Len(Value)
Select Case Align
Case 0
    Value = SetupString(Value, DestPB.ScaleWidth): ValLen = Len(Value)
    Y = (DestPB.ScaleHeight / 2) - 8
    X = ((DestPB.ScaleWidth / 2) - ((16 * ValLen / 2))) - 16
    For Ip = 1 To ValLen
        CurChar = Mid(Value, Ip, 1)
        BitBlt DestPB.hDC, X + (Ip * 16), Y, 16, 16, GFXPBARRAY(CInt(CurChar)).hDC, 0, 0, SRCCOPY
    Next
Case 1
    Y = (DestPB.ScaleHeight / 2) - 8
    X = ((DestPB.ScaleWidth / 2) - ((16 * ValLen / 2))) - 16
    For Ip = 1 To ValLen
        CurChar = Mid(Value, Ip, 1)
        BitBlt DestPB.hDC, X + (Ip * 16), Y, 16, 16, GFXPBARRAY(CInt(CurChar)).hDC, 0, 0, SRCCOPY
    Next
End Select
End Sub

Public Sub LoadGFX(GFXPBARRAY As Object, Optional CustomColor As Boolean, Optional BackColor As Long, Optional DiodeColor As Long, Optional LitDiodeColor As Long)
Dim CurChar As String, Str1to16(1 To 16) As String, R As Byte, cr As Byte, Number As Byte

For Number = 0 To 9
Select Case Number

Case 0
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00$$$$$$$$$$$$00"
 Str1to16(3) = "0$0$$$$$$$$$$0$0"
 Str1to16(4) = "0$$0000000000$$0"
 Str1to16(5) = "0$$0000000000$$0"
 Str1to16(6) = "0$$0000000000$$0"
 Str1to16(7) = "0$$0000000000$$0"
 Str1to16(8) = "0$0##########0$0"
 Str1to16(9) = "00############00"
Str1to16(10) = "0$000000000000$0"
Str1to16(11) = "0$$0000000000$$0"
Str1to16(12) = "0$$0000000000$$0"
Str1to16(13) = "0$$0000000000$$0"
Str1to16(14) = "0$0$$$$$$$$$$0$0"
Str1to16(15) = "00$$$$$$$$$$$$00"
Str1to16(16) = "0000000000000000"
 
Case 1
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00############00"
 Str1to16(3) = "0#0##########0$0"
 Str1to16(4) = "0##0000000000$$0"
 Str1to16(5) = "0##0000000000$$0"
 Str1to16(6) = "0##0000000000$$0"
 Str1to16(7) = "0##0000000000$$0"
 Str1to16(8) = "0#0##########0$0"
 Str1to16(9) = "00############00"
Str1to16(10) = "0#000000000000$0"
Str1to16(11) = "0##0000000000$$0"
Str1to16(12) = "0##0000000000$$0"
Str1to16(13) = "0##0000000000$$0"
Str1to16(14) = "0#0##########0$0"
Str1to16(15) = "00############00"
Str1to16(16) = "0000000000000000"

Case 2
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00$$$$$$$$$$$$00"
 Str1to16(3) = "0#0$$$$$$$$$$0$0"
 Str1to16(4) = "0##0000000000$$0"
 Str1to16(5) = "0##0000000000$$0"
 Str1to16(6) = "0##0000000000$$0"
 Str1to16(7) = "0##0000000000$$0"
 Str1to16(8) = "0#0$$$$$$$$$$0$0"
 Str1to16(9) = "00$$$$$$$$$$$$00"
Str1to16(10) = "0$000000000000#0"
Str1to16(11) = "0$$0000000000##0"
Str1to16(12) = "0$$0000000000##0"
Str1to16(13) = "0$$0000000000##0"
Str1to16(14) = "0$0$$$$$$$$$$0#0"
Str1to16(15) = "00$$$$$$$$$$$$00"
Str1to16(16) = "0000000000000000"

Case 3
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00$$$$$$$$$$$$00"
 Str1to16(3) = "0#0$$$$$$$$$$0$0"
 Str1to16(4) = "0##0000000000$$0"
 Str1to16(5) = "0##0000000000$$0"
 Str1to16(6) = "0##0000000000$$0"
 Str1to16(7) = "0##0000000000$$0"
 Str1to16(8) = "0#0$$$$$$$$$$0$0"
 Str1to16(9) = "00$$$$$$$$$$$$00"
Str1to16(10) = "0#000000000000$0"
Str1to16(11) = "0##0000000000$$0"
Str1to16(12) = "0##0000000000$$0"
Str1to16(13) = "0##0000000000$$0"
Str1to16(14) = "0#0$$$$$$$$$$0$0"
Str1to16(15) = "00$$$$$$$$$$$$00"
Str1to16(16) = "0000000000000000"
 
Case 4
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00############00"
 Str1to16(3) = "0$0##########0$0"
 Str1to16(4) = "0$$0000000000$$0"
 Str1to16(5) = "0$$0000000000$$0"
 Str1to16(6) = "0$$0000000000$$0"
 Str1to16(7) = "0$$0000000000$$0"
 Str1to16(8) = "0$0$$$$$$$$$$0$0"
 Str1to16(9) = "00$$$$$$$$$$$$00"
Str1to16(10) = "0#000000000000$0"
Str1to16(11) = "0##0000000000$$0"
Str1to16(12) = "0##0000000000$$0"
Str1to16(13) = "0##0000000000$$0"
Str1to16(14) = "0#0##########0$0"
Str1to16(15) = "00############00"
Str1to16(16) = "0000000000000000"

Case 5
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00$$$$$$$$$$$$00"
 Str1to16(3) = "0$0$$$$$$$$$$0#0"
 Str1to16(4) = "0$$0000000000##0"
 Str1to16(5) = "0$$0000000000##0"
 Str1to16(6) = "0$$0000000000##0"
 Str1to16(7) = "0$$0000000000##0"
 Str1to16(8) = "0$0$$$$$$$$$$0#0"
 Str1to16(9) = "00$$$$$$$$$$$$00"
Str1to16(10) = "0#000000000000$0"
Str1to16(11) = "0##0000000000$$0"
Str1to16(12) = "0##0000000000$$0"
Str1to16(13) = "0##0000000000$$0"
Str1to16(14) = "0#0$$$$$$$$$$0$0"
Str1to16(15) = "00$$$$$$$$$$$$00"
Str1to16(16) = "0000000000000000"

Case 6
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00$$$$$$$$$$$$00"
 Str1to16(3) = "0$0$$$$$$$$$$0#0"
 Str1to16(4) = "0$$0000000000##0"
 Str1to16(5) = "0$$0000000000##0"
 Str1to16(6) = "0$$0000000000##0"
 Str1to16(7) = "0$$0000000000##0"
 Str1to16(8) = "0$0$$$$$$$$$$0#0"
 Str1to16(9) = "00$$$$$$$$$$$$00"
Str1to16(10) = "0$000000000000$0"
Str1to16(11) = "0$$0000000000$$0"
Str1to16(12) = "0$$0000000000$$0"
Str1to16(13) = "0$$0000000000$$0"
Str1to16(14) = "0$0$$$$$$$$$$0$0"
Str1to16(15) = "00$$$$$$$$$$$$00"
Str1to16(16) = "0000000000000000"
 
Case 7
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00$$$$$$$$$$$$00"
 Str1to16(3) = "0#0$$$$$$$$$$0$0"
 Str1to16(4) = "0##0000000000$$0"
 Str1to16(5) = "0##0000000000$$0"
 Str1to16(6) = "0##0000000000$$0"
 Str1to16(7) = "0##0000000000$$0"
 Str1to16(8) = "0#0##########0$0"
 Str1to16(9) = "00############00"
Str1to16(10) = "0#000000000000$0"
Str1to16(11) = "0##0000000000$$0"
Str1to16(12) = "0##0000000000$$0"
Str1to16(13) = "0##0000000000$$0"
Str1to16(14) = "0#0##########0$0"
Str1to16(15) = "00############00"
Str1to16(16) = "0000000000000000"
 
Case 8
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00$$$$$$$$$$$$00"
 Str1to16(3) = "0$0$$$$$$$$$$0$0"
 Str1to16(4) = "0$$0000000000$$0"
 Str1to16(5) = "0$$0000000000$$0"
 Str1to16(6) = "0$$0000000000$$0"
 Str1to16(7) = "0$$0000000000$$0"
 Str1to16(8) = "0$0$$$$$$$$$$0$0"
 Str1to16(9) = "00$$$$$$$$$$$$00"
Str1to16(10) = "0$000000000000$0"
Str1to16(11) = "0$$0000000000$$0"
Str1to16(12) = "0$$0000000000$$0"
Str1to16(13) = "0$$0000000000$$0"
Str1to16(14) = "0$0$$$$$$$$$$0$0"
Str1to16(15) = "00$$$$$$$$$$$$00"
Str1to16(16) = "0000000000000000"

Case 9
 Str1to16(1) = "0000000000000000"
 Str1to16(2) = "00$$$$$$$$$$$$00"
 Str1to16(3) = "0$0$$$$$$$$$$0$0"
 Str1to16(4) = "0$$0000000000$$0"
 Str1to16(5) = "0$$0000000000$$0"
 Str1to16(6) = "0$$0000000000$$0"
 Str1to16(7) = "0$$0000000000$$0"
 Str1to16(8) = "0$0$$$$$$$$$$0$0"
 Str1to16(9) = "00$$$$$$$$$$$$00"
Str1to16(10) = "0#000000000000$0"
Str1to16(11) = "0##0000000000$$0"
Str1to16(12) = "0##0000000000$$0"
Str1to16(13) = "0##0000000000$$0"
Str1to16(14) = "0#0$$$$$$$$$$0$0"
Str1to16(15) = "00$$$$$$$$$$$$00"
Str1to16(16) = "0000000000000000"

End Select
GFXPBARRAY(Number).AutoRedraw = True
GFXPBARRAY(Number).ScaleMode = 3

    For cr = 1 To 16
        For R = 1 To 16
            CurChar = Mid(Str1to16(cr), R, 1)
            If Not CustomColor Then
                If CurChar = "0" Then GFXPBARRAY(Number).PSet (R - 1, cr - 1), &H8000&
                If CurChar = "$" Then GFXPBARRAY(Number).PSet (R - 1, cr - 1), vbGreen
                If CurChar = "#" Then GFXPBARRAY(Number).PSet (R - 1, cr - 1), &H7000&
            Else
                If CurChar = "0" Then GFXPBARRAY(Number).PSet (R - 1, cr - 1), BackColor
                If CurChar = "$" Then GFXPBARRAY(Number).PSet (R - 1, cr - 1), LitDiodeColor
                If CurChar = "#" Then GFXPBARRAY(Number).PSet (R - 1, cr - 1), DiodeColor
            End If
        Next
    Next
Next
End Sub

Function SetupString(Str1 As String, Width As Integer) As String
Dim Caps As Integer: Caps = Width \ 16
If Len(Str1) < Caps Then
Dim OString As String: OString = ""
For I = 1 To Caps - Len(Str1)
OString = OString & "0"
Next
SetupString = OString & Str1
Else
SetupString = Str1
End If
End Function
