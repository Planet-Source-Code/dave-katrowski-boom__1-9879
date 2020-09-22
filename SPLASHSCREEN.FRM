VERSION 5.00
Begin VB.Form SplashScreen 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   3135
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      MousePointer    =   11  'Hourglass
      ScaleHeight     =   2835
      ScaleWidth      =   4275
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   2850
         MousePointer    =   11  'Hourglass
         Picture         =   "SplashScreen.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   1365
         TabIndex        =   3
         Top             =   2565
         Width           =   1365
      End
      Begin VB.PictureBox pbSP1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   1320
         MousePointer    =   11  'Hourglass
         Picture         =   "SplashScreen.frx":03C3
         ScaleHeight     =   1695
         ScaleWidth      =   1545
         TabIndex        =   2
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dave Katrowski"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         MousePointer    =   11  'Hourglass
         TabIndex        =   5
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "v0.0a0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3120
         MousePointer    =   11  'Hourglass
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   345
         Left            =   -45
         Top             =   2505
         Width           =   4440
      End
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   0
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Label2.Caption = "v" & App.Major & "." & App.Minor & "a" & App.Revision
End Sub

