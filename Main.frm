VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A1E5518C-0442-11D4-853E-44F1F6C00000}#18.0#0"; "BEATDISPLAY.OCX"
Begin VB.Form fMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "BOOM!"
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   10395
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   3960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1620
      TabIndex        =   95
      Top             =   60
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CDB 
      Left            =   240
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   960
      TabIndex        =   88
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   87
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "×"
      Height          =   195
      Left            =   60
      TabIndex        =   32
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   2040
      Picture         =   "Main.frx":030A
      ScaleHeight     =   6000
      ScaleWidth      =   8100
      TabIndex        =   0
      Top             =   240
      Width           =   8100
      Begin VB.CheckBox Check6 
         Caption         =   "Rev"
         Height          =   255
         Index           =   5
         Left            =   6390
         TabIndex        =   126
         Top             =   3465
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Rev"
         Height          =   255
         Index           =   4
         Left            =   6390
         TabIndex        =   125
         Top             =   2745
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Rev"
         Height          =   255
         Index           =   3
         Left            =   6390
         TabIndex        =   124
         Top             =   2025
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Rev"
         Height          =   255
         Index           =   2
         Left            =   6390
         TabIndex        =   123
         Top             =   1305
         Width           =   615
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Rev"
         Height          =   255
         Index           =   1
         Left            =   6390
         TabIndex        =   122
         Top             =   585
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Dist"
         Height          =   255
         Index           =   5
         Left            =   5790
         TabIndex        =   121
         Top             =   3465
         Width           =   735
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Dist"
         Height          =   255
         Index           =   4
         Left            =   5790
         TabIndex        =   120
         Top             =   2745
         Width           =   735
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Dist"
         Height          =   255
         Index           =   3
         Left            =   5790
         TabIndex        =   119
         Top             =   2025
         Width           =   735
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Dist"
         Height          =   255
         Index           =   2
         Left            =   5790
         TabIndex        =   118
         Top             =   1305
         Width           =   735
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Dist"
         Height          =   255
         Index           =   1
         Left            =   5790
         TabIndex        =   117
         Top             =   585
         Width           =   735
      End
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   120
         TabIndex        =   116
         Top             =   4140
         Width           =   1575
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H000000C0&
         Height          =   135
         Left            =   1720
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   115
         Top             =   4740
         Width           =   135
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00008000&
         Height          =   360
         Index           =   5
         Left            =   4560
         ScaleHeight     =   300
         ScaleWidth      =   1170
         TabIndex        =   94
         Top             =   5355
         Width           =   1230
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   131
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00008000&
         Height          =   375
         Index           =   4
         Left            =   4560
         ScaleHeight     =   315
         ScaleWidth      =   1170
         TabIndex        =   93
         Top             =   4995
         Width           =   1230
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   130
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00008000&
         Height          =   375
         Index           =   3
         Left            =   4560
         ScaleHeight     =   315
         ScaleWidth      =   1170
         TabIndex        =   92
         Top             =   4635
         Width           =   1230
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   129
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00008000&
         Height          =   375
         Index           =   2
         Left            =   4560
         ScaleHeight     =   315
         ScaleWidth      =   1170
         TabIndex        =   91
         Top             =   4275
         Width           =   1230
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   128
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00008000&
         Height          =   375
         Index           =   1
         Left            =   4560
         ScaleHeight     =   315
         ScaleWidth      =   1170
         TabIndex        =   90
         Top             =   3915
         Width           =   1230
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   0.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   75
         Index           =   0
         Left            =   3690
         TabIndex        =   101
         Top             =   5715
         Width           =   2190
      End
      Begin VB.CommandButton Command1 
         Caption         =   "> Wav 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3780
         TabIndex        =   31
         Top             =   5355
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "> Wav 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3780
         TabIndex        =   30
         Top             =   4995
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "> Wav 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3780
         TabIndex        =   29
         Top             =   4635
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "> Wav 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3780
         TabIndex        =   28
         Top             =   4275
         Width           =   795
      End
      Begin VB.CommandButton Command1 
         Caption         =   "> Wav 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3780
         TabIndex        =   27
         Top             =   3915
         Width           =   795
      End
      Begin VB.CommandButton Command7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   3
         Left            =   3690
         TabIndex        =   113
         Top             =   3915
         Width           =   105
      End
      Begin VB.CommandButton Command7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Index           =   1
         Left            =   5790
         TabIndex        =   112
         Top             =   3915
         Width           =   90
      End
      Begin VB.CommandButton Command7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   0.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   75
         Index           =   2
         Left            =   3690
         TabIndex        =   114
         Top             =   3855
         Width           =   2190
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Delay"
         Height          =   255
         Index           =   1
         Left            =   4560
         TabIndex        =   102
         Top             =   580
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reverb"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Offset"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000040&
         Height          =   1095
         Left            =   1680
         ScaleHeight     =   1035
         ScaleWidth      =   1860
         TabIndex        =   89
         Top             =   4695
         Width           =   1920
         Begin VB.Line Line1 
            BorderColor     =   &H000000C0&
            X1              =   -150
            X2              =   2130
            Y1              =   525
            Y2              =   525
         End
      End
      Begin VB.FileListBox File1 
         Height          =   870
         Left            =   1665
         Pattern         =   "*.wav"
         TabIndex        =   21
         Top             =   3825
         Width           =   1950
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Delay"
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   111
         Top             =   3460
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Delay"
         Height          =   255
         Index           =   4
         Left            =   4560
         TabIndex        =   110
         Top             =   2740
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Delay"
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   109
         Top             =   2020
         Width           =   735
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Delay"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   108
         Top             =   1300
         Width           =   735
      End
      Begin BeatDisplayer.beatdisplay beatdisplay1 
         Height          =   495
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   873
      End
      Begin MSComctlLib.Slider Slider5 
         Height          =   240
         Index           =   1
         Left            =   5280
         TabIndex        =   103
         Top             =   585
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   130
         Max             =   250
         SelStart        =   130
         TickFrequency   =   20
         Value           =   130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reverb"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Offset"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   48
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reverb"
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   8
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Offset"
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   49
         Top             =   2040
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reverb"
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   14
         Top             =   2760
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Offset"
         Height          =   195
         Index           =   4
         Left            =   2520
         TabIndex        =   50
         Top             =   2760
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reverb"
         Height          =   195
         Index           =   5
         Left            =   3240
         TabIndex        =   18
         Top             =   3480
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Offset"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   51
         Top             =   3480
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         Caption         =   "R"
         Height          =   255
         Index           =   5
         Left            =   1035
         TabIndex        =   100
         Top             =   3465
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Caption         =   "R"
         Height          =   255
         Index           =   4
         Left            =   1035
         TabIndex        =   99
         Top             =   2745
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Caption         =   "R"
         Height          =   255
         Index           =   3
         Left            =   1035
         TabIndex        =   98
         Top             =   2025
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Caption         =   "R"
         Height          =   255
         Index           =   2
         Left            =   1035
         TabIndex        =   97
         Top             =   1305
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Caption         =   "R"
         Height          =   255
         Index           =   1
         Left            =   1035
         TabIndex        =   96
         Top             =   585
         Width           =   495
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   130
         TabIndex        =   20
         Top             =   3840
         Width           =   1550
      End
      Begin VB.PictureBox DMKLED1 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6120
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   85
         Top             =   4170
         Width           =   540
      End
      Begin VB.PictureBox DMKLED2 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6675
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   84
         Top             =   4170
         Width           =   240
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   9
         Left            =   3360
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   83
         Top             =   6240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   8
         Left            =   3000
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   82
         Top             =   6240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   7
         Left            =   2640
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   81
         Top             =   6240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   6
         Left            =   2280
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   80
         Top             =   6240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   5
         Left            =   1920
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   79
         Top             =   6240
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   4
         Left            =   3360
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   78
         Top             =   6000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   3
         Left            =   3000
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   77
         Top             =   6000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   2
         Left            =   2640
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   76
         Top             =   6000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   1
         Left            =   2280
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   75
         Top             =   6000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox LED 
         AutoRedraw      =   -1  'True
         Height          =   300
         Index           =   0
         Left            =   1920
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   74
         Top             =   6000
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton Command3 
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   73
         Top             =   3500
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   72
         Top             =   2780
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   71
         Top             =   2060
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   70
         Top             =   1340
         Width           =   495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   69
         Top             =   620
         Width           =   495
      End
      Begin VB.CommandButton cmesup 
         Caption         =   ">"
         Height          =   255
         Left            =   7560
         TabIndex        =   68
         Top             =   5430
         Width           =   255
      End
      Begin VB.CommandButton cmesdown 
         Caption         =   "<"
         Height          =   255
         Left            =   7320
         TabIndex        =   67
         Top             =   5430
         Width           =   255
      End
      Begin VB.CommandButton tmesup 
         Caption         =   ">"
         Height          =   255
         Left            =   7560
         TabIndex        =   65
         Top             =   5190
         Width           =   255
      End
      Begin VB.CommandButton tmesdown 
         Caption         =   "<"
         Height          =   255
         Left            =   7320
         TabIndex        =   64
         Top             =   5190
         Width           =   255
      End
      Begin VB.PictureBox pbS 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   7455
         Picture         =   "Main.frx":1067
         ScaleHeight     =   465
         ScaleWidth      =   420
         TabIndex        =   57
         Top             =   3960
         Width           =   420
      End
      Begin VB.PictureBox pDULIT 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   525
         Left            =   -600
         Picture         =   "Main.frx":1AD5
         ScaleHeight     =   465
         ScaleWidth      =   420
         TabIndex        =   56
         Top             =   6000
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox pULIT 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   525
         Left            =   360
         Picture         =   "Main.frx":2543
         ScaleHeight     =   465
         ScaleWidth      =   420
         TabIndex        =   55
         Top             =   6000
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox pLIT 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   525
         Left            =   -120
         Picture         =   "Main.frx":2FB1
         ScaleHeight     =   465
         ScaleWidth      =   420
         TabIndex        =   54
         Top             =   6000
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox pbP 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   7005
         Picture         =   "Main.frx":3A1F
         ScaleHeight     =   465
         ScaleWidth      =   420
         TabIndex        =   53
         Top             =   3960
         Width           =   420
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   240
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   423
         _Version        =   393216
         Min             =   5000
         Max             =   80000
         SelStart        =   5000
         TickStyle       =   3
         Value           =   5000
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   240
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   423
         _Version        =   393216
         Min             =   5000
         Max             =   80000
         SelStart        =   5000
         TickStyle       =   3
         Value           =   5000
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   240
         Index           =   3
         Left            =   1800
         TabIndex        =   13
         Top             =   2040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   423
         _Version        =   393216
         Min             =   5000
         Max             =   80000
         SelStart        =   5000
         TickStyle       =   3
         Value           =   5000
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   240
         Index           =   4
         Left            =   1800
         TabIndex        =   17
         Top             =   2760
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   423
         _Version        =   393216
         Min             =   5000
         Max             =   80000
         SelStart        =   5000
         TickStyle       =   3
         Value           =   5000
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   240
         Index           =   5
         Left            =   1800
         TabIndex        =   34
         Top             =   3480
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   423
         _Version        =   393216
         Min             =   5000
         Max             =   80000
         SelStart        =   5000
         TickStyle       =   3
         Value           =   5000
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   240
         Index           =   1
         Left            =   7320
         TabIndex        =   35
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   240
         Index           =   2
         Left            =   7320
         TabIndex        =   36
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   240
         Index           =   3
         Left            =   7320
         TabIndex        =   37
         Top             =   2040
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   240
         Index           =   4
         Left            =   7320
         TabIndex        =   38
         Top             =   2760
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   240
         Index           =   5
         Left            =   7320
         TabIndex        =   39
         Top             =   3480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   -5000
         Max             =   0
         TickStyle       =   3
      End
      Begin VB.PictureBox sULIT 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   525
         Left            =   840
         Picture         =   "Main.frx":448D
         ScaleHeight     =   465
         ScaleWidth      =   420
         TabIndex        =   58
         Top             =   6000
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox sDLIT 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   525
         Left            =   1320
         Picture         =   "Main.frx":4EFB
         ScaleHeight     =   465
         ScaleWidth      =   420
         TabIndex        =   59
         Top             =   6000
         Visible         =   0   'False
         Width           =   480
      End
      Begin MSComctlLib.Slider Slider4 
         Height          =   255
         Left            =   6120
         TabIndex        =   45
         Top             =   4710
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         LargeChange     =   1
         Min             =   50
         Max             =   240
         SelStart        =   160
         TickFrequency   =   10
         Value           =   160
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   225
         Index           =   1
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   397
         _Version        =   393216
         Min             =   10
         Max             =   40
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   240
         Index           =   2
         Left            =   4080
         TabIndex        =   10
         Top             =   1305
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   10
         Max             =   40
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   240
         Index           =   3
         Left            =   4080
         TabIndex        =   11
         Top             =   2025
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   10
         Max             =   40
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   240
         Index           =   5
         Left            =   4080
         TabIndex        =   19
         Top             =   3465
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   10
         Max             =   40
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   240
         Index           =   4
         Left            =   4080
         TabIndex        =   15
         Top             =   2745
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   10
         Max             =   40
         SelStart        =   10
         TickFrequency   =   10
         Value           =   10
      End
      Begin MSComctlLib.Slider Slider5 
         Height          =   240
         Index           =   2
         Left            =   5280
         TabIndex        =   104
         Top             =   1305
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   130
         Max             =   250
         SelStart        =   130
         TickFrequency   =   20
         Value           =   130
      End
      Begin MSComctlLib.Slider Slider5 
         Height          =   240
         Index           =   3
         Left            =   5280
         TabIndex        =   105
         Top             =   2025
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   130
         Max             =   250
         SelStart        =   130
         TickFrequency   =   20
         Value           =   130
      End
      Begin MSComctlLib.Slider Slider5 
         Height          =   240
         Index           =   4
         Left            =   5280
         TabIndex        =   106
         Top             =   2745
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   130
         Max             =   250
         SelStart        =   130
         TickFrequency   =   20
         Value           =   130
      End
      Begin MSComctlLib.Slider Slider5 
         Height          =   240
         Index           =   5
         Left            =   5280
         TabIndex        =   107
         Top             =   3465
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   423
         _Version        =   393216
         Min             =   130
         Max             =   250
         SelStart        =   130
         TickFrequency   =   20
         Value           =   130
      End
      Begin BeatDisplayer.beatdisplay beatdisplay3 
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   873
      End
      Begin BeatDisplayer.beatdisplay beatdisplay4 
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   2280
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   873
      End
      Begin BeatDisplayer.beatdisplay beatdisplay2 
         Height          =   495
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   873
      End
      Begin BeatDisplayer.beatdisplay beatdisplay5 
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   3000
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   873
      End
      Begin VB.Image Image1 
         Height          =   210
         Index           =   1
         Left            =   240
         Picture         =   "Main.frx":5969
         Top             =   1320
         Width           =   210
      End
      Begin VB.Image Image1 
         Height          =   210
         Index           =   0
         Left            =   240
         Picture         =   "Main.frx":5B0F
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo"
         Height          =   255
         Index           =   0
         Left            =   6045
         TabIndex        =   46
         Top             =   4500
         Width           =   1845
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   6000
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Counter"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6075
         TabIndex        =   86
         Top             =   3930
         Width           =   900
      End
      Begin VB.Shape Shape4 
         Height          =   540
         Left            =   6075
         Top             =   3930
         Width           =   885
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   6675
         X2              =   6675
         Y1              =   4185
         Y2              =   4415
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   270
         Left            =   6120
         Top             =   4170
         Width           =   810
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6840
         TabIndex        =   66
         Top             =   5190
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6840
         TabIndex        =   63
         Top             =   5430
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Current"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   62
         Top             =   5430
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   61
         Top             =   5190
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Meassure:"
         Height          =   255
         Left            =   6000
         TabIndex        =   60
         Top             =   4980
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Vol"
         Height          =   255
         Index           =   4
         Left            =   6840
         TabIndex        =   44
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Vol"
         Height          =   255
         Index           =   3
         Left            =   6840
         TabIndex        =   43
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Vol"
         Height          =   255
         Index           =   2
         Left            =   6840
         TabIndex        =   42
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Vol"
         Height          =   255
         Index           =   1
         Left            =   6840
         TabIndex        =   41
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Vol"
         Height          =   255
         Index           =   0
         Left            =   6840
         TabIndex        =   40
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Pit."
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   16
         Top             =   3480
         Width           =   255
      End
      Begin VB.Image Image1 
         Height          =   210
         Index           =   4
         Left            =   240
         Picture         =   "Main.frx":5C9F
         Top             =   3480
         Width           =   210
      End
      Begin VB.Label Label2 
         Caption         =   "Pit."
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   2760
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   210
         Index           =   3
         Left            =   240
         Picture         =   "Main.frx":5E46
         Top             =   2760
         Width           =   210
      End
      Begin VB.Label Label2 
         Caption         =   "Pit."
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   2040
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   210
         Index           =   2
         Left            =   240
         Picture         =   "Main.frx":5FDB
         Top             =   2040
         Width           =   210
      End
      Begin VB.Label Label2 
         Caption         =   "Pit."
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   2
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Pit."
         Height          =   255
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   4800
      Picture         =   "Main.frx":6177
      ScaleHeight     =   1170
      ScaleWidth      =   1740
      TabIndex        =   33
      Top             =   1800
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2520
      Left            =   3240
      Picture         =   "Main.frx":6CC4
      ScaleHeight     =   2520
      ScaleWidth      =   4545
      TabIndex        =   52
      Top             =   1920
      Width           =   4545
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub beatdisplay1_Klick()
SEQ(1, curmes) = beatdisplay1.outDis
End Sub
Private Sub beatdisplay2_Klick()
SEQ(2, curmes) = beatdisplay2.outDis
End Sub
Private Sub beatdisplay3_Klick()
SEQ(3, curmes) = beatdisplay3.outDis
End Sub
Private Sub beatdisplay4_Klick()
SEQ(4, curmes) = beatdisplay4.outDis
End Sub
Private Sub beatdisplay5_Klick()
SEQ(5, curmes) = beatdisplay5.outDis
End Sub
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check4(Index).Value = 1 Then Check4(Index).Value = 0
End Sub

Private Sub Check4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1(Index).Value = 1 Then Check1(Index).Value = 0
End Sub
Private Sub cmesdown_Click()
If curmes = 1 Then Exit Sub
curmes = curmes - 1: Label8 = curmes
beatdisplay1.inDis SEQ(1, curmes)
beatdisplay2.inDis SEQ(2, curmes)
beatdisplay3.inDis SEQ(3, curmes)
beatdisplay4.inDis SEQ(4, curmes)
beatdisplay5.inDis SEQ(5, curmes)
End Sub

Private Sub cmesup_Click()
If curmes = tmes Then Exit Sub
curmes = curmes + 1: Label8 = curmes
beatdisplay1.inDis SEQ(1, curmes)
beatdisplay2.inDis SEQ(2, curmes)
beatdisplay3.inDis SEQ(3, curmes)
beatdisplay4.inDis SEQ(4, curmes)
beatdisplay5.inDis SEQ(5, curmes)
End Sub

Private Sub Command1_Click(Index As Integer) ': On Error Resume Next
If File1.Path = "" Or File1.Filename = "" Then Exit Sub
Dim ffL As Integer, LByte() As Byte: ffL = FreeFile
Picture5(Index).BackColor = &H40&
ReDim LByte(FileLen(File1.Path & "\" & File1.Filename))
Open File1.Path & "\" & File1.Filename For Binary As #ffL
Get #ffL, , LByte
Close #ffL: ffL = FreeFile
Open App.Path & "\temp" & Index & ".wav" For Binary As #ffL
Put #ffL, , LByte
Close #ffL

WavePaths(Index) = App.Path & "\temp" & Index & ".wav"
DS.LoadWavToChannel Index, App.Path & "\temp" & Index & ".wav"
Label10(Index).Caption = "Generating Effects...": Label10(Index).Refresh
Distort Index
Reverse Index
DisRev Index
DS.LoadWavToChannel Index + 5, App.Path & "\tempdis" & Index & ".wav"
DS.LoadWavToChannel Index + 10, App.Path & "\temprev" & Index & ".wav"
DS.LoadWavToChannel Index + 15, App.Path & "\tdisrev" & Index & ".wav"

Slider1(Index).Value = DS.GetFrequency(Index)
Label10(Index).Caption = ""
Picture5(Index).BackColor = &H8000&
doPlot Picture5(Index), WavePaths(Index), (Rnd * 1) + 2
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Unload Me
End Sub

Private Sub Command3_Click(Index As Integer)
If Command3(Index).Caption = "On" Then
cON(Index) = False
Command3(Index).Caption = "Off"
Else
cON(Index) = True
Command3(Index).Caption = "On"
End If
End Sub

Private Sub Command4_Click()
'On Error GoTo Hell
CDB.Filename = vbNullString
CDB.DialogTitle = "Save your groove..."
CDB.ShowSave
FileOUT Me, CDB.Filename
Hell:
End Sub

Private Sub Command5_Click(): Dim cne As Integer
'On Error GoTo Hell
CDB.Filename = vbNullString
CDB.DialogTitle = "Open a phat patch..."
CDB.ShowOpen

curmes = 1: tmes = 1: mes = 1: tick = 0
Label8 = curmes: Label9 = tmes

beatdisplay1.inDis "0000000000000000"
beatdisplay2.inDis "0000000000000000"
beatdisplay3.inDis "0000000000000000"
beatdisplay4.inDis "0000000000000000"
beatdisplay5.inDis "0000000000000000"

For cne = 1 To 5: For m = 1 To 16
SEQ(cne, m) = "0000000000000000"
Next
cON(cne) = True: Slider1(cne).Value = 5000: Slider2(cne).Value = 10
Slider3(cne).Value = 0: Check1(cne).Value = 0: Check2(cne).Value = 0: Check5(cne).Value = 0: Check6(cne).Value = 0
Picture5(cne).Cls: DS.ClearBuffer cne
Next

Slider4.Value = 160

FileIN Me, CDB.Filename
'Hell:
End Sub

Private Sub Command6_Click(): Dim cne As Integer
curmes = 1: tmes = 1: mes = 1: tick = 0
Label8 = curmes: Label9 = tmes

beatdisplay1.inDis "0000000000000000"
beatdisplay2.inDis "0000000000000000"
beatdisplay3.inDis "0000000000000000"
beatdisplay4.inDis "0000000000000000"
beatdisplay5.inDis "0000000000000000"

For cne = 1 To 5: For m = 1 To 16
SEQ(cne, m) = "0000000000000000"
Next
cON(cne) = True: Slider1(cne).Value = 5000: Slider2(cne).Value = 5
Slider3(cne).Value = 0: Check1(cne).Value = 0: Check2(cne).Value = 0: Check5(cne).Value = 0: Check6(cne).Value = 0
Picture5(cne).Cls: DS.ClearBuffer cne: WavePaths(cne) = ""
Next
ReDim W1BYTE(0)
ReDim W2BYTE(0)
ReDim W3BYTE(0)
ReDim W4BYTE(0)
ReDim W5BYTE(0)
Slider4.Value = 160
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
doPlot Picture4, File1.Path & "\" & File1.Filename, plotState
lastplot = File1.Path & "\" & File1.Filename
End Sub

Private Sub Form_Load(): Static c As Integer, m As Integer
DS.Initialize_Engine Me.Hwnd

center Picture1
center Picture2
center Picture3

Me.WindowState = 2
curmes = 1: tmes = 1: mes = 1: tick = 0

beatdisplay1.inDis "0000000000000000"
beatdisplay2.inDis "0000000000000000"
beatdisplay3.inDis "0000000000000000"
beatdisplay4.inDis "0000000000000000"
beatdisplay5.inDis "0000000000000000"

For c = 1 To 5: For m = 1 To 16
SEQ(c, m) = "0000000000000000"
Next: cON(c) = True: Next

LoadGFX LED(): DrawLED DMKLED1, "1", LED(), 0: DrawLED DMKLED2, "1", LED(), 1
CDB.Flags = &H2: CDB.Filter = "Boom  - Patch Files (*.PAT)|*.pat"
P.tempo = Slider4.Value: plotState = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Picture1.Visible = False
Picture3.Visible = False
Command2.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False

center Picture2
Picture2.Visible = True
fMain.Cls
DS.Pause 5000
fMain.BackColor = &HFFFFFF
Picture2.Visible = False
Picture3.Visible = True
fMain.Cls
DS.Pause 5000
DS.Terminate_Engine
On Error Resume Next
For I = 1 To 5
Kill App.Path & "\tdisrev" & I & ".wav"
Kill App.Path & "\tempdis" & I & ".wav"
Kill App.Path & "\temprev" & I & ".wav"
Kill App.Path & "\temp" & I & ".wav"
Next
Unload SplashScreen
Set SplashScreen = Nothing
Unload Me
Set fMain = Nothing
End
End Sub


Private Sub pbP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Enabled = True Then Exit Sub
pbP.Picture = pDULIT.Picture
End Sub

Private Sub pbP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cv: If Timer1.Enabled = True Then Exit Sub
pbP.Picture = pLIT.Picture
For I = 1 To 5: For cv = 1 To 16
If SEQ(I, cv) = "" Then SEQ(I, cv) = "0000000000000000"
Next: Next
Timer1.Enabled = True
End Sub

Private Sub pbS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Enabled = False Then Exit Sub
pbS.Picture = sDLIT.Picture
End Sub

Private Sub pbS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Enabled = False Then Exit Sub
pbS.Picture = sULIT.Picture
pbP.Picture = pULIT.Picture
Timer1.Enabled = False: mes = 1: tick = 0
DrawLED DMKLED1, "1", LED(), 0
DrawLED DMKLED2, "1", LED(), 1
End Sub

Private Sub Picture5_Click(Index As Integer)
If WavePaths(Index) = "" Then Exit Sub
doPlot Picture4, WavePaths(Index), plotState
lastplot = WavePaths(Index)
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Picture6.BackColor = &HC0& Then
Picture6.BackColor = &HFF: plotState = 1
doPlot Picture4, lastplot, plotState
Else
Picture6.BackColor = &HC0&: plotState = 0
doPlot Picture4, lastplot, plotState
End If
End Sub

Private Sub Slider3_Change(Index As Integer)
DS.SetVolume Index, Slider3(Index).Value
End Sub

Private Sub Slider3_Scroll(Index As Integer)
DS.SetVolume Index, Slider3(Index).Value
End Sub

Private Sub Slider4_Change()
P.tempo = Slider4.Value + 3
End Sub

Private Sub Slider4_Scroll()
P.tempo = Slider4.Value + 3
End Sub

Private Sub Timer1_Timer() ': On Error Resume Next
Dim s As Byte: For s = 1 To 2
tick = tick + 1: DrawLED DMKLED2, CStr(GetCounterBeat(tick)), LED(), 1

If tick = 17 Then tick = 1: mes = mes + 1: DrawLED DMKLED1, CStr(mes), LED(), 0: DrawLED DMKLED2, CStr(GetCounterBeat(tick)), LED(), 1
If mes > tmes Then mes = 1: DrawLED DMKLED1, CStr(mes), LED(), 0

For I = 1 To 5
If -(I And 1) Then DS.SetFrequency I, Slider1(I).Value: DS.SetFrequency I + 5, Slider1(I).Value: DS.SetFrequency I + 10, Slider1(I).Value: DS.SetFrequency I + 15, Slider1(I).Value

If cON(I) And Mid(SEQ(I, mes), tick, 1) = 1 Then
If Check2(I).Value = 1 Then DS.SetFrequency I, RndRange(Slider1(I).Value - 2000, Slider1(I).Value + 2000): DS.SetFrequency I + 5, RndRange(Slider1(I).Value - 2000, Slider1(I).Value + 2000): DS.SetFrequency I + 10, RndRange(Slider1(I).Value - 2000, Slider1(I).Value + 2000): DS.SetFrequency I + 15, RndRange(Slider1(I).Value - 2000, Slider1(I).Value + 2000): GoTo PlaY
If Check3(I).Value = 1 Then DS.SetFrequency I, RndRange(Slider1(I).Value - 10000, Slider1(I).Value + 10000): DS.SetFrequency I + 5, RndRange(Slider1(I).Value - 10000, Slider1(I).Value + 10000): DS.SetFrequency I + 10, RndRange(Slider1(I).Value - 10000, Slider1(I).Value + 10000): DS.SetFrequency I + 15, RndRange(Slider1(I).Value - 10000, Slider1(I).Value + 10000)
PlaY:
If Check6(I).Value = 0 And Check5(I).Value = 0 Then
If Check1(I).Value = 0 And Check4(I).Value = 0 Then DS.StopSound I: DS.PlaySound I: GoTo out
If Check1(I).Value = 1 Then DS.StopSound I: DS.PlaySound I: DS.PlayEcho I, 2, Slider2(I).Value: GoTo out
If Check4(I).Value = 1 Then DS.StopSound I: DS.PlaySound I: DS.PlayEcho I, 2, Slider5(I).Value: GoTo out
End If
If Check6(I).Value = 1 And Check5(I).Value = 1 Then
If Check1(I).Value = 0 And Check4(I).Value = 0 Then DS.StopSound I + 15: DS.PlaySound I + 15: GoTo out
If Check1(I).Value = 1 Then DS.StopSound I: DS.PlaySound I + 15: DS.PlayEcho I + 15, 2, Slider2(I).Value: GoTo out
If Check4(I).Value = 1 Then DS.StopSound I: DS.PlaySound I + 15: DS.PlayEcho I + 15, 2, Slider5(I).Value: GoTo out
End If
If Check5(I).Value = 1 Then
If Check1(I).Value = 0 And Check4(I).Value = 0 Then DS.StopSound I + 5: DS.PlaySound I + 5: GoTo out
If Check1(I).Value = 1 Then DS.StopSound I: DS.PlaySound I + 5: DS.PlayEcho I + 5, 2, Slider2(I).Value: GoTo out
If Check4(I).Value = 1 Then DS.StopSound I: DS.PlaySound I + 5: DS.PlayEcho I + 5, 2, Slider5(I).Value: GoTo out
End If
If Check6(I).Value = 1 Then
If Check1(I).Value = 0 And Check4(I).Value = 0 Then DS.StopSound I + 10: DS.PlaySound I + 10: GoTo out
If Check1(I).Value = 1 Then DS.StopSound I + 10: DS.PlaySound I + 10: DS.PlayEcho I + 10, 2, Slider2(I).Value: GoTo out
If Check4(I).Value = 1 Then DS.StopSound I + 10: DS.PlaySound I + 10: DS.PlayEcho I + 10, 2, Slider5(I).Value: GoTo out
End If
End If

out:
Next

st = timeGetTime: While timeGetTime < st + BpmToInterval(P.tempo): DoEvents: Wend

If Timer1.Enabled = False Then Exit Sub
If s = 2 Then s = 1
Next
End Sub

Private Sub tmesdown_Click()
If tmes = 1 Then Exit Sub
tmes = tmes - 1: Label9 = tmes
End Sub

Private Sub tmesup_Click()
If tmes = 16 Then Exit Sub
tmes = tmes + 1: Label9 = tmes
End Sub


