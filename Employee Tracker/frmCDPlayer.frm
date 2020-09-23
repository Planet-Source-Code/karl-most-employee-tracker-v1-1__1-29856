VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmCDPlayer 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " CD Player "
   ClientHeight    =   2025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5160
   Icon            =   "frmCDPlayer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEject 
      Caption         =   "Eject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   10
      Top             =   1620
      Width           =   1080
   End
   Begin MCI.MMControl MMC 
      Height          =   495
      Left            =   30
      TabIndex        =   9
      Top             =   2430
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   393216
      DeviceType      =   "CDAudio"
      FileName        =   ""
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   4335
      Picture         =   "frmCDPlayer.frx":1CCA
      ScaleHeight     =   795
      ScaleWidth      =   750
      TabIndex        =   6
      Top             =   1155
      Width           =   750
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   30
      Picture         =   "frmCDPlayer.frx":3994
      ScaleHeight     =   720
      ScaleWidth      =   750
      TabIndex        =   5
      Top             =   0
      Width           =   750
   End
   Begin VB.PictureBox PicPause 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3015
      Picture         =   "frmCDPlayer.frx":565E
      ScaleHeight     =   735
      ScaleWidth      =   720
      TabIndex        =   4
      Top             =   285
      Width           =   720
   End
   Begin VB.PictureBox PicStop 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3735
      Picture         =   "frmCDPlayer.frx":7130
      ScaleHeight     =   735
      ScaleWidth      =   645
      TabIndex        =   3
      Top             =   285
      Width           =   645
   End
   Begin VB.PictureBox PicPlay 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2295
      Picture         =   "frmCDPlayer.frx":8B3A
      ScaleHeight     =   735
      ScaleWidth      =   720
      TabIndex        =   2
      Top             =   285
      Width           =   720
   End
   Begin VB.PictureBox PicNext 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1650
      Picture         =   "frmCDPlayer.frx":A60C
      ScaleHeight     =   735
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   285
      Width           =   720
   End
   Begin VB.PictureBox PicBack 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   960
      Picture         =   "frmCDPlayer.frx":C016
      ScaleHeight     =   735
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   285
      Width           =   795
   End
   Begin VB.Label lblTrack 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2730
      TabIndex        =   8
      Top             =   1245
      Width           =   270
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Track:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2055
      TabIndex        =   7
      Top             =   1260
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   8
      Height          =   855
      Left            =   870
      Top             =   240
      Width           =   3585
   End
End
Attribute VB_Name = "frmCDPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdEject_Click()
                                                               '  err%
    MMC.Command = "Eject"
                                                                                '  err%
End Sub

Private Sub Form_Load()
                                                                           '  err%
     MMC.Command = "Open"
                                                                                               '  err%
                                                                                       '  err%
End Sub


Private Sub MMC_StatusUpdate()
                                                              '  err%
     lblTrack.Caption = MMC.Track
                                                                              '  err%
End Sub

Private Sub PicBack_Click()
                                                                '  err%
     MMC.Command = "Prev"
                                                                                              '  err%
                                                                               '  err%
End Sub

Private Sub PicNext_Click()
                                                                  '  err%
     MMC.Command = "Next"
End Sub

Private Sub PicPause_Click()
                                                               '  err%
     MMC.Command = "Pause"
                                                                              '  err%
End Sub

Private Sub PicPlay_Click()
                                                               '  err%
     MMC.Command = "Play"
                                                                              '  err%
End Sub

Private Sub PicStop_Click()
                                                               '  err%
   MMC.Command = "Stop"
                                                                                  '  err%
End Sub
