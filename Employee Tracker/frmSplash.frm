VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4050
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4140
      Left            =   -60
      TabIndex        =   0
      Top             =   -75
      Width           =   7155
      Begin MCI.MMControl MMC1 
         Height          =   330
         Left            =   1485
         TabIndex        =   3
         Top             =   3240
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         PlayEnabled     =   -1  'True
         DeviceType      =   "WaveAudio"
         FileName        =   "C:\WINDOWS.0\Media\Windows XP Shutdown.wav"
      End
      Begin VB.Timer Timer1 
         Interval        =   3500
         Left            =   240
         Top             =   3330
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Copyright 2001 Archive Software Solutions inc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3210
         TabIndex        =   2
         Top             =   3780
         Width           =   3735
      End
      Begin VB.Image Image2 
         Height          =   1365
         Left            =   2115
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1965
      End
      Begin VB.Image Image1 
         Height          =   1035
         Left            =   3120
         Picture         =   "frmSplash.frx":044E
         Stretch         =   -1  'True
         Top             =   2130
         Width           =   1485
      End
      Begin VB.Image imgLogo 
         Height          =   1350
         Left            =   4140
         Picture         =   "frmSplash.frx":0890
         Stretch         =   -1  'True
         Top             =   1005
         Width           =   1350
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         Caption         =   "Employee Tracker 1.1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   90
         TabIndex        =   1
         Top             =   165
         Width           =   6660
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub Form_Load()
MMC1.Command = "Open"
MMC1.Command = "Play"
End Sub


Private Sub Timer1_Timer()
MMC1.Command = "Open"
MMC1.Command = "Play"
     frmLogin.Show
     Me.Hide
     Unload Me
End Sub
