VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMain 
   Caption         =   "Employee Tracker 1.1"
   ClientHeight    =   8085
   ClientLeft      =   1770
   ClientTop       =   1485
   ClientWidth     =   11025
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   6990
      Picture         =   "frmMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Help."
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdDateTime 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   6120
      Picture         =   "frmMain.frx":210C
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Insert current Date and Time."
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdCDPlayer 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   5250
      Picture         =   "frmMain.frx":3DD6
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "open Employee Tracker CD Player."
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdCalc 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   4380
      Picture         =   "frmMain.frx":5AA0
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Open the calculator."
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   3510
      Picture         =   "frmMain.frx":776A
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Print"
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   2640
      Picture         =   "frmMain.frx":9434
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Delete selected record."
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdUpdate 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   1770
      Picture         =   "frmMain.frx":B0FE
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Save selected record."
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdAddNew 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   900
      Picture         =   "frmMain.frx":CDC8
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Add a new record."
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdLogoff 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   30
      Picture         =   "frmMain.frx":EA92
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Logoff"
      Top             =   0
      Width           =   885
   End
   Begin MSAdodcLib.Adodc EMPData 
      Height          =   330
      Left            =   1800
      Top             =   7275
      Visible         =   0   'False
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmMain.frx":1075C
      OLEDBString     =   $"frmMain.frx":107F7
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "EmployeeRecords"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   630
      Top             =   7650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Arial"
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9525
      TabIndex        =   49
      Top             =   3585
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9540
      TabIndex        =   48
      Top             =   4155
      Width           =   1080
   End
   Begin VB.TextBox txtDate 
      DataField       =   "Date"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   8400
      TabIndex        =   46
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtNotes 
      DataField       =   "Notes"
      DataSource      =   "EMPData"
      Height          =   1935
      Left            =   5385
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   5685
      Width           =   5175
   End
   Begin VB.Frame Frame2 
      Height          =   1425
      Left            =   600
      TabIndex        =   41
      Top             =   5610
      Width           =   4695
      Begin VB.TextBox txtEmail 
         DataField       =   "Email address"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1560
         TabIndex        =   20
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label21 
         Caption         =   "E-Mail Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   585
      TabIndex        =   31
      Top             =   2760
      Width           =   10305
      Begin VB.TextBox txtMobile 
         DataField       =   "Mobile"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6735
         TabIndex        =   19
         Top             =   2130
         Width           =   1815
      End
      Begin VB.TextBox txtFax 
         DataField       =   "Fax"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6735
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtExt 
         DataField       =   "Work Ext"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6720
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtWorkPhone 
         DataField       =   "Work Phone"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6705
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtEveningPhone 
         DataField       =   "Evening Phone"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6720
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtDayTimePhone 
         DataField       =   "Daytime Phone"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   6720
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtZipCode 
         DataField       =   "Zip Code"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox CboState 
         DataField       =   "State"
         DataSource      =   "EMPData"
         Height          =   315
         ItemData        =   "frmMain.frx":10892
         Left            =   1080
         List            =   "frmMain.frx":10894
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtCity 
         DataField       =   "City"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1095
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtLine2 
         DataField       =   "Line 2"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   705
         Width           =   4095
      End
      Begin VB.TextBox txtAddress 
         DataField       =   "Address"
         DataSource      =   "EMPData"
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label20 
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   43
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   42
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   " Ext"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   40
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Work Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   39
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Eveining Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5340
         TabIndex        =   38
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label15 
         Caption         =   "Daytime Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   37
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   35
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Line 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtEmployeeID 
      DataField       =   "Employee ID"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   8400
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtSupervisorID 
      DataField       =   "Supervisor ID"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtSupervisor 
      DataField       =   "Supervisor Name"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtSocialSecurity 
      DataField       =   "Social Security #"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   8400
      TabIndex        =   5
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtSalary 
      DataField       =   "Salary"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   4800
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtDateHired 
      DataField       =   "Date Hired"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtPosition 
      DataField       =   "Position"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   8400
      TabIndex        =   2
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtLastName 
      DataField       =   "Last Name"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   4815
      TabIndex        =   1
      Top             =   1545
      Width           =   2175
   End
   Begin VB.TextBox txtFirstName 
      DataField       =   "First Name"
      DataSource      =   "EMPData"
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1545
      Width           =   2175
   End
   Begin MSComctlLib.ImageList ImgLstToolbar1 
      Left            =   0
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10896
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12570
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1424A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17BFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":198D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B5B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D28C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2291A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label23 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7860
      TabIndex        =   47
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5430
      TabIndex        =   45
      Top             =   5445
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "Employee ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   30
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Supervisor ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   29
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Supervisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Social Security"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   27
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Date Hired"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7605
      TabIndex        =   24
      Top             =   1575
      Width           =   750
   End
   Begin VB.Label Label2 
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3840
      TabIndex        =   23
      Top             =   1560
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   22
      Top             =   1560
      Width           =   945
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_cmd_Logoff 
         Caption         =   "&Logoff"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cmd_Add 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnu_cmd_Edit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnu_cmd_Delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnu_cmd_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnu_cmd_Print 
         Caption         =   "&Print"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cmd_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_Accessories 
      Caption         =   "A&ccessories"
      Begin VB.Menu mnu_cmd_Calculator 
         Caption         =   "Calculat&or"
      End
      Begin VB.Menu mnu_cmd_CD_Player 
         Caption         =   "CD Pla&yer"
      End
   End
   Begin VB.Menu mnu_Format 
      Caption         =   "For&mat"
      Begin VB.Menu mnu_cmd_Font 
         Caption         =   "Fo&nt"
      End
      Begin VB.Menu mnu_cmd_DateTime 
         Caption         =   "Insert Date and Time"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_cmd_How_To 
         Caption         =   "&How To"
      End
      Begin VB.Menu mnu_cmd_About 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
'API Function Declaration
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Selected As Integer
Dim Focused As Boolean

Private Sub cmdAddNew_Click()
 On Error GoTo AddErr
 txtFirstName.SetFocus
  EMPData.Recordset.AddNew
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdBack_Click()
    On Error GoTo BackErr:
EMPData.Recordset.MovePrevious
    If EMPData.Recordset.BOF Then
        EMPData.Recordset.MoveNext
    End If
Exit Sub
BackErr:
    MsgBox Err.Description
End Sub

Private Sub cmdCalc_Click()
X = Shell("C:\Windows.0\System32\calc.exe", 3)
End Sub

Private Sub cmdCDPlayer_Click()
frmCDPlayer.Show
End Sub

Private Sub cmdDateTime_Click()
currentDateTime = Now
txtDate.Text = Now
End Sub

Private Sub cmdDelete_Click()
On Error GoTo DeleteErr
  With EMPData.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdHelp_Click()
frmHelp.Show
End Sub

Private Sub cmdLogoff_Click()
Unload frmMain
frmLogin.Show
End Sub

Private Sub cmdNext_Click()
    On Erro GoTo NextErr:
EMPData.Recordset.MoveNext
    If EMPData.Recordset.EOF Then
        EMPData.Recordset.MovePrevious
    End If
Exit Sub
NextErr:
    MsgBox Err.Description
End Sub

Private Sub cmdPrint_Click()
MsgBox ("This feature has not yey been implemented."), vbOKOnly, "Comming Soon"
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo UpdateErr
  EMPData.Recordset.UpdateBatch adAffectAll
  EMPData.Recordset.MoveFirst
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub
Private Sub mnu_cmd_About_Click()
     frmAbout.Show
End Sub

Private Sub mnu_cmd_Calculator_Click()
'Open the windows calculator
X = Shell("C:\Windows.0\System32\calc.exe", 3)
End Sub

Private Sub mnu_cmd_CD_Player_Click()
frmCDPlayer.Show                                                                                   '  err%
End Sub

Private Sub mnu_cmd_DateTime_Click()                                                             '  err%
currentDateTime = Now
txtDate.Text = Now                                                                             '  err%
End Sub

Private Sub mnu_cmd_Exit_Click()
Unload frmMain
    End                                                                                          '  err%                                                                                         '  err%
End Sub

Private Sub mnu_cmd_Font_Click()
                                                         
     CD1.ShowFont

    txtFirstName.Text = CD1.FontBold
     txtFirstName.Text = CD1.FontItalic
     txtFirstName.Text = CD1.FontName
     txtFirstName.Text = CD1.FontSize
   txtFirstName.Text = CD1.FontStrikethru
    txtFirstName.Text = CD1.FontUnderline

9     txtLastName.Text = CD1.FontBold
10    txtLastName.Text = CD1.FontItalic
11    txtLastName.Text = CD1.FontName
12    txtLastName.Text = CD1.FontSize
13    txtLastName.Text = CD1.FontStrikethru
14    txtLastName.Text = CD1.FontUnderline

15    txtPosition.Text = CD1.FontBold
16    txtPosition.Text = CD1.FontItalic
17    txtPosition.Text = CD1.FontName
18    txtPosition.Text = CD1.FontSize
19    txtPosition.Text = CD1.FontStrikethru
20    txtPosition.Text = CD1.FontUnderline

21    txtDateHired.Text = CD1.FontBold
22    txtDateHired.Text = CD1.FontItalic
23    txtDateHired.Text = CD1.FontName
24    txtDateHired.Text = CD1.FontSize
25    txtDateHired.Text = CD1.FontStrikethru
26    txtDateHired.Text = CD1.FontUnderline

27    txtSalary.Text = CD1.FontBold
28    txtSalary.Text = CD1.FontItalic
29    txtSalary.Text = CD1.FontName
30    txtSalary.Text = CD1.FontSize
31    txtSalary.Text = CD1.FontStrikethru
32    txtSalary.Text = CD1.FontUnderline

33    txtSocialSecurity.Text = CD1.FontBold
34    txtSocialSecurity.Text = CD1.FontItalic
35    txtSocialSecurity.Text = CD1.FontName
36    txtSocialSecurity.Text = CD1.FontSize
37    txtSocialSecurity.Text = CD1.FontStrikethru
38    txtSocialSecurity.Text = CD1.FontUnderline

39    txtSupervisor.Text = CD1.FontBold
40    txtSupervisor.Text = CD1.FontItalic
41    txtSupervisor.Text = CD1.FontName
42    txtSupervisor.Text = CD1.FontSize
43    txtSupervisor.Text = CD1.FontStrikethru
44    txtSupervisor.Text = CD1.FontUnderline

45    txtSupervisorID.Text = CD1.FontBold
46    txtSupervisorID.Text = CD1.FontItalic
47    txtSupervisorID.Text = CD1.FontName
48    txtSupervisorID.Text = CD1.FontSize
49    txtSupervisorID.Text = CD1.FontStrikethru
50    txtSupervisorID.Text = CD1.FontUnderline

51    txtEmployeeID.Text = CD1.FontBold
52    txtEmployeeID.Text = CD1.FontItalic
53    txtEmployeeID.Text = CD1.FontName
54    txtEmployeeID.Text = CD1.FontSize
55    txtEmployeeID.Text = CD1.FontStrikethru
56    txtEmployeeID.Text = CD1.FontUnderline

57    txtAddress.Text = CD1.FontBold
58    txtAddress.Text = CD1.FontItalic
59    txtAddress.Text = CD1.FontName
60    txtAddress.Text = CD1.FontSize
61    txtAddress.Text = CD1.FontStrikethru
62    txtAddress.Text = CD1.FontUnderline

63    txtLine2.Text = CD1.FontBold
64    txtLine2.Text = CD1.FontItalic
65    txtLine2.Text = CD1.FontName
66    txtLine2.Text = CD1.FontSize
67    txtLine2.Text = CD1.FontStrikethru
68    txtLine2.Text = CD1.FontUnderline

69    txtCity.Text = CD1.FontBold
70    txtCity.Text = CD1.FontItalic
71    txtCity.Text = CD1.FontName
72    txtCity.Text = CD1.FontSize
73    txtCity.Text = CD1.FontStrikethru
74    txtCity.Text = CD1.FontUnderline

75    txtZipCode.Text = CD1.FontBold
76    txtZipCode.Text = CD1.FontItalic
77    txtZipCode.Text = CD1.FontName
78    txtZipCode.Text = CD1.FontSize
79    txtZipCode.Text = CD1.FontStrikethru
80    txtZipCode.Text = CD1.FontUnderline

81    txtDayTimePhone.Text = CD1.FontBold
82    txtDayTimePhone.Text = CD1.FontItalic
83    txtDayTimePhone.Text = CD1.FontName
84    txtDayTimePhone.Text = CD1.FontSize
85    txtDayTimePhone.Text = CD1.FontStrikethru
86    txtDayTimePhone.Text = CD1.FontUnderline

87    txtEveningPhone.Text = CD1.FontBold
88    txtEveningPhone.Text = CD1.FontItalic
89    txtEveningPhone.Text = CD1.FontName
90    txtEveningPhone.Text = CD1.FontSize
91    txtEveningPhone.Text = CD1.FontStrikethru
92    txtEveningPhone.Text = CD1.FontUnderline

93    txtMobile.Text = CD1.FontBold
94    txtMobile.Text = CD1.FontItalic
95    txtMobile.Text = CD1.FontName
96    txtMobile.Text = CD1.FontSize
97    txtMobile.Text = CD1.FontStrikethru
98    txtMobile.Text = CD1.FontUnderline

99    txtFax.Text = CD1.FontBold
100   txtFax.Text = CD1.FontItalic
101   txtFax.Text = CD1.FontName
102   txtFax.Text = CD1.FontSize
103   txtFax.Text = CD1.FontStrikethru
104   txtFax.Text = CD1.FontUnderline

105   txtWorkPhone.Text = CD1.FontBold
106   txtWorkPhone.Text = CD1.FontItalic
107   txtWorkPhone.Text = CD1.FontName
108   txtWorkPhone.Text = CD1.FontSize
109   txtWorkPhone.Text = CD1.FontStrikethru
110   txtWorkPhone.Text = CD1.FontUnderline

111   txtExt.Text = CD1.FontBold
112   txtExt.Text = CD1.FontItalic
113   txtExt.Text = CD1.FontName
114   txtExt.Text = CD1.FontSize
115   txtExt.Text = CD1.FontStrikethru
116   txtExt.Text = CD1.FontUnderline

117   txtEmail.Text = CD1.FontBold
118   txtEmail.Text = CD1.FontItalic
119   txtEmail.Text = CD1.FontName
120   txtEmail.Text = CD1.FontSize
121   txtEmail.Text = CD1.FontStrikethru
122   txtEmail.Text = CD1.FontUnderline

123   txtNotes.Text = CD1.FontBold
124   txtNotes.Text = CD1.FontItalic
125   txtNotes.Text = CD1.FontName
126   txtNotes.Text = CD1.FontSize
127   txtNotes.Text = CD1.FontStrikethru
128   txtNotes.Text = CD1.FontUnderline                                                                                 '  err%
End Sub

Private Sub mnu_cmd_How_To_Click()                                                 '  err%
frmHelp.Show                                                                                         '  err%
End Sub

Private Sub mnu_cmd_Logoff_Click()
Unload frmMain
frmLogin.Show                                                                        '  err%
End Sub

Private Sub mnu_cmd_Print_Click()
Beep
    MsgBox ("This function has not yet been implemented."), vbOKOnly, "Comming Soon"                                                                             '  err%
End Sub

Private Sub EMPData_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub EMPData_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  EMPData.Caption = "Record: " & CStr(EMPData.Recordset.AbsolutePosition)
End Sub

Private Sub EMPData_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
    Case adRsnAddNew
    
    Case adRsnClose
    
    Case adRsnDelete
    
    Case adRsnFirstChange
    
    Case adRsnMove
    
    Case adRsnRequery
    
    Case adRsnResynch
    
    Case adRsnUndoAddNew
    
    Case adRsnUndoDelete
    
    Case adRsnUndoUpdate
    
    Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub
