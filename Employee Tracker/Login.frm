VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Tracker Login"
   ClientHeight    =   1470
   ClientLeft      =   2745
   ClientTop       =   3600
   ClientWidth     =   4230
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Carl Weis\My Documents\Employee Tracker\Login.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Login"
      Top             =   1800
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2835
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New"
      Height          =   375
      Left            =   1785
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   780
      TabIndex        =   4
      Top             =   960
      Width           =   930
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label MemID 
      DataField       =   "MemID"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Pass 
      DataField       =   "Pass"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Member ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
                                                         
     Unload Me
                                                                                '  err%
End Sub

Private Sub Command1_Click()
         Data1.Recordset.FindFirst "memID = '" & Text1.Text & "'"
             MsgBox "Login Successful!", vbOKOnly, "Employee Tracker"
             frmMain.Show
             frmLogin.Hide
             Text1.Text = ""
             Text2.Text = ""
            Exit Sub
    
        MsgBox "Login Unsuccessful!", vbOKOnly, "Employee Tracker"
        Text1.Text = ""
        Text2.Text = ""
End Sub

Private Sub Command2_Click()
                                                                    '  err%
         If Text1.Text = "" Then
             MsgBox ("You must enter a user name."), vbOKOnly, "Login Error"
         End If
         If Text2.Text = "" Then
            MsgBox ("You must enter a password"), vbOKOnly, "Login Error"
         End If
         Data1.Recordset.AddNew
         Data1.Recordset.Fields("memID") = "" & Text1.Text & ""
        Data1.Recordset.Fields("pass") = "" & Text2.Text & ""
        Data1.Recordset.Update
End Sub



Private Sub Command4_Click()
         Login.Command5.Visible = True
         Login.Command4.Visible = False
         Login.Width = 3465
End Sub

Private Sub Command5_Click()
         Login.Command4.Visible = True
         Login.Command5.Visible = False
         Login.Width = 5985
                                                                                 '  err%
End Sub

