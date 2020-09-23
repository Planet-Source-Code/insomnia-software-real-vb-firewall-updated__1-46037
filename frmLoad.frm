VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   0  'None
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2.5
   ScaleMode       =   5  'Inch
   ScaleWidth      =   4.948
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   3000
      Top             =   1560
   End
   Begin CyberSentry.ProgressCntrl pbLoad 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your computer is now protected from the internet!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1320
      Left            =   157
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   6810
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Professional 1.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1920
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Visible = True
pbLoad.Caption = "Loading CyberSentry Personal Firewall"
pbLoad.Max = 10
DoEvents
Label2.Caption = App.Comments & " " & App.Major & "." & App.Minor & App.Revision
pbLoad.Value = 1
DoEvents
If GetSetting(App.Title, "Settings", "HideWinServices", "True") = "True" Then frmMain.Check1.Value = vbChecked
pbLoad.Value = 2
DoEvents
If GetSetting(App.Title, "Settings", "InetOnlyServices", "True") = "True" Then frmMain.Check2.Value = vbChecked
pbLoad.Value = 3
DoEvents
Open App.Path & "\system.dat" For Input As #1
Do While EOF(1) = False
Line Input #1, a$
strSysFiles = strSysFiles & a$ & vbCrLf
DoEvents
Loop
Close #1
pbLoad.Value = 4
DoEvents
'strSysFiles = Replace(strSysFiles, vbCrLf, " ")
pbLoad.Value = 5
DoEvents
Load frmMinimize
pbLoad.Value = 6
DoEvents
Load frmExit
pbLoad.Value = 7
DoEvents
Load frmAlertApp
pbLoad.Value = 8
DoEvents
Load frmMain
pbLoad.Value = 9
DoEvents
frmMain.SysTray.InTray = True
pbLoad.Value = 10
DoEvents
Label1.Visible = False
Label2.Visible = False
pbLoad.Visible = False
Label3.Visible = True
DoEvents
Timer1.Enabled = True


End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub


