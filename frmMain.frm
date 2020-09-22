VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "System Tray"
   ClientHeight    =   510
   ClientLeft      =   4395
   ClientTop       =   1545
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   720
      Top             =   120
   End
   Begin VB.Label lblNothing 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hokay, nothing to see here, people, let's just move along..."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   435
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   6975
      WordWrap        =   -1  'True
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuStartAnimation 
         Caption         =   "Start &Animation"
      End
      Begin VB.Menu mnuStopAnimation 
         Caption         =   "&Stop Animation"
      End
      Begin VB.Menu mnuSelectIcon 
         Caption         =   "Select &Icon"
      End
      Begin VB.Menu SEP01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Window"
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SysTray As CSysTray
Attribute SysTray.VB_VarHelpID = -1
Private Sub Form_Load()
    LoadForm
End Sub
Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        SysTray.MinToSysTray
        Timer1.Enabled = True
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Timer1.Enabled = False
    SysTray.RemoveFromSysTray

End Sub
Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub
Private Sub mnuSelectIcon_Click()
    ' display the form that lets the user select an .ANI file
    frmSearch.Show
End Sub
Private Sub SysTray_RButtonUP()

    'Display popup menu when user presses the right mouse button on the System Tray icon
    PopupMenu Me.mnuPopup
    
End Sub
Private Sub mnuRestore_Click()

    'This restores the BIT Manager application
    Timer1.Enabled = False
    Me.WindowState = vbNormal
    Me.Show
    App.TaskVisible = True
    SysTray.RemoveFromSysTray
   
End Sub
Private Sub mnuExit_Click()

    SysTray.RemoveFromSysTray
    End
    
End Sub
Private Sub mnuStartAnimation_Click()

    Timer1.Enabled = True
    
End Sub
Private Sub mnuStopAnimation_Click()

    Timer1.Enabled = False
    SysTray.ChangeIcon gstrSysTrayIcon
    'SysTray.ChangeIcon App.Path & "\globe.ico"
    
End Sub
Private Sub Timer1_Timer()

    SysTray.ChangeIcon gstrSysTrayIcon
    'SysTray.ChangeIcon App.Path & "\globe.ani"
    
End Sub
Private Sub LoadForm()
    frmSearch.Show
    Set SysTray = New CSysTray
    Set SysTray.SourceWindow = Me
    SysTray.ChangeIcon gstrSysTrayIcon
    'SysTray.ChangeIcon App.Path & "\globe.ani"
    SysTray.ToolTip = Me.Caption
    SysTray.MinToSysTray
    Timer1.Enabled = True
End Sub
