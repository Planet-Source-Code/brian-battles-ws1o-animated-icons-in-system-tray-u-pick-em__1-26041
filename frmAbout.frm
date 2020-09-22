VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   About..."
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "Enough Already!"
      Height          =   405
      Left            =   2805
      TabIndex        =   3
      Top             =   2040
      Width           =   1530
   End
   Begin VB.Label lblBB 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      ForeColor       =   &H00000080&
      Height          =   660
      Left            =   53
      TabIndex        =   2
      Top             =   1245
      Width           =   7035
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Click here to see the original project on www.planetsourcecode.com!"
      Top             =   615
      Width           =   7050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   7020
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const cstrAbout  As String = "A Cool System Tray Animation Project" & vbCrLf & "Written by Todd Herman in VB 6.0"
Private Const cstrPscUrl As String = "http://www.planetsourcecode.com/xq/ASP/txtCodeId.26016/lngWId.1/qx/vb/scripts/ShowCode.htm"
Private Const cstrBB     As String = "Modified August 9, 2001, By Brian Battles WS1O" & vbCrLf & "brianb@cmtelephone.com"
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    lblAbout.Caption = cstrAbout
    lblURL.Caption = "Click Here to See Todd's Original Code!"
    lblBB.Caption = cstrBB
End Sub
Private Sub lblURL_Click()
    ShellExecute Me.hWnd, vbNullString, cstrPscUrl, vbNullString, "C:\", SW_SHOWMAXIMIZED
    Unload Me
End Sub
