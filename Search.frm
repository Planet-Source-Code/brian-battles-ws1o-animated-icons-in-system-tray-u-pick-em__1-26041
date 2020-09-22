VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Find Animated Cursor Files (.ANI)"
   ClientHeight    =   7455
   ClientLeft      =   1680
   ClientTop       =   2235
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Search.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "  Folders and Files Found   "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1560
      TabIndex        =   3
      Top             =   285
      Visible         =   0   'False
      Width           =   3540
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   390
         Left            =   1395
         TabIndex        =   6
         Top             =   930
         Width           =   840
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folders"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   450
         TabIndex        =   8
         Top             =   255
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFolderCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   15
         TabIndex        =   7
         Top             =   480
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblProgsCount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   1935
         TabIndex        =   5
         Top             =   480
         Width           =   1530
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Animated Cursors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   1905
         TabIndex        =   4
         Top             =   255
         Width           =   1590
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   390
      Left            =   5715
      TabIndex        =   1
      Top             =   7080
      Width           =   780
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6690
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   6240
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Find"
      Height          =   390
      Left            =   4890
      TabIndex        =   0
      Top             =   7080
      Width           =   780
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   165
      TabIndex        =   9
      Top             =   6510
      Width           =   6195
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancel          As Boolean
Dim iPHits           As Integer
Dim iFHits           As Integer
Dim bFindAllFiles    As Boolean
Dim BeginTime        As Variant
Dim EndTime          As Variant
Dim SearchTime       As Variant
Sub ListSubDirs(Path)
    
    On Error GoTo Err_ListSubDirs
    
    Dim D()      As Variant
    Dim Count    As Integer
    Dim I        As Integer
    Dim DirName  As String
    Dim MyName   As String
    Dim MyEXE    As String
    Dim F        As Integer
    Dim Matched  As Boolean
    Dim InitAttr As Long
    
    Count = 0
    DirName = Dir(Path, vbDirectory) ' Get first directory name
    'Iterate through PATH, caching all subdirectories in D()
    Do While DirName <> ""
        GoSub CheckForCancel
        If DirName <> "." And DirName <> ".." Then
            If DirName <> "pagefile.sys" Then    ' (for NT)
                InitAttr = GetAttr(Path & DirName)
                InitAttr = InitAttr And 16
                If InitAttr = vbDirectory Then
                    Matched = False
                    Count = Count + 1   ' Increment counter.
                    ReDim Preserve D(Count)    ' Resize the array.
                    D(Count) = DirName
                    iFHits = iFHits + 1
                    lblFolderCount.Caption = iFHits
                End If
            End If
        End If
        DirName = Dir   ' Get another directory name.
        DoEvents
    Loop
    For I = 1 To Count
        GoSub CheckForCancel
        ' Now recursively iterate through each cached subdirectory.
        MyName = Path & D(I) & "\"
        Call ListSubDirs(MyName)
        ' AND see if any EXE files
        MyEXE = Dir(MyName)
        Do While MyEXE <> ""
            If bCancel Then
                Exit Sub
            End If
            If UCase(Right(MyEXE, 3)) = "ANI" Then
                List1.AddItem MyName & MyEXE
                iPHits = iPHits + 1
                lblProgsCount.Caption = iPHits
            Else
            End If
            MyEXE = Dir
            DoEvents
        Loop
        DoEvents
    Next I
    Exit Sub
    
Exit_ListSubDirs:
    
    Exit Sub
    
Err_ListSubDirs:
    
    Select Case Err
        Case 0, 52
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & " - Advisory..."
            Resume Exit_ListSubDirs
    End Select
    
CheckForCancel:
    
    DoEvents
    If bCancel Then
        If MsgBox("Okay to end the search?", 36) = vbYes Then
            Exit Sub
        Else
            bCancel = False
            Return
        End If
    End If
    Return
    
End Sub
Private Sub cmdCancel_Click()
    
    On Error GoTo Err_cmdCancel_Click
    
    bCancel = True
    
Exit_cmdCancel_Click:
    
    Exit Sub
    
Err_cmdCancel_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & " - Advisory..."
            Resume Exit_cmdCancel_Click
    End Select
    
End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Err_cmdCancel_MouseMove
    
    Screen.MousePointer = vbDefault
    
Exit_cmdCancel_MouseMove:
    
    Exit Sub
    
Err_cmdCancel_MouseMove:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & " - Advisory..."
            Resume Exit_cmdCancel_MouseMove
    End Select
    
End Sub
Private Sub cmdExit_Click()
    
    On Error GoTo Err_cmdExit_Click
    
    Unload Me
    DoEvents
    End
    
Exit_cmdExit_Click:
    
    Exit Sub
    
Err_cmdExit_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & " - Advisory..."
            Resume Exit_cmdExit_Click
    End Select
    
End Sub
Private Sub cmdRun_Click()
    
    On Error GoTo Err_cmdRun_Click
    
    BeginTime = Now()
    Me.Caption = "Find Animated Cursor Files (.ANI)    - " & Format$(Now(), "DDDD,  MMMM d, yyyy")
    cmdRun.Visible = False
    Frame1.Visible = True
    bCancel = False
    List1.Clear
    iPHits = 0
    iFHits = 0
    lblProgsCount.Caption = iPHits
    lblFolderCount.Caption = iFHits
    lblResults.Caption = "" ' # progs found
    Screen.MousePointer = vbHourglass
    DoEvents
    Call ListSubDirs("C:\")    ' Call ListSubDirs
    cmdRun.Visible = True
    Frame1.Visible = False
    Screen.MousePointer = vbDefault
    EndTime = Now()
    SearchTime = Format((EndTime - BeginTime), "n:ss")
    lblResults.Caption = iPHits & " .ANI files listed, search took " & SearchTime
    
Exit_cmdRun_Click:
    
    Exit Sub
    
Err_cmdRun_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & " - Advisory..."
            Resume Exit_cmdRun_Click
    End Select
    
End Sub
Private Sub Form_Load()
    
    Dim recIn
    
    On Error Resume Next
    
    Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2
    Frame1.Move (Screen.Width / 2 - Me.Width / 2) / 4, (Screen.Height / 2 - Me.Height / 2) * 2.5
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Err_Frame1_MouseMove
    
    Screen.MousePointer = vbHourglass
    
Exit_Frame1_MouseMove:
    
    Exit Sub
    
Err_Frame1_MouseMove:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & " - Advisory..."
            Resume Exit_Frame1_MouseMove
    End Select
    
End Sub
Private Sub List1_Click()
    
    On Error GoTo Err_List1_Click
    
    gstrSysTrayIcon = List1.Text
    Unload Me
    
Exit_List1_Click:
    
    Exit Sub
    
Err_List1_Click:
    
    Select Case Err
        Case 0
            Resume Next
        Case Else
            MsgBox "Error Code: " & Err.Number & vbCrLf & vbCrLf & Err.Description, vbInformation, App.Title & " - Advisory..."
            Resume Exit_List1_Click
    End Select
    
End Sub
