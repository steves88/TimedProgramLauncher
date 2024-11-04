VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Launcher"
   ClientHeight    =   4440
   ClientLeft      =   8280
   ClientTop       =   6330
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7335
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   435
      Left            =   120
      TabIndex        =   17
      Top             =   3900
      Width           =   1035
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   6180
      TabIndex        =   12
      Top             =   3900
      Width           =   1035
   End
   Begin TabDlg.SSTab tab1 
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6482
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Operation"
      TabPicture(0)   =   "frmMain.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCountDown"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNextLaunch"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCurrentTime"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdStart"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Configuration"
      TabPicture(1)   =   "frmMain.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "cmdSaveSettings"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "Window Style"
         Height          =   1155
         Left            =   -74880
         TabIndex        =   13
         Top             =   1860
         Width           =   6855
         Begin VB.OptionButton obMinimizedNoFocus 
            Caption         =   "Minimized Without Focus"
            Height          =   315
            Left            =   240
            TabIndex        =   19
            Top             =   660
            Width           =   2415
         End
         Begin VB.OptionButton obNormalFocus 
            Caption         =   "Normal With Focus"
            Height          =   315
            Left            =   2820
            TabIndex        =   18
            Top             =   300
            Width           =   2415
         End
         Begin VB.OptionButton obNormalNoFocus 
            Caption         =   "Normal Without Focus"
            Height          =   315
            Left            =   2820
            TabIndex        =   16
            Top             =   660
            Width           =   2415
         End
         Begin VB.OptionButton obMaximized 
            Caption         =   "Maximized"
            Height          =   315
            Left            =   5340
            TabIndex        =   15
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton obMinimizedFocus 
            Caption         =   "Minimized With Focus"
            Height          =   315
            Left            =   240
            TabIndex        =   14
            Top             =   300
            Width           =   2415
         End
      End
      Begin VB.CommandButton cmdSaveSettings 
         Caption         =   "&Save Settings"
         Height          =   435
         Left            =   -69780
         TabIndex        =   8
         Top             =   3120
         Width           =   1755
      End
      Begin VB.Frame Frame1 
         Caption         =   "Launch Details"
         Height          =   1275
         Left            =   -74880
         TabIndex        =   2
         Top             =   420
         Width           =   6855
         Begin VB.TextBox txtLaunchInterval 
            Height          =   315
            Left            =   2100
            TabIndex        =   5
            Top             =   300
            Width           =   675
         End
         Begin VB.TextBox txtProgramName 
            Height          =   315
            Left            =   2100
            TabIndex        =   4
            Top             =   720
            Width           =   4275
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            Height          =   315
            Left            =   6420
            TabIndex        =   3
            Top             =   720
            Width           =   315
         End
         Begin VB.Label lblLaunchInterval 
            Caption         =   "Launch Interval (secs):"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblExeName 
            Caption         =   "Program Path\Name:"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   780
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Start"
         Height          =   615
         Left            =   5460
         TabIndex        =   1
         Top             =   2940
         Width           =   1515
      End
      Begin VB.Label lblCurrentTime 
         Alignment       =   1  'Right Justify
         Caption         =   "hh:mm:ss"
         Height          =   255
         Left            =   5640
         TabIndex        =   11
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lblNextLaunch 
         Caption         =   "Next Launch:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   300
         TabIndex        =   10
         Top             =   1620
         Width           =   2715
      End
      Begin VB.Label lblCountDown 
         Caption         =   "99999"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3060
         TabIndex        =   9
         Top             =   1620
         Width           =   1215
      End
   End
   Begin VB.Timer LaunchTimer 
      Interval        =   1000
      Left            =   3240
      Top             =   3900
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   3900
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer clockTimer 
      Interval        =   1000
      Left            =   4260
      Top             =   3900
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const TIME_CONV_FACTOR As Double = 1 / 24 / 60 / 60

Private mLaunchInterval As Long
Private mLastTime As Date
Private mAppSettings As AppSettings
Private mCountdown As Long

Private Sub cmdAbout_Click()
   On Error Resume Next
   frmAbout.Show
End Sub

Private Sub Form_Load()
   On Error Resume Next
   
   ' Center the form
   Me.Left = (Screen.Width - Me.Width) \ 2
   Me.Top = (Screen.Height - Me.Height) \ 2
   
   lblCurrentTime.Caption = Format(Now, "h:mm:ss AMPM")
   clockTimer.Enabled = True
   LaunchTimer.Enabled = False
   Set mAppSettings = New AppSettings
   If mAppSettings.Initialize Then
      txtLaunchInterval.Text = mAppSettings.LaunchInterval
      mLaunchInterval = mAppSettings.LaunchInterval
      lblCountDown.Caption = CStr(mLaunchInterval)
      txtProgramName.Text = mAppSettings.ProgramName
      Select Case mAppSettings.WindowStyle
         Case vbMaximizedFocus
            obMaximized.Value = True
         
         Case vbMinimizedFocus
            obMinimizedFocus.Value = True
         
         Case vbMinimizedNoFocus
            obMinimizedNoFocus.Value = True
         
         Case vbNormalFocus
            obNormalFocus.Value = True
            
         Case vbNormalNoFocus
            obNormalNoFocus.Value = True
               
         Case Else
            obMinimizedNoFocus.Value = True
      End Select
   Else
      MsgBox "An error occurred while loading the program initialization settings. The program will continue with default settings.", vbExclamation, APP_NAME
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim Response As Long
   
   On Error Resume Next
   
   If UnloadMode = 0 And LaunchTimer.Enabled = True Then
      Cancel = 1
   Else
      If mAppSettings.Dirty Then
         Response = MsgBox("You have un-saved configuration settings. Would you like to save your changed before exiting the program?", vbQuestion + vbYesNoCancel + vbDefaultButton1, APP_NAME)
         Select Case Response
            Case vbYes
               cmdSaveSettings_Click
               ' The dirty flag will be false if the Save operation was successful.
               ' If not successful, then don't close the program. Give the user a
               ' chance to fix the problem before exiting.
               If mAppSettings.Dirty = False Then
                  Cancel = 0
               Else
                  Cancel = 1
               End If
            
            Case vbNo
               Cancel = 0
            
            Case vbCancel
               Cancel = 1
         End Select
      Else
         Cancel = 0
      End If
   End If
End Sub

Private Sub clockTimer_Timer()
   On Error Resume Next
   lblCurrentTime.Caption = Format(Now, "h:mm:ss AMPM")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   End
End Sub

Private Sub LaunchTimer_Timer()
   On Error Resume Next
   mCountdown = mCountdown - (1 - (Abs(mLastTime - Now) * TIME_CONV_FACTOR))
   If mCountdown >= 0 Then
      lblCountDown.Caption = CStr(mCountdown)
      frmMain.Caption = "Program Launcher (" & CStr(mCountdown) & ")"
      If mCountdown > mAppSettings.RedZone Then
         lblNextLaunch.ForeColor = &H80000012
         lblCountDown.ForeColor = &H80000012
      Else
         lblNextLaunch.ForeColor = &HC0&
         lblCountDown.ForeColor = &HC0&
      End If
   Else
      lblNextLaunch.ForeColor = &HC0&
      lblCountDown.ForeColor = &HC0&
      lblCountDown.Caption = "0"
      frmMain.Caption = "Program Launcher (0)"
   End If
   If Now >= mLastTime + (mLaunchInterval * TIME_CONV_FACTOR) Then
      mLastTime = Now()
      LaunchProgram
      mCountdown = mLaunchInterval
      lblNextLaunch.ForeColor = &H80000012
      lblCountDown.ForeColor = &H80000012
      lblCountDown.Caption = CStr(mLaunchInterval)
      frmMain.Caption = "Program Launcher (" & CStr(mLaunchInterval) & ")"
   End If
End Sub

Private Sub LaunchProgram()
   On Error GoTo errorHandler
   Shell mAppSettings.ProgramName, mAppSettings.WindowStyle
   Exit Sub
errorHandler:
   On Error Resume Next
   If mAppSettings.BeepOnLaunchError Then
      Beep
   End If
End Sub

Private Sub cmdBrowse_Click()
   On Error GoTo errorHandler
   With CommonDialog1
      .DialogTitle = "Select Program to Run"
      .InitDir = "c:\"
      .Filter = "Exe Files (*.exe)|*.exe|Batch Files (*.bat)|*.bat|Cmd Files (*.cmd)|*.cmd"
      .FilterIndex = 1
      .ShowOpen
      If .CancelError = True Then
         Exit Sub
      Else
         txtProgramName.Text = .FileName
      End If
   End With
   Exit Sub
errorHandler:
   
End Sub

Private Sub cmdExit_Click()
   On Error Resume Next
   Unload Me
End Sub

Private Sub cmdSaveSettings_Click()
   
   On Error Resume Next
   
   ' Refresh the settings.
   ' Note: The window style property of the AppSettings object
   '       is set on the fly when the user chooses a diffent option
   '       button.
   mAppSettings.LaunchInterval = CLng(txtLaunchInterval.Text)
   mAppSettings.ProgramName = Trim(txtProgramName.Text)
   
   If mAppSettings.Save Then
      MsgBox "Settings successfully saved.", vbInformation, APP_NAME
   Else
      MsgBox "The following error occurred while saving the application settings: " & vbCrLf & Err.Description, vbExclamation, APP_NAME
   End If
End Sub

Private Sub cmdStart_Click()
   On Error Resume Next
   If LaunchTimer.Enabled Then
      frmMain.Caption = "Program Launcher (" & txtLaunchInterval.Text & ")"
      lblCountDown.Caption = txtLaunchInterval.Text
      LaunchTimer.Enabled = False
      cmdStart.Caption = "&Start"
      EnableControls
   Else
      mCountdown = mLaunchInterval
      mLastTime = Now()
      LaunchTimer.Enabled = True
      cmdStart.Caption = "&Stop"
      DisableControls
   End If
   
End Sub

Public Function IsValidPath(ByVal Path As String) As Boolean
   Dim ReturnCode As Variant
   Dim CurrentDirectory As Variant
   
   On Error GoTo errorHandler
   
   ' If the path string is null, then don't bother to check
   ' anything, just return False.
   If Path = "" Then
      IsValidPath = False
      Exit Function
   End If
   
   CurrentDirectory = CurDir
   
   ' Check the path to see if it is valid. An error will occur
   ' if the path submitted is not valid. Otherwise, set return
   ' value to True, and change the directory back to where it
   ' was before the ChDir call.
   ChDir Path
   IsValidPath = True
   On Error Resume Next
   ChDir CurrentDirectory
   Exit Function
errorHandler:
   IsValidPath = False
   On Error Resume Next
   ChDir CurrentDirectory
End Function

Private Sub ValidateInteger(ByRef KeyAscii As Integer)
' Check to make sure that the value is numeric
   If KeyAscii < 48 Or KeyAscii > 57 Then
      ' check for back space and minus sign
      If KeyAscii <> 8 And KeyAscii <> 45 Then
         KeyAscii = 0
      End If
   End If
End Sub

Private Function GetChildPath(ByVal Path As String) As String
   Dim Index As Integer
   Dim c As String
   Dim s As String
   Dim StrLen As Integer
   
   On Error GoTo errorHandler
   
   StrLen = Len(Path)
   
   ' Check for a root only path (i.e., C:\, D:\, ...) and an invalid path.
   ' If found, don't return anything.
   If StrLen > 3 And IsValidPath(Path) Then
   ' Step through the string backwards until the first backslash
   ' charater is found. Return the portion of the string up
   ' to that point.
   For Index = StrLen To 1 Step -1
      c = Mid(Path, Index, 1)
      If c <> "\" Then
         s = c & s
      Else
         GetChildPath = Trim(s)
         Exit For
      End If
   Next Index
   Else
      GetChildPath = ""
   End If
   Exit Function
errorHandler:
   GetChildPath = ""
End Function

Private Function FileIncludesPath(ByVal FilePath As String) As Boolean

   On Error GoTo errorHandler
   
   If InStr(1, FilePath, "\") > 0 Then
      FileIncludesPath = True
   Else
      FileIncludesPath = False
   End If
   Exit Function
errorHandler:
   
End Function

Private Function GetBareFileName(ByVal FileName As String) As String
   Dim StrLen As Integer
   Dim Index As Integer
   Dim fName As String
   Dim c As String
   
   On Error GoTo errorHandler
   
   StrLen = Len(FileName)
   If StrLen > 0 Then
      If FileIncludesPath(FileName) Then
         For Index = StrLen To 1 Step -1
            c = Mid(FileName, Index, 1)
            If c <> "\" Then
               fName = c & fName
            Else
               Exit For
            End If
         Next Index
         GetBareFileName = fName
      Else
         GetBareFileName = FileName
      End If
   Else
      GetBareFileName = ""
   End If
   Exit Function
errorHandler:
   GetBareFileName = ""
End Function

Private Sub DisableControls()
   On Error Resume Next
   tab1.TabEnabled(1) = False
   cmdExit.Enabled = False
   cmdAbout.Enabled = False
End Sub

Private Sub EnableControls()
   On Error Resume Next
   tab1.TabEnabled(1) = True
   cmdExit.Enabled = True
   cmdAbout.Enabled = True
End Sub

Private Sub obMaximized_Click()
   On Error Resume Next
   mAppSettings.WindowStyle = vbMaximizedFocus
End Sub

Private Sub obMinimizedFocus_Click()
   On Error Resume Next
   mAppSettings.WindowStyle = vbMinimizedFocus
End Sub

Private Sub obMinimizedNoFocus_Click()
   On Error Resume Next
   mAppSettings.WindowStyle = vbMinimizedNoFocus
End Sub

Private Sub obNormalFocus_Click()
   On Error Resume Next
   mAppSettings.WindowStyle = vbNormalFocus
End Sub

Private Sub obNormalNoFocus_Click()
   On Error Resume Next
   mAppSettings.WindowStyle = vbNormalNoFocus
End Sub

Private Sub txtLaunchInterval_Change()
   On Error Resume Next
   mLaunchInterval = CLng(txtLaunchInterval.Text)
   lblCountDown.Caption = CStr(mLaunchInterval)
End Sub

Private Sub txtLaunchInterval_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   ValidateInteger KeyAscii
End Sub

Private Sub txtProgramName_LostFocus()
   On Error Resume Next
   If txtProgramName.Text <> "" Then
      If Not FileExists(txtProgramName.Text) Then
         MsgBox "The path\file name you specified does not exist. Please try again.", vbExclamation, APP_NAME
         txtProgramName.SetFocus
      Else
         mAppSettings.ProgramName = txtProgramName.Text
      End If
   End If
End Sub

Private Function FileExists(ByVal FilePath As String) As Boolean
   
   On Error GoTo errorHandler
   
   If Dir(FilePath) <> "" Then
      FileExists = True
   Else
      FileExists = False
   End If
   Exit Function
errorHandler:
   FileExists = False
End Function


