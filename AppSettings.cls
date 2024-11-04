VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const INI_FILE = "\TimedProgramLauncher.ini"

Private mLaunchInterval As Long
Private mProgramName As String
Private mWindowStyle As VbAppWinStyle
Private mDirty As Boolean
Private mRedZone As Long
Private mBeepOnLaunchError As Boolean

Public Property Get LaunchInterval() As Long
   LaunchInterval = mLaunchInterval
End Property

Public Property Let LaunchInterval(ByVal vdata As Long)
   On Error Resume Next
   If mLaunchInterval <> vdata Then
      mLaunchInterval = vdata
      mDirty = True
   End If
End Property

Public Property Get ProgramName() As String
   ProgramName = mProgramName
End Property

Public Property Let ProgramName(ByVal vdata As String)
   On Error Resume Next
   If mProgramName <> Trim(vdata) Then
      mProgramName = Trim(vdata)
      mDirty = True
   End If
End Property

Public Property Get WindowStyle() As VbAppWinStyle
   WindowStyle = mWindowStyle
End Property

Public Property Let WindowStyle(ByVal vdata As VbAppWinStyle)
   On Error Resume Next
   If mWindowStyle <> vdata Then
      mWindowStyle = vdata
      mDirty = True
   End If
End Property

Public Property Get Dirty() As Boolean
   Dirty = mDirty
End Property

Public Property Get RedZone() As Long
   RedZone = mRedZone
End Property

Public Property Get BeepOnLaunchError() As Boolean
   BeepOnLaunchError = mBeepOnLaunchError
End Property

Public Function Save() As Boolean
   Dim FileName As String
   Dim FileNum As Integer
   
   On Error GoTo errorHandler
   
   FileName = App.Path & INI_FILE
   FileNum = FreeFile
   Open FileName For Output Access Write Lock Write As FileNum
   Print #FileNum, "[AppSettings]"
   Print #FileNum, "LaunchInterval="; CStr(LaunchInterval)
   Print #FileNum, "ProgramName="; ProgramName
   Print #FileNum, "WindowStyle="; CStr(WindowStyle)
   Print #FileNum, "RedZone="; CStr(mRedZone)
   Print #FileNum, "BeepOnLaunchError="; CStr(CLng(mBeepOnLaunchError))
   Close FileNum
   Save = True
   mDirty = False
   Exit Function
errorHandler:
   Save = False
   On Error Resume Next
   Close FileNum
End Function

Public Function Initialize() As Boolean
   Dim FileName As String
   Dim ExeName As String * 256
   Dim WinStyle As String * 256
   Dim s As String
   Dim ReturnCode As Long
   
   On Error GoTo errorHandler
   FileName = App.Path & INI_FILE
   If FileExists(FileName) Then
      LaunchInterval = GetPrivateProfileInt("AppSettings", "LaunchInterval", 120, FileName)
      WindowStyle = GetPrivateProfileInt("AppSettings", "WindowStyle", 0, FileName)
      mBeepOnLaunchError = GetPrivateProfileInt("AppSettings", "BeepOnLaunchError", -1, FileName)
      mRedZone = GetPrivateProfileInt("AppSettings", "RedZone", 10, FileName)
      
      ReturnCode = GetPrivateProfileString("AppSettings", "ProgramName", "", ExeName, 256, FileName)
      If ReturnCode > 0 Then
         ProgramName = TrimNulls(ExeName)
      End If
   Else
      MsgBox "Cannot find existing program initialization settings. Defaults will be used.", vbInformation, APP_NAME
      Save
   End If
   Initialize = True
   mDirty = False
   Exit Function
errorHandler:
   Initialize = False
End Function

Private Sub Class_Initialize()
   On Error Resume Next
   LaunchInterval = 120
   ProgramName = ""
   WindowStyle = vbMinimizedNoFocus
   mRedZone = 10
   mBeepOnLaunchError = True
   mDirty = False
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

Private Function TrimNulls(ByVal s As String) As String
   Dim NullPos As Long
   
   On Error GoTo errorHandler
   
   NullPos = InStr(1, s, Chr$(0))
   TrimNulls = Left(s, NullPos - 1)
   
   Exit Function
errorHandler:
   TrimNulls = ""
End Function
