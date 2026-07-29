VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Text            =   "Notify Title"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Text            =   "Save files"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4560
      Top             =   2280
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   975
      Left            =   3960
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Notify"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show SysTray"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "error icon"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "warning icon"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "info icon"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "no icon"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4215
   End
   Begin VB.Menu zmnuSysTrayDemo 
      Caption         =   ""
      Begin VB.Menu mnuFile 
         Caption         =   "Start Shooting"
         Index           =   0
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Stop Shooting"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   +{F3}
      End
      Begin VB.Menu mnuFile 
         Caption         =   "View Screenshots"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Purge Screenshots"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Preferences"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "About this App"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit"
         Index           =   6
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const NOTIFYICON_VERSION = &H3

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

'icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long
   
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
   Alias "Shell_NotifyIconA" _
  (ByVal dwMessage As Long, _
   lpData As NOTIFYICONDATA) As Long
   
Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)
   
   'begin code for shot
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Public Function KillFolder(ByVal FullPath As String) _
   As Boolean

'PARAMETER: FullPath = FullPath of Folder to Delete

'RETURNS:   True is successful, false otherwise

'EXAMPLE:   'KillFolder("D:\MyOldFiles")

'******************************************
On Error Resume Next
Dim oFso As New Scripting.FileSystemObject

'deletefolder method does not like the "\"
'at end of fullpath

If Right(FullPath, 1) = "\" Then FullPath = _
    Left(FullPath, Len(FullPath) - 1)

If oFso.FolderExists(FullPath) Then
    
    'Setting the 2nd parameter to true
    'forces deletion of read-only files
    oFso.DeleteFolder FullPath, True
    
    KillFolder = Err.Number = 0 And _
      oFso.FolderExists(FullPath) = False
End If

End Function
Private Sub GetShot(lWindowhWnd As Long)
    Dim nLeft As Long
    Dim nTop As Long
    Dim nWidth As Long
    Dim nHeight As Long
    Dim rRect As RECT
    Dim bm As BITMAP
    Dim lWindowhDC As Long
    'check if the screenshot folder exists.
    If Dir(App.Path & "\screenshots", vbDirectory) = "" Then
      'directory doesn't exist...create it
      MkDir App.Path & "\screenshots"
    End If
    'resize the picturebox to screen resolution
    Picture1.Width = Screen.Width
    Picture1.Height = Screen.Height
    
    Hide
    Picture1.Cls
    Set Picture1.Picture = Nothing
    GetWindowRect lWindowhWnd, rRect
    lWindowhDC = GetWindowDC(lWindowhWnd)
    '// Get coordinates
    nLeft = 0
    nTop = 0
    nWidth = rRect.Right - rRect.Left
    nHeight = rRect.Bottom - rRect.Top
    '// Blt to frm.Picture1
    BitBlt Picture1.hDC, 0, 0, nWidth, nHeight, lWindowhDC, nLeft, nTop, vbSrcCopy
    '// Del DC
    ReleaseDC lWindowhWnd, lWindowhDC
    '// set picture
    Picture1.Picture = Picture1.Image
   ' Show
    '//save picture file name
    Text2.Text = App.Path & Replace(Replace(Replace("\screenshots\" + FormatDateTime(Now), ":", "-"), " ", "_"), "/", "-")
     '//save BMP
    SavePicture Picture1.Picture, Text2.Text
   '//Convert to JPG
    Dim ImgF As WIACtl.ImageFile
    Dim ImgP As WIACtl.ImageProcess
    Set ImgF = New WIACtl.ImageFile
    ImgF.LoadFile Text2.Text
    Set ImgP = New WIACtl.ImageProcess
    With ImgP
    .Filters.Add .FilterInfos!Convert.FilterID
    .Filters.Item(1).Properties!FormatID.Value = wiaFormatJPEG
    .Filters.Item(1).Properties!Quality.Value = 70
    Set ImgF = .Apply(ImgF)
    End With
    ImgF.SaveFile Text2.Text & ".jpg"
    Kill Text2.Text
End Sub
'end of code for shot

      


Private Sub Form_Load()

   Text1.Text = "ScreenRifle notification"
   Text3.Text = "ScreenRifle"
   Command1.Caption = "Add Systray Icon"
   Command2.Caption = "Show Balloon Tip"
   Command2.Enabled = False
   Option1(0).Caption = "no icon"
   Option1(1).Caption = "information icon"
   Option1(2).Caption = "warning icon"
   Option1(3).Caption = "error icon"
   Option1(1).Value = True
   
   'settings
   Timer1.Enabled = False
   Form1.Visible = False
   Call Command1_Click
   'end settings
End Sub


Private Sub Form_Unload(Cancel As Integer)

  'Remove the icon added to the taskbar
   ShellTrayRemove
   
  'remove subclassing
   UnSubClass Me.hwnd
       
  'ensure unloading proceeds
   Cancel = False
   
End Sub


Private Sub Command1_Click()
   
   Call ShellTrayAdd
   Command2.Enabled = True
   
End Sub


Private Sub Command2_Click()
   
   ShellTrayModifyTip GetSelectedOptionIndex()
   
End Sub


Private Sub ShellTrayAdd()
   
   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = Me.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .dwState = NIS_SHAREDICON
      .hIcon = Form1.Icon
      .szTip = "ScreenRifle" & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
      .uCallbackMessage = WM_MYHOOK
   End With
   
  'add the icon ...
   If Shell_NotifyIcon(NIM_ADD, nid) = 1 Then
   
     '... and inform the system of the
     'NOTIFYICON version in use
      Call Shell_NotifyIcon(NIM_SETVERSION, nid)
      
     'prepare to receive the systray messages
      SubClass Me.hwnd
      
   End If
       
End Sub

Private Sub ShellTrayRemove()

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = Form1.hwnd
      .uID = APP_SYSTRAY_ID
   End With
   
   Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub


Private Sub ShellTrayModifyTip(nIconIndex As Long)

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hwnd = Form1.hwnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      .szInfoTitle = Text3.Text & vbNullChar
      .szInfo = Text1.Text & vbNullChar
   End With

   Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub


Private Sub UnSubClass(hwnd As Long)

  'restore the default message handling
  'before exiting
   If defWindowProc <> 0 Then
      SetWindowLong hwnd, GWL_WNDPROC, defWindowProc
      defWindowProc = 0
   End If
   
End Sub

Private Sub SubClass(hwnd As Long)

  'assign our own window message
  'procedure (WindowProc)
   On Error Resume Next
   defWindowProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
   
End Sub


Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0 structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub


Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function


Private Function GetSelectedOptionIndex() As Long

  'returns the selected item index from
  'an option button array. Use in place
  'of multiple If...Then statements!
  'If your array contains more elements,
  'just append them to the test condition,
  'setting the multiplier to the button's
  'negative -index.
   GetSelectedOptionIndex = Option1(0).Value * 0 Or _
                            Option1(1).Value * -1 Or _
                            Option1(2).Value * -2 Or _
                            Option1(3).Value * -3
End Function

'System Tray Right-click menu
Private Sub mnuFile_Click(Index As Integer)
'Which item was clicked?
   Select Case Index
      'Clicked Start Shooting
      Case 0:
      Timer1.Enabled = True
      mnuFile(0).Enabled = False
      mnuFile(1).Enabled = True

      Option1(1).Enabled = True
      Text3.Text = "Shooting started" & vbNullChar
      Text1.Text = "Screenshot shooting has started." & vbNullChar
      ShellTrayModifyTip GetSelectedOptionIndex()
      'clicked Stop Shooting
      Case 1:
      Timer1.Enabled = False
      mnuFile(0).Enabled = True
      mnuFile(1).Enabled = False
      Option1(1).Enabled = True
      Text3.Text = "Shooting stopped" & vbNullChar
      Text1.Text = "Shooting has stopped. Your screenshots have been saved." & vbNullChar
      ShellTrayModifyTip GetSelectedOptionIndex()
      'clicked View Screenshots
      Case 2:
      If Dir(App.Path & "\screenshots", vbDirectory) = "" Then
      Option1(1).Enabled = True
      Text3.Text = "ScreenRifle" & vbNullChar
      Text1.Text = "No screenshots were found." & vbNullChar
      ShellTrayModifyTip GetSelectedOptionIndex()
      Else
      Shell "explorer " & App.Path & "\screenshots", vbMaximizedFocus
      End If
      'clicked Purge
      Case 3:
      'check if dir exists
      If Dir(App.Path & "\screenshots", vbDirectory) = "" Then
      'directory doesn't exist...
      Option1(3).Enabled = True
      Text3.Text = "ScreenRifle" & vbNullChar
      Text1.Text = "Nothing deleted. No screenshots were found." & vbNullChar
      'directory exists
      Else
      Option1(2).Enabled = True
      Text3.Text = "ScreenRifle" & vbNullChar
      'ask the user
      Dim a As String
      a = MsgBox("Are you sure you want to DELETE all saved screenshots? You CANNOT undo this.", vbYesNo, "ScreenRifle")
      'user wants to kill
      If a = vbYes Then
      'check if shooter running. its not
      If Timer1.Enabled = False Then
      Text1.Text = "Your saved screenshots have successfully been deleted." & vbNullChar
      KillFolder (App.Path & "\screenshots")
      'shooter running
      Else
      Text1.Text = "Your saved screenshots have successfully been deleted. Shooter is still running." & vbNullChar
      Timer1.Enabled = False
      KillFolder (App.Path & "\screenshots")
      Timer1.Enabled = True
      End If
      Else
      Text1.Text = "No screenshots were deleted." & vbNullChar
      End If
      End If
      ShellTrayModifyTip GetSelectedOptionIndex()
    
      Case 4, 5:
         
         MsgBox "Called from File " & mnuFile(Index).Caption
      
      Case 6:
      
        'Executing 'Unload Me' from within a
        'menu event invoked from a systray icon
        'will cause a GPF. The proper way to
        'terminate under these circumstances
        'is to send a WM_CLOSE message to the
        'form. The form will process the
        'message as though the user had selected
        'Close from the sysmenu, invoking the
        'normal chain of shutdown events, removing
        'the tray icon, terminating the subclassing
        'cleanly and ultimately preventing the GPF.
        '
        'This code can also be called directly from
        'the form's menu as well, so no special coding
        'is required to differentiate between an end
        'command from a popup systray menu, or from
        'a normal form menu.
        '
        'The UnloadMode of QueryUnload/UnloadMode
        'will equal vbFormControlMenu when this
        'close method is used.
        If Timer1.Enabled = True Then
        Timer1.Enabled = False
        mnuFile(0).Enabled = True
        mnuFile(1).Enabled = False
      Option1(3).Enabled = True
      Text3.Text = "ScreenRifle" & vbNullChar
      Text1.Text = "Shooting has stopped. Your screenshots have been saved. Closing the program." & vbNullChar
      Else
      Option1(1).Enabled = True
      Text3.Text = "ScreenRifle" & vbNullChar
      Text1.Text = "Closing the program" & vbNullChar
      End If
      ShellTrayModifyTip GetSelectedOptionIndex()
    
      Call PostMessage(Me.hwnd, WM_CLOSE, 0&, ByVal 0&)
      Case Else
   End Select
              
End Sub
'Timer for screen shooting
Private Sub Timer1_Timer()
Call GetShot(GetDesktopWindow)
End Sub
