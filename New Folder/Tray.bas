Attribute VB_Name = "Tray"
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DllGetVersion Lib "shell32" (ByRef pdvi As DLLVERSIONINFO) As Long

' Used for cbSize if old DLL installed
Public Type NOTIFYICONDATA1
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

'C++ DECLARE for Version 5 Updates.
'    DWORD dwState; //Version 5.0
'    DWORD dwStateMask; //Version 5.0
'    TCHAR szInfo[256]; //Version 5.0
'    union {
'        UINT  uTimeout; //Version 5.0
'        UINT  uVersion; //Version 5.0
'    } DUMMYUNIONNAME;
'    TCHAR szInfoTitle[64]; //Version 5.0
'    DWORD dwInfoFlags; //Version 5.0
'        End Type

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Integer
    'uVersion As Integer
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type


Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
End Type

        
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2    '10
Private Const NIF_TIP = &H4     '100

Private Const NIF_INFO = &H8 '&H8    '1000
'Private Const NIF_INFO = &H10 '&H8    '1000
'Private Const NIF_INFO = &H20 '&H8    '1000
'Private Const NIF_INFO = &H40 '&H8    '1000
'Private Const NIF_INFO = &H80 '&H8    '1000
'Private Const NIF_INFO = &H100 '&H8    '1000

Private Const NIF_STATE = &H8  '11111

'Public Const SW_SHOWMINIMIZED = 2
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_MINIMIZE = 6
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNORMAL = 1

'Private Const WS_VISIBLE = &H10000000
Private Const GWL_STYLE = -16
'Private Const WS_MINIMIZE = &H20000000
Const WS_BORDER = &H800000
Const WS_CAPTION = &HC00000
Const WS_CHILD = &H40000000
Const WS_CHILDWINDOW = &H40000000
Const WS_CLIPCHILDREN = &H2000000
Const WS_CLIPSIBLINGS = &H4000000
Const WS_DISABLED = &H8000000
Const WS_DLGFRAME = &H400000
Const WS_GROUP = &H20000
Const WS_HSCROLL = &H100000
Const WS_ICONIC = &H20000000
Const WS_MAXIMIZE = &H1000000
Const WS_MAXIMIZEBOX = &H10000
Const WS_MINIMIZE = &H20000000
Const WS_MINIMIZEBOX = &H20000
Const WS_OVERLAPPED = &H0
Const WS_OVERLAPPEDWINDOW = &HCF0000
Const WS_POPUP = &H80000000
Const WS_POPUPWINDOW = &H80880000
Const WS_SIZEBOX = &H40000
Const WS_SYSMENU = &H80000
Const WS_TABSTOP = &H10000
Const WS_THICKFRAME = &H40000
Const WS_TILED = &H0
Const WS_TILEDWINDOW = &HCF0000
Const WS_VISIBLE = &H10000000
Const WS_VSCROLL = &H200000
Private Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205
Private Const WM_MOUSEMOVE = &H200

Sub SendToTray(vbForm As Form, hwnd As Long, TrayCaption As String, HideHWND As Boolean, HideForm As Boolean, IconH As Long)
    On Error Resume Next
    Dim TC As String
    If Len(TrayCaption) > 63 Then
        TC = Left$(TrayCaption, 40) & ".." & Right$(TrayCaption, 20)
    Else
        TC = TrayCaption
    End If
    
    Dim xTray As NOTIFYICONDATA
        With xTray
            
            If IsV5Compat = False Then
                Dim x As NOTIFYICONDATA1
                .cbSize = Len(x)        ' only send old data
                .hwnd = vbForm.hwnd
                .uId = vbNull
                .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
                .ucallbackMessage = WM_MOUSEMOVE
                .hIcon = IconH     ' handle of iCon
                .szTip = TC & vbNullChar
            Else
                ' Send the whole shibam
                frmMain.Caption = "Using NIF_INFO"
                
                .cbSize = Len(xTray)
                .hwnd = vbForm.hwnd
                .uId = vbNull
                .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_INFO
                .ucallbackMessage = WM_MOUSEMOVE
                .hIcon = IconH     ' handle of iCon
                '.szTip = TrayCaption & vbNullChar
                .dwInfoFlags = NIF_INFO
                .szInfo = "Window Caption: " & vbCrLf & TrayCaption & vbNullChar
                .uTimeout = 10000
                .szInfoTitle = "Lance's System Tray Minimizer" & vbNullChar
            End If
        
        End With
        
        
        If HideForm = True Then
        '    App.TaskVisible = False
            vbForm.Hide
        End If
            
            
        ' Save the Style settings for when resetting it.
        Dim xStyle As Long
        xStyle = GetWindowLong(hwnd, GWL_STYLE)
        vbForm.Tag = Trim$(Str$(xStyle))
        
        If HideHWND = True Then
            ' First, Minimize it if not already
            
            If Not (xStyle And WS_MINIMIZE) Then
                xStyle = (xStyle Xor WS_MINIMIZE)
                    
'                SetWindowLong hWnd, GWL_STYLE, xStyle
 '               DoEvents    ' let it minimize
                ShowWindow hwnd, SW_MINIMIZE
                DoEvents
            End If
            
            
            If (xStyle And WS_VISIBLE) Then
                ' Is visible, take it out
                'xStyle = (xStyle Xor WS_VISIBLE)
                'SetWindowLong hWnd, GWL_STYLE, xStyle
                ShowWindow hwnd, SW_HIDE
            End If
        End If
        
        ' Cool, now add the Tray icon
        Shell_NotifyIcon NIM_ADD, xTray
        


End Sub
        
Sub UpdateTray(vbForm As Form, TrayCaption As String)
    On Error Resume Next
    Dim xTray As NOTIFYICONDATA
        With xTray
            .cbSize = Len(xTray)
            .hwnd = vbForm.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .ucallbackMessage = WM_MOUSEMOVE
            .hIcon = vbForm.ICON    ' handle of iCon
            .szTip = TrayCaption & vbNullChar
        End With
      Shell_NotifyIcon NIM_MODIFY, xTray

End Sub

Sub KillTray_And_RestoreHwnd(vbForm As Form, hwnd As Long, ShowForm As Boolean, SetHWNDFocus As Boolean)
    On Error Resume Next
    Dim xTray As NOTIFYICONDATA
        With xTray
            .cbSize = Len(xTray)
            .hwnd = vbForm.hwnd
            .uId = vbNull
        End With
        
        ' Remove the Tray Icon
        Shell_NotifyIcon NIM_DELETE, xTray
        
        ' Restore the Window and set focus
        
        Dim xStyle As Long
        xStyle = Val(Trim$(vbForm.Tag))
        
        If (xStyle And WS_VISIBLE) Then
            ShowWindow hwnd, SW_SHOW
        End If
        
        If (xStyle And WS_MINIMIZE) Then
            ShowWindow hwnd, SW_SHOWMINIMIZED
        ElseIf (xStyle And WS_MAXIMIZE) Then
            ShowWindow hwnd, SW_SHOWMAXIMIZED
        Else
            ShowWindow hwnd, SW_SHOWNORMAL
        End If
        
        If ShowForm = True Then
            vbForm.Show
        End If
        
        If SetHWNDFocus = True Then
            SetForegroundWindow hwnd
            SetFocus hwnd
        End If
        
        
End Sub

Private Function IsV5Compat() As Boolean
  'Purpose: Get Version info of Shell32.dll (and other main DLLs..all same version#)
  Dim x As DLLVERSIONINFO
  x.cbSize = Len(x)
  DllGetVersion x
  IsV5Compat = False: Exit Function
  If x.dwMajorVersion >= 5 Then IsV5Compat = True Else IsV5Compat = False
  frmMain.Caption = "Version " & x.dwMajorVersion & "." & x.dwMinorVersion
End Function

