VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "$TRAY FORM$"
   ClientHeight    =   705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1155
      Top             =   315
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Visible         =   0   'False
      Begin VB.Menu mnuRestoreALL 
         Caption         =   "Restore &ALL Windows"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestoreWindow 
         Caption         =   "&Restore This Window"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xIconH As Long
Public xHWND As Long
Public APos As Integer

Private Sub Form_Load()
    Me.Visible = False
End Sub
Sub RemME()
    Tray.KillTray_And_RestoreHwnd Me, xHWND, False, True
    With Formz(APos)
        .inUse = False
        Set .vbForm = Nothing
        End With
    frmMain.ReBuildWindowList
    Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Me.Visible = False Then ' minimzed to tarks bar
                Static lngMsg As Long
                On Error Resume Next
                lngMsg = X / Screen.TwipsPerPixelX
                
                Select Case lngMsg
                '            capture Double Click and Right Click
                    Case WM_LBUTTONUP
'                            Me.WindowState = Normal
'                            Me.Visible = True
'                            SysTray.FormOnTop Me, True
                             RemME
                             
                    Case WM_RBUTTONUP
                            PopupMenu mnuOptions
                            ' Popup the &File menu (invisible normally)
                    End Select
        End If ' if me.visible = false
End Sub

Public Function SendToTray()
    Me.Caption = Me.Caption ' & "(" & APos & ")"
    Tray.SendToTray Me, xHWND, Me.Caption, True, True, xIconH
    End Function

Private Sub mnuAbout_Click()
    On Error Resume Next
    frmSplash.Show
    'MsgBox "Window Hider v" & App.Major & "." & App.Minor & "." & App.Revision & " (c) Lance Meyrick 2001", vbInformation, "About"
    End Sub

Private Sub mnuRestoreALL_Click()
    frmMain.ReleaseAll
    End Sub

Private Sub mnuRestoreWindow_Click()
    RemME
    End Sub

Private Sub Timer1_Timer()
  ' Check to see if window is still minimized
  Dim WL As Long
  WL = GetWindowLong(Me.xHWND, GWL_STYLE)
  
  If WL = 0 Then
    ' Window Handle gone.
    ' Kill Tray and unload form
    'set values to zero so no window is effected by mistake
    Me.Tag = 0      ' loose WindowStyle
    Me.xHWND = 0    ' set Handle to zero
    RemME
  End If
  
  If (WL And WS_VISIBLE) = True Then
    ' Window is visible... damnit
    RemME
  End If
    
End Sub
