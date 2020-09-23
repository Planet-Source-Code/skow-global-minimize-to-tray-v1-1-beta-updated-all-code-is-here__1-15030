VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Minimizer v1.0 (c) Lance Meyrick 2001"
   ClientHeight    =   2070
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   285
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":0CCA
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image DefaultIcon 
      Height          =   480
      Left            =   1485
      Picture         =   "frmMain.frx":0DA4
      Top             =   2100
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "Restore Window"
         Begin VB.Menu mnuMinWinList 
            Caption         =   "Current Windows Minimized:"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuBlank1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuWindow 
            Caption         =   "(none)"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
 frmOnTop.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim CurOpen As Integer
 ReleaseAll
 Unload frmOnTop
 Unload frmTray
 End
End Sub

Private Sub mnuAbout_Click()
   On Error Resume Next
   frmSplash.Show
End Sub

Private Sub mnuExit_Click()
 Unload Me
End Sub

Sub ReleaseAll()
 
 Dim OC As String
 OC = Me.Caption
 Me.Caption = "Rebuilding All Windows.."
 For I = 1 To 999 'frmTray.Count + 1
  
    If Formz(I).inUse = True Then
            Formz(I).vbForm.RemME
            DoEvents
    End If
 Next
 Me.Caption = OC
End Sub

Sub ReBuildWindowList()
 ' Build Window list.
 
 ' First, remove old list
 On Error Resume Next
 For I = 1 To mnuWindow.Count + 1      ' max of 999.. never happen but ah well
    Unload mnuWindow(I)
    Next
    
 ' Now, Build a new list
 Dim CurrentPos As Integer
 CurrentPos = 0
 
 For I = 1 To 999
    If Formz(I).inUse = True Then
        Load mnuWindow(CurPos)
        With mnuWindow(CurPos)
            .Caption = Formz(I).vbForm.Caption
            .Enabled = True
            .Tag = Trim$(Str$(I))
        End With
        CurPos = CurPos + 1
    End If
Next

If CurPos = 0 Then
    ' No windows.
    Load mnuWindow(CurPos)
        With mnuWindow(CurPos)
            .Caption = "(empty)"
            .Enabled = False
            .Tag = "0"
        End With
End If

End Sub

Private Sub mnuOptions_Click()
 frmOptions.Show
End Sub

Private Sub mnuWindow_Click(Index As Integer)
    ' Window Clicked in List.
    Dim Msg As String
    With Formz(Val(Trim$(mnuWindow(Index).Tag)))
        Msg = Msg & .vbForm.Caption & vbCrLf
        Msg = Msg & "----------------------------------------------" & vbCrLf & vbCrLf
        Msg = Msg & "Put Away at: " & .SentAwayTime & vbCrLf
        'Msg = Msg & "ThreadID=" & .ThreadID & " hWND=" & .hwnd & vbCrLf & vbCrLf
        Dim WS As String
        Dim X As Long
        X = Val(Trim$(.vbForm.Tag))
        If (X And WS_MAXIMIZE) Then
            WS = "Maximized"
        ElseIf (X And WS_MINIMIZE) Then
            WS = "Minimized"
        Else
            WS = "Normal (not maximized)"
        End If
        
        Msg = Msg & "Window State: " & WS & vbCrLf & vbCrLf
        Msg = Msg & "Restore this window?"
    
        If MsgBox(Msg, vbYesNo Or vbQuestion, "Restore Window?") = vbYes Then
            .vbForm.RemME
        End If
    End With

End Sub
