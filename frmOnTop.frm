VERSION 5.00
Begin VB.Form frmOnTop 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   ControlBox      =   0   'False
   Icon            =   "frmOnTop.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   915
      Top             =   1785
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2280
      Top             =   1215
   End
End
Attribute VB_Name = "frmOnTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public LastHWND As Long
Private NewRECT As RECT
Private oldRECT As RECT
Private newHWND As Long
Private MMTF As Boolean
Private Sub Form_Load()
 With Me
    .Left = 0 ' frmMain.Width - (3 * (Me.Width)) + frmMain.Left
    .Top = 0 '
    .Width = 0
    .Height = 0
    'Me.Refresh
    'DoEvents
    Me.Show
    Me.Refresh
    
End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'frmMain.Caption = "MouseDown"
 GWL.BuildButton Me, True
 MMTF = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 0 Then
 'frmMain.Caption = "MouseMove"
  With Me
    If (X < 0) Or (X > (.Width / 15)) Or (Y < 0) Or (Y > (.Height / 15)) Then
        ' outside
        If MMTF = True Then
                GWL.BuildButton Me, False
                MMTF = False
        End If
    Else
        ' inside
        If MMTF = False Then
                GWL.BuildButton Me, True
                MMTF = True
        End If
    End If
  End With
  hWndontop Me.hWND, True
End If
End Sub

'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' If Button = vbLeftButton Then LM.WINDOW_DragWindow Me
'End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
'    frmMain.Test LastHWND
'    Misc.SetFocus LastHWND
'Exit Sub
'Check that the cursor is over the button
With Me
 'frmMain.Caption = "Mouseup"
If Button = vbRightButton Then
'    If (X < 0) Or (X > .Width) Or (Y < 0) Or (Y > .Height) Then
    If (X < 0) Or (X > (.Width / 15)) Or (Y < 0) Or (Y > (.Height / 15)) Then
        ' outside
    Else
        'inside
        PopupMenu frmMain.mnuWindowList
    End If

    GWL.GetPos LastHWND, Me
    Exit Sub
End If


        GWL.GetPos LastHWND, Me
        GWL.hWndontop Me.hWND, True
    
'    If (X < 0) Or (X > .Width) Or (Y < 0) Or (Y > .Height) Then
     If (X < 0) Or (X > (.Width / 15)) Or (Y < 0) Or (Y > (.Height / 15)) Then
        ' Outside
           Misc.SetFocus LastHWND
     Else
        ' Inside
         CreateNewWindow LastHWND
          Me.Visible = False
        'SetParent Me.hWnd, LastHWND
        'Misc.SetFocus LastHWND
 
        LastHWND = 0
        newHWND = 0
    End If
End With
End Sub
Sub CreateNewWindow(bHWND As Long)
 
 'frmMain.Caption = "CreateNewWindow"
 Dim APos As Integer
    APos = Misc.GetNextFreeFormz
    Formz(APos).inUse = True
    Set Formz(APos).vbForm = New frmTray
    Formz(APos).hWND = bHWND
    Formz(APos).SentAwayTime = Format$(Time$, "h:mm.ss AMPM") & " on " & zDate.GetDateStr
    Dim PID As Long
    Formz(APos).ThreadID = Misc.GetWindowThreadProcessId(bHWND, PID)
    'Dim X As New frmTray

    'bHWND = GetParentsHWND(bHWND)
    
    
    With Formz(APos).vbForm
        Relatives.GetRelatives bHWND, Formz(APos).vbForm
        .xHWND = bHWND
        .APos = APos
        .xIconH = zICON.GetIconHandle(bHWND)
        If .xIconH = 0 Then .xIconH = frmMain.DefaultIcon.Picture.Handle
        
        .Caption = GetWindowCaption(bHWND)
        .SendToTray
    End With

 ' Check for Problems
 Dim W As Long
 W = GetWindowLong(LastHWND, GWL_STYLE)
 
 'If (W And WS_VISIBLE) = WS_VISIBLE Then
 '       MsgBox "This application is not supported", vbInformation, "Cannot hide window"
 '       Tray.KillTray_And_RestoreHwnd Formz(APos).vbForm, bHWND, False, True
 '       Unload Formz(APos).vbForm
 '       Formz(APos).inUse = False
 '       Misc.SetFocus bHWND
 'End If

 frmMain.ReBuildWindowList  ' add to Menu
End Sub

Private Sub Timer1_Timer()
    ' OK, Bug Fix time
' The problem:
' -----------------------------------------
' when you click on the button, the ForeGroundWindow is set to
' the hWnd of this form. However if you close the app, sometimes your
' app will gain focus without wanting it. We must detect this then hide
' the form until another window is generated
   
' Find the hWnd of the Current ForeGround
 newHWND = GetForegroundWindow
    
' If newHWND is the hWnd of this form, check lastHWND
If newHWND = Me.hWND Then
    ' Yes, the focus is on this window.
    Dim GW As Long
    GW = GetWindowLong(LastHWND, GWL_STYLE)
    If (GW And WS_VISIBLE) = WS_VISIBLE And (GW And WS_MINIMIZE) <> WS_MINIMIZE Then
            ' The Window is Visible. And NOT minimized
            ' The user is clicking on the button
            'Misc.SetFocus LastHWND
            Exit Sub
    Else
            ' The Window is NOT visible or is minimized. Hide this window
            ' Most likley reason, user closed app.
            LastHWND = 0
            Me.Visible = False
            Exit Sub
    End If
End If
        
        Misc.GetWindowRect newHWND, NewRECT
        
        If (oldRECT.Right <> NewRECT.Right) Or (oldRECT.Top <> NewRECT.Top) Or (newHWND <> LastHWND) Then
        
        LastHWND = newHWND
        GetWindowRect newHWND, oldRECT
        
        If GetParent(newHWND) <> 0 Then
            ' Window has a parent.
            ' Do not allow this window to be sent to system tray
            Me.Visible = False
            Exit Sub
        End If
        
        If Misc.IsValid(newHWND) = False Then
            Me.Visible = False
            Exit Sub
        End If
        
        GWL.BuildButton Me
        
        Dim W As Long
        W = GetWindowLong(newHWND, GWL_STYLE)
        
        If (W And WS_VISIBLE) <> WS_VISIBLE Then
                ' Not Visible
                Me.Visible = False
        Else
                Me.Visible = True
        End If
        
        GWL.GetPos newHWND, Me
        
        If Me.Visible = True Then GWL.hWndontop Me.hWND, True
        
        Misc.SetFocus newHWND
    End If
End Sub

Private Sub Timer2_Timer()
 Dim X As Long
 X = GetWindowLong(Me.hWND, GWL_EXSTYLE)
 If (X And WS_EX_TOPMOST) <> WS_EX_TOPMOST Then
    If Me.Visible = True Then GWL.hWndontop Me.hWND, True
    Misc.SetFocus LastHWND
 End If
End Sub
