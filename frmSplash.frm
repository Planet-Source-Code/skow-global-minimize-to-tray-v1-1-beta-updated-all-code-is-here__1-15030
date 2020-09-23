VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3300
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   5130
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   45
         Top             =   3450
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1935
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   2820
         Width           =   1305
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   2445
         Left            =   45
         ScaleHeight     =   2415
         ScaleWidth      =   5010
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   255
         Width           =   5040
         Begin VB.Label Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Scrollling Text :~)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   930
            Left            =   45
            TabIndex        =   3
            Top             =   750
            UseMnemonic     =   0   'False
            Width           =   4905
            WordWrap        =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  FillText
  Me.Text1.AutoSize = True
  Me.Text1.Top = Me.Picture1.Height
End Sub

Sub A(T As String)
    Me.Text1.Caption = Me.Text1.Caption & T & vbCrLf
End Sub


Sub FillText()
    Me.Text1.Caption = ""
    A "LMWARE PRESENTS"
    A ""
    A ""
    A ""
    A "Global Minimize To Tray. Version " & App.Major & "." & App.Minor & "." & App.Revision
    A ""
    A "Coded By and © Lance Meyrick 2001"
    A ""
    A "Big Thanks to all who gave me help finding bugs."
    A "and to Irredescent for giving me the idea. :)"
    A ""
    A ""
    A "Program Details"
    A ""
    A "This Program was designed to allow you to minimize any window to the system tray and rebuild the window with ease."
    A "I have purposly designed it to NOT work with ICQ, Child Windows and Dialogs (Ie Start>Run dialog)"
    A "I plan to also deny any application already on the system tray but that is yet to come."
    A ""
    A "Questions and Answers"
    A ""
    A "Q. What if I minimize My Computer, then Run My Computer again?"
    A ""
    A "A. By Default, Windows will take the minimized My Computer and re-use it. This has been thought through. After 1-2 seconds, the tray icon will disapear and the . button will apear on My Computer again. As too if you close a window that is minimized."
    A ""
    A ""
    A "Q. Can an application be sent to the tray more then once?"
    A ""
    A "A. Not really. It can be done, but this program will remove the tray icon when it detects the window is not hidden anymore. That and the . button will not be visible if a tray icon is present for that window."
    A ""
    A ""
    A "Q. What is the value for NIF_INFO?"
    A ""
    A "A. Well I'm glad you asked. According to Visual C++ 6.0 SP4, there is no such value. But MSDN uses it in NOTIFYICONDATA structure. So does the help files have a constant that the actual header files don't?.. God knows.. Anyway, Good ol' API guide (www.vbapi.com) had it. And might I say, the old place on the net that had the value. (it is &H10)"
    A ""
    A ""
    A "Thanks for reading!!"
    A ""
    A ""
    A "Coded By Lance Meyrick 2001"
    A ""
    A ""
    A ""
    A "" ' restart scrolling again :~)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If Button <> 0 Then Timer1_Timer
 If Button <> 0 Then Me.Text1.Top = Me.Text1.Top - 50
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button <> 0 Then Me.Text1.Top = Me.Text1.Top - 50
End Sub

Public Sub Timer1_Timer()
 With Me.Text1
    .Top = .Top - (Me.Picture1.Height / 150)
    If .Top + .Height <= 0 Then .Top = Me.Picture1.Height
 End With
End Sub
