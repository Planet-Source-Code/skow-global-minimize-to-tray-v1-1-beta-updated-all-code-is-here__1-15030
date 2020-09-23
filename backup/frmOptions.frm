VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   360
      Left            =   4035
      TabIndex        =   2
      Top             =   3135
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2910
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5265
      Begin VB.CheckBox chkUseBalloon 
         Caption         =   "Use Balloon Captions when sending an app to system tray"
         Height          =   285
         Left            =   210
         TabIndex        =   1
         Top             =   375
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Tray.IsV5Compat = False Then
        Me.chkUseBalloon.Enabled = False
        Me.chkUseBalloon.Value = 0
Else
        Me.chkUseBalloon.Enabled = True
        Me.chkUseBalloon.Value = CStr(GetSetting("MIN2TRAY", "Options", "UseBalloon", "0"))
End If
        
 
End Sub
