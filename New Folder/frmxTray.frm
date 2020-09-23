VERSION 5.00
Begin VB.Form frmxTray 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "$Tray Caption$"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   675
      Left            =   570
      Top             =   855
      Width           =   585
   End
End
Attribute VB_Name = "frmxTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Windowz(999) As WindowType
Private WindowCount As Long

Public Function CheckForWindow() 'xWindow As WindowType) As Boolean
    'Returns TRUE if window is already here.
    '    Sort method = WindowType.IconHandle
    For I = 1 To WindowCount
        With Windowz(I)
            If .WindowIconHandle = xWindow.WindowIconHandle Then
                CheckForWindow = True
                Exit Function
             
            End If
        End With
    Next
    
    CheckForWindow = False
    ' the Icon is not here.
    
    
End Function
