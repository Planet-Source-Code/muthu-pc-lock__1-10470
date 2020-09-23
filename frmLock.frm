VERSION 5.00
Begin VB.Form frmLock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Locked"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   120
      Top             =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmLock.frx":030A
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Dim AltDown, CtrlDown, ShiftDown
AltDown = (Shift And vbAltMask) > 0
CtrlDown = (Shift And vbCtrlMask) > 0
ShiftDown = (Shift And vbShiftMask) > 0

    If CtrlDown And AltDown And ShiftDown Then

    Unload frmLock
    frmVery.Show
    
    End If

End Sub

Private Sub Form_Load()
    
    Label1.Caption = "This PC is in use and has been locked" _
    & Chr(13) & "The PC can only be unlocked by" _
    & " " & frmPassword.U & Chr(13) _
    & Chr(13) & Chr(13) _
    & "Press Control + Shift + Alt to Unlock "

    frmLock.Top = (Screen.Height - frmLock.Height) / 2
    frmLock.Left = (Screen.Width - frmLock.Width) / 2

    DisableCtrlAltDelete (True)
End Sub

Private Sub Timer1_Timer()

Dim a, b

a = Rnd * FrmScreen.Height / 2
b = Rnd * FrmScreen.Width / 2

frmLock.Move b, a, frmLock.Width, frmLock.Height

End Sub
