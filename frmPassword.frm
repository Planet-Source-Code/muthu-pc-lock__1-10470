VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Screen"
   ClientHeight    =   1650
   ClientLeft      =   4350
   ClientTop       =   5730
   ClientWidth     =   4680
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4680
   Begin VB.TextBox txtUser 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      HideSelection   =   0   'False
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox U 
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "x"
      TabIndex        =   0
      Top             =   600
      Width           =   2400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   1
      Top             =   270
      Width           =   510
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    frmVery.Label2.Caption = frmPassword.Text1.Text
    Me.Visible = False
    
    FrmScreen.Show
    frmLock.Show , FrmScreen

End Sub

Private Sub Form_Load()
    
    DisableCtrlAltDelete (True)
    frmPassword.Show vbModaless, FrmScreen
   
    Command1.Enabled = False
    
    frmPassword.Top = (Screen.Height - frmPassword.Height) / 2
    frmPassword.Left = (Screen.Width - frmPassword.Width) / 2

    Call GetCurrentUser
    txtUser.Text = U
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = 0 Then
 Cancel = True
End If

'Dim Msg, Response   ' Declare variables.
'   Msg = "Are you sre you want to quit this program?"
'   Response = MsgBox(Msg, vbQuestion + vbOKCancel, "Enquiry")
'   Select Case Response
'      Case vbCancel   ' Don't allow close.
'         Cancel = -1
'         Case vbOK
'   End Select
End Sub

Private Sub Text1_Change()
    
    If Text1.Text = "" Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If

End Sub
