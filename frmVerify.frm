VERSION 5.00
Begin VB.Form frmVery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Verification"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmVerify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&End"
      Height          =   325
      Left            =   3360
      TabIndex        =   3
      Top             =   630
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   325
      Left            =   3360
      TabIndex        =   2
      Top             =   210
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "x"
      TabIndex        =   1
      Top             =   600
      Width           =   2275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter your secret password "
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1965
   End
End
Attribute VB_Name = "frmVery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    If Text1.Text = frmPassword.Text1.Text Then
        'MsgBox "Password verified"
        DisableCtrlAltDelete (False)
        End
    Else
        MsgBox "Incorrect password, please re-enter", vbInformation, "Password Verification"
        Text1.SetFocus
        Text1.Text = ""
        
End If
End Sub

Private Sub Command2_Click()
Unload Me
frmLock.Show
End Sub

Private Sub Form_Load()
    
    Command1.Enabled = False
    frmVery.Top = (Screen.Height - frmVery.Height) / 2
    frmVery.Left = (Screen.Width - frmVery.Width) / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmLock.Show , FrmScreen

End Sub

Private Sub Text1_Change()
    If Text1.Text = "" Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If

End Sub
