VERSION 5.00
Begin VB.Form FrmScreen 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   3855
   ClientTop       =   4725
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "FrmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    frmPassword.Show vbModaless, FrmScreen
    
End Sub
