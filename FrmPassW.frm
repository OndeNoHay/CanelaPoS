VERSION 5.00
Begin VB.Form FrmPassW 
   Caption         =   "XXX"
   ClientHeight    =   285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleWidth      =   1590
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpassword 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmpassw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        AñadeArtic = False
        Unload Me
    End If
End Sub

Private Sub txtpassword_Change()
    If txtpassword.Text = "aula" Or txtpassword.Text = "salas" Or txtpassword.Text = "ropita" Then
        AñadeArtic = True
        Unload Me
    End If
    
End Sub
