VERSION 5.00
Begin VB.Form FrmPrinterIcon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso de Impresión"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   2880
      Top             =   1320
   End
   Begin VB.CommandButton CmdImprimiendo 
      Caption         =   "Imprimiendo"
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   0
      Picture         =   "FrmPrinterIcon.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "FrmPrinterIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdImprimiendo_Click()
    Timer1.Enabled = False
    Unload Me
    
End Sub

Private Sub Form_Load()
    'SkinMe Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
