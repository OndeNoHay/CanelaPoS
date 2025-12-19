VERSION 5.00
Begin VB.Form frmalarma 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3480
      Top             =   2640
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUITA LAS ALARMAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2175
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RECUERDA:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   7695
   End
End
Attribute VB_Name = "frmalarma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim segundos As Integer
Dim Cr, Cg, Cb As Integer

Private Sub Form_Load()
Cr = 255
Cg = 100
Cb = 100
If BlAlarmaQuitar = True Then
    MsgAlarma = "QUITAR LAS ALARMAS"
Else
    MsgAlarma = "PONER LAS ALARMAS"
End If

Label2.Caption = MsgAlarma
End Sub

Private Sub Timer1_Timer()
    If Cr > 100 And Cr <= 255 Then
        Cr = Cr - 2
        Cg = Cg + 2
    Else
        Cg = 100
        Cr = 255
    End If
    Me.BackColor = RGB(Cr, Cg, Cb)
End Sub

Private Sub Timer2_Timer()
    Dim xx
    xx = Me.hWnd
    If segundos = 3 Then
        segundos = 0
        BlAlarmaQuitar = True

        Unload Me
    Else
        segundos = segundos + 1
    End If
End Sub
