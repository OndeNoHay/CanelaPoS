VERSION 5.00
Begin VB.Form frmEncargo 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Encargo de Flamenca"
   ClientHeight    =   8910
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   ScaleHeight     =   15.716
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   5.768
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmEncargo.frx":0000
      Left            =   120
      List            =   "frmEncargo.frx":000D
      TabIndex        =   35
      Text            =   "3 copias"
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   16
      Left            =   1200
      TabIndex        =   33
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   1125
      Index           =   15
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   32
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   14
      Left            =   1200
      TabIndex        =   31
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   13
      Left            =   1200
      TabIndex        =   30
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   12
      Left            =   1200
      TabIndex        =   29
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   1200
      TabIndex        =   28
      Top             =   5040
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   1200
      TabIndex        =   27
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   1200
      TabIndex        =   26
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   25
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   7
      Left            =   1200
      TabIndex        =   24
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   23
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   22
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   21
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   20
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   19
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   18
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir Recibo"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8280
      Width           =   1695
   End
   Begin VB.TextBox Encargo 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   17
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CANELA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   18
      Left            =   0
      TabIndex        =   37
      Top             =   480
      Width           =   3180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Encargo Flamenca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   17
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   3180
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Comentarios"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   16
      Top             =   6480
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "A cuenta"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   7800
      Width           =   1500
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   5.292
      Y1              =   13.547
      Y2              =   13.547
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   -0.423
      X2              =   5.292
      Y1              =   6.773
      Y2              =   6.773
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Foto Nº"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Encuentro"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   13
      Top             =   6120
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ancho Manga"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   12
      Top             =   5760
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Largo Manga"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Largo"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cadera"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cintura"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pecho"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fecha entrega"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Precio"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   -0.423
      X2              =   5.292
      Y1              =   4.022
      Y2              =   4.022
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nº Traje"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dirección"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Teléfono"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apellidos"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1500
   End
End
Attribute VB_Name = "frmEncargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
Dim x As Integer
x = Left(Combo1.Text, 1)
 For i = 1 To x
    Me.PrintForm
    Printer.EndDoc
 Next i
 Unload Me
End Sub

Private Sub Encargo_GotFocus(Index As Integer)
Encargo(Index).SelStart = 0
Encargo(Index).SelLength = Len(Encargo(Index).Text)
End Sub

