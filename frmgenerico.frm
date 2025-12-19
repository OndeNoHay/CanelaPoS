VERSION 5.00
Begin VB.Form frmgenerico 
   BackColor       =   &H00FF0000&
   Caption         =   "Añadir Artículo Genérico"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5160
      Width           =   1815
   End
   Begin VB.ListBox LstTipos 
      Height          =   1815
      ItemData        =   "frmgenerico.frx":0000
      Left            =   1920
      List            =   "frmgenerico.frx":0002
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Index           =   1
      Left            =   2520
      TabIndex        =   37
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Cinturón"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Bolso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Plata"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Bisutería"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Estola"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF8080&
      Caption         =   "&Complementos"
      Enabled         =   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Index           =   3
      Left            =   7080
      TabIndex        =   35
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
      Begin VB.TextBox TxtGoyseCode 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Vestido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Falda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Camiseta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Conj. Falda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Conj. Pantalón"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "Código Goyse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2520
         Width           =   1575
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF8080&
      Caption         =   "Ropa &Goyse"
      Enabled         =   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Index           =   2
      Left            =   4800
      TabIndex        =   34
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton Option1 
         BackColor       =   &H000000FF&
         Caption         =   "Flamenca"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   40
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Conj. Pantalón"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Conj. Falda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Camiseta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Falda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Vestido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Index           =   0
      Left            =   360
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Cadena"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Anillo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Pulsera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Colgante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Pendiente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Otros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF8080&
      Caption         =   "&Ropa Genérica"
      Enabled         =   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF8080&
      Caption         =   "&Plata"
      Enabled         =   0   'False
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtprecio 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Tipo de Artículo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Precio del Artículo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmgenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    If BlGoyse = True Then GoyseCode = TxtGoyseCode
    
    Dumresp = txtprecio
    If Len(Dumresp) = 0 Then Exit Sub
    If InStr(1, Dumresp, ".", 1) > 0 Then
        Dumprecio = Val(Dumresp)
    ElseIf InStr(1, Dumresp, ",") > 0 Then
        Dumprecio = CCur(Dumresp)
    Else
        Dumprecio = Val(Dumresp)
    End If
'Option1_Click
    If Dumprecio = 0 Then
        If MsgBox("El precio indicado es 0€. ¿Es correcto?", vbYesNo) = vbYes Then
            NumArtVend = NumArtVend + 1
            Venta.PoneArticuloGenerico Dumprecio
        Else
            MsgBox ("Artículo genérico no añadido")
        End If
    Else
            NumArtVend = NumArtVend + 1
            Venta.PoneArticuloGenerico Dumprecio
    
    End If
Unload Me
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dumtipo = "Bisutería"
    BlGoyse = False
    CargaTipos 0
End Sub

Private Sub LstTipos_Click()
    Dumtipo = Dumtipo & "-" & LstTipos.Text
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index >= 0 And Index <= 5 Then
        LstTipos.Visible = True
        LstTipos.Top = Option1(Index).Top + Frame1(0).Top
        LstTipos.Left = Option1(Index).Left + Frame1(0).Left + 1500
        CargaTipos Index
        LstTipos.Visible = True
    Else
        LstTipos.Visible = False
    End If
    If Index = 24 Then
        frmEncargo.Show 1
    End If
    Dumtipo = Option1(Index).Caption
    txtprecio.SetFocus
End Sub
Private Sub CargaTipos(NumDum As Integer)
    LstTipos.Clear
    Select Case NumDum
        Case 0
            LstTipos.AddItem "liso"
            LstTipos.AddItem "coral"
            LstTipos.AddItem "negro"
            LstTipos.AddItem "nacar"
            LstTipos.AddItem "avalon"
            LstTipos.AddItem "otros"
            
        Case 1
            LstTipos.AddItem "electroform"
            LstTipos.AddItem "aros"
            LstTipos.AddItem "piedras"
            LstTipos.AddItem "otras"
        Case 2
            LstTipos.AddItem "liso"
            LstTipos.AddItem "coral"
            LstTipos.AddItem "negro"
            LstTipos.AddItem "nacar"
            LstTipos.AddItem "avalon"
            LstTipos.AddItem "otros"
        Case 3
            LstTipos.AddItem "torno"
            LstTipos.AddItem "colgar"
        Case 4
            LstTipos.AddItem "rígida"
            LstTipos.AddItem "flexible"
            LstTipos.AddItem "caucho"
        Case 5
            LstTipos.AddItem "alfiler"
            LstTipos.AddItem "gemelos"

    End Select
End Sub
Private Sub Option2_Click(Index As Integer)
Select Case Index
    Case 0
        Frame1(0).Visible = True
        Frame1(1).Visible = False
        Frame1(2).Visible = False
        Frame1(3).Visible = False
        'Option1(0).SetFocus
        'Option1_Click 0
        LstTipos.Top = Option1(Index).Top + Frame1(0).Top
        LstTipos.Left = Option1(Index).Left + Frame1(0).Left + 1500
        
        'CargaTipos 0
        DumCode = "Plata"
    Case 1
        Frame1(0).Visible = False
        Frame1(1).Visible = True
        Frame1(2).Visible = False
        Frame1(3).Visible = False
        LstTipos.Visible = False
        
        Option1(6).Value = True
        Option1_Click 6
        DumCode = "Complementos"
    Case 2
        Frame1(0).Visible = False
        Frame1(1).Visible = False
        Frame1(2).Visible = True
        Frame1(3).Visible = False
        LstTipos.Visible = False
        
        Option1(12).Value = True
        Option1_Click 12
        DumCode = "Genérica"

    Case 3
        Frame1(0).Visible = False
        Frame1(1).Visible = False
        Frame1(2).Visible = False
        Frame1(3).Visible = True
        LstTipos.Visible = False
        BlGoyse = True
        GoyseCode = TxtGoyseCode
        DumCode = "Goyse"
        Option1(18).Value = True
    
End Select
End Sub

Private Sub txtprecio_Change()
If Val(txtprecio) > 0 Then
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    Option2(2).Enabled = True
    Option2(3).Enabled = True
    cmdAceptar.Enabled = True
End If
End Sub
