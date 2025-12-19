VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmAddArt 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "articulos"
   ClientHeight    =   7350
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   490
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBuscarfoto 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4800
      TabIndex        =   47
      Top             =   4920
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   7320
      Left            =   5640
      ScaleHeight     =   484
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   636
      TabIndex        =   46
      Top             =   0
      Width           =   9600
   End
   Begin VB.TextBox txtFields 
      DataField       =   "precioventa"
      Height          =   315
      Index           =   11
      Left            =   2040
      TabIndex        =   44
      Top             =   4920
      Width           =   2655
   End
   Begin VB.ComboBox CombPorciento 
      Height          =   315
      Left            =   4080
      TabIndex        =   43
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox comrebaja 
      Height          =   315
      ItemData        =   "frmaddart.frx":0000
      Left            =   4560
      List            =   "frmaddart.frx":0002
      TabIndex        =   7
      Text            =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "precioventa"
      Height          =   315
      Index           =   5
      Left            =   2040
      TabIndex        =   6
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CheckBox chkdepo 
      BackColor       =   &H00FF0000&
      Caption         =   "&Depósito"
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
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ag&regar"
      Height          =   300
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   2370
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdmove 
      BackColor       =   &H00FFC0C0&
      Caption         =   ">>"
      Height          =   375
      Index           =   0
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdmove 
      BackColor       =   &H00FFC0C0&
      Caption         =   ">>>>"
      Height          =   375
      Index           =   1
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdmove 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<<"
      Height          =   375
      Index           =   2
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdmove 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<<<<"
      Height          =   375
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdmove 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<"
      Height          =   375
      Index           =   4
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdmove 
      BackColor       =   &H00FFC0C0&
      Caption         =   ">"
      Height          =   375
      Index           =   5
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6720
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   1320
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   4095
      Begin VB.ComboBox cmbveces 
         Height          =   315
         Left            =   2160
         TabIndex        =   31
         Text            =   "1"
         Top             =   120
         Width           =   615
      End
      Begin VB.CommandButton cmdduplicar 
         Caption         =   "Duplicar"
         Height          =   375
         Left            =   3000
         TabIndex        =   29
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Duplicar este articulo X"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ComboBox cmbtalla 
      Height          =   315
      ItemData        =   "frmaddart.frx":0004
      Left            =   2040
      List            =   "frmaddart.frx":0006
      TabIndex        =   12
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ComboBox cmbcolor 
      Height          =   315
      ItemData        =   "frmaddart.frx":0008
      Left            =   2040
      List            =   "frmaddart.frx":000A
      TabIndex        =   9
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cmbtipo 
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   900
      Width           =   1455
   End
   Begin VB.ComboBox cmbmayor 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   465
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   7560
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   15270
      TabIndex        =   27
      Top             =   7170
      Width           =   15270
   End
   Begin VB.TextBox txtFields 
      DataField       =   "descuento"
      Height          =   315
      Index           =   10
      Left            =   2040
      TabIndex        =   15
      Top             =   4455
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "extra"
      Height          =   315
      Index           =   9
      Left            =   2040
      TabIndex        =   14
      Top             =   4020
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "talla"
      Height          =   315
      Index           =   8
      Left            =   3600
      TabIndex        =   13
      Top             =   3585
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      DataField       =   "color"
      Height          =   315
      Index           =   7
      Left            =   2880
      TabIndex        =   10
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "fechacompra"
      Height          =   315
      Index           =   6
      Left            =   2040
      TabIndex        =   8
      Top             =   2700
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "precioventa"
      Height          =   315
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Top             =   1785
      Width           =   2055
   End
   Begin VB.TextBox txtFields 
      DataField       =   "preciocompra"
      Height          =   315
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Top             =   1335
      Width           =   2055
   End
   Begin VB.TextBox txtFields 
      DataField       =   "tipo"
      Height          =   315
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   900
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      DataField       =   "codigo"
      Height          =   315
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   465
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      DataField       =   "idart"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   60
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Foto (F12):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   11
      Left            =   360
      TabIndex        =   45
      Top             =   4920
      Width           =   1395
   End
   Begin VB.Label lbrepe 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1320
      TabIndex        =   42
      Top             =   5280
      Width           =   4095
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "Precio Rebajado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   41
      Top             =   2280
      Width           =   1755
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "descuento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   10
      Left            =   600
      TabIndex        =   26
      Top             =   4455
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "extra:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   585
      TabIndex        =   25
      Top             =   4020
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "talla:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   585
      TabIndex        =   24
      Top             =   3585
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   585
      TabIndex        =   23
      Top             =   3135
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "fechacompra:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   600
      TabIndex        =   22
      Top             =   2700
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "precioventa:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   600
      TabIndex        =   21
      Top             =   1785
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "preciocompra:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   585
      TabIndex        =   20
      Top             =   1335
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "tipo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   600
      TabIndex        =   19
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   585
      TabIndex        =   18
      Top             =   465
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "idart:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   600
      TabIndex        =   16
      Top             =   60
      Width           =   1395
   End
End
Attribute VB_Name = "FrmAddArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dumprov As String
Dim secs As Integer
Dim RsAddArt As Recordset
Dim Añadiendo As Boolean
Dim tempstr As String
Dim Repetidos As Integer
Dim PorCiento As Integer
Dim RutaImgWeb As String


Private Sub cmbcolor_Click()
    txtFields(7).Text = cmbcolor.Text
End Sub

Private Sub cmbmayor_Click()
    'txtFields(2) = cmbmayor.Text & "/"
    txtFields(1).SetFocus
    txtFields(1).SelStart = Len(txtFields(1).Text)
    txtFields(1).SelLength = 0
    'txtFields(2).SelText
End Sub

Private Sub cmbtalla_Change()
    txtFields(8).Text = cmbcolor.Text
End Sub



Private Sub cmbtipo_LostFocus()
    txtFields(2) = cmbtipo
    txtFields(2).SetFocus

End Sub

Private Sub CmdBuscarfoto_Click()
    CommonDialog1.Filter = "jpg(*.jpg)|*.jpg"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName = "" Then Exit Sub
    'locate is position of the last "\" length is the full path length
    Dim Rutaimagen As String
    Rutaimagen = CommonDialog1.FileName
    MsgBox Rutaimagen
    txtFields(11) = Right(Rutaimagen, (Len(Rutaimagen) - InStr(1, Rutaimagen, "\img")))

End Sub

Private Sub cmdduplicar_Click()
If Len(txtFields(2).Text) < 2 Then
    MsgBox ("Añadir código de mayorista")
    Exit Sub
End If
    
    With RsAddArt
For i = 1 To cmbveces.Text
    On Error Resume Next
    'MsgBox (!idart) '= txtFields(0)
    .AddNew
    'MsgBox (!idart) '= txtFields(0)
    !Codigo = cmbmayor.Text & txtFields(1)
    !Tipo = txtFields(2)
    !PrecioCompra = txtFields(3)
    !PrecioVenta = txtFields(4)
    !preciorebaja = txtFields(5)
    !fechacompra = txtFields(6)
    !Color = txtFields(7)
    !talla = txtFields(8)
    !extra = txtFields(9)
    !Descuento = txtFields(10)
    !foto = txtFields(11)
    !fotoweb = "http://canelamoda.es" & Replace(txtFields(11), "\", "/")
     .Update
     .MoveLast
    Next i
End With

End Sub

Private Sub cmdmove_Click(Index As Integer)
   On Error GoTo sehodio
    ActualizaArts
    BorraTxtfields
    With RsAddArt
        Select Case Index
            Case 0
                For i = 0 To 9
                    .MoveNext
                Next i
            Case 1
                For i = 0 To 99
                    .MoveNext
                Next i
            Case 2
                For i = 0 To 9
                    .MovePrevious
                Next i
            Case 3
                For i = 0 To 99
                    .MovePrevious
                Next i
            Case 4
                .MovePrevious
            Case 5
                .MoveNext
        End Select
        PoneArts
    End With
    Exit Sub
sehodio:
    MsgBox ("No hay más datos para mostrar")
End Sub


Private Sub Command1_Click()
cmdiag.ShowColor
txtFields(7).BackColor = cmdiag.Color
txtFields(7).Text = cmdiag.Color
End Sub


Private Sub CombPorciento_Click()
    PorCiento = CombPorciento.Text
    txtFields(4).Text = (Val(txtFields(3).Text) * (PorCiento + 100)) / 100
    txtFields(4).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 107 Then
        tempstr = txtFields(2).Text 'Right(txtFields(2).Text, Len(txtFields(2).Text) - 4)
        BuscaArticuloIgual
    End If
    Select Case KeyCode
        Case vbKeyF1
            Repetidos = 1
        Case vbKeyF2
            Repetidos = 2
        Case vbKeyF3
            Repetidos = 3
        Case vbKeyF4
            Repetidos = 4
        Case vbKeyF5
            Repetidos = 5
        Case vbKeyF6
            Repetidos = 6
        Case vbKeyF7
            Repetidos = 7
        Case vbKeyF8
            Repetidos = 8
        Case vbKeyF9
            Repetidos = 9
        Case vbKeyF10
            Repetidos = 10
        Case vbKeyF12
            FotoAddArt = True
            Picture1.Picture = Nothing
            If DirFoto <> "" Then Kill App.Path & DirFoto
            FrmWebCam.Show 1
            txtFields(11).Text = DirFoto
            Picture1.Picture = LoadPicture(App.Path & DirFoto)
            DirFoto = ""
            FotoAddArt = False
    End Select
    lbrepe = "Total: " & Repetidos
End Sub

Private Sub Form_Load()
    RutaImgWeb = "http://www.canelamoda.es/ps/img/p/"
    Repetidos = 1
    fechacompra = Format(Date, "Short Date")
    LlenaCmbs
    Set RsAddArt = bdtienda.OpenRecordset("Select * from articulos order by idart")
    RsAddArt.MoveLast
    'PoneArts
    cmbcolor.AddItem "liso"
    cmbcolor.AddItem "estampado"
    cmbcolor.AddItem "blanco"
    cmbcolor.AddItem "negro"
    cmbcolor.AddItem "rojo"
    cmbcolor.AddItem "verde"
    cmbcolor.AddItem "amarillo"
    cmbcolor.AddItem "azul"
    cmbcolor.AddItem "rosa"
    cmbcolor.AddItem "fucsia"
    cmbcolor.AddItem "gris"
    cmbcolor.AddItem "naranja"
    cmbcolor.AddItem "celeste"
    cmbcolor.AddItem "turquesa"
    cmbcolor.AddItem "caldera"
    
    cmbtalla.AddItem "S"
    cmbtalla.AddItem "M"
    cmbtalla.AddItem "L"
    cmbtalla.AddItem "XL"
    cmbtalla.AddItem "36"
    cmbtalla.AddItem "38"
    cmbtalla.AddItem "40"
    cmbtalla.AddItem "42"
    cmbtalla.AddItem "44"
    cmbtalla.AddItem "46"
    cmbtalla.AddItem "48"
    cmbtalla.AddItem "50"
    cmbtalla.AddItem "52"
    cmbtalla.AddItem "54"
    cmbtalla.AddItem "56"
    
    For i = 2 To 10
        cmbveces.AddItem i
    Next i
    For i = 0 To 100 Step 5
        comrebaja.AddItem i
    Next i
    CombPorciento.Text = 100
    PorCiento = 100
    For i = 50 To 250 Step 10
        CombPorciento.AddItem i
    Next i
        CombPorciento.Text = 105
        PorCiento = CombPorciento.Text

    'Set Data1.Recordset = RsAddArt
End Sub
Private Sub PoneArts()
On Error Resume Next
With RsAddArt
    txtFields(0) = !Idart
    cmbmayor.Text = Left(!Codigo, 3)
    txtFields(1) = Right(!Codigo, Len(!Codigo) - 4)
    txtFields(2) = !Tipo
    txtFields(3) = !PrecioCompra
    txtFields(4) = !PrecioVenta
    txtFields(5) = !preciorebaja
    txtFields(6) = !fechacompra
    txtFields(7) = !Color
    If !deposito = True Then chkdepo.Value = 1 Else chkdepo.Value = 0
    txtFields(8) = !talla
    txtFields(9) = !extra
    txtFields(10) = !Descuento
    txtFields(11) = !foto
    Picture1.Picture = Nothing
    If Exists(App.Path & !foto) = True Then
        Picture1.Picture = LoadPicture(App.Path & !foto)
    Else
        Picture1.Picture = Nothing
    End If
End With
End Sub
Private Sub BorraTxtfields()
On Error Resume Next
With RsAddArt
    txtFields(0) = ""
    txtFields(2) = ""
    txtFields(3) = ""
    txtFields(4) = ""
    txtFields(5) = ""
    txtFields(6) = ""
    txtFields(7) = ""
    txtFields(8) = ""
    txtFields(9) = ""
    txtFields(10) = ""
    txtFields(11) = ""
    cmbmayor.Text = ""
    cmbtipo.Text = ""
    cmbcolor.Text = ""
    'chkdepo.Value = 0
    cmbtalla.Text = ""
    txtFields(3).ForeColor = vbBlack

End With
End Sub
Private Sub ActualizaArts()
On Error Resume Next
If Len(txtFields(2).Text) < 2 Then
    MsgBox ("Añadir código de mayorista")
    Exit Sub
End If
For i = 1 To Repetidos
    With RsAddArt
        If Añadiendo = True Then
            .AddNew
        Else
            .Edit
        End If
        '!idart = txtFields(0)
        !Codigo = cmbmayor.Text & "/" & txtFields(1)
        !Tipo = txtFields(2)
        !PrecioCompra = txtFields(3)
        !PrecioVenta = txtFields(4)
        !preciorebaja = txtFields(5)
        !fechacompra = txtFields(6)
        !Color = txtFields(7)
        If chkdepo.Value = 1 Then !deposito = True Else !deposito = False
        !talla = txtFields(8)
        !extra = txtFields(9)
        !Descuento = txtFields(10)
        !foto = txtFields(11)
        !fotoweb = "http://canelamoda.es" & Replace(txtFields(11), "\", "/")
         .Update
         'Añadiendo = False
    End With
Next i
Repetidos = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub



Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    dumprov = cmbmayor.Text
    ActualizaArts
    
    Añadiendo = True
  RsAddArt.MoveLast
    On Error Resume Next
    'dumprov = Left(txtFields(2).Text, 3)
  
  BorraTxtfields
  'txtFields(0) = RsAddArt!Idart + 1
    txtFields(6).Text = Format(Date, "Short Date")
    cmbmayor.Text = dumprov
    'txtFields(2).Text = dumprov
    txtFields(1).SetFocus
    txtFields(1).SelStart = Len(txtFields(1).Text)
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With Data1.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub



Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()
    secs = secs + 1
    If secs >= 1200 Then
        secs = 0
        Unload Me
    End If
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    If Index <> 2 Then
    txtFields(Index).SelStart = 0
    txtFields(Index).SelLength = Len(txtFields(Index))
    End If
End Sub
Private Sub LlenaCmbs()
    Dim rstipos As Recordset
    Set rstipos = bdtienda.OpenRecordset("tipo")
    Do Until rstipos.EOF
        cmbtipo.AddItem rstipos!Tipo
        rstipos.MoveNext
    Loop
    rstipos.Close
    Dim rsmayor As Recordset
    Set rsmayor = bdtienda.OpenRecordset("proveedor")
    Do Until rsmayor.EOF
        cmbmayor.AddItem rsmayor!Codigo
'        cmbmayor.ItemData(cmbmayor.Index) = rsmayor!codigo
        rsmayor.MoveNext
    Loop
    rsmayor.Close
   
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
    
    'If Index = 2 Then BuscaArticuloIgual
    txtFields(Index + 1).SetFocus
End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
    On Error Resume Next
    If Index = 3 Or Index = 4 Or Index = 5 Then
        txtFields(Index).Text = Replace(txtFields(Index).Text, ".", ",")
    End If
    dumprov = Right(txtFields(1).Text, Len(txtFields(1).Text) - 4)
    If Index = 3 Then
'        txtFields(4).Text = (Int(Val(txtFields(3).Text)) * 2) + 1
        Dim Dumprecio As Currency
        Dumprecio = CCur(txtFields(3).Text)
        txtFields(4).Text = Format((Dumprecio * (PorCiento + 100)) / 100, "#00.0#")
        If comrebaja.Text <> "0" And comrebaja.Text <> "" Then
            Dim porc As Currency
            porc = ((CCur(txtFields(3).Text) * comrebaja.Text) / 100)
            If Index = 3 Then txtFields(5).Text = ((CCur(txtFields(3).Text) - porc) * 2) + 1
            txtFields(3).ForeColor = vbRed
            txtFields(3).Text = CCur(txtFields(3).Text) - porc
        End If
    End If
    'tempstr = Right(txtFields(2).Text, Len(txtFields(2).Text) - 4)
'    tempstr = txtFields(2).Text
    'If Index = 2 Then BuscaArticuloIgual
   
    'ActualizaArts
End Sub
Private Sub BuscaArticuloIgual()
    Dim strsql As String
    strsql = "Select* from articulos where codigo like '*" & tempstr & "*'"
    MsgBox (strsql)
'    RsBuscaArticulos.Close
    Set RsBuscaArticulos = bdtienda.OpenRecordset(strsql)
    'If RsBuscaArticulos.EOF = True Then RsBuscaArticulos.MoveFirst
    On Error GoTo sehodio
    RsBuscaArticulos.MoveLast
    If RsBuscaArticulos.RecordCount > 0 Then
        If RsBuscaArticulos.RecordCount = 1 Then
            PoneDatos
            Exit Sub
        ElseIf RsBuscaArticulos.RecordCount > 1 Then
            FrmBuscaArticulos.Show 1
            RsBuscaArticulos.MoveFirst
            Do Until RsBuscaArticulos!Idart = IdArtBuscado
                RsBuscaArticulos.MoveNext
            Loop
            PoneDatos
        End If
    End If
Exit Sub
sehodio:
End Sub

Private Sub PoneDatos()
On Error Resume Next

With RsBuscaArticulos
    txtFields(2) = Right(!Codigo, Len(!Codigo) - 4)
    txtFields(3) = !Tipo
    txtFields(4) = !PrecioCompra
    txtFields(5) = !PrecioVenta
    txtFields(6) = !fechacompra
    txtFields(7) = !Color
    If !deposito = True Then chkdepo.Value = 1 Else chkdepo.Value = 0
    txtFields(8) = !talla
    txtFields(9) = !extra
    txtFields(10) = !Descuento
    txtFields(11) = !foto
    Picture1.Picture = LoadPicture(App.Path & !foto)
End With
End Sub

