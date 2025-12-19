VERSION 5.00
Begin VB.Form frmZoom 
   Caption         =   "Zoom"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actualizar"
      Height          =   615
      Left            =   2760
      TabIndex        =   32
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Venta"
      Height          =   2295
      Left            =   0
      TabIndex        =   17
      Top             =   1800
      Width           =   6855
      Begin VB.Label lbPrecioFinal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbIdVenta 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label LbFechaVenta 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbTelefono 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label LbApellidos 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   27
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label LbNombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   26
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label LbIdcliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3600
         TabIndex        =   24
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Id Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3600
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3600
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Apellidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Id Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Articulo"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CheckBox chkvendido 
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkApartado 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lbTipo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3480
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lbcodigo 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Id Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Vendido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Apartado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Precio compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha compra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lbId 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbPreciocompra 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Lbprecioventa 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbFechacompra 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsZoom As Recordset

Private Sub CmdActualizar_Click()
    With RsZoom
        .Edit
        If chkvendido.Value = 1 Then !vendido = True Else !vendido = False
        If chkApartado.Value = 1 Then !apartado = True Else !apartado = False
        .Update
    End With

End Sub

Private Sub Form_Load()
'On Error Resume Next
    'If IdArtSelec = 0 Then Unload Me
    If Permiso = True Then
        lbPreciocompra.Visible = True
        Label1(3).Visible = True
    End If
    IdZoom = IdArtSelec
    If IdZoom = 0 Or IdZoom = 1 Then
        If InputBox("Escriba el número del artículo", "Número de Artículo", IdZoom) = "" Then
            Exit Sub
        Else
            IdZoom = InputBox("Escriba el número del artículo", "Número de Artículo", IdZoom)
        End If
    End If
    Set RsZoom = bdtienda.OpenRecordset("articulos")
    With RsZoom
        .Index = "idart"
        .Seek "=", IdZoom
        
        lbId = !Idart
        If !vendido = True Then chkvendido.Value = 1 Else chkvendido.Value = 0
        If !apartado = True Then chkApartado.Value = 1 Else chkApartado.Value = 0
        lbTipo = !Tipo
        'lbPreciocompra = !PrecioCompra
        Lbprecioventa = !PrecioVenta
        lbFechacompra = !fechacompra
    End With
    
    Dim RsInfo As Recordset
   ' Dim FechaPres
   ' FechaPres = Format(Date - 15, "mm/dd/yy")
    Set RsInfo = bdtienda.OpenRecordset("Select clientes.idcliente, nombre, apellidos," _
    & " telefono, direccion, venta.idventa, venta.fecha, venta.acuenta, venta.total, detalleventa.idart" _
    & " , detalleventa.preciofinal from clientes" _
    & " inner join (venta inner join detalleventa on venta.idventa = detalleventa.idventa)" _
    & " on clientes.idcliente = venta.idcliente" _
    & " where detalleventa.idart = " & IdZoom)
    If RsInfo.EOF = True Then Exit Sub
    'RsInfo.MoveLast
    'MsgBox (RsInfo.RecordCount)

    'If chkvendido.Value = True Or chkApartado.Value = True Then
        With RsInfo
            LbIdcliente = !IdCliente
            LbNombre = !Nombre
            LbApellidos = !apellidos
            lbTelefono = !telefono
            LbFechaVenta = !Fecha
            lbIdVenta = !IdVenta
            lbPrecioFinal = !PrecioFinal
        End With
    'End If
End Sub

