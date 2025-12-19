VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInformeProveedor 
   Caption         =   "Informe por Proveedor"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Ver Todos por fechas"
      Height          =   375
      Index           =   3
      Left            =   7680
      TabIndex        =   21
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ver Todos los articulos"
      Height          =   375
      Index           =   2
      Left            =   5640
      TabIndex        =   20
      Top             =   120
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ver por Fechas"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Ver Todos"
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Value           =   -1  'True
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dtpick2 
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19726337
      CurrentDate     =   38518
   End
   Begin MSComCtl2.DTPicker dtpick1 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19726337
      CurrentDate     =   38518
   End
   Begin VB.ListBox lstprovee 
      Height          =   5130
      ItemData        =   "FrmInformeProveedor.frx":0000
      Left            =   120
      List            =   "FrmInformeProveedor.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lbresult 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   19
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label lbresult 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   18
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lbresult 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   17
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lbresult 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   16
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lbresult 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   15
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lbresult 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Diferencia Ventas hechas - compras"
      Height          =   255
      Index           =   6
      Left            =   2160
      TabIndex        =   13
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Ventas hechas:"
      Height          =   255
      Index           =   5
      Left            =   2160
      TabIndex        =   12
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Suma de Ventas:"
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Porcentaje Vendidos:"
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   10
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Suma de compras:"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Nº de prendas:"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Desde:"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Selecciona Proveedor"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmInformeProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrProvee As String
Dim PrecioCompra As Currency
Dim PrecioVenta As Currency
Dim TotalVentas As Currency
Dim TotalArticulos As Integer
Dim ArticulosVendidos As Integer
Dim PorcentajeVendidos As Integer
Dim TotalVentasHechas As Currency
Dim Fecha1 As Date
Dim Fecha2 As Date


Private Sub dtpick1_Click()
    Fecha1 = dtpick1.Value

End Sub

Private Sub dtpick2_Click()
    Fecha2 = dtpick2.Value

End Sub

Private Sub Form_Load()
    PoneProveedores
End Sub
Private Sub PoneProveedores()
    Dim rsProvee As Recordset
    Set rsProvee = bdtienda.OpenRecordset("proveedor")
    With rsProvee
        lstprovee.AddItem "Todos"
        lstprovee.ItemData(lstprovee.NewIndex) = 0
        Do Until .EOF = True
            lstprovee.AddItem !proveedor
            lstprovee.ItemData(lstprovee.NewIndex) = !id1
            .MoveNext
        Loop
    End With
    rsProvee.Close
End Sub

Private Sub lstprovee_Click()
    Dim rsProvee As Recordset
    If lstprovee.ItemData(lstprovee.ListIndex) <> 0 Then
        Set rsProvee = bdtienda.OpenRecordset("Select * from proveedor where id1 = " & lstprovee.ItemData(lstprovee.ListIndex))
        StrProvee = rsProvee!codigo
    Else
        StrProvee = ""
    End If
    PoneResultados
End Sub
Private Sub PoneResultados()
    PrecioCompra = 0
    PrecioVenta = 0
    TotalVentas = 0
    ArticulosVendidos = 0
    TotalArticulos = 0
    
    lbresult(0) = ""
    lbresult(1) = ""
    lbresult(2) = ""
    lbresult(3) = ""
    lbresult(4) = ""
    lbresult(5) = ""
            On Error Resume Next
    Fecha1 = Format(dtpick1.Value, "mm,dd,yyyy")
    Fecha2 = Format(dtpick2.Value, "mm,dd,yyyy")
    Dim RsArtXProvee As Recordset
    If Option1(0).Value = True Then
        Set RsArtXProvee = bdtienda.OpenRecordset("Select * from articulos where codigo like '" & _
        StrProvee & "*'")
    ElseIf Option1(1).Value = True Then
        Set RsArtXProvee = bdtienda.OpenRecordset("Select * from articulos where codigo like '" & _
        StrProvee & "*' and fechacompra between #" & Fecha1 & "# and  #" & Fecha2 & "#")
    ElseIf Option1(2).Value = True Then
        Set RsArtXProvee = bdtienda.OpenRecordset("Select * from articulos")
    
    ElseIf Option1(3).Value = True Then
        Set RsArtXProvee = bdtienda.OpenRecordset("Select * from articulos where fechacompra between #" & Fecha1 & "# and  #" & Fecha2 & "# order by fechacompra")
'        Set RsArtXProvee = bdtienda.OpenRecordset("Select * from articulos where fechacompra between " & dtpick1.Value & " and  " & dtpick2.Value & "")
    End If
    With RsArtXProvee
        Do Until .EOF = True
            'MsgBox (!fechacompra)
            PrecioCompra = PrecioCompra + !PrecioCompra
            PrecioVenta = PrecioVenta + !PrecioVenta
            If !vendido = True Then
                TotalVentas = TotalVentas + !PrecioVenta
                ArticulosVendidos = ArticulosVendidos + 1
            End If
            .MoveNext
        Loop
        TotalArticulos = .RecordCount
    End With
    lbresult(0) = TotalArticulos
    lbresult(1) = PrecioCompra
    lbresult(2) = PrecioVenta
    lbresult(3) = ArticulosVendidos
    lbresult(4) = PrecioVenta - TotalVentas
    lbresult(5) = Format((ArticulosVendidos * 100) / TotalArticulos, "##0.0")
    
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Label2(0).Visible = False
            Label2(1).Visible = False
            dtpick1.Visible = False
            dtpick2.Visible = False
        Case 1
            Label2(0).Visible = True
            Label2(1).Visible = True
            dtpick1.Visible = True
            dtpick2.Visible = True
        Case 2
            Label2(0).Visible = False
            Label2(1).Visible = False
            dtpick1.Visible = False
            dtpick2.Visible = False
        Case 3
            Label2(0).Visible = True
            Label2(1).Visible = True
            dtpick1.Visible = True
            dtpick2.Visible = True
    End Select
End Sub
