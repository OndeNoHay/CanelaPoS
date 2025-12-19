VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportar 
   Caption         =   "Importar prendas de otra base de datos"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   369
   ScaleMode       =   0  'User
   ScaleWidth      =   705.336
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdImportar 
      Caption         =   "Importar"
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox TxtUltimo 
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox TxtPrimero 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton CmdAbrir 
      Caption         =   "Abrir Datos"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Base de datos MS Access(*.mdb)|*.mdb"
   End
   Begin MSDBGrid.DBGrid DBGridImp 
      Bindings        =   "frmImportar.frx":0000
      Height          =   4215
      Left            =   120
      OleObjectBlob   =   "frmImportar.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
   Begin VB.Data Data1 
      Caption         =   "Dataimp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Width           =   10095
   End
   Begin VB.Line Line1 
      X1              =   14.849
      X2              =   683.063
      Y1              =   352
      Y2              =   352
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta el artículo: "
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Importar desde el artículo: "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
End
Attribute VB_Name = "frmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RutaBDImp As String
Dim BDimportar As Database
Dim RSimportar As Recordset
Dim RsArticulos2 As Recordset
Dim RSSeleccionados As Recordset

Private Sub CmdAbrir_Click()
    CommonDialog1.DialogTitle = "Seleccione la Base de Datos..."

    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.ShowOpen
    RutaBDImp = CommonDialog1.FileName
    
    If RutaBDImp = "" Then Exit Sub
    Set BDimportar = OpenDatabase(RutaBDImp)
    Set RSimportar = BDimportar.OpenRecordset("articulos")
    Set Data1.Recordset = RSimportar
    MsgBox RutaBDImp
End Sub

Private Sub CmdImportar_Click()
    On Error GoTo sehodio
    Set RsArticulos2 = bdtienda.OpenRecordset("articulos")
    Set RSSeleccionados = BDimportar.OpenRecordset("Select * from articulos where idart between " & Txtprimero.Text & " and " & Txtultimo.Text)
    With RSSeleccionados
    .MoveLast
    If MsgBox("Desea importar " & .RecordCount & " prendas?", vbYesNo) = vbNo Then Exit Sub
    .MoveFirst
    Do Until .EOF
        RsArticulos2.AddNew
            RsArticulos2!idarticulo = !idarticulo
            RsArticulos2!Codigo = !Codigo
            RsArticulos2!Tipo = !Tipo
            RsArticulos2!PrecioCompra = !PrecioCompra
            RsArticulos2!PrecioVenta = !PrecioVenta
            RsArticulos2!preciorebaja = !preciorebaja
            RsArticulos2!Descuento = !Descuento
            RsArticulos2!Color = !Color
            RsArticulos2!talla = !talla
            RsArticulos2!extra = !extra
            RsArticulos2!vendido = !vendido
            RsArticulos2!apartado = !apartado
            RsArticulos2!inventario = !inventario
            RsArticulos2!ACuenta = !ACuenta
            RsArticulos2!IdCliente = !IdCliente
            RsArticulos2!devuelto = !devuelto
            RsArticulos2!fechacompra = !fechacompra
            RsArticulos2!fechaventa = !fechaventa
            RsArticulos2!id1 = !id1
            RsArticulos2!deposito = !deposito
        RsArticulos2.Update
        .MoveNext
    Loop
    End With
    MsgBox ("Hecho")
    Exit Sub
sehodio:
    MsgBox ("No se ha podido realizar la operación." & vbCrLf & "Error nº: " & Err.Number & vbCrLf & "Descripción: " & Err.Description)
    
End Sub

