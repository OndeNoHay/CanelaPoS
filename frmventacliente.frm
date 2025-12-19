VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmventacliente 
   BackColor       =   &H00FF0000&
   Caption         =   "Ventas por Cliente"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txttotales 
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
      Height          =   420
      Left            =   7320
      TabIndex        =   11
      Top             =   6840
      Width           =   1200
   End
   Begin VB.TextBox txttarjeta 
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
      Height          =   420
      Left            =   5520
      TabIndex        =   10
      Top             =   6840
      Width           =   1200
   End
   Begin VB.TextBox txtefectivo 
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
      Height          =   420
      Left            =   3840
      TabIndex        =   6
      Top             =   6840
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   6840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20971521
      CurrentDate     =   38340
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   7200
      Width           =   11295
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   9720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "rsarticulo"
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Data Data 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   4440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "rsarticulo"
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbregistros 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Encontrados: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4920
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmventacliente.frx":0000
      Height          =   4695
      Left            =   0
      OleObjectBlob   =   "frmventacliente.frx":0013
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmventacliente.frx":09E6
      Height          =   1575
      Left            =   0
      OleObjectBlob   =   "frmventacliente.frx":09FA
      TabIndex        =   3
      Top             =   4800
      Width           =   11175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Suma Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   3
      Left            =   7320
      TabIndex        =   9
      Top             =   6600
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Con tarjeta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   2
      Left            =   5640
      TabIndex        =   8
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "En efectivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   1
      Left            =   3840
      TabIndex        =   7
      Top             =   6600
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Compras desde el :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   6600
      Width           =   2550
   End
End
Attribute VB_Name = "frmventacliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msSortCol As String
Dim mbCtrlKey
Dim secsalir As Integer
Dim StrBusca As String
Dim VerEnVenta As Boolean
Dim RsVentaXCliente As Recordset
Dim RsDetalVentaXCliente As Recordset
Dim RsTotalVentaXCliente As Recordset
Dim RsArtAnulacion As Recordset
    Dim dummyid As Integer
    Dim dummysql As String
Dim FechaVentas


Private Sub DBGrid1_BeforeDelete(Cancel As Integer)
'    dummyid = DBGrid1.Text
'   dummysql = "Select vendido, apartado, devuelto from articulos inner join detalleventa " _
'   & " inner join (venta on venta.idventa = detalleventa.idventa) on detalleventa.idart" _
'   & " = articulos.idart where idventa = " & dummyid
''    dummyid = DBGrid1.Text
''   dummysql = "Select vendido, apartado, devuelto from articulos inner join detalleventa " _
''   & " inner join (venta on venta.idventa = detalleventa.idventa) on detalleventa.idart" _
''   & " = articulos.idart where idventa = " & dummyid
'   MsgBox (dummysql)
'    Set RsDetalVentaXCliente = bdtienda.OpenRecordset(dummysql)
'    With RsDetalVentaXCliente
'        .MoveLast
'        .MoveFirst
'        MsgBox (.RecordCount)
'    End With
    On Error GoTo sehodio
    With RsDetalVentaXCliente
        .MoveFirst
        Do Until .EOF = True
            .Edit
            !vendido = False
            !apartado = False
            .Update
            .MoveNext
        Loop
    End With
    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
        
End Sub

Private Sub DBGrid1_DblClick()
    On Error Resume Next
    dummyid = DBGrid1.Text
    dummysql = "Select detalleventa.idart, codigo, tipo, extra, precioventa," _
    & " preciofinal, vendido, apartado from detalleventa inner join articulos on" _
    & " detalleventa.idart = articulos.idart where idventa = " & DBGrid1.Text
    'MsgBox (dummysql)
    Set RsDetalVentaXCliente = bdtienda.OpenRecordset(dummysql)
    Set Data1.Recordset = RsDetalVentaXCliente
   ' RsDetalVentaXCliente.Close
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo sehodio
    
    If Data.RecordsetType = vbRSTypeTable Then Exit Sub
    
    'comprueba el uso de la tecla ctrl para orden descendente
    If mbCtrlKey Then
        msSortCol = "[" & Data.Recordset(ColIndex).Name & "] desc"
        mbCtrlKey = 0 'actualiza
    Else
        msSortCol = "[" & Data.Recordset(ColIndex).Name & "]"
    End If
    OrdenaColumnas
    msSortCol = vbNullString 'actualiza

    Exit Sub
sehodio:
    MsgBox (Err.Number & Chr(13) & Err.Description)
    MsgBox ("dbgrid1_headclick")
    
End Sub


Private Sub OrdenaColumnas()
    On Error GoTo SortErr

    Dim recRecordset1 As Recordset, recRecordset2 As Recordset
    Dim SortStr As String

    If Data.RecordsetType = vbRSTypeTable Then
        Beep
        MsgBox "Imposible ordenar un Recordset de tipo Table", 48
        Exit Sub
    End If

    Set recRecordset1 = Data.Recordset                        'copia el recordset
    
    If Len(msSortCol) = 0 Then
        SortStr = InputBox("Escriba la columna de orden:")
        If Len(SortStr) = 0 Then Exit Sub
    Else
        SortStr = msSortCol
    End If
    
    Screen.MousePointer = vbHourglass
    recRecordset1.Sort = SortStr
    
    'establece el orden
    Set recRecordset2 = recRecordset1.OpenRecordset(recRecordset1.Type)
    Set Data.Recordset = recRecordset2
    
    Screen.MousePointer = vbDefault
    Exit Sub

SortErr:
    Screen.MousePointer = vbDefault
    MsgBox "Error:" & Err & " " & Err.Description
    MsgBox ("ordenacolumnas")

End Sub




Private Sub DtPicker1_Change()
FechaVentas = Format(DTPicker1.Value, "mm/dd/yy")
    Set RsVentaXCliente = bdtienda.OpenRecordset("select idventa, fecha, tarjeta, total, pagado, acuenta, descuento from venta where idcliente= " & IdCliente & " and fecha > #" & FechaVentas & "#")
   ' MsgBox (RsVentaXCliente.RecordCount)
    Set Data.Recordset = RsVentaXCliente
    DBGrid1.Columns(1).Width = 2000
SumaTotalesVentas

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then
End Sub

Private Sub Form_Load()

    FechaVentas = Format(Now - 180, "mm/dd/yy")
    Set RsCliente = bdtienda.OpenRecordset("select * from clientes where idcliente = " & IdCliente)
    Set RsVentaXCliente = bdtienda.OpenRecordset("select idventa, fecha, tarjeta, total, pagado, acuenta, descuento from venta where idcliente= " & IdCliente & " and fecha > #" & FechaVentas & "#")
    Set Data.Recordset = RsVentaXCliente
    DBGrid1.Columns(1).Width = 2000
    DTPicker1.Value = Date - 180
    SumaTotalesVentas
End Sub
Private Sub SumaTotalesVentas()
    Set RsTotalVentaXCliente = bdtienda.OpenRecordset("Select sum(total)as totalefectivo, sum(tarjeta) as totaltarjeta from venta where idcliente = " & IdCliente & " and fecha > #" & FechaVentas & "#")
'    If RsTotalVentaXCliente.RecordCount <= 1 Then Exit Sub
    'If RsTotalVentaXCliente.EOF = True Then Exit Sub
    On Error Resume Next
    txtefectivo.Text = RsTotalVentaXCliente!totalefectivo
    txttarjeta.Text = RsTotalVentaXCliente!totaltarjeta
    txttotales.Text = CCur(txtefectivo) + CCur(txttarjeta)
    'MsgBox (RsTotalVentaXCliente.RecordCount)
End Sub
Private Sub Form_Resize()
On Error Resume Next
'DBGrid1.Width = Me.Width - 300
'DBGrid1.Height = Me.Height - 2000
Frame1.Top = Me.Height - 1300
Frame1.Width = Me.Width - 300
End Sub

