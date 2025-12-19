VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmApartados 
   BackColor       =   &H00FF8080&
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Anular apartados"
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   8040
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmApartados.frx":0000
      Height          =   7215
      Left            =   120
      OleObjectBlob   =   "FrmApartados.frx":0014
      TabIndex        =   0
      Top             =   720
      Width           =   10575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lbmodo 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "FrmApartados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsApart As Recordset

Private Sub CmdAnular_Click()
    Dim RsApart As Recordset
    Dim RsVentapagado As Recordset

If MsgBox("¿Seguro que desea anular todos los apartados?", vbYesNo, "Anulando apartados") = vbYes Then
    FechaPres = Format(Date - 15, "mm/dd/yy")
    Set RsApart = bdtienda.OpenRecordset("select * from articulos WHERE apartado = -1")
    RsApart.MoveLast
    RsApart.MoveFirst
    'MsgBox (RsApart.RecordCount)
    With RsApart
        Do Until .EOF
            .Edit
            !apartado = False
            .Update
            .MoveNext
        Loop
    End With
    RsApart.MoveLast
    'MsgBox (RsApart.RecordCount)
    Set RsVentapagado = bdtienda.OpenRecordset("select * from venta where pagado = false")
    With RsVentapagado
        .MoveLast
        .MoveFirst
        MsgBox (.RecordCount)
        Do Until .EOF
            .Edit
            !pagado = True
            .Update
            .MoveNext
        Loop
        
    End With
    
End If

End Sub

Private Sub Form_Load()
    Select Case ModoFrmApartado
        Case "apartados"
            DBGrid1.Left = 10
            DBGrid1.Width = Me.Width - 100
            Set RsApart = bdtienda.OpenRecordset("Select clientes.idcliente, nombre, apellidos," _
            & " telefono, direccion, venta.fecha, venta.acuenta, venta.total, detalleventa.idart" _
            & " from clientes" _
            & " inner join (venta inner join detalleventa on venta.idventa = detalleventa.idventa)" _
            & " on clientes.idcliente = venta.idcliente" _
            & " where venta.pagado = false order by venta.fecha")
            Set Data1.Recordset = RsApart
            DBGrid1.Columns(0).Width = 550
            DBGrid1.Columns(5).Width = 2000
            DBGrid1.Columns(6).Width = 700
            DBGrid1.Columns(7).Width = 700
            lbmodo = "Artículos Apartados"
        Case "prestamos"
            DBGrid1.Left = 10
            DBGrid1.Width = Me.Width - 100
            Dim sssql As String
            sssql = "Select prestamo.idcliente, prestamo.idart," _
            & " prestamo.fecha, prestamo.comentario, clientes.nombre, clientes.apellidos," _
            & " clientes.telefono " _
            & " from prestamo " _
            & " inner join clientes " _
            & " on clientes.idcliente = prestamo.idcliente" 'on articulos.idart = prestamo.idart"
            'MsgBox (sssql)
            
            Set RsApart = bdtienda.OpenRecordset(sssql)
            '& " inner join articulos prestamo.idart = articulos.idart")
            Set Data1.Recordset = RsApart
            DBGrid1.Columns(0).Width = 550
            DBGrid1.Columns(2).Width = 2000
            DBGrid1.Columns(3).Width = 2000
            DBGrid1.Columns(5).Width = 2000
            DBGrid1.Columns(6).Width = 700
'            DBGrid1.Columns(7).Width = 700
        
            lbmodo = "Artículos Prestados"
        End Select
    ModoFrmApartado = ""
   ' MsgBox (Data1.Recordset.RecordCount)
End Sub
