VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCajaDiaria 
   BackColor       =   &H00C00000&
   Caption         =   "Caja Diaria"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "FrmCajaDiaria.frx":0000
      Height          =   1695
      Left            =   0
      OleObjectBlob   =   "FrmCajaDiaria.frx":0014
      TabIndex        =   2
      Top             =   5280
      Width           =   11895
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   11595
      TabIndex        =   1
      Top             =   6960
      Width           =   11655
      Begin VB.CommandButton CmdListado 
         Caption         =   "Listado"
         Height          =   255
         Left            =   6960
         TabIndex        =   21
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdcambiafecha 
         Caption         =   "Cambiar fecha de venta"
         Height          =   495
         Left            =   9480
         TabIndex        =   13
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdborraventa 
         Caption         =   "Borra Venta"
         Height          =   495
         Left            =   8400
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdchart 
         Caption         =   "Chart"
         Height          =   495
         Left            =   10800
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "Cambiar Tarjeta-Efectivo"
         Height          =   495
         Left            =   6960
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106496001
         CurrentDate     =   38197
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   8400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ropa"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   25
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "plata"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   24
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lbplata 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Lbropa 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lbtotalcardtarde 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   600
         Width           =   105
      End
      Begin VB.Label lbtotalcardmorning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   19
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarde"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   6000
         TabIndex        =   18
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mañana"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   17
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Totales"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   16
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lbcashtarde 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6000
         TabIndex        =   15
         Top             =   240
         Width           =   105
      End
      Begin VB.Label lbcashmañana 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   14
         Top             =   240
         Width           =   105
      End
      Begin VB.Label lbtarjeta 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbefectivo 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarjeta"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbtotal 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   2400
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Día"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmCajaDiaria.frx":09E7
      Height          =   5295
      Left            =   0
      OleObjectBlob   =   "FrmCajaDiaria.frx":09FB
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "FrmCajaDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalMorning, TotalTarde As Currency
Dim TotalCashMorning, TotalCashTarde, TotalCardMorning, TotalCardTarde, TotalSilver, TotalClothing As Currency


Private Sub cmdborraventa_Click()
    Dim idenVenta As Long
    Dim RsTemp As Recordset
    On Error GoTo sehodio
    DBGrid1.Col = 0
    If DBGrid1.Text = "" Then GoTo sehodio
    If DBGrid1.Col <> 0 Then GoTo sehodio
    idenVenta = DBGrid1.Text
    If MsgBox("¿Desea borrar la venta " & idenVenta & "?", vbYesNo) = vbNo Then Exit Sub
    Set RsTemp = bdtienda.OpenRecordset("Select * from venta where idventa = " & idenVenta)
        If RsTemp.EOF = False Then
            RsTemp.Delete
        End If
    Set RsTemp = bdtienda.OpenRecordset("Select * from arqueo where idventa = " & idenVenta)
        If RsTemp.EOF = False Then
            RsTemp.Delete
        End If
    Set RsTemp = bdtienda.OpenRecordset("Select * from detalleventa where idventa = " & idenVenta)
        If RsTemp.EOF = False Then
            Do Until RsTemp.EOF = True
                RsTemp.Delete
            Loop
        End If
    MuestraResultado
    Exit Sub
sehodio:
    MsgBox ("No se ha podido completar la operación. Señale de nuevo la venta")
End Sub

Private Sub cmdcambiafecha_Click()
    Dim idenVenta As Long
    Dim RsTemp As Recordset
    Dim OldDate, NewDate As Date
    On Error GoTo sehodio
    DBGrid1.Col = 0
    If DBGrid1.Text = "" Then GoTo sehodio
    If DBGrid1.Col <> 0 Then GoTo sehodio
    idenVenta = DBGrid1.Text
    If MsgBox("¿Desea cambiar la fecha de la venta " & idenVenta & "?", vbYesNo) = vbNo Then Exit Sub
    
    Set RsTemp = bdtienda.OpenRecordset("Select * from venta where idventa = " & idenVenta)
        OldDate = RsTemp!fecha
        NewDate = InputBox("Escriba la fecha que desea para esta venta (dia-mes-año)", , OldDate)
        If IsDate(NewDate) = False Or Year(NewDate) <> Year(OldDate) Then
            MsgBox ("Por favor, inténtelo de nuevo con una fecha válida")
            Exit Sub
        End If
        RsTemp.Edit
        RsTemp!fecha = NewDate
        RsTemp.Update
    Set RsTemp = bdtienda.OpenRecordset("Select * from arqueo where idventa = " & idenVenta)
        OldDate = RsTemp!fecha
        RsTemp.Edit
        RsTemp!fecha = NewDate
        RsTemp.Update
       
    MuestraResultado
    Exit Sub
sehodio:
    MsgBox ("No se ha podido completar la operación. Señale de nuevo la venta")
End Sub

Private Sub cmdCambiar_Click()
    Dim idenVenta As Long
    Dim RsTemp As Recordset
    Dim tempcantidad As Currency
    On Error GoTo sehodio
    DBGrid1.Col = 0
    If DBGrid1.Text = "" Then GoTo sehodio
    If DBGrid1.Col <> 0 Then GoTo sehodio
    idenVenta = DBGrid1.Text
    Set RsTemp = bdtienda.OpenRecordset("Select * from venta where idventa = " & idenVenta)
        tempcantidad = RsTemp!Tarjeta
        RsTemp.Edit
        RsTemp!Tarjeta = RsTemp!total
        RsTemp!total = tempcantidad
        RsTemp.Update
    Set RsTemp = bdtienda.OpenRecordset("Select * from arqueo where idventa = " & idenVenta)
        tempcantidad = RsTemp!Tarjeta
        RsTemp.Edit
        RsTemp!Tarjeta = RsTemp!caja
        RsTemp!caja = tempcantidad
        RsTemp.Update
       
    MuestraResultado
    Exit Sub
sehodio:
    MsgBox ("No se ha podido completar la operación. Señale de nuevo la venta")
End Sub


Private Sub cmdchart_Click()
frmchart.Show
End Sub

Private Sub CmdListado_Click()
ImprimeCajaDiaria
End Sub

Private Sub DtPicker1_Change()
txtFecha = DTPicker1.Value
'Calendar1.Visible = False
MuestraResultado

End Sub

Private Sub dtpicker1_Click()
txtFecha = DTPicker1.Value
'Calendar1.Visible = False
MuestraResultado
End Sub

Private Sub cmdeligedia_Click()
'Calendar1.Visible = True
End Sub
Private Sub ImprimeCajaDiaria()
    Dim fechainicio, FechaFinal
    SelectPrinter "Axiohm A793 CLASS 7193 Full"

    fechainicio = InputBox("Escriba la fecha de inicio", "Fecha de Inicio", Date - 7) 'Format(DTPicker1.Value, "Short Date")
    FechaFinal = InputBox("Escriba la fecha final", "Fecha Final", Date)
    Dim dummysql As String
    Dim NumDias As Integer
    NumDias = CDate(FechaFinal) - CDate(fechainicio)
    Dim ArqueoXdia() As Dinerito
    ReDim ArqueoXdia(NumDias)
    For i = 0 To NumDias
    
        dummysql = "Select arqueo.idventa, caja, arqueo.tarjeta, arqueo.fecha, movimiento, concepto, clientes.nombre, clientes.apellidos, venta.acuenta, " _
        & " clientes.telefono from arqueo, clientes, venta where venta.idventa = arqueo.idventa " _
        & " and clientes.idcliente = venta.idcliente and arqueo.fecha like '*" & CDate(fechainicio) + i & "*'" _
        & " ORDER BY arqueo.fecha"
        'MsgBox (dummysql)
         Set RsArqueo = bdtienda.OpenRecordset(dummysql)
        Set Data1.Recordset = RsArqueo
        DTPicker1.Value = CDate(fechainicio) + i
        On Error GoTo sehodio
        Dim dumtarjeta As Currency
        Dim DumTarjeta2 As Currency
        With RsArqueo
            '.MoveLast
            '.MoveFirst
            If .RecordCount = 0 Then
                'MsgBox ("No hay datos de " & FechaHoy)
            Else
                Do Until .EOF
                        If IsNull(!Tarjeta) Then
                            DumTarjeta2 = 0
                        Else
                            DumTarjeta2 = !Tarjeta
                        End If
                        If !movimiento = True Then
                            ArqueoXdia(i).Arqueo = Arqueo + !caja + DumTarjeta2
                            dumtarjeta = dumtarjeta + DumTarjeta2
                        Else
                            ArqueoXdia(i).Arqueo = ArqueoXdia(i).Arqueo - !caja
                        End If
                        If Hour(!fecha) < 15 Then
                            ArqueoXdia(i).TotalCashMorning = ArqueoXdia(i).TotalCashMorning + !caja
                            ArqueoXdia(i).TotalCardMorning = ArqueoXdia(i).TotalCardMorning + !Tarjeta
                            ArqueoXdia(i).TotalMorning = ArqueoXdia(i).TotalMorning + !caja + !Tarjeta
                        Else
                            ArqueoXdia(i).TotalCashTarde = ArqueoXdia(i).TotalCashTarde + !caja
                            ArqueoXdia(i).TotalCardTarde = ArqueoXdia(i).TotalCardTarde + !Tarjeta
                            ArqueoXdia(i).TotalTarde = ArqueoXdia(i).TotalTarde + !caja + !Tarjeta
                        End If
                    
                    .MoveNext
                Loop
            End If
        End With
    Next i
    
    Dim Texto
    Dim TotalEfectivo As Currency
    Dim TotalTarjeta As Currency
    
    For i = 0 To UBound(ArqueoXdia)
        TotalEfectivo = TotalEfectivo + ArqueoXdia(i).TotalCashMorning + ArqueoXdia(i).TotalCashTarde
        TotalTarjeta = TotalTarjeta + ArqueoXdia(i).TotalCardMorning + ArqueoXdia(i).TotalCardTarde
        Texto = Texto & vbCrLf
        Texto = Texto & "Dia: " & CDate(fechainicio) + i & vbCrLf
        Texto = Texto & "Total Mañana: " & ArqueoXdia(i).TotalMorning & vbCrLf
        Texto = Texto & Chr(9) & "Efectivo: " & ArqueoXdia(i).TotalCashMorning & vbCrLf
        Texto = Texto & Chr(9) & "Tarjeta: " & ArqueoXdia(i).TotalCardMorning & vbCrLf
        Texto = Texto & "Total Tarde: " & ArqueoXdia(i).TotalTarde & vbCrLf
        Texto = Texto & Chr(9) & "Efectivo: " & ArqueoXdia(i).TotalCashTarde & vbCrLf
        Texto = Texto & Chr(9) & "Tarjeta: " & ArqueoXdia(i).TotalCardTarde & vbCrLf
        Texto = Texto & "Total día Efectivo: " & ArqueoXdia(i).TotalCashMorning + ArqueoXdia(i).TotalCashTarde & vbCrLf
        Texto = Texto & "Total día Tarjeta: " & ArqueoXdia(i).TotalCardMorning + ArqueoXdia(i).TotalCardTarde & vbCrLf
        
        Texto = Texto & "-------------------------" & vbCrLf
    Next i
        Texto = Texto & "*************************" & vbCrLf
        Texto = Texto & "RECUENTO DE TOTALES" & vbCrLf
        Texto = Texto & "Total efectivo: " & TotalEfectivo & vbCrLf
        Texto = Texto & "Total tarjeta: " & TotalTarjeta & vbCrLf
        Texto = Texto & "Suma de totales: " & TotalEfectivo + TotalTarjeta & vbCrLf
        If UBound(ArqueoXdia) < 4 Then
            MsgBox (Texto)
        Else
            MsgBox ("El listado es demasiado largo para ponerlo en pantalla")
        End If
        If MsgBox("¿Desea imprimir el listado?", vbYesNo) = vbYes Then
                ImprimeTexto (Texto)
                Printer.EndDoc
        End If
    
    Exit Sub
sehodio:
    MsgBox Err.Number & Err.Description
End Sub
Private Sub MuestraResultado()
    FechaHoy = DTPicker1.Value 'Format(DTPicker1.Value, "Short Date")
    Dim dummysql As String
    
  '  dummysql = "Select arqueo.idventa, caja, arqueo.tarjeta, " _
    & " arqueo.fecha, movimiento, concepto, artconcepto.codigo, clientes.nombre, clientes.apellidos, venta.acuenta, " _
    & " clientes.telefono from arqueo, clientes, venta, artconcepto where venta.idventa = arqueo.idventa " _
    & " and clientes.idcliente = venta.idcliente and artconcepto.idart=detalleventa.idart and artconcepto.idventa = venta.idventa and arqueo.fecha like '*" & FechaHoy & "*'" _
    & " ORDER BY arqueo.fecha"
    'MsgBox dummysql
    '*******************************************************************************************
    'la siguiente es la buena
    
    dummysql = "Select arqueo.idventa, arqueo.caja, arqueo.tarjeta, arqueo.fecha, arqueo.movimiento, arqueo.concepto, clientes.nombre, clientes.apellidos, venta.acuenta, " _
    & " clientes.telefono from arqueo, clientes, venta where venta.idventa = arqueo.idventa " _
    & " and clientes.idcliente = venta.idcliente and arqueo.fecha like '*" & FechaHoy & "*'" _
    & " ORDER BY arqueo.fecha"
    
    'dummysql = "select arqueo.idventa, venta.idventa from arqueo, venta where venta.idventa=arqueo.idventa"
    
'    dummysql = "select arqueo.idventa, arqueo.fecha, arqueo.caja, arqueo.movimiento," _
'    & " arqueo.concepto, clientes.nombre from arqueo, venta, clientes where arqueo.fecha" _
'    & " like '*" & FechaHoy & "*' and arqueo.idventa = venta.idventa and clientes.idcliente" _
'    & " = venta.idcliente"
    
'    dummysql = "Select arqueo.idventa, arqueo.fecha, arqueo.caja, arqueo.movimiento, venta.idcliente, venta.total," _
'    & " venta.pagado from arqueo INNER JOIN venta ON arqueo.idventa = venta.idventa" _
'    & " where arqueo.fecha like '*" & FechaHoy & "*'"
    'MsgBox (dummysql)
    Set RsArqueo = bdtienda.OpenRecordset(dummysql)
    TotalCashMorning = 0
    TotalCashTarde = 0
    TotalCardMorning = 0
    TotalCardTarde = 0
    TotalMorning = 0
    TotalTarde = 0
    TotalSilver = 0
    TotalClothing = 0
    Set Data1.Recordset = RsArqueo
    'On Error Resume Next
    Dim dumtarjeta As Currency
    Dim DumTarjeta2 As Currency
'    Dim RsSilver As Recordset
'    Set RsSilver = bdtienda.OpenRecordset("Select * from detalleventa")
'    RsSilver.Filter "idart", "18234"
    With RsArqueo
        If .RecordCount = 0 Then Exit Sub
        .MoveLast
        .MoveFirst
        If .RecordCount = 0 Then
            'MsgBox ("No hay datos de " & FechaHoy)
        Else
            Do Until .EOF
                    If IsNull(!Tarjeta) Then
                        DumTarjeta2 = 0
                    Else
                        DumTarjeta2 = !Tarjeta
                    End If
                    If !movimiento = True Then
                        Arqueo = Arqueo + !caja + DumTarjeta2
                        dumtarjeta = dumtarjeta + DumTarjeta2
                    Else
                        Arqueo = Arqueo - !caja
                    End If
                    If Hour(!fecha) < 15 Then
'¡Dim , TotalCashTarde, TotalCardMorning, TotalCardTarde As Currency
                        TotalCashMorning = TotalCashMorning + !caja
                        TotalCardMorning = TotalCardMorning + !Tarjeta
                        TotalMorning = TotalMorning + !caja + !Tarjeta
                    Else
                        TotalCashTarde = TotalCashTarde + !caja
                        TotalCardTarde = TotalCardTarde + !Tarjeta
                        TotalTarde = TotalTarde + !caja + !Tarjeta
                    End If
'                    If !Codigo = "Plata" Then
'                        TotalSilver = TotalSilver + !caja + !Tarjeta
'                    Else
                        TotalClothing = TotalClothing + !caja + !Tarjeta
'                    End If
                .MoveNext
            Loop
        End If
    End With
    lbcashmañana = TotalCashMorning
    lbcashtarde = TotalCashTarde
    lbtotalcardmorning = TotalCardMorning
    lbtotalcardtarde = TotalCardTarde
    lbplata = TotalSilver
    Lbropa = TotalClothing
        'MsgBox (VentasHoy & " ventas y " & ApartadosHoy & " apartados.")
        'MsgBox (Arqueo & " €")
        lbtotal = Arqueo & " €"
        lbefectivo = Arqueo - dumtarjeta
        lbtarjeta = dumtarjeta
    Arqueo = 0
    VentasHoy = 0
    DBGrid1.Columns(0).Width = 700
    DBGrid1.Columns(1).Width = 700
    DBGrid1.Columns(2).Width = 700
    DBGrid1.Columns(3).Width = 1800
    DBGrid1.Columns(4).Width = 400
    DBGrid1.Columns(5).Width = 1500
    DBGrid1.Columns(8).Width = 700

End Sub
Private Sub DBGrid1_DblClick()
    Dim dummyid As String
    On Error Resume Next
    dummyid = DBGrid1.Text
'    dummysql = "Select venta.idventa, venta.idcliente, clientes.nombre from venta inner join clientes on" _
'    & " venta.idcliente = clientes.idcliente where venta.idventa = " & dummyid
'    dummysql = "Select venta.idventa, venta.idcliente, clientes.nombre, detalleventa.preciofinal" _
'    & " from venta inner join clientes on venta.idcliente = clientes.idcliente" _
'    & " inner join detalleventa on venta.idventa = detalleventa.idventa where detalleventa.idventa = " & dummyid

'dummysql = "SELECT d.idventa, d.idcliente, d.preciofinal, a.idart, a.tipo FROM detalleventa d INNER JOIN articulos a" _
'    & " ON d.idart = a.idart WHERE d.idventa = " & dummyid

dummysql = "SELECT d.idventa, d.idcliente, d.preciofinal, a.idart, a.tipo FROM" _
& " detalleventa d INNER JOIN articulos a ON d.idart = a.idart WHERE d.idventa = " & dummyid

    
   ' MsgBox (dummysql)
    Set RsDetalVentaXCliente = bdtienda.OpenRecordset(dummysql)
    Set Data2.Recordset = RsDetalVentaXCliente
    'MsgBox (RsDetalVentaXCliente!IdVenta & " / " & RsDetalVentaXCliente!IdCliente & " / " & RsDetalVentaXCliente!PrecioFinal & " / " & RsDetalVentaXCliente!Idart & " / " & RsDetalVentaXCliente!Tipo)
    'MsgBox (Data2.Recordset.RecordCount)
    DBGrid2.Refresh

End Sub

Private Sub Form_Load()
    txtFecha = Date
    DTPicker1.Value = Date
    'Calendar1.Value = Date
    MuestraResultado
End Sub

Private Sub Form_Resize()
DBGrid1.Width = Me.Width - 500
DBGrid2.Width = Me.Width - 500
'Picture1.Top = Me.Height - 1000 'Picture1.Height
'Calendar1.Top = Me.Height - (Calendar1.Height + 1015)
End Sub

