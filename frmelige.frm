VERSION 5.00
Begin VB.Form Elige 
   BackColor       =   &H00FF8080&
   Caption         =   "Elige"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCajaDiaria 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja Diaria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdVale 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Vale de compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton CmdSinIva 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Calcula Art. sin IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Añadir artículos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Flamenca"
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton CmdWebCam 
      Caption         =   "Camara"
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CmdPruebaPrint 
      Caption         =   "Prueba Printer"
      Height          =   855
      Left            =   6120
      TabIndex        =   18
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton CmdPrinter 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Info de Artículo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdfacturas 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Facturas"
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdInformeProveedor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Informe Por Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Hacer Inventario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   3600
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Command7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdVerTablas 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ver Todas las Tablas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   3000
   End
   Begin VB.CommandButton cmdBorrarCliente 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Borrar Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdresultados 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Resultados"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Artículos Prestados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Artículos Apartados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Etiquetas"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdelimina 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Eliminar Artículo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Informe"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ver Totales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdsalida 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sale de Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdentradaextra 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Entra en Caja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdventa 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   3240
      Picture         =   "frmelige.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   570
   End
End
Attribute VB_Name = "Elige"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FechaHoy As Date
Dim VentasHoy As Integer
Dim ApartadosHoy As Integer
Dim Pass As String

Private Sub cmdapartado_Click()

End Sub

Private Sub cmdcaja_Click()
    FrmCajaDiaria.Show
'    frmarqueo.Show
'    On Error Resume Next
'    FechaHoy = InputBox("¿Qué fecha quiere ver?")
'    Set RsArqueo = bdtienda.OpenRecordset("arqueo")
'    On Error Resume Next
'    With RsArqueo
'        .MoveLast
'        .MoveFirst
'        If .RecordCount = 0 Then
'            MsgBox ("No hay datos de " & FechaHoy)
'        Else
'            Do Until .EOF
'                If Format(!Fecha, "Short Date") = FechaHoy Then
'                    If !movimiento = True Then
'                        Arqueo = Arqueo + !caja
'                    Else
'                        Arqueo = Arqueo - !caja
'                    End If
'                End If
''                If Format(!Fecha, "Short Date") = FechaHoy Then
''                    If !acuenta > 0 Then
''                        Arqueo = Arqueo + !acuenta
''                        ApartadosHoy = ApartadosHoy + 1
''                    Else
''                        Arqueo = Arqueo + !total
''                        VentasHoy = VentasHoy + 1
''                    End If
''                End If
'                .MoveNext
'            Loop
'        End If
'    End With
'        'MsgBox (VentasHoy & " ventas y " & ApartadosHoy & " apartados.")
'        MsgBox (Arqueo & " €")
'    Arqueo = 0
'    VentasHoy = 0
End Sub


Private Sub cmdBorrarCliente_Click()
If MsgBox("¿Desea borrar los datos de un cliente?", vbYesNo, "ATENCIÓN: BORRANDO DATOS") = vbNo Then Exit Sub
Dim TempId
Dim RsBorraCliente As Recordset
On Error GoTo sehodio

TempId = InputBox("Escriba el código de cliente", "Código de Cliente")
Set RsBorraCliente = bdtienda.OpenRecordset("select * from venta where idcliente like " & TempId & " and pagado = false")
'RsBorraCliente.MoveLast
'RsBorraCliente.MoveFirst

If RsBorraCliente.EOF = False Then
    MsgBox ("El cliente " & TempId & " tiene prendas apartadas." & Chr(13) _
        & "Solucione este problema y vuelva a intentar el borrado.")
    Exit Sub
End If
Set RsBorraCliente = bdtienda.OpenRecordset("select * from prestamo where idcliente like " & TempId)
'RsBorraCliente.MoveLast
'RsBorraCliente.MoveFirst
If RsBorraCliente.EOF = False Then
    MsgBox ("El cliente " & TempId & " tiene prendas en préstamo." & Chr(13) _
        & "Solucione este problema y vuelva a intentar el borrado.")
    Exit Sub
End If
Set RsBorraCliente = bdtienda.OpenRecordset("Select * from clientes where idcliente = " & TempId)
RsBorraCliente.MoveLast
RsBorraCliente.MoveFirst
If RsBorraCliente.EOF = False Then
    If MsgBox("¿Desea continuar con el borrado de los datos de: " & Chr(13) _
    & Chr(13) & Chr(9) & RsBorraCliente!Nombre & " " & RsBorraCliente!apellidos, vbYesNo, "ATENCIÓN: BORRANDO DATOS") = vbYes Then
        BorraVentas TempId
        RsBorraCliente.Delete
        'RsBorraCliente.Update
    MsgBox ("Cliente Borrado")
    End If
End If



Exit Sub
sehodio:
MsgBox (Err.Number & Err.Description & "Operación Cancelada")
End Sub
Private Sub BorraVentas(ByVal tid As Integer)
Dim RsBorrarVentas As Recordset
Set RsBorrarVentas = bdtienda.OpenRecordset("select * from venta where idcliente=" & tid)
On Error Resume Next
Do Until RsBorrarVentas.EOF
    RsBorrarVentas.Delete
    RsBorrarVentas.MoveNext
Loop
End Sub

Private Sub CmdCajaDiaria_Click()
    FrmCajaDiaria.Show
End Sub

Private Sub cmdElimina_Click()
Dim numartic As Integer
Dim rselimina As Recordset
Dim dummy As String

On Error GoTo sehodio
numartic = InputBox("Escriba el número del artículo para eliminar")
If numartic = 0 Then Exit Sub
Set rselimina = bdtienda.OpenRecordset("select * from articulos where idart = " & numartic)
dummy = "desea eliminar el artículo" & Chr(13)
dummy = dummy & "número: " & Chr(9) & rselimina!idArt
dummy = dummy & Chr(13) & "tipo: " & Chr(9) & rselimina!tipo
dummy = dummy & Chr(13) & "precio: " & Chr(9) & rselimina!PrecioVenta
dummy = dummy & Chr(13) & "vendido: " & Chr(9) & rselimina!vendido

If MsgBox(dummy, vbYesNo) = vbYes Then
    rselimina.Delete
End If

Exit Sub
sehodio:
MsgBox ("No se ha podido eliminar el artículo")
End Sub

Private Sub cmdentradaextra_Click()
    On Error GoTo sehodio
    Dim fechasale As String
    Dim cantidad As Currency
    Dim Concepto As String
    
    fechasale = InputBox("¿Qué fecha tiene el apunte?" & Chr(13) & "(Aceptar=hoy)") = ""
    If fechasale = True Then fechasale = FechaTrabajo
    
    cantidad = InputBox("¿Qué cantidad extra entra en caja?")
    If cantidad = 0 Then
        MsgBox ("Debe indicar  una cantidad")
        Exit Sub
    End If
    
    Concepto = InputBox("¿Qué concepto desea asignar a la entrada?")
    If Concepto = "" Then
        MsgBox ("Debe asigar un concepto a cada entrada")
        Exit Sub
    End If
'    Set RsArqueo = bdtienda.OpenRecordset("arqueo")
'
'    With RsArqueo
'        .AddNew
'        !IdVenta = 0
'        !Fecha = fechasale
'        !caja = cantidad
'        !Concepto = Concepto
'        !movimiento = False
'        .Update
'    End With
        Set RsVenta = bdtienda.OpenRecordset("venta")
    
    With RsVenta
        .AddNew
        !IdCliente = 2
        !total = cantidad
        !pagado = True
        !pago = 0
        MueveCaja !IdVenta, FechaTrabajo, !total, , True, Concepto
        .Update
        .MoveLast
     End With
    Exit Sub
sehodio:
    MsgBox ("Ha ocurrido un error. Inténtelo de nuevo")
    

End Sub

Private Sub cmdfacturas_Click()
frmFacturas.Show
End Sub

Private Sub cmdInformeProveedor_Click()
    FrmInformeProveedor.Show
End Sub

Private Sub CmdPrinter_Click()
    'FrmPrinters.Show
End Sub

Private Sub CmdPruebaPrint_Click()
    'ImprimeTicket ""
    'MandaCodigo 2
    'ImprimeAncho "    CANELA"
    ImprimeTicket "123456789012345678901234567890123456789012345678901234567890"
'    Printer.ScaleMode = 6
'
'    Printer.PaintPicture Image1.Picture, 0, 0, 10, 10
'
'
'    Printer.CurrentX = 12
'    Printer.CurrentY = 0
'    Printer.Print "Canela"
'
'    Printer.CurrentX = 12
'    Printer.CurrentY = 3
'    Printer.Print "Pl. Iglesia, 12"
'
'    Printer.CurrentX = 12
'    Printer.CurrentY = 6
'    Printer.Print "21800 Moguer"
'
'    Printer.CurrentX = 12
'    Printer.CurrentY = 9
'    Printer.Print "959 37 33 58"

'    Dim i As Integer
'    For i = 0 To 80
'        Printer.Line (i, 0)-(i, 20)
'        Printer.Line (0, i)-(80, i)
'    Next i
    Printer.EndDoc
End Sub

Private Sub cmdresultados_Click()
    frmResultados.Show
End Sub

Private Sub cmdsalida_Click()
    On Error GoTo sehodio
    Dim fechasale As String
    Dim cantidad As Currency
    Dim Concepto As String
    
    fechasale = InputBox("¿Qué fecha tiene el apunte?" & Chr(13) & "(Aceptar=hoy)") = ""
    If fechasale = True Then fechasale = FechaTrabajo
    
    cantidad = InputBox("¿Qué Cantidad?")
    If cantidad = 0 Then
        MsgBox ("Debe indicar  una cantidad")
        Exit Sub
    End If
    
    Concepto = InputBox("¿Qué concepto desea asignar a la salida?")
    If Concepto = "" Then
        MsgBox ("Debe asigar un concepto a cada salida")
        Exit Sub
    End If
'    Set RsArqueo = bdtienda.OpenRecordset("arqueo")
'
'    With RsArqueo
'        .AddNew
'        !IdVenta = 0
'        !Fecha = fechasale
'        !caja = cantidad
'        !Concepto = Concepto
'        !movimiento = False
'        .Update
'    End With
        Set RsVenta = bdtienda.OpenRecordset("venta")
    
    With RsVenta
        .AddNew
        !IdCliente = 2
        !total = 0 - cantidad
        !pagado = True
        !pago = 0
        MueveCaja !IdVenta, FechaTrabajo, !total, , False, Concepto
        .Update
        .MoveLast
     End With
    Exit Sub
sehodio:
    MsgBox ("Ha ocurrido un error. Inténtelo de nuevo")
    
    
End Sub

Private Sub CmdSinIva_Click()
    On Error GoTo sehodio
    Dim RsIvaArt As Recordset
    Dim indice As Double
    Dim TotalRegistros As Integer
    Set RsIvaArt = bdtienda.OpenRecordset("articulos")
    RsIvaArt.MoveLast
    TotalRegistros = RsIvaArt.RecordCount
    RsIvaArt.MoveFirst
    Dim TipoIva As Double
    Dim Divisor As Double
    TipoIva = Int(InputBox("Cuál es el tipo de IVA?", "IVA", 18))
    Divisor = ((TipoIva + 100) / 100)
    Screen.MousePointer = vbHourglass
    With RsIvaArt
        Do While Not RsIvaArt.EOF
            .Edit
            !iva = TipoIva
            !siniva = !PrecioVenta / Divisor
            .Update
            indice = indice + 1
            Me.Caption = CInt((indice * 100) / TotalRegistros) & " % completado"
            DoEvents
            .MoveNext
        Loop
    End With
    RsIvaArt.Close
    MsgBox ("Se ha completado la operación")
    Me.Caption = "Elige"
    Screen.MousePointer = vbNormal
Exit Sub
sehodio:
MsgBox ("No se ha completado la operación" & Chr(13) & Err.Description)

End Sub

Private Sub cmdVale_Click()
    If MsgBox("¿Desea hacer un vale de canje?", vbYesNo, "Vale") = vbYes Then
        HaceVale (InputBox("Importe del vale", "Importe", "0,00"))
    End If
End Sub

Private Sub cmdventa_Click()
    Venta.Show
End Sub

Private Sub cmdVerTablas_Click()
FrmVerTablas.Show
End Sub

Private Sub CmdWebCam_Click()
'FrmWebCam.Show
End Sub

Private Sub Command1_Click()
frmtotales.Show
End Sub

Private Sub Command10_Click()
    frmEncargo.Show
End Sub

Private Sub Command11_Click()
    FrmAddArt2.Show
End Sub



Private Sub Command3_Click()
    frminforme.Show
End Sub

Private Sub Command4_Click()
    FrmEtiquetas.Show
End Sub

Private Sub Command5_Click()
    ModoFrmApartado = "apartados"
    FrmApartados.Show
End Sub

Private Sub Command6_Click()
    ModoFrmApartado = "prestamos"
    FrmApartados.Show

End Sub


Private Sub Command7_Click()
Dim num As Integer
Dim RsDetal As Recordset
Dim inicio As Integer
Dim totaleuros As Currency
inicio = InputBox("inicio")
    Set RsVenta = bdtienda.OpenRecordset("venta")
    With RsVenta
        .MoveFirst
        Do Until .EOF = True
        If !IdVenta >= inicio Then
            Set RsDetal = bdtienda.OpenRecordset("select * from detalleventa where idventa = " & RsVenta!IdVenta)
            
            If RsDetal.EOF = True Then
                num = num + 1
                totaleuros = totaleuros + !total
            End If
        End If
        .MoveNext
        Loop
    End With
    MsgBox ("Ventas sin detalleventa = " & num)
    MsgBox ("total = " & totaleuros)
End Sub

Private Sub Command8_Click()
FrmInventario.Show
End Sub


Private Sub Command9_Click()
frmZoom.Show
End Sub

Private Sub Form_Activate()
    OcultaBotones
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 57 Then
        If LCase(Pass) = LCase("salas") Then
            FrmCajaDiaria.Show
        End If
        If LCase(Pass) = LCase("aula") Then
            MuestraBotones
        End If

        
        If LCase(Pass) = LCase("salas8") Then frmResultados.Show
        
        If LCase(Pass) = LCase("tablas") Then cmdVerTablas.Visible = True
        If LCase(Pass) = LCase("impor") Then frmImportar.Show

        Pass = ""
    Else
        Pass = Pass & Chr(KeyCode)
    End If
    If Len(Pass) > 6 Then Pass = ""
    If KeyCode = vbKeyF11 Then FrmAddArt2.Show
    
Dim CtrlDown
    CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = 65 And CtrlDown Then
        AbrirCajon
    End If
'***********************************************
'ESTO TAMBIÉN VA, CREO
'    If Shift = 2 And KeyCode = vbKeyA Then
'        Ctrl+A pressed
'    End If
'***********************************************
'ESTO TAMBIÉN VA, CREO
'    ShiftTest = Shift And 7
'    Select Case ShiftTest
'        Case 1
'            Debug.Print "Shift"
'        Case 2
'            Debug.Print "Ctrl"
'        Case 4
'            Debug.Print "Alt"
'        Case 3
'            Debug.Print "Shift + Ctrl"
'        Case 5
'            Debug.Print "Shift + Alt"
'        Case 6
'            Debug.Print "Ctrl + Alt"
'        Case 7
'            Debug.Print "Shift + Ctrl + Alt"
'    End Select
    
End Sub

Private Sub Form_Load()
    
    Set bdtienda = OpenDatabase(App.Path & "\canela.mdb")
    FechaTrabajo = HaceFecha(Now)
    
    ' ========== NUEVO: Inicializar módulo PrestaShop ==========
'    If InicializarModuloPS() Then
'        ' Opcional: mostrar mensaje de éxito
'        Debug.Print "? Conectado con PrestaShop"
'    Else
'        MsgBox "Advertencia: No se pudo conectar con PrestaShop" & vbCrLf & _
'               "El sistema funcionará solo con datos locales", vbExclamation, "Modo Offline"
'    End If
    ' ==========================================================
    
    BuscarPrestados
    BuscarApartados
    BlAlarmaQuitar = True
End Sub
Private Sub BuscarApartados()
    Dim RsApart As Recordset
    Dim FechaPres
    FechaPres = Format(Date - 15, "mm/dd/yy")
    Set RsApart = bdtienda.OpenRecordset("Select clientes.idcliente, nombre, apellidos," _
    & " telefono, direccion, venta.fecha, venta.acuenta, venta.total, detalleventa.idart" _
    & " from clientes" _
    & " inner join (venta inner join detalleventa on venta.idventa = detalleventa.idventa)" _
    & " on clientes.idcliente = venta.idcliente" _
    & " where venta.pagado = false and venta.fecha < #" & FechaPres & "# order by venta.fecha")
    If RsApart.EOF Then
        Timer2.Enabled = False
        Command5.ToolTipText = "No hay artículos apartados"
    Else
        Timer2.Enabled = True
        Command5.ToolTipText = "¡¡¡Hay artículos apartados desde hace 15 días!!!"
    End If
End Sub
Private Sub BuscarPrestados()
Dim RsPrestados As Recordset
Dim FechaPres
FechaPres = Format(Date - 2, "mm/dd/yy")
Set RsPrestados = bdtienda.OpenRecordset("Select * from prestamo where fecha < #" & FechaPres & "#")
If RsPrestados.EOF Then
    Timer1.Enabled = False
    Exit Sub
    Command6.ToolTipText = ""
Else
    RsPrestados.MoveLast
    RsPrestados.MoveFirst
    If RsPrestados.RecordCount > 0 Then Timer1.Enabled = True
    Command6.ToolTipText = "¡¡¡Hay prendas prestadas desde hace 2 días!!!"
End If
End Sub

Private Sub Timer1_Timer()
    Command6.BackColor = RGB(Int(Rnd * 255) + 50, 50, 50)
End Sub

Private Sub Timer2_Timer()
    Command5.BackColor = RGB(Int(Rnd * 255) + 50, 50, 50)

End Sub
Private Sub MuestraBotones()
    cmdInformeProveedor.Visible = True
    cmdfacturas.Visible = True
    Command8.Visible = True
    Command1.Visible = True
    cmdVerTablas.Visible = True

End Sub
Private Sub OcultaBotones()
'    cmdInformeProveedor.Visible = False
'    cmdfacturas.Visible = False
'    Command8.Visible = False
'    Command1.Visible = False
'    cmdVerTablas.Visible = False
'

End Sub
