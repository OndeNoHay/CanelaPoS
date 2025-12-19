VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmchart 
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   11655
      Begin VB.CommandButton cmdmostrar 
         Caption         =   "Mostrar"
         Height          =   615
         Left            =   8400
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
      Begin VB.ListBox lstpormes 
         Height          =   2595
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   7215
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   9000
         TabIndex        =   3
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmdXAño 
         Caption         =   "Ingresos - Gastos X Año"
         Height          =   615
         Left            =   8880
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
      End
   End
   Begin MSChart20Lib.MSChart Chart 
      Height          =   5295
      Left            =   840
      OleObjectBlob   =   "frmchart.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "frmchart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RsResultado As Recordset
Dim RsTotalXMes As Recordset
Dim TotalPunto As Currency
Dim SerieDato As Integer
Dim DatoNum As Integer
Dim RsSumaCoste As Recordset
Dim RsSumaVenta As Recordset


Private Sub HaceTabla()
Dim dummy As TableDef
Dim max As Long
Dim wrk As Workspace
Dim newtabla As TableDef
Dim IdxEjer As Index
Dim RsDuplica As Recordset
'On Error GoTo merde
'Dim duplo As String
'duplo = "Select * into " & numero & " from TablaOrigenUsuario where 1 = 1"
'dbjuego.Execute duplo
'dbjuego.Execute "CREATE INDEX idejercicio on " & numero & " (idejercicio)"


    Set newtabla = bdtienda.CreateTableDef("TotalXMes") 'nombrecompleto)
'
    With newtabla
        .Fields.Append .CreateField("mes", dbText)
        .Fields.Append .CreateField("año", dbText)
        .Fields.Append .CreateField("efectivo", dbCurrency)
        .Fields.Append .CreateField("tarjeta", dbCurrency)
        .Fields.Append .CreateField("numarqueos", dbInteger)
        
'        .Fields.Append .CreateField("pasa", dbBoolean)
'        .Fields.Append .CreateField("veces", dbBoolean)
'        .Fields.Append .CreateField("numfrases", dbInteger)
'        .Fields.Append .CreateField("fecha", dbDate)
    bdtienda.TableDefs.Append newtabla
''         Dim rsdum As Recordset
''        Set rsdum = dbjuego.OpenRecordset(Tablausuar)
''        rsdum.AddNew
''        rsdum!idunidad = 0
''        rsdum.Update
'    IdxEjer.Fields.Append .CreateField("idejercicio")
'        .Indexes.Append IdxEjer
 '       .Indexes.Refresh
'
'
    End With
'añadetemas numero, grupousuario 'nombrecompleto

End Sub
Private Sub CalculaXMes()
On Error GoTo sehodio
Dim SumaEfectivo As Currency
Dim SumaTarjeta As Currency
Dim Mes As String
Dim Año As String
Dim SumaAno As Currency
Set RsResultado = bdtienda.OpenRecordset("select * from arqueo order by fecha")
Set RsTotalXMes = bdtienda.OpenRecordset("totalxmes")
Do Until RsTotalXMes.EOF
    RsTotalXMes.Delete
    RsTotalXMes.MoveNext
Loop
Dim contador As Integer
With RsResultado
    .MoveFirst
    Do Until .EOF
'        If Año = Year(!Fecha) Then
            If Mes = "" Then
                Mes = Month(!fecha)
                Año = Year(!fecha)
            ElseIf Mes <> Month(!fecha) Then
                RsTotalXMes.AddNew
                RsTotalXMes!efectivo = SumaEfectivo
                RsTotalXMes!Tarjeta = SumaTarjeta
                RsTotalXMes!Mes = Mes
                RsTotalXMes!Año = Año
                RsTotalXMes!numarqueos = contador
                lstpormes.AddItem Mes & "/" & Año & Chr(9) & " = " & Chr(9) & SumaEfectivo & " + " & SumaTarjeta & Chr(9) & "=" & Chr(9) & SumaEfectivo + SumaTarjeta
                RsTotalXMes.Update
                SumaEfectivo = 0
                SumaTarjeta = 0
                contador = 0
                Mes = ""

            ElseIf Mes = Month(!fecha) Then
                If IsNull(!caja) = False Then
                    If !caja < 0 Then
                        SumaEfectivo = SumaEfectivo - !caja
                    Else
                        SumaEfectivo = SumaEfectivo + !caja
                    End If
                End If
                If IsNull(!Tarjeta) = False Then SumaTarjeta = SumaTarjeta + !Tarjeta
                contador = contador + 1
                .MoveNext
            End If
'        Else
'            Año = Year(!Fecha)
'            Mes = Month(!Fecha)
'            SumaEfectivo = 0
'            SumaTarjeta = 0
'        End If
    Loop
                RsTotalXMes.AddNew
                RsTotalXMes!efectivo = SumaEfectivo
                RsTotalXMes!Tarjeta = SumaTarjeta
                RsTotalXMes!Mes = Mes
                RsTotalXMes!Año = Año
                RsTotalXMes!numarqueos = contador
'                lstpormes.AddItem Mes & "/" & Año & " = " & SumaEfectivo & " + " & SumaTarjeta
                lstpormes.AddItem Mes & "/" & Año & Chr(9) & " = " & Chr(9) & SumaEfectivo & " + " & SumaTarjeta & Chr(9) & "=" & Chr(9) & SumaEfectivo + SumaTarjeta
                RsTotalXMes.Update
                SumaEfectivo = 0
                SumaTarjeta = 0
                contador = 0
                Mes = ""
End With
Exit Sub
sehodio:
    If Err.Number = 3078 Then
        HaceTabla
        CalculaXMes
    End If
End Sub

Private Sub Chart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Chart.ToolTipText = TotalPunto
'TotalPunto = 0

End Sub

Private Sub Chart_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error Resume Next
    DatoNum = DataPoint - 1
    SerieDato = Series
    
    Dim contador As Integer
    With RsResultado
        .MoveFirst
        Do Until contador = DatoNum
            contador = contador + 1
            .MoveNext
        Loop
        Select Case SerieDato
            Case 2
                TotalPunto = !Tarjeta
            Case 1
                TotalPunto = !efectivo
            Case 3
                TotalPunto = !efectivo + !Tarjeta
        End Select
        
End With

End Sub

Private Sub cmdmostrar_Click()
lstpormes.Clear
CalculaXMes
Set RsResultado = bdtienda.OpenRecordset("totalxmes")
Dim i As Integer
i = 1

'    Chart.RowCount = 0
'    Chart.AllowDynamicRotation = True
'    Chart.RandomFill = False
'    Dim xnota() As Integer
'    Chart.RowCount = 1
 
    Dim dumx As Integer
    RsResultado.MoveLast
    RsResultado.MoveFirst
    Do Until RsResultado.EOF
        dumx = dumx + 1
        RsResultado.MoveNext
    Loop
     'dumx = RsResultado.RecordCount
    
ReDim datos(0 To dumx + 1, 1 To 4)
    RsResultado.MoveFirst
Dim x As Integer
'x = 1
With RsResultado
    Do Until .EOF
        datos(x, 1) = !Mes & "/" & !Año
        datos(x, 2) = !efectivo
        datos(x, 3) = !Tarjeta
        datos(x, 4) = !efectivo + !Tarjeta
        x = x + 1
        .MoveNext
    Loop
End With
'Chart.chartType = 2
Chart.ColumnLabel = "Ingresos"
Chart.ShowLegend = True
Chart.AllowDithering = True
Chart.AllowSelections = True
Chart.ChartData = datos
'    Chart.ColumnCount = 2
'    For y = 0 To dumx - 1
'
'        Chart.Row = i
'        Chart.RowLabel = RsResultado!Mes & "/" & RsResultado!Año
'        Chart.Column = 1
'    '    Chart.Data = xnota(i)
'        If IsNull(RsResultado!efectivo) = True Then
'            Chart.Data = 0
'        Else
'            Chart.DataGrid.SetData i, 1, RsResultado!efectivo, 0
'            'Chart.Data = RsResultado!efectivo
'        End If
'        If IsNull(RsResultado!Tarjeta) = True Then
'            Chart.Data = 0
'        Else
'            Chart.DataGrid.SetData i, 2, RsResultado!Tarjeta, 0
'            'Chart.Data = RsResultado!Tarjeta
'        End If
'
'
'        Chart.RowCount = i
'        RsResultado.MoveNext
'    'Loop
'    Next y
'    Chart.Refresh


End Sub

Private Sub cmdprint_Click()
' this code uses the clipboard so it is a good idea to clear its
    ' contents first
    Clipboard.Clear


    ' this copys the chart to the clipboard as a metafile
    Chart.EditCopy


    ' this sets the orientation of the printer to landscape
    Printer.Orientation = vbPRORLandscape


    ' this sends the metafile to the printer and scale the picture to
    '  the paper size
    Printer.PaintPicture Clipboard.GetData(3), 0, 0, _
          Printer.ScaleWidth, Printer.ScaleHeight


    ' tells the printer to print
    Printer.EndDoc


End Sub

Private Sub cmdXAño_Click()
Dim Añobusca As String
Dim FechaInicio, FechaFin As String
Dim Ssql As String
Dim Beneficios() As Currency
Dim TotalCoste() As Currency
Dim TotalIngreso() As Currency

Dim x As Integer
For i = 2003 To Year(Now)
    FechaInicio = "01/01/" & i
    FechaFin = "31/12/" & i
    
    Ssql = "select sum(preciocompra)as coste from articulos where fechacompra between #" & FechaInicio & "# and #" & FechaFin & "#"
    Set RsSumaCoste = bdtienda.OpenRecordset(Ssql)
    
    Ssql = "select sum(total)as SumaVenta from venta where fecha between #" & FechaInicio & "# and #" & FechaFin & "#"
    
    Set RsSumaVenta = bdtienda.OpenRecordset(Ssql)
    ReDim Preserve Beneficios(x)
    Beneficios(x) = RsSumaVenta!sumaVenta - RsSumaCoste!coste
    ReDim Preserve TotalCoste(x)
    TotalCoste(x) = RsSumaCoste!coste
    ReDim Preserve TotalIngreso(x)
    TotalIngreso(x) = RsSumaVenta!sumaVenta
    x = x + 1
    
    'MsgBox ("Año: " & i & Chr(13) & " Suma Ventas: " & RsSumaVenta!sumaVenta & Chr(13) & " Suma Compra : " & RsSumaCoste!coste & Chr(13) & " Resultado: " & RsSumaVenta!sumaVenta - RsSumaCoste!coste)
Next i
RsSumaCoste.Close
RsSumaVenta.Close


ReDim datos(0 To x + 1, 1 To 4)
Dim Y As Integer
'x = 1
For Y = 0 To x - 1
        datos(Y, 1) = "31/12/" & 2003 + Y
        datos(Y, 2) = TotalIngreso(Y) - TotalCoste(Y)
        datos(Y, 3) = TotalCoste(Y)
        datos(Y, 4) = TotalIngreso(Y)
        
        
Next Y

'Chart.chartType = 2
Chart.ColumnLabel = "Ingresos"
Chart.ShowLegend = True
Chart.AllowDithering = True
Chart.AllowSelections = True
Chart.ChartData = datos
End Sub

Private Sub Form_Load()
Chart.Left = 0
Chart.Width = Me.Width
End Sub

Private Sub Form_Resize()
Chart.Width = Me.Width
Chart.Height = Me.Height - Frame1.Height - 1000 '2000
Frame1.Top = Me.Height - (Frame1.Height + 500)
'cmdmostrar.Top = Me.Height - 1800
'cmdprint.Top = Me.Height - 1800
'cmdXAño.Top = Me.Height - 1800
'lstpormes.Top = Me.Height - lstpormes.Height - 1200
End Sub
