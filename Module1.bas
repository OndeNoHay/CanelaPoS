Attribute VB_Name = "Module1"
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Option Explicit
Global bdtienda As Database
Global RsArticulo As Recordset ' Recordset
Global RsCliente As Recordset
Global RsVenta As Recordset
Global RsDetalVenta As Recordset
Global RsApartado As Recordset
Global RsArtApartado As Recordset
Global RsDetalApartado As Recordset
Global SqlArticulos As String
Global IdArtSelec As Double
Global IdCliSelec As Double
Global NumArtVend As Double
Global SumaTotal As Currency
Global DescuentoTotal As Currency
Global ModoBusca As String
Global IdVenta As Double
Global IdCliente As Double
Global FormaPago As Integer
Global NuevoCliente As Boolean
Global AñadeArtic As Boolean
Global VentaApartado As Boolean
Global Modo As String
Global IdArtApart() As Double
Global PreFinalApart() As Currency
Global CodigoBusca As String
Global VerTodos As Boolean
Global Arqueo As Currency
Global RsArqueo As Recordset
Global FechaTrabajo As Date
Global hablando As Boolean
Global PagoTarjeta As Currency
Global RsPrestamo As Recordset
Global ModoFrmApartado As String
Global Permiso As Boolean
Global IdVentaApartado As Double

Global ArtApartParaPagar As Double
Global BlAlarmaQuitar As Boolean
Global MsgAlarma As String
Global IdZoom As Double
Global VerTabla As String

Global RsBuscaArticulos As Recordset
Global IdArtBuscado As Double


Global Dumresp As String
Global Dumprecio As Currency
Global Dumtipo As String
Global DumCode As String

'para genericos de goyse
Global BlGoyse As Boolean
Global GoyseCode As String

Type GenData
    IdVenta As Double
    Modo As String
    fecha As Date
    IdCliente As Double
    FormaPago As String
    SumaTotal As Currency
    IVATotal As Currency
    BaseTotal As Currency
    ACuenta As Currency
    Tarjeta As Currency
    Nombre As String
End Type
Type ArtTicket
    Idart As Double
    Descripcion As String
    Precio As Currency
    Descuento As Currency
    PrecioFinal As Currency
End Type
Global Header As GenData
Global Ticket() As ArtTicket

Global IdCliFoto As Double
Type Dinerito
    Arqueo As Currency
    TotalMorning As Currency
    TotalTarde As Currency
    TotalCashMorning As Currency
    TotalCashTarde As Currency
    TotalCardMorning As Currency
    TotalCardTarde As Currency
End Type

Global FotoAddArt As Boolean
Global DirFoto As String

Global TicketRegalo As Boolean



Public Sub DevuelvePrestamo()
    Dim x As Double
    Dim dummsg As String
    Set RsPrestamo = bdtienda.OpenRecordset("select * from prestamo where idcliente like " & IdCliente)
    If RsPrestamo.RecordCount > 0 Then
        If MsgBox("Tiene articulos prestados desde " & Chr(13) _
            & Chr(9) & RsPrestamo!fecha & Chr(13) _
            & "¿Quiere devolverlos?", vbYesNo) = vbYes Then
            With RsPrestamo
                .MoveFirst
                dummsg = "Anote las referencias de los artículos " & Chr(13) & "si los quiere comprar: " & Chr(13) & Chr(13)
                dummsg = dummsg & "Artículo:" & Chr(9) & "Fecha del préstamo:" & Chr(13)
                dummsg = dummsg & "-----------" & Chr(9) & "-------------------"
                Do Until .EOF
                    dummsg = dummsg & Chr(13) & !Idart & Chr(9) & !fecha
                    .MoveNext
                Loop
                If MsgBox(dummsg, vbYesNo) = vbYes Then
                    .MoveFirst
                    Do Until .EOF
                        .Delete
                        .MoveNext
                    Loop
                    MsgBox ("Devolución completada. Artículos disponibles para venta")
                    BlAlarmaQuitar = False
                    frmalarma.Show 1
                    'MsgAlarma = "QUITAR LAS ALARMAS"
                Else
                    MsgBox ("Devolución Cancelada")
                End If
            End With
        
        End If
    End If

End Sub

Public Sub MueveCaja(Venta As Double, fecha As Date, Importe As Currency, _
    Optional Tarjeta As Currency, Optional Entra As Boolean, Optional Concepto As String)
    Set RsArqueo = bdtienda.OpenRecordset("arqueo")
    With RsArqueo
        .AddNew
        'If Entra = False Then !movimiento = False
        
        !IdVenta = Venta
'        If PagoTarjeta = SumaTotal Then Importe = 0 'Importe - PagoTarjeta
'        If PagoTarjeta = 0 Then
'            !caja = Importe
'        ElseIf PagoTarjeta < SumaTotal And PagoTarjeta > 0 Then
'            !caja = Importe - PagoTarjeta
'        End If
        !caja = Importe
        !Tarjeta = PagoTarjeta
        !Concepto = Concepto
        !fecha = HaceFecha(fecha)
        .Update
    End With
End Sub
Public Sub PlayWave(varWave)
   sndPlaySound varWave, 1
End Sub


Sub DrawBarcode(ByVal bc_string As String, obj As Object)
'Thanks to someone on PSC to give me information about BarCode
Dim xpos!
Dim Y1!
Dim Y2!
Dim dw%
Dim Th!
Dim tw
Dim new_string$
    If bc_string = "" Then obj.Cls: Exit Sub

Dim bc(90) As String
    bc(1) = "1 1221"

    bc(2) = "1 1221"
    bc(48) = "11 221"
    bc(49) = "21 112"
    bc(50) = "12 112"
    bc(51) = "22 111"
    bc(52) = "11 212"
    bc(53) = "21 211"
    bc(54) = "12 211"
    bc(55) = "11 122"
    bc(56) = "21 121"
    bc(57) = "12 121"
    bc(65) = "211 12"
    bc(66) = "121 12"
    bc(67) = "221 11"
    bc(68) = "112 12"
    bc(69) = "212 11"
    bc(70) = "122 11"
    bc(71) = "111 22"
    bc(72) = "211 21"
    bc(73) = "121 21"
    bc(74) = "112 21"
    bc(75) = "2111 2"
    bc(76) = "1211 2"
    bc(77) = "2211 1"
    bc(78) = "1121 2"
    bc(79) = "2121 1"
    bc(80) = "1221 1"
    bc(81) = "1112 2"
    bc(82) = "2112 1"
    bc(83) = "1212 1"
    bc(84) = "1122 1"
    bc(85) = "2 1112"
    bc(86) = "1 2112"
    bc(87) = "2 2111"
    bc(88) = "1 1212"
    bc(89) = "2 1211"
    bc(90) = "1 2211"
    bc(32) = "1 2121"
    bc(35) = ""
    bc(36) = "1 1 1 11"
    bc(37) = "11 1 1 1"
    bc(43) = "1 11 1 1"
    bc(45) = "1 1122"
    bc(47) = "1 1 11 1"
    bc(46) = "2 1121"
    bc(64) = ""
    bc(42) = "1 1221"
    bc_string = UCase(bc_string)
    obj.ScaleMode = 3
    obj.Cls
    obj.Picture = Nothing
    dw = CInt(obj.ScaleHeight / 40)
    If dw < 1 Then dw = 1
    Th = obj.TextHeight(bc_string)
    tw = obj.TextWidth(bc_string)
    new_string = Chr$(1) & bc_string & Chr$(2)
    Y1 = obj.ScaleTop
    Y2 = obj.ScaleTop + obj.ScaleHeight - 1.5 * Th
    obj.Width = 1.1 * Len(new_string) * (15 * dw) * obj.Width / obj.ScaleWidth
    xpos = obj.ScaleLeft
    Dim n, c, bc_pattern$, i
    For n = 1 To Len(new_string)
        c = Asc(Mid$(new_string, n, 1))
        If c > 90 Then c = 0
        bc_pattern$ = bc(c)
        For i = 1 To Len(bc_pattern$)
            Select Case Mid$(bc_pattern$, i, 1)
                Case " "
                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                xpos = xpos + dw
                Case "1"
                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                xpos = xpos + dw
                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &H0&, BF
                xpos = xpos + dw
                Case "2"
                obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
                xpos = xpos + dw
                obj.Line (xpos, Y1)-(xpos + 2 * dw, Y2), &H0&, BF
                xpos = xpos + 2 * dw
            End Select
        Next
    Next
    obj.Line (xpos, Y1)-(xpos + 1 * dw, Y2), &HFFFFFF, BF
    xpos = xpos + dw
    obj.Width = (xpos + dw) * obj.Width / obj.ScaleWidth
    obj.CurrentX = (obj.ScaleWidth - tw) / 2
    obj.CurrentY = Y2 + 0.25 * Th
    obj.Print bc_string

End Sub
Public Function SelectPrinter(ByVal printer_name As String) As Boolean
Dim i As Integer
    'SelectPrinter = True
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = printer_name Or Printers(i).DeviceName = printer_name & " en CANELA" Then
            Set Printer = Printers(i)
            SelectPrinter = True
            Exit For
        End If
    Next i
End Function

Public Sub HaceTicket()
    Dim TextoTicket As String
    TextoTicket = TextoTicket & Chr(13) & CentraTexto("www.canelamoda.es ", 44)
    
    TextoTicket = TextoTicket & Chr(13) & CentraTexto("Pl. Iglesia, 12", 44)
    
    TextoTicket = TextoTicket & Chr(13) & CentraTexto("21800 Moguer - 959 373 358", 44)

    TextoTicket = TextoTicket & Chr(13) & CentraTexto(CStr(Header.fecha), 44)
    
    TextoTicket = TextoTicket & Chr(13) & Header.FormaPago
    
    TextoTicket = TextoTicket & Chr(13) & "Nº Venta:" & Header.IdVenta & "           Cliente: " & Header.IdCliente

    TextoTicket = TextoTicket & Chr(13)
    
    If TicketRegalo = True Then
            TextoTicket = TextoTicket & Chr(13) & CentraTexto("TICKET REGALO ", 44)
    End If
    
    Printer.FontSize = 5
    Dim x As Integer
    Dim Y As Integer
    Dim i As Integer
    x = 0
    Y = 22
    For i = 1 To UBound(Ticket)
            Dim textoart As String
            Dim textoBase As String
            Dim textoIVA As String
            
            If TicketRegalo = True Then
                textoart = HaceStringArt(Ticket(i).Idart, Ticket(i).Descripcion, "¡Regalo!")
            Else
                textoart = HaceStringArt(Ticket(i).Idart, Ticket(i).Descripcion, Format(Ticket(i).PrecioFinal, "0.00") & "€")
            End If
            TextoTicket = TextoTicket & Chr(13) & textoart
    Next i
    TextoTicket = TextoTicket & Chr(13) & HaceStringArt("0000", "Bolsa", "0,05€")
    Y = Y + 2
    
    TextoTicket = TextoTicket & Chr(13)
    
    Printer.FontSize = 6
    If TicketRegalo = True Then
        
        textoart = HaceStringArt("TOTAL: ", "", "¡Regalo!")
        
        TextoTicket = TextoTicket & Chr(13) & textoart
        TextoTicket = TextoTicket & Chr(13) & "----------------------------------------" '& Chr(13)
        If Header.Tarjeta > 0 Then
            textoart = HaceStringArt("Pagado con Tarjeta: ", "", "¡Regalo!")
            TextoTicket = TextoTicket & Chr(13) & textoart
            TextoTicket = TextoTicket & Chr(13) & "----------------------------------------" '& Chr(13)
        End If
        If Header.ACuenta > 0 Then
            textoart = HaceStringArt("Entregado a Cuenta: ", "", "¡Regalo!")
            TextoTicket = TextoTicket & Chr(13) & textoart
            TextoTicket = TextoTicket & Chr(13) & "----------------------------------------" '& Chr(13)
        End If
    Else
        textoBase = HaceStringArt("Base imponible: ", "", Format(Header.BaseTotal, "0.00") & "€")
        
        textoIVA = HaceStringArt("IVA (21%): ", "", Format(Header.IVATotal, "0.00") & "€")
        
        textoart = HaceStringArt("TOTAL: ", "", Format(Header.SumaTotal, "0.00") & "€")
        
        TextoTicket = TextoTicket & Chr(13) & textoBase & Chr(13) & textoIVA & Chr(13) & textoart
        
        TextoTicket = TextoTicket & Chr(13) & "----------------------------------------" '& Chr(13)
        If Header.Tarjeta > 0 Then
            textoart = HaceStringArt("Pagado con Tarjeta: ", "", Format(Header.Tarjeta, "0.00") & "€")
            TextoTicket = TextoTicket & Chr(13) & textoart
            TextoTicket = TextoTicket & Chr(13) & "----------------------------------------" '& Chr(13)
        End If
        If Header.ACuenta > 0 Then
            textoart = HaceStringArt("Entregado a Cuenta: ", "", Format(Header.ACuenta, "0.00") & "€")
            TextoTicket = TextoTicket & Chr(13) & textoart
            TextoTicket = TextoTicket & Chr(13) & "----------------------------------------" '& Chr(13)
        End If
    End If
    Y = Y + 5
    
    Printer.FontSize = 6
    TextoTicket = TextoTicket & Chr(13) & "No se aceptan devoluciones pasados 7 días."
    
    TextoTicket = TextoTicket & Chr(13) & "No se aceptan devoluciones sin etiquetas."
    TextoTicket = TextoTicket & Chr(13) & "Cualquier devolución se hará con vale."
    
    Y = Y + 2
    
    TextoTicket = TextoTicket & Chr(13) & "No se devuelven artículos de fiesta."
    
    Y = Y + 2
    If Header.Modo = "Apartado" Then
        TextoTicket = TextoTicket & Chr(13) & "Las prendas apartadas se guardan 1 mes."
    End If
    
    If Header.Modo = "PRESTAMO" Then
        TextoTicket = TextoTicket & Chr(13) & "LAS PRENDAS PRESTADAS SE DEVUELVEN EL MISMO DIA"
    End If
    
    TextoTicket = TextoTicket & Chr(13) & "Conserve este ticket."
    'MsgBox TextoTicket
    'espacios en blanco para impresora generica sin corte
    TextoTicket = TextoTicket & Chr(13)
    TextoTicket = TextoTicket & Chr(13)
    TextoTicket = TextoTicket & Chr(13)
    TextoTicket = TextoTicket & Chr(13)
    TextoTicket = TextoTicket & Chr(13)
    TextoTicket = TextoTicket & Chr(13)
    TextoTicket = TextoTicket & Chr(13)
    
    ImprimeTicket TextoTicket
    Printer.EndDoc
'     Header.ACuenta = 0
'     Header.fecha = Now
'     Header.FormaPago = ""
'     Header.IdCliente = 0
'     Header.IdVenta = 0
'     Header.Modo = ""
'     Header.Nombre = ""
'     Header.SumaTotal = 0
'     Header.Tarjeta = 0
     TicketRegalo = False
    'Erase Ticket
End Sub
Private Function HaceStringArt(Idart, Descr, Importe As String) As String
    Dim anchura As Integer
    Dim Relleno As Integer
    Dim i As Integer
    Printer.ScaleMode = 4
    anchura = 40
    HaceStringArt = Idart & "-" & Descr
    Relleno = anchura - Len(HaceStringArt) - Len(Importe)
    
    For i = 0 To Relleno
        HaceStringArt = HaceStringArt & "_"
    Next i
    HaceStringArt = HaceStringArt & Importe
    Printer.ScaleMode = 6
    
End Function

Public Function HaceFecha(Temp As Variant) As Date
    Dim dia
    Dim Mes
    Dim Año
    Dim hora
    Dim minutos
    Dim segundos
    dia = Day(Temp)
    Mes = Month(Temp)
    Año = Year(Temp)
    hora = Hour(Temp)
    minutos = Minute(Temp)
    segundos = Second(Temp)
    HaceFecha = CDate(dia & "/" & Mes & "/" & Año & " " & hora & ":" & minutos & ":" & segundos)
    
End Function

Public Sub AbrirCajon()
On Error GoTo sehodio
    Open "COM1" For Output Access Write As #1

    Print #1, Chr(27) & Chr(112) & Chr(0); ""
    Close #1
Exit Sub
sehodio:
MsgBox ("No se ha podido abrir el cajon" & vbCrLf & Err.Number & " / " & Err.Description)
End Sub

Function Exists(f$) As Integer      'COMPRUEBA si existe un archivo
    On Error Resume Next
    Dim x&
    x& = FileLen(f$)
    If x& Then Exists% = True
End Function


                
'**************************************
' Name: Export Database to .CSV [DAO]
' Description:This function converts database table to .csv (comma separated values) file. ENJOY!
' By: Milos Todorovic
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=28631&lngWId=1'for details.'**************************************

Public Function ExportToCSV(DaoRecordset As Recordset, CsvFilePath As String) As Boolean
 
Dim Csv As Long
Dim CsvText As String
Dim CsvDatabase As Database
Dim CsvRecordset As Recordset
On Error GoTo CSV_Error
 ExportToCSV = True
' Set CsvDatabase = OpenDatabase(DatabaseFileName)
' Set CsvRecordset = CsvDatabase.OpenRecordset(DatabaseTableName)
 If Not DaoRecordset.RecordCount >= 1 Then Exit Function
 If Not DaoRecordset.EOF Then DaoRecordset.MoveFirst
 
 Randomize
 
 
 Open CsvFilePath For Append As #2
 Dim Scv As Double
 For Scv = 0 To DaoRecordset.Fields.Count - 1
    CsvText = CsvText & DaoRecordset.Fields(Scv).Name & ";"
 Next
 CsvText = Left(CsvText, Len(CsvText) - 1)
 Print #2, CsvText
 Close #2
 CsvText = ""
 Dim n
 Do While Not DaoRecordset.EOF
    Open CsvFilePath For Append As #3
    n = 0
    For Scv = 0 To DaoRecordset.Fields.Count - 1
        If DaoRecordset(Scv).Name = "cantidad" Then
            CsvText = CsvText & Chr(34) & CInt(Int((6 * Rnd()) + 1)) & Chr(34) & ";"
        ElseIf DaoRecordset(Scv).Name = "iva" Then
            CsvText = CsvText & Chr(34) & "53" & Chr(34) & ";"
            
        Else
            CsvText = CsvText & Chr(34) & DaoRecordset(n) & Chr(34) & ";"
'        Else
'            CsvText = CsvText & DaoRecordset(Csv) & ";"
        End If
        n = n + 1
    Next
    CsvText = Left(CsvText, Len(CsvText) - 1)
    Print #3, CsvText
    Close #3
    CsvText = ""
    DaoRecordset.MoveNext
 Loop
 Exit Function
 
CSV_Error:

    MsgBox (Err.Description)
 ExportToCSV = False
 
End Function

Public Sub HaceVale(Importe As Currency)
    Dim TextoTicket As String
    Dim textoart As String
    TextoTicket = TextoTicket & Chr(13) & CentraTexto("www.canelamoda.es", 44)
    
    TextoTicket = TextoTicket & Chr(13) & CentraTexto("Pl. Iglesia, 12", 44)
    
    TextoTicket = TextoTicket & Chr(13) & CentraTexto("21800 Moguer - 959 373 358", 44)

    TextoTicket = TextoTicket & Chr(13) & CentraTexto(CStr(Date), 44)
    
    TextoTicket = TextoTicket & Chr(13) & Header.FormaPago
    
    TextoTicket = TextoTicket & Chr(13) & CentraTexto("VALE DE COMPRA", 44)
    
    TextoTicket = TextoTicket & Chr(13)
    
    Printer.FontSize = 5
    Dim x As Integer
    Dim Y As Integer
    Dim i As Integer
    x = 0
    Y = 22
    
    Y = Y + 2
    
    TextoTicket = TextoTicket & Chr(13)
    
    Printer.FontSize = 6
    textoart = HaceStringArt("TOTAL: ", "", Format(Importe, "0.00") & "€")
    
    TextoTicket = TextoTicket & Chr(13) & textoart
    TextoTicket = TextoTicket & Chr(13) & "----------------------------------------" '& Chr(13)
    
    Y = Y + 5
    
    Printer.FontSize = 6
    TextoTicket = TextoTicket & Chr(13) & "Este vale caduca en 6 meses."
    
    TextoTicket = TextoTicket & Chr(13) & "Último día de validez: " & Date + 183
    
    TextoTicket = TextoTicket & Chr(13) & "No se aceptan devoluciones pasados 7 días."
    
    TextoTicket = TextoTicket & Chr(13) & "No se aceptan devoluciones sin etiquetas."
    
    TextoTicket = TextoTicket & Chr(13) & "Cualquier devolución se hará con vale."
    
    Y = Y + 2
    
    TextoTicket = TextoTicket & Chr(13) & "No se devuelven artículos de fiesta."
    
    Y = Y + 2

    
    TextoTicket = TextoTicket & Chr(13) & "Conserve este vale para su canje."
    
    MsgBox TextoTicket
    
    ImprimeTicket TextoTicket
    Printer.EndDoc
     Header.ACuenta = 0
     Header.fecha = Now
     Header.FormaPago = ""
     Header.IdCliente = 0
     Header.IdVenta = 0
     Header.Modo = ""
     Header.Nombre = ""
     Header.SumaTotal = 0
     Header.Tarjeta = 0
    Erase Ticket
End Sub
