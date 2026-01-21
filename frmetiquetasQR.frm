VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{8856F961-340A-11D0-A96B-00C04FD705A2}#1.0#0"; "mswebbrw.dll"
Begin VB.Form FrmEtiquetasQR
   Caption         =   "Etiquetas PrestaShop con QR"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   ScaleHeight     =   139.965
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   231.775
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1
      Height          =   15
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   15
      ExtentX         =   26
      ExtentY         =   26
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox DBGrid1PB 
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   13035
      TabIndex        =   26
      Top             =   1440
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmetiquetasQR.frx":0000
         Height          =   5535
         Left            =   -120
         OleObjectBlob   =   "frmetiquetasQR.frx":0013
         TabIndex        =   27
         Top             =   0
         Width           =   13215
      End
   End
   Begin VB.CommandButton Command3
      Caption         =   "Buscar en PrestaShop"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   7080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Text            =   "12345678901"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   3000
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   21
      Top             =   7080
      Width           =   855
   End
   Begin VB.Data Data 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "rsarticulo"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Intervalo Impresi�n"
      Height          =   1335
      Left            =   9600
      TabIndex        =   13
      Top             =   0
      Width           =   3495
      Begin VB.TextBox Txtultimo 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   330
         Width           =   735
      End
      Begin VB.TextBox Txtprimero 
         Height          =   285
         Left            =   840
         TabIndex        =   17
         Top             =   330
         Width           =   735
      End
      Begin VB.CommandButton Command1
         Caption         =   "Imprime con QR"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lbnumero 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   720
         Width           =   105
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame2
      Caption         =   "Tama�o Etiqueta"
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1935
      Begin VB.TextBox TxtMargenInteriorV
         Height          =   285
         Left            =   1320
         TabIndex        =   29
         Text            =   "1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox TxtMargenInteriorH
         Height          =   285
         Left            =   600
         TabIndex        =   28
         Text            =   "1"
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox Cmbalto
         Height          =   315
         ItemData        =   "frmetiquetasQR.frx":09E6
         Left            =   840
         List            =   "frmetiquetasQR.frx":09E8
         TabIndex        =   11
         Text            =   "29.7"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox Cmbancho
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Text            =   "52.5"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7
         Caption         =   "Márgenes int. H/V:"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4
         Caption         =   "Alto"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1
         Caption         =   "Ancho"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Condiciones Impresi�n"
      Height          =   1335
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      Begin VB.TextBox TxtMargenSuperior
         Height          =   285
         Left            =   4800
         TabIndex        =   24
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chknum 
         Caption         =   "N�mero de Etiquetas:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtNumEtiq 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Text            =   "27"
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Empezar a imprimir en:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox Cmbfila 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4560
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbcolumna 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Margen Superior"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fila"
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Columna"
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "imprime con c�digo"
      Height          =   615
      Left            =   10800
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   8880
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   120
      Picture         =   "frmetiquetasQR.frx":09EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "FrmEtiquetasQR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Declaración de API de Windows para Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Tipo para cada etiqueta a imprimir (expande combinaciones)
Private Type EtiquetaImpresion
    idProducto As Long
    EAN13 As String
    NombreProducto As String
    Talla As String
    PrecioConIVA As Currency
    idCombinacion As Long
End Type

Dim Numetiqhor As Integer
Dim Numetiqver As Integer
Dim AltoEtiq, AnchoEtiq As Integer
Dim RsArtImpr As DAO.Recordset
Dim PasaPrimerNum As Boolean
Dim MargenSuperior As Integer
Dim MargenInteriorH As Integer  ' Margen horizontal interior de la etiqueta (mm)
Dim MargenInteriorV As Integer  ' Margen vertical interior de la etiqueta (mm)

' Variables para productos de PrestaShop
Dim productosPS() As ProductoPrestaShop
Dim numProductosPS As Integer
Dim etiquetasParaImprimir() As EtiquetaImpresion
Dim numEtiquetas As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
    cmbcolumna.Enabled = True
    Cmbfila.Enabled = True
Else
    cmbcolumna.Enabled = False
    Cmbfila.Enabled = False
End If
End Sub

Private Sub chknum_Click()
If chknum.Value = 1 Then
    TxtNumEtiq.Enabled = True
Else
    TxtNumEtiq.Enabled = False
End If

End Sub

Private Sub Command1_Click()
ImprimeEtiquetas
End Sub
Private Sub ImprimeEtiquetas()
    Dim Contahoriz, Contaverti As Integer
    Dim NumImpresa As Integer
    Dim indiceEtiqueta As Integer

    On Error GoTo sehodio

    ' Verificar que hay etiquetas para imprimir
    If numEtiquetas = 0 Then
        MsgBox "No hay etiquetas para imprimir. Primero debe buscar productos en PrestaShop.", vbExclamation
        Exit Sub
    End If

    ' Configurar dimensiones de etiquetas
    AnchoEtiq = Val(Cmbancho.Text)
    AltoEtiq = Val(Cmbalto.Text)
    MargenInteriorH = Val(TxtMargenInteriorH.Text)
    MargenInteriorV = Val(TxtMargenInteriorV.Text)

    Numetiqver = Int(297 / AltoEtiq)
    Numetiqhor = Int(210 / AnchoEtiq)  ' Usar el ancho exacto configurado
    Contahoriz = 0
    Contaverti = 0

    Dim x, Y As Integer
    x = 2
    Y = MargenSuperior

    ' Configurar posición inicial si se especifica
    If Check1.Value = 1 Then
        Contahoriz = Val(cmbcolumna.Text) - 1
        Contaverti = Val(Cmbfila.Text) - 1
        For i = 1 To Contahoriz
            x = x + AnchoEtiq
        Next i
        For i = 1 To Contaverti
            Y = Y + AltoEtiq
        Next i
    End If

    ' Configurar impresora A4
    Printer.ScaleMode = 6  ' Milímetros
    NumImpresa = 0
    indiceEtiqueta = 1

    ' Imprimir todas las etiquetas
    Do While indiceEtiqueta <= numEtiquetas
        ' Calcular área útil dentro de la etiqueta (con márgenes interiores)
        Dim xInicio As Integer, yInicio As Integer
        Dim anchoUtil As Integer, altoUtil As Integer

        xInicio = x + MargenInteriorH
        yInicio = Y + MargenInteriorV
        anchoUtil = AnchoEtiq - (MargenInteriorH * 2)
        altoUtil = AltoEtiq - (MargenInteriorV * 2)

        ' Generar e imprimir código QR
        Dim ean13 As String
        ean13 = Trim(etiquetasParaImprimir(indiceEtiqueta).EAN13)

        ' Tamaño del QR (cuadrado) basado en el alto útil
        Dim qrSize As Integer
        qrSize = Int(altoUtil * 0.7)  ' 70% del alto útil
        If qrSize < 10 Then qrSize = 10   ' Mínimo 10mm
        If qrSize > 25 Then qrSize = 25   ' Máximo 25mm

        ' Generar QR code y obtener imagen
        Dim qrPicture As Object
        Set qrPicture = GenerarQRCode(ean13, qrSize)

        If Not qrPicture Is Nothing Then
            ' Imprimir QR code en esquina superior izquierda
            Printer.PaintPicture qrPicture, xInicio, yInicio, qrSize, qrSize
        Else
            ' Si falla el QR, imprimir código en texto plano como fallback
            Printer.FontName = "Arial"
            Printer.FontSize = 8
            Printer.CurrentX = xInicio
            Printer.CurrentY = yInicio
            Printer.Print "EAN: " & ean13
        End If

        ' Imprimir precio a la derecha del QR, en la parte superior
        Printer.FontName = "Arial"
        Printer.FontSize = 14
        Printer.FontBold = True
        Printer.CurrentX = xInicio + qrSize + 2  ' 2mm de separación del QR
        Printer.CurrentY = yInicio
        Printer.Print "PVP: " & Format(etiquetasParaImprimir(indiceEtiqueta).PrecioConIVA, "0.00") & Chr(128)

        ' Imprimir nombre del producto a la derecha del QR, debajo del precio
        Printer.FontBold = False
        Printer.FontSize = 8
        Printer.CurrentX = xInicio + qrSize + 2  ' Alineado con el precio
        Printer.CurrentY = yInicio + 6  ' Debajo del precio

        Dim nombreTruncado As String
        Dim maxChars As Integer
        Dim espacioDisponible As Integer
        espacioDisponible = anchoUtil - qrSize - 2  ' Espacio a la derecha del QR
        maxChars = Int(espacioDisponible / 2)  ' Aproximadamente 2mm por carácter
        If maxChars < 10 Then maxChars = 10  ' Mínimo 10 caracteres
        nombreTruncado = Left(etiquetasParaImprimir(indiceEtiqueta).NombreProducto, maxChars)

        If etiquetasParaImprimir(indiceEtiqueta).Talla <> "" Then
            Printer.Print nombreTruncado & " - " & etiquetasParaImprimir(indiceEtiqueta).Talla
        Else
            Printer.Print nombreTruncado
        End If

        ' Avanzar a siguiente posición
        Contahoriz = Contahoriz + 1
        x = x + AnchoEtiq
        NumImpresa = NumImpresa + 1
        indiceEtiqueta = indiceEtiqueta + 1

        ' Si llegamos al final de la fila
        If Contahoriz >= Numetiqhor Then
            Contahoriz = 0
            Contaverti = Contaverti + 1
            x = 2

            ' Si llegamos al final de la página
            If Contaverti >= Numetiqver Then
                Printer.NewPage
                Contaverti = 0
                Y = MargenSuperior
            Else
                Y = Y + AltoEtiq
            End If
        End If
    Loop

    Printer.EndDoc
    MsgBox "Impresión completada: " & NumImpresa & " etiquetas", vbInformation, "Etiquetas PrestaShop"
    Exit Sub

sehodio:
    MsgBox "Error de impresión: " & Err.Number & Chr(13) & Err.Description, vbCritical
End Sub
Private Sub ImprimeCodigo()
Dim Contahoriz, Contaverti As Integer
Contahoriz = 0
Contaverti = 0
Dim x, Y As Integer
x = 5
Y = 10

Printer.ScaleMode = 6
Do Until Contaverti = Numetiqver
    Do Until Contahoriz = Numetiqhor
        Printer.FontName = "IDAutomationHC39M"
        Printer.CurrentX = x
        Printer.CurrentY = Y
        Printer.Print "*1234567890*" & Contahoriz & Contaverti
        Printer.FontName = "Arial"
        Printer.CurrentX = x
        Printer.CurrentY = Y + 10
        Printer.Print "Etiqueta horizontal n� " & Contahoriz
        Printer.CurrentX = x
        Printer.CurrentY = Y + 14
        Printer.Print "Etiqueta vertical n�" & Contaverti
        Contahoriz = Contahoriz + 1
        x = x + (210 / Numetiqhor)
    Loop
    Contahoriz = 0
    Contaverti = Contaverti + 1
    If Y >= 260 Then
        Printer.NewPage
        Y = 10
    End If
    Y = Y + (290 / Numetiqver)
    x = 5
Loop
Printer.EndDoc

End Sub

Private Sub Command2_Click()
Numetiqhor = 2
Numetiqver = 14
ImprimeCodigo
End Sub

Private Sub Command3_Click()
    ' Buscar productos en PrestaShop por rango de IDs
    Dim idInicio As Long
    Dim idFin As Long
    Dim i As Integer
    Dim j As Integer
    Dim producto As ProductoPrestaShop

    On Error GoTo ErrorHandler

    ' Validar que se han introducido los IDs
    If Trim(TxtPrimero.Text) = "" Or Trim(TxtUltimo.Text) = "" Then
        MsgBox "Por favor, introduzca el rango de IDs de productos (Desde/Hasta)", vbExclamation
        Exit Sub
    End If

    idInicio = CLng(Val(TxtPrimero.Text))
    idFin = CLng(Val(TxtUltimo.Text))

    If idInicio < 1 Or idFin < idInicio Then
        MsgBox "Rango de IDs inválido. El ID final debe ser mayor o igual que el inicial.", vbExclamation
        Exit Sub
    End If

    ' Mostrar mensaje de espera
    Me.MousePointer = vbHourglass
    lbnumero.Caption = "Buscando en PrestaShop..."
    DoEvents

    ' Buscar productos en PrestaShop
    numProductosPS = BuscarProductosPorRangoID(idInicio, idFin, productosPS)

    If numProductosPS = 0 Then
        Me.MousePointer = vbDefault
        lbnumero.Caption = "No se encontraron productos"
        MsgBox "No se encontraron productos activos con stock en el rango especificado.", vbInformation
        Exit Sub
    End If

    ' Expandir productos con combinaciones en etiquetas individuales
    ReDim etiquetasParaImprimir(1 To 1000) ' Máximo 1000 etiquetas
    numEtiquetas = 0

    For i = 1 To numProductosPS
        producto = productosPS(i)

        If producto.TieneCombinaciones And producto.NumCombinaciones > 0 Then
            ' Crear una etiqueta por cada combinación (talla)
            For j = 1 To producto.NumCombinaciones
                If producto.Combinaciones(j).stock > 0 Then
                    numEtiquetas = numEtiquetas + 1
                    etiquetasParaImprimir(numEtiquetas).idProducto = producto.idProducto
                    etiquetasParaImprimir(numEtiquetas).EAN13 = producto.EAN
                    etiquetasParaImprimir(numEtiquetas).NombreProducto = producto.Nombre
                    etiquetasParaImprimir(numEtiquetas).Talla = producto.Combinaciones(j).Talla
                    etiquetasParaImprimir(numEtiquetas).PrecioConIVA = producto.PrecioConIVA
                    etiquetasParaImprimir(numEtiquetas).idCombinacion = producto.Combinaciones(j).idCombinacion
                End If
            Next j
        Else
            ' Producto estándar - una sola etiqueta
            numEtiquetas = numEtiquetas + 1
            etiquetasParaImprimir(numEtiquetas).idProducto = producto.idProducto
            etiquetasParaImprimir(numEtiquetas).EAN13 = producto.EAN
            etiquetasParaImprimir(numEtiquetas).NombreProducto = producto.Nombre
            etiquetasParaImprimir(numEtiquetas).Talla = ""
            etiquetasParaImprimir(numEtiquetas).PrecioConIVA = producto.PrecioConIVA
            etiquetasParaImprimir(numEtiquetas).idCombinacion = 0
        End If
    Next i

    ' Redimensionar array al tamaño real
    If numEtiquetas > 0 Then
        ReDim Preserve etiquetasParaImprimir(1 To numEtiquetas)
    End If

    ' Poblar el grid con los datos
    PoblarGridConProductos

    ' Actualizar contador
    lbnumero.Caption = "Productos: " & numProductosPS & " | Etiquetas: " & numEtiquetas

    Me.MousePointer = vbDefault
    MsgBox "Se encontraron " & numProductosPS & " productos." & vbCrLf & _
           "Total de etiquetas a imprimir: " & numEtiquetas, vbInformation, "Búsqueda completada"

    Exit Sub

ErrorHandler:
    Me.MousePointer = vbDefault
    MsgBox "Error al buscar productos: " & Err.Description, vbCritical
End Sub

Private Sub DBGrid1_Click()

    'PasaPrimerNum = Not PasaPrimerNum
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If PasaPrimerNum = True Then
    TxtPrimero = DBGrid1.Text
    PasaPrimerNum = False
Else
    TxtUltimo = DBGrid1.Text
    PasaPrimerNum = True
End If

End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler

AnchoEtiq = Cmbancho
AltoEtiq = Cmbalto
Numetiqhor = 3
Numetiqver = 12
'MargenSuperior = 3
MargenSuperior = Int(TxtMargenSuperior.Text)

For i = 1 To Numetiqhor
    cmbcolumna.AddItem i
Next i
For i = 1 To Numetiqver
    Cmbfila.AddItem i
Next i

' NO cargar datos de BD local - los productos se buscarán en PrestaShop
numProductosPS = 0
numEtiquetas = 0

' Crear recordset vacío temporal para el grid
Set RsArtImpr = CrearRecordsetVacio()

If Not RsArtImpr Is Nothing Then
    Set Data.Recordset = RsArtImpr
End If

' Cargar HTML con generador de QR en WebBrowser
Dim htmlPath As String
htmlPath = App.Path & "\qr_generator.html"
If Dir(htmlPath) <> "" Then
    WebBrowser1.Navigate htmlPath
Else
    MsgBox "Advertencia: No se encontró qr_generator.html. La generación de códigos QR no funcionará.", vbExclamation
End If

Exit Sub

ErrorHandler:
    MsgBox "Error al inicializar formulario: " & Err.Description, vbCritical
End Sub

Private Sub Text1_Change()
Call DrawBarcode(Text1, Picture1)
End Sub

Private Sub TxtMargenSuperior_Change()
    On Error GoTo sehodio
    MargenSuperior = Int(TxtMargenSuperior.Text)
    Exit Sub
sehodio:
    MsgBox "El margen no se ha podido fijar. Actualmente est� en 0mm"
    MargenSuperior = 0
End Sub

Private Sub Txtultimo_Change()
On Error Resume Next
    Dim canti As Integer
    canti = Val(TxtUltimo.Text) - Val(TxtPrimero.Text) + 1
    If canti <= 0 Then lbnumero = "": Exit Sub
    lbnumero = "Rango: " & canti & " IDs"
End Sub

' ========================================================================
' GENERACIÓN DE IMÁGENES DE CÓDIGOS DE BARRAS
' ========================================================================


Private Sub Form_Unload(Cancel As Integer)
    ' Limpiar tabla temporal al cerrar el formulario
    On Error Resume Next

    ' Desvincular recordset del control Data
    Set Data.Recordset = Nothing

    ' Cerrar recordset si está abierto
    If Not RsArtImpr Is Nothing Then
        RsArtImpr.Close
        Set RsArtImpr = Nothing
    End If

    ' Esperar un momento para asegurar liberación de recursos
    DoEvents

    ' Eliminar tabla temporal si existe
    Dim i As Integer
    For i = 0 To bdtienda.TableDefs.Count - 1
        If bdtienda.TableDefs(i).Name = "TempEtiquetasPS" Then
            bdtienda.TableDefs.Delete "TempEtiquetasPS"
            Exit For
        End If
    Next i
End Sub

'******************************************************************************
'* FUNCIÓN: CrearRecordsetVacio
'* PROPÓSITO: Crea un recordset vacío DAO para el DBGrid usando tabla temporal
'******************************************************************************
Private Function CrearRecordsetVacio() As DAO.Recordset
    On Error GoTo ErrorHandler

    Dim Rs As DAO.Recordset
    Dim tblDef As DAO.TableDef
    Dim fld As DAO.Field
    Dim tablaExiste As Boolean
    Dim i As Integer

    ' Verificar si la tabla temporal ya existe
    tablaExiste = False
    For i = 0 To bdtienda.TableDefs.Count - 1
        If bdtienda.TableDefs(i).Name = "TempEtiquetasPS" Then
            tablaExiste = True
            Exit For
        End If
    Next i

    If tablaExiste Then
        ' Si la tabla existe, simplemente limpiarla (más rápido y evita bloqueos)
        bdtienda.Execute "DELETE FROM TempEtiquetasPS", dbFailOnError
    Else
        ' Crear nueva tabla temporal solo si no existe
        Set tblDef = bdtienda.CreateTableDef("TempEtiquetasPS")

        ' Agregar campos
        Set fld = tblDef.CreateField("idProducto", dbLong)
        tblDef.Fields.Append fld

        Set fld = tblDef.CreateField("EAN13", dbText, 50)
        tblDef.Fields.Append fld

        Set fld = tblDef.CreateField("Nombre", dbText, 200)
        tblDef.Fields.Append fld

        Set fld = tblDef.CreateField("Talla", dbText, 50)
        tblDef.Fields.Append fld

        Set fld = tblDef.CreateField("Precio", dbCurrency)
        tblDef.Fields.Append fld

        ' Agregar la tabla a la BD
        bdtienda.TableDefs.Append tblDef
    End If

    ' Abrir recordset sobre la tabla temporal (ahora siempre vacía)
    Set Rs = bdtienda.OpenRecordset("TempEtiquetasPS", dbOpenTable)

    Set CrearRecordsetVacio = Rs
    Exit Function

ErrorHandler:
    MsgBox "Error al crear recordset: " & Err.Description, vbCritical
    Set CrearRecordsetVacio = Nothing
End Function

'******************************************************************************
'* FUNCIÓN: PoblarGridConProductos
'* PROPÓSITO: Llena el DBGrid con las etiquetas a imprimir
'******************************************************************************
Private Sub PoblarGridConProductos()
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim RsTemp As DAO.Recordset

    ' Cerrar y desvincular recordset anterior del control Data
    On Error Resume Next
    Set Data.Recordset = Nothing
    On Error GoTo ErrorHandler

    If Not RsArtImpr Is Nothing Then
        RsArtImpr.Close
        Set RsArtImpr = Nothing
    End If

    ' Crear nuevo recordset usando tabla temporal
    Set RsArtImpr = CrearRecordsetVacio()

    ' Verificar que se creó correctamente
    If RsArtImpr Is Nothing Then
        MsgBox "No se pudo crear el recordset para mostrar los productos", vbExclamation
        Exit Sub
    End If

    ' Agregar cada etiqueta al recordset
    For i = 1 To numEtiquetas
        With RsArtImpr
            .AddNew
            !idProducto = etiquetasParaImprimir(i).idProducto
            !EAN13 = etiquetasParaImprimir(i).EAN13
            !Nombre = etiquetasParaImprimir(i).NombreProducto
            !Talla = etiquetasParaImprimir(i).Talla
            !Precio = etiquetasParaImprimir(i).PrecioConIVA
            .Update
        End With
    Next i

    ' Vincular al grid
    If RsArtImpr.RecordCount > 0 Then
        RsArtImpr.MoveFirst
    End If
    Set Data.Recordset = RsArtImpr

    Exit Sub

ErrorHandler:
    MsgBox "Error al poblar grid: " & Err.Description & " (Err #" & Err.Number & ")", vbCritical
End Sub

'******************************************************************************
'* FUNCIÓN: GenerarQRCode
'* PROPÓSITO: Genera un código QR usando JavaScript y retorna la imagen
'* PARÁMETROS:
'*   - texto: El texto a codificar en el QR (ej: EAN13)
'*   - tamanoMM: Tamaño del QR en milímetros
'* RETORNA: Object (StdPicture) con la imagen del QR, o Nothing si falla
'******************************************************************************
Private Function GenerarQRCode(ByVal texto As String, ByVal tamanoMM As Integer) As Object
    On Error GoTo ErrorHandler

    Dim base64Data As String
    Dim tempFile As String
    Dim pic As Object
    Dim retryCount As Integer
    Dim maxRetries As Integer
    Dim i As Integer
    Dim debugMsg As String

    Set GenerarQRCode = Nothing
    maxRetries = 50
    retryCount = 0

    ' Debug: verificar estado del WebBrowser
    debugMsg = "WebBrowser ReadyState: " & WebBrowser1.ReadyState & vbCrLf

    ' Esperar a que el WebBrowser esté listo (ReadyState = 4 = READYSTATE_COMPLETE)
    Do While WebBrowser1.ReadyState <> 4 And retryCount < maxRetries
        Sleep 100  ' Esperar 100ms
        DoEvents
        retryCount = retryCount + 1
    Loop

    If WebBrowser1.ReadyState <> 4 Then
        debugMsg = debugMsg & "WebBrowser no está listo" & vbCrLf
        MsgBox debugMsg, vbExclamation, "Debug QR"
        Exit Function
    End If

    debugMsg = debugMsg & "WebBrowser listo" & vbCrLf

    ' Verificar que JavaScript está listo
    retryCount = 0
    Dim jsReady As Boolean
    jsReady = False

    Do While retryCount < maxRetries
        On Error Resume Next
        jsReady = False
        If Not WebBrowser1.Document Is Nothing Then
            If Not WebBrowser1.Document.parentWindow Is Nothing Then
                jsReady = WebBrowser1.Document.parentWindow.qrReady
            End If
        End If
        On Error GoTo ErrorHandler

        If jsReady = True Then Exit Do
        Sleep 100
        DoEvents
        retryCount = retryCount + 1
    Loop

    If Not jsReady Then
        debugMsg = debugMsg & "JavaScript no está listo" & vbCrLf
        MsgBox debugMsg, vbExclamation, "Debug QR"
        Exit Function
    End If

    debugMsg = debugMsg & "JavaScript listo" & vbCrLf

    ' Convertir mm a píxeles (aproximadamente 3.78 píxeles por mm para 96 DPI)
    Dim tamanoPixeles As Integer
    tamanoPixeles = Int(tamanoMM * 3.78)
    If tamanoPixeles < 50 Then tamanoPixeles = 50  ' Mínimo 50px

    debugMsg = debugMsg & "Tamaño: " & tamanoPixeles & "px" & vbCrLf
    debugMsg = debugMsg & "Texto: " & texto & vbCrLf

    ' Llamar a la función JavaScript para generar QR
    On Error Resume Next
    base64Data = ""
    base64Data = WebBrowser1.Document.parentWindow.GenerateQRCode(texto, tamanoPixeles)

    If Err.Number <> 0 Then
        debugMsg = debugMsg & "Error JS: " & Err.Description & vbCrLf
        On Error GoTo ErrorHandler
        MsgBox debugMsg, vbCritical, "Debug QR"
        Exit Function
    End If
    On Error GoTo ErrorHandler

    If base64Data = "" Or Left(base64Data, 5) = "ERROR" Then
        debugMsg = debugMsg & "Base64 vacío o error: " & Left(base64Data, 50) & vbCrLf
        MsgBox debugMsg, vbExclamation, "Debug QR"
        Exit Function
    End If

    debugMsg = debugMsg & "Base64 OK (len=" & Len(base64Data) & ")" & vbCrLf

    ' Crear archivo temporal para la imagen con timestamp único
    tempFile = Environ("TEMP") & "\qr_temp_" & Format(Now, "hhnnssms") & "_" & Int(Rnd * 10000) & ".bmp"

    ' Decodificar base64 y guardar como archivo
    If Not Base64ToFile(base64Data, tempFile) Then
        debugMsg = debugMsg & "Error al guardar archivo" & vbCrLf
        MsgBox debugMsg, vbCritical, "Debug QR"
        Exit Function
    End If

    debugMsg = debugMsg & "Archivo guardado: " & tempFile & vbCrLf

    ' Cargar imagen desde archivo
    On Error Resume Next
    Set pic = LoadPicture(tempFile)

    If Err.Number <> 0 Then
        debugMsg = debugMsg & "Error LoadPicture: " & Err.Description & vbCrLf
        On Error GoTo ErrorHandler
        MsgBox debugMsg, vbCritical, "Debug QR"
        Exit Function
    End If
    On Error GoTo ErrorHandler

    If pic Is Nothing Then
        debugMsg = debugMsg & "pic es Nothing después de LoadPicture" & vbCrLf
        MsgBox debugMsg, vbCritical, "Debug QR"
        Exit Function
    End If

    debugMsg = debugMsg & "LoadPicture OK" & vbCrLf

    ' Eliminar archivo temporal
    On Error Resume Next
    Kill tempFile
    On Error GoTo ErrorHandler

    Set GenerarQRCode = pic

    ' Debug: Mostrar solo en la primera etiqueta
    Static primeraVez As Boolean
    If Not primeraVez Then
        MsgBox debugMsg & "¡QR generado con éxito!", vbInformation, "Debug QR"
        primeraVez = True
    End If

    Exit Function

ErrorHandler:
    debugMsg = debugMsg & "Error: " & Err.Number & " - " & Err.Description & vbCrLf
    MsgBox debugMsg, vbCritical, "Error QR"
    Set GenerarQRCode = Nothing
End Function

'******************************************************************************
'* FUNCIÓN: Base64ToFile
'* PROPÓSITO: Convierte una cadena base64 a un archivo
'* PARÁMETROS:
'*   - base64String: Cadena en formato base64 (data:image/bmp;base64,...)
'*   - outputFile: Ruta del archivo de salida
'* RETORNA: Boolean - True si tuvo éxito
'******************************************************************************
Private Function Base64ToFile(ByVal base64String As String, ByVal outputFile As String) As Boolean
    On Error GoTo ErrorHandler

    Dim stream As Object
    Dim xmlDoc As Object
    Dim node As Object

    Base64ToFile = False

    ' Eliminar el prefijo "data:image/bmp;base64," si existe
    If InStr(base64String, "base64,") > 0 Then
        base64String = Mid(base64String, InStr(base64String, "base64,") + 7)
    End If

    ' Crear XML DOM para decodificar base64
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set node = xmlDoc.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64String

    ' Crear stream para guardar los bytes
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1  ' adTypeBinary
    stream.Open
    stream.Write node.nodeTypedValue
    stream.SaveToFile outputFile, 2  ' adSaveCreateOverWrite
    stream.Close

    Base64ToFile = True

    Set stream = Nothing
    Set node = Nothing
    Set xmlDoc = Nothing
    Exit Function

ErrorHandler:
    Base64ToFile = False
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Set stream = Nothing
    Set node = Nothing
    Set xmlDoc = Nothing
End Function
