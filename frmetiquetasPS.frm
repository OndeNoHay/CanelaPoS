VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form FrmEtiquetas 
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   ScaleHeight     =   139.965
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   231.775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox DBGrid1PB 
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   13035
      TabIndex        =   26
      Top             =   1440
      Width           =   13095
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmetiquetasPS.frx":0000
         Height          =   5535
         Left            =   -120
         OleObjectBlob   =   "frmetiquetasPS.frx":0013
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
         Caption         =   "Imprime con logo"
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
      Begin VB.ComboBox Cmbalto 
         Height          =   315
         ItemData        =   "frmetiquetasPS.frx":09E6
         Left            =   840
         List            =   "frmetiquetasPS.frx":09E8
         TabIndex        =   11
         Text            =   "35"
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox Cmbancho 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Text            =   "70"
         Top             =   360
         Width           =   735
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
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4800
         TabIndex        =   24
         Text            =   "15"
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
      Picture         =   "frmetiquetasPS.frx":09EA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "FrmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    AnchoEtiq = Cmbancho
    AltoEtiq = Cmbalto

    Numetiqver = Int(297 / AltoEtiq)
    Numetiqhor = Int(210 / AnchoEtiq)
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
        ' Imprimir logo de Canela (esquina superior izquierda)
        Printer.PaintPicture Image1.Picture, x, Y, 10, 10

        ' Imprimir código de barras (parte superior derecha)
        ' IMPORTANTE: El scanner puede necesitar mayor tamaño y zona de silencio
        '
        ' Opciones de fuentes (en orden de preferencia):
        ' 1. IDAutomationHC39M (Code 39) - Más compatible con scanners
        ' 2. Libre Barcode 128 Text (Code 128) - Requiere tamaño mayor
        '
        ' Ver instrucciones: INSTALAR_FUENTE_EAN13.md

        Dim fuenteDisponible As String
        Dim usarCode39 As Boolean
        On Error Resume Next

        ' PRIMERO: Intentar Code 39 (más compatible con scanners)
        Printer.FontName = "IDAutomationHC39M"
        fuenteDisponible = Printer.FontName
        usarCode39 = (fuenteDisponible = "IDAutomationHC39M")

        If Not usarCode39 Then
            ' Si no está Code 39, intentar Code128
            Printer.FontName = "Libre Barcode 128 Text"
            fuenteDisponible = Printer.FontName

            ' Si tampoco está Code128, intentar EAN13
            If fuenteDisponible <> "Libre Barcode 128 Text" Then
                Printer.FontName = "Libre Barcode EAN13 Text"
                fuenteDisponible = Printer.FontName
            End If
        End If

        On Error GoTo sehodio

        ' Configurar fuente y tamaño según disponibilidad
        If usarCode39 Then
            ' Code 39 requiere asteriscos y tamaño mayor
            Printer.FontSize = 32
            Printer.CurrentX = x + 15
            Printer.CurrentY = Y
            Printer.Print "*" & etiquetasParaImprimir(indiceEtiqueta).EAN13 & "*"
        ElseIf fuenteDisponible = "Libre Barcode 128 Text" Or fuenteDisponible = "Libre Barcode EAN13 Text" Then
            ' Code 128 / EAN13 sin asteriscos pero tamaño MAYOR para scanner
            Printer.FontSize = 36
            Printer.CurrentX = x + 15
            Printer.CurrentY = Y
            Printer.Print etiquetasParaImprimir(indiceEtiqueta).EAN13
        Else
            ' Fallback a Arial (no escaneable)
            Printer.FontName = "Arial"
            Printer.FontSize = 8
            Printer.CurrentX = x + 18
            Printer.CurrentY = Y + 1
            Printer.Print etiquetasParaImprimir(indiceEtiqueta).EAN13
        End If

        ' Imprimir número legible debajo del código de barras
        Printer.FontName = "Arial"
        Printer.FontSize = 6
        Printer.CurrentX = x + 18
        Printer.CurrentY = Y + 9
        Printer.Print etiquetasParaImprimir(indiceEtiqueta).EAN13

        ' Imprimir nombre del producto (debajo del logo)
        Printer.FontName = "Arial"
        Printer.FontSize = 7
        Printer.CurrentX = x
        Printer.CurrentY = Y + 12
        Dim nombreTruncado As String
        nombreTruncado = Left(etiquetasParaImprimir(indiceEtiqueta).NombreProducto, 30)
        If etiquetasParaImprimir(indiceEtiqueta).Talla <> "" Then
            Printer.Print nombreTruncado & " - " & etiquetasParaImprimir(indiceEtiqueta).Talla
        Else
            Printer.Print nombreTruncado
        End If

        ' Imprimir precio con IVA (parte inferior)
        Printer.CurrentX = x
        Printer.CurrentY = Y + 17
        Printer.FontSize = 12
        Printer.FontBold = True
        ' Usar Chr(128) para euro en VB6, o usar formato sin símbolo
        Printer.Print "PVP: " & Format(etiquetasParaImprimir(indiceEtiqueta).PrecioConIVA, "0.00") & Chr(128)
        Printer.FontBold = False

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
MargenSuperior = Int(Text2.Text)

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

' Verificar si hay fuentes de código de barras instaladas
Dim fuenteTest As String
Dim tieneFuenteBarcode As Boolean
On Error Resume Next

' Intentar Code128 (recomendada)
Printer.FontName = "Libre Barcode 128 Text"
fuenteTest = Printer.FontName
tieneFuenteBarcode = (fuenteTest = "Libre Barcode 128 Text")

' Si no está Code128, intentar EAN13
If Not tieneFuenteBarcode Then
    Printer.FontName = "Libre Barcode EAN13 Text"
    fuenteTest = Printer.FontName
    tieneFuenteBarcode = (fuenteTest = "Libre Barcode EAN13 Text")
End If

On Error GoTo ErrorHandler

' Mostrar mensaje inicial
Dim mensaje As String
mensaje = "Formulario de etiquetas PrestaShop" & vbCrLf & vbCrLf & _
          "1. Introduzca el rango de IDs de productos" & vbCrLf & _
          "2. Haga clic en 'Buscar en PrestaShop'" & vbCrLf & _
          "3. Revise los productos encontrados" & vbCrLf & _
          "4. Haga clic en 'Imprime con logo'"

' Advertencia si no tiene fuente de código de barras
If Not tieneFuenteBarcode Then
    mensaje = mensaje & vbCrLf & vbCrLf & _
              "⚠️ ADVERTENCIA: Fuente de código de barras NO instalada" & vbCrLf & _
              "Los códigos NO serán escaneables." & vbCrLf & vbCrLf & _
              "RECOMENDADO: Libre Barcode 128 Text (GRATIS)" & vbCrLf & _
              "https://fonts.google.com/specimen/Libre+Barcode+128+Text" & vbCrLf & vbCrLf & _
              "ALTERNATIVA: Libre Barcode EAN13 Text" & vbCrLf & _
              "(requiere EAN13 con checksum válido)" & vbCrLf & vbCrLf & _
              "Ver instrucciones: INSTALAR_FUENTE_EAN13.md"
    MsgBox mensaje, vbExclamation, "Etiquetas PrestaShop"
Else
    MsgBox mensaje, vbInformation, "Etiquetas PrestaShop"
End If

Exit Sub

ErrorHandler:
    MsgBox "Error al inicializar formulario: " & Err.Description, vbCritical
End Sub

Private Sub Text1_Change()
Call DrawBarcode(Text1, Picture1)
End Sub

Private Sub Text2_Change()
    On Error GoTo sehodio
    MargenSuperior = Int(Text2.Text)
    Exit Sub
sehodio:
    MsgBox "El margen no se ha podido fijar. Actualmente est� en 3px"
    MargenSuperior = 3
End Sub

Private Sub Txtultimo_Change()
On Error Resume Next
    Dim canti As Integer
    canti = Val(TxtUltimo.Text) - Val(TxtPrimero.Text) + 1
    If canti <= 0 Then lbnumero = "": Exit Sub
    lbnumero = "Rango: " & canti & " IDs"
End Sub

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
