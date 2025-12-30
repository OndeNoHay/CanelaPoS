Attribute VB_Name = "ModuloIntegracion"
'******************************************************************************
'* MÓDULO: ModuloIntegracion.bas
'* PROPÓSITO: Funciones de integración entre POS local y PrestaShop
'* AUTOR: Claude Code
'* FECHA: 2025-12-29
'******************************************************************************

Option Explicit

'--- Variables globales para tracking de productos PrestaShop ---
Type ArticuloPrestaShop
    idArtLocal As Long              ' ID del artículo en la BD local
    idProductoPS As Long            ' ID del producto en PrestaShop
    idCombinacionPS As Long         ' ID de combinación en PrestaShop (0 si no tiene)
    CodigoBuscado As String         ' Código usado para buscar
    SincronizarStock As Boolean     ' True si debe sincronizarse con PS
End Type

' Colección de artículos pendientes de sincronización
Private articulosVenta() As ArticuloPrestaShop
Private numArticulosVenta As Integer

'******************************************************************************
'* FUNCIÓN: InicializarIntegracion
'* PROPÓSITO: Inicializa el sistema de integración PrestaShop
'* Debe llamarse al iniciar la aplicación
'******************************************************************************
Public Sub InicializarIntegracion()
    On Error Resume Next

    ' Cargar configuración
    CargarConfiguracion

    ' Inicializar logging si está habilitado
    If logHabilitado Then
        InicializarLog
        EscribirLog "Sistema de integración PrestaShop iniciado"
    End If

    ' Limpiar logs antiguos (mantener 30 días)
    LimpiarLogsAntiguos 30

    ' Inicializar array de artículos
    ReDim articulosVenta(0)
    numArticulosVenta = 0
End Sub

'******************************************************************************
'* FUNCIÓN: BuscarProductoPrestaShop
'* PROPÓSITO: Busca un producto en PrestaShop y lo agrega a la BD local si se encuentra
'* PARÁMETROS:
'*   - codigo: Código del producto (referencia o EAN)
'* RETORNA: ID del artículo local (0 si no se encuentra o hay error)
'******************************************************************************
Public Function BuscarProductoPrestaShop(ByVal Codigo As String) As Long
    On Error GoTo ErrorHandler

    Dim producto As ProductoPrestaShop
    Dim idArtLocal As Long

    BuscarProductoPrestaShop = 0

    ' Verificar si la integración está habilitada
    If Not IntegracionHabilitada Or Not BuscarEnPrestaShop Then
        Exit Function
    End If

    ' Validar código
    If Trim(Codigo) = "" Then Exit Function

    LogBusquedaProducto Codigo, False

    ' Buscar en PrestaShop
    producto = ModuloPrestaShop.BuscarProductoPorCodigo(Codigo)

    ' Si no se encuentra, salir sin error
    If Not producto.encontrado Then
        LogBusquedaProducto Codigo, False
        Exit Function
    End If

    LogBusquedaProducto Codigo, True, producto.idProducto, producto.Nombre

    ' Producto encontrado - crear en BD local
    idArtLocal = CrearArticuloDesdePrestaShop(producto, Codigo)

    If idArtLocal > 0 Then
        ' Registrar para sincronización de stock posterior
        RegistrarArticuloParaSincronizacion idArtLocal, producto.idProducto, producto.idCombinacion, Codigo
        BuscarProductoPrestaShop = idArtLocal
    End If

    Exit Function

ErrorHandler:
    LogError "Error en BuscarProductoPrestaShop: " & Err.Description
    BuscarProductoPrestaShop = 0
End Function

'******************************************************************************
'* FUNCIÓN: CrearArticuloDesdePrestaShop
'* PROPÓSITO: Crea un artículo temporal en la BD local desde datos de PrestaShop
'******************************************************************************
Private Function CrearArticuloDesdePrestaShop(producto As ProductoPrestaShop, Codigo As String) As Long
    On Error GoTo ErrorHandler

    Dim Rs As Recordset
    Dim nuevoIdArt As Long

    CrearArticuloDesdePrestaShop = 0

    ' Abrir tabla de artículos
    Set Rs = bdtienda.OpenRecordset("articulos", dbOpenTable)

    With Rs
        .AddNew

        ' Generar un ID temporal único (negativo para identificar artículos de PrestaShop)
        ' Esto evita conflictos con IDs reales de la BD local
        nuevoIdArt = (producto.idProducto * 10000 + IIf(producto.idCombinacion > 0, producto.idCombinacion, 1))
        !Idart = nuevoIdArt

        ' Datos del producto
        !Codigo = Codigo
        !Tipo = Left(producto.Nombre, 50)  ' Nombre del producto (truncado a 50 chars)
        !Color = ""  ' Por defecto vacio
        !talla = ""  ' Por defecto vacio

        ' Precios
        !PrecioVenta = producto.PrecioConIVA
        !siniva = producto.PrecioSinIVA
        !PrecioCompra = producto.PrecioConIVA / 2.5 ' Asumir que precio sin IVA es precio de compra

        ' IVA
        If producto.PorcentajeIVA > 0 Then
            !iva = CInt(producto.PorcentajeIVA)
        Else
            !iva = 21  ' IVA por defecto España
        End If

        ' Stock
        !cantidad = producto.StockDisponible

        ' Estados
        !vendido = False
        !apartado = False
        !devuelto = False
        !inventario = False
        !deposito = False

        ' Descripción en campo extra (para referencia)
        If producto.TieneCombinaciones Then
            !extra = "PS_ID:" & producto.idProducto & "_" & producto.idCombinacion & " [COMBO]"
        Else
            !extra = "PS_ID:" & producto.idProducto & " [SIMPLE]"
        End If

        ' Fechas
        !fechacompra = Date

        .Update
    End With

    ' Cerrar recordset de insercion
    Rs.Close
    Set Rs = Nothing

    ' Verificar que el articulo se creo correctamente
    Set Rs = bdtienda.OpenRecordset("SELECT idart FROM articulos WHERE idart = " & nuevoIdArt)
    If Rs.EOF Then
        LogError "CRITICO: Articulo no encontrado despues de .Update. ID: " & nuevoIdArt
        CrearArticuloDesdePrestaShop = 0
    Else
        CrearArticuloDesdePrestaShop = nuevoIdArt
        LogInfo "Artículo creado desde PrestaShop - ID Local: " & nuevoIdArt & " | ID PS: " & producto.idProducto
        LogDebug "Verificacion BD exitosa - Articulo existe con idart=" & Rs!Idart
    End If

    Rs.Close
    Set Rs = Nothing

    Exit Function

ErrorHandler:
    LogError "Error en CrearArticuloDesdePrestaShop: " & Err.Description
    CrearArticuloDesdePrestaShop = 0
    If Not Rs Is Nothing Then Rs.Close
    Set Rs = Nothing
End Function

'******************************************************************************
'* FUNCIÓN: RegistrarArticuloParaSincronizacion
'* PROPÓSITO: Registra un artículo para sincronización de stock posterior
'******************************************************************************
Private Sub RegistrarArticuloParaSincronizacion(idArtLocal As Long, idProductoPS As Long, _
                                               idCombinacionPS As Long, Codigo As String)
    On Error Resume Next

    numArticulosVenta = numArticulosVenta + 1
    ReDim Preserve articulosVenta(numArticulosVenta)

    With articulosVenta(numArticulosVenta)
        .idArtLocal = idArtLocal
        .idProductoPS = idProductoPS
        .idCombinacionPS = idCombinacionPS
        .CodigoBuscado = Codigo
        .SincronizarStock = True
    End With

    LogDebug "Artículo registrado para sincronización: " & idArtLocal
End Sub

'******************************************************************************
'* FUNCIÓN: SincronizarStockVendido
'* PROPÓSITO: Sincroniza el stock de todos los artículos vendidos con PrestaShop
'* Debe llamarse después de completar una venta
'******************************************************************************
Public Sub SincronizarStockVendido()
    On Error Resume Next

    Dim i As Integer
    Dim resultado As ResultadoActualizacion
    Dim exitos As Integer
    Dim fallos As Integer

    ' Verificar si la integración y actualización automática están habilitadas
    If Not IntegracionHabilitada Or Not ActualizarStockAutomatico Then
        LimpiarRegistrosSincronizacion
        Exit Sub
    End If

    ' Verificar si hay artículos para sincronizar
    If numArticulosVenta = 0 Then Exit Sub

    LogInfo "Iniciando sincronización de stock - " & numArticulosVenta & " artículo(s)"

    exitos = 0
    fallos = 0

    ' Procesar cada artículo
    For i = 1 To numArticulosVenta
        If articulosVenta(i).SincronizarStock Then
            ' Actualizar stock en PrestaShop (decrementar en 1)
            resultado = ActualizarStock(articulosVenta(i).idProductoPS, 1, articulosVenta(i).idCombinacionPS)

            If resultado.exito Then
                exitos = exitos + 1
                LogSincronizacionStock articulosVenta(i).idProductoPS, resultado.stockAnterior, _
                                      resultado.stockNuevo, True

                ' Eliminar el artículo temporal de la BD local (IDs negativos)
                If articulosVenta(i).idArtLocal < 0 Then
                    EliminarArticuloTemporal articulosVenta(i).idArtLocal
                End If
            Else
                fallos = fallos + 1
                LogSincronizacionStock articulosVenta(i).idProductoPS, 0, 0, False, resultado.MensajeError
            End If
        End If
    Next i

    LogInfo "Sincronización completada - Éxitos: " & exitos & " | Fallos: " & fallos

    ' Limpiar registros
    LimpiarRegistrosSincronizacion
End Sub

'******************************************************************************
'* FUNCIÓN: EliminarArticuloTemporal
'* PROPÓSITO: Elimina un artículo temporal de la BD local (creado desde PrestaShop)
'******************************************************************************
Private Sub EliminarArticuloTemporal(Idart As Long)
    On Error Resume Next

    Dim Rs As Recordset

    ' Solo eliminar artículos con ID negativo (temporales de PrestaShop)
    If Idart >= 0 Then Exit Sub

    Set Rs = bdtienda.OpenRecordset("SELECT * FROM articulos WHERE idart = " & Idart)

    If Not Rs.EOF Then
        Rs.Delete
        LogDebug "Artículo temporal eliminado: " & Idart
    End If

    Rs.Close
    Set Rs = Nothing
End Sub

'******************************************************************************
'* FUNCIÓN: LimpiarRegistrosSincronizacion
'* PROPÓSITO: Limpia el array de artículos para sincronización
'******************************************************************************
Private Sub LimpiarRegistrosSincronizacion()
    ReDim articulosVenta(0)
    numArticulosVenta = 0
End Sub

'******************************************************************************
'* FUNCIÓN: CancelarVenta
'* PROPÓSITO: Cancela una venta y limpia los artículos temporales
'* Debe llamarse si se cancela una venta con artículos de PrestaShop
'******************************************************************************
Public Sub CancelarVenta()
    On Error Resume Next

    Dim i As Integer

    ' Eliminar artículos temporales
    For i = 1 To numArticulosVenta
        If articulosVenta(i).idArtLocal < 0 Then
            EliminarArticuloTemporal articulosVenta(i).idArtLocal
        End If
    Next i

    LogInfo "Venta cancelada - Artículos temporales eliminados"

    ' Limpiar registros
    LimpiarRegistrosSincronizacion
End Sub

'******************************************************************************
'* FUNCIÓN: VerificarStockPrestaShop
'* PROPÓSITO: Verifica el stock actual de un producto en PrestaShop
'* PARÁMETROS:
'*   - codigo: Código del producto
'* RETORNA: Stock disponible (-1 si hay error o no se encuentra)
'******************************************************************************
Public Function VerificarStockPrestaShop(ByVal Codigo As String) As Long
    On Error GoTo ErrorHandler

    Dim producto As ProductoPrestaShop

    VerificarStockPrestaShop = -1

    If Not IntegracionHabilitada Then Exit Function

    ' Buscar producto
    producto = BuscarProductoPorCodigo(Codigo)

    If producto.encontrado Then
        VerificarStockPrestaShop = producto.StockDisponible
        LogDebug "Stock verificado para " & Codigo & ": " & producto.StockDisponible
    Else
        LogDebug "Producto no encontrado para verificar stock: " & Codigo
    End If

    Exit Function

ErrorHandler:
    LogError "Error en VerificarStockPrestaShop: " & Err.Description
    VerificarStockPrestaShop = -1
End Function

'******************************************************************************
'* FUNCIÓN: ObtenerEstadoIntegracion
'* PROPÓSITO: Devuelve el estado actual de la integración (para mostrar en UI)
'******************************************************************************
Public Function ObtenerEstadoIntegracion() As String
    Dim estado As String

    If IntegracionHabilitada Then
        estado = "PrestaShop: ACTIVO"
        If numArticulosVenta > 0 Then
            estado = estado & " (" & numArticulosVenta & " art. pendientes sync)"
        End If
    Else
        estado = "PrestaShop: DESACTIVADO"
    End If

    ObtenerEstadoIntegracion = estado
End Function

'******************************************************************************
'* FUNCIÓN: FinalizarIntegracion
'* PROPÓSITO: Finaliza el sistema de integración
'* Debe llamarse al cerrar la aplicación
'******************************************************************************
Public Sub FinalizarIntegracion()
    On Error Resume Next

    ' Cerrar log
    If logHabilitado Then
        EscribirLog "Sistema de integración PrestaShop finalizado"
        CerrarLog
    End If

    ' Limpiar registros
    LimpiarRegistrosSincronizacion
End Sub
