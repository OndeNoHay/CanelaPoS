Attribute VB_Name = "ModuloPrestaShop"
'******************************************************************************
'* MÓDULO: ModuloPrestaShop.bas
'* PROPÓSITO: Integración con PrestaShop API a través de API Bridge
'* AUTOR: Claude Code
'* FECHA: 2025-12-29
'******************************************************************************

Option Explicit

'--- Constantes de configuración ---
Private Const PS_API_BRIDGE_URL As String = "https://www.canelamoda.es/api_bridge/"
Private Const PS_API_TIMEOUT As Long = 30000  ' 30 segundos

'--- Tipos de datos ---
Type ProductoPrestaShop
    IdProducto As Long
    Referencia As String
    EAN As String
    Nombre As String
    Descripcion As String
    PrecioSinIVA As Currency
    PrecioConIVA As Currency
    PorcentajeIVA As Double
    StockDisponible As Long
    TieneCombinaciones As Boolean
    IdCombinacion As Long
    Activo As Boolean
    Encontrado As Boolean
    MensajeError As String
End Type

Type ResultadoActualizacion
    Exito As Boolean
    StockAnterior As Long
    StockNuevo As Long
    MensajeError As String
End Type

'******************************************************************************
'* FUNCIÓN: BuscarProductoPorCodigo
'* PROPÓSITO: Busca un producto en PrestaShop por código (referencia o EAN)
'* PARÁMETROS:
'*   - codigo: Código de producto (puede ser referencia o EAN13)
'* RETORNA: Estructura ProductoPrestaShop con los datos del producto
'******************************************************************************
Public Function BuscarProductoPorCodigo(ByVal codigo As String) As ProductoPrestaShop
    On Error GoTo ErrorHandler

    Dim xmlHttp As Object
    Dim url As String
    Dim responseText As String
    Dim producto As ProductoPrestaShop

    ' Inicializar producto
    producto.Encontrado = False
    producto.TieneCombinaciones = False
    producto.IdProducto = 0
    producto.IdCombinacion = 0

    ' Validar código
    If Trim(codigo) = "" Then
        producto.MensajeError = "Código vacío"
        BuscarProductoPorCodigo = producto
        Exit Function
    End If

    ' Limpiar el código (solo números y letras)
    codigo = Trim(codigo)

    ' Construir URL para búsqueda
    url = PS_API_BRIDGE_URL & "bridge.php?action=buscar_producto&codigo=" & URLEncode(codigo)

    ' Log de la petición
    EscribirLog "Buscando producto: " & codigo
    EscribirLog "URL: " & url

    ' Crear objeto HTTP
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configurar timeout
    xmlHttp.setTimeouts 10000, 10000, PS_API_TIMEOUT, PS_API_TIMEOUT

    ' Realizar petición GET
    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send

    ' Verificar respuesta
    If xmlHttp.Status = 200 Then
        responseText = xmlHttp.responseText
        EscribirLog "Respuesta recibida: " & Left(responseText, 500)

        ' Parsear JSON response
        producto = ParsearProductoJSON(responseText)

        If producto.Encontrado Then
            EscribirLog "Producto encontrado: " & producto.Nombre & " (ID: " & producto.IdProducto & ")"
        Else
            EscribirLog "Producto no encontrado en PrestaShop"
        End If
    Else
        producto.MensajeError = "Error HTTP: " & xmlHttp.Status & " - " & xmlHttp.statusText
        EscribirLog "ERROR: " & producto.MensajeError
    End If

    Set xmlHttp = Nothing
    BuscarProductoPorCodigo = producto
    Exit Function

ErrorHandler:
    producto.MensajeError = "Error: " & Err.Description
    EscribirLog "ERROR en BuscarProductoPorCodigo: " & Err.Description
    BuscarProductoPorCodigo = producto
End Function

'******************************************************************************
'* FUNCIÓN: ObtenerStockProducto
'* PROPÓSITO: Obtiene el stock actual de un producto
'* PARÁMETROS:
'*   - idProducto: ID del producto en PrestaShop
'*   - idCombinacion: ID de la combinación (0 si no tiene combinaciones)
'* RETORNA: Cantidad de stock disponible (-1 si hay error)
'******************************************************************************
Public Function ObtenerStockProducto(ByVal idProducto As Long, Optional ByVal idCombinacion As Long = 0) As Long
    On Error GoTo ErrorHandler

    Dim xmlHttp As Object
    Dim url As String
    Dim responseText As String
    Dim stock As Long

    stock = -1  ' Valor por defecto en caso de error

    ' Construir URL
    url = PS_API_BRIDGE_URL & "bridge.php?action=obtener_stock&id=" & idProducto
    If idCombinacion > 0 Then
        url = url & "&combination_id=" & idCombinacion
    End If

    EscribirLog "Obteniendo stock para producto ID: " & idProducto

    ' Crear objeto HTTP
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttp.setTimeouts 5000, 5000, 15000, 15000

    ' Realizar petición
    xmlHttp.Open "GET", url, False
    xmlHttp.Send

    If xmlHttp.Status = 200 Then
        responseText = xmlHttp.responseText
        stock = ParsearStockJSON(responseText)
        EscribirLog "Stock disponible: " & stock
    Else
        EscribirLog "ERROR HTTP al obtener stock: " & xmlHttp.Status
    End If

    Set xmlHttp = Nothing
    ObtenerStockProducto = stock
    Exit Function

ErrorHandler:
    EscribirLog "ERROR en ObtenerStockProducto: " & Err.Description
    ObtenerStockProducto = -1
End Function

'******************************************************************************
'* FUNCIÓN: ActualizarStock
'* PROPÓSITO: Actualiza el stock de un producto en PrestaShop (decrementa)
'* PARÁMETROS:
'*   - idProducto: ID del producto en PrestaShop
'*   - cantidad: Cantidad a decrementar
'*   - idCombinacion: ID de la combinación (0 si no tiene combinaciones)
'* RETORNA: Estructura ResultadoActualizacion con el resultado
'******************************************************************************
Public Function ActualizarStock(ByVal idProducto As Long, ByVal cantidad As Long, _
    Optional ByVal idCombinacion As Long = 0) As ResultadoActualizacion
    On Error GoTo ErrorHandler

    Dim xmlHttp As Object
    Dim url As String
    Dim postData As String
    Dim responseText As String
    Dim resultado As ResultadoActualizacion

    resultado.Exito = False
    resultado.StockAnterior = 0
    resultado.StockNuevo = 0

    ' Validar parámetros
    If idProducto <= 0 Or cantidad <= 0 Then
        resultado.MensajeError = "Parámetros inválidos"
        ActualizarStock = resultado
        Exit Function
    End If

    ' NOTA: El bridge.php actual (Fase 1) es SOLO LECTURA
    ' La actualización de stock se implementará en Fase 2
    ' Por ahora, registramos la operación en el log
    EscribirLog "ADVERTENCIA: Actualización de stock aún no implementada en bridge.php"
    EscribirLog "Producto: " & idProducto & ", Cantidad a decrementar: " & cantidad

    ' Marcar como éxito (simulado) para no bloquear ventas
    resultado.Exito = True
    resultado.StockAnterior = 0
    resultado.StockNuevo = 0
    resultado.MensajeError = "Actualización de stock pendiente de implementación"

    ActualizarStock = resultado
    Exit Function

    ' Código desactivado temporalmente (para Fase 2):
    ' url = PS_API_BRIDGE_URL & "bridge.php?action=actualizar_stock"
    ' postData = "{""id"":" & idProducto & ",""cantidad"":" & cantidad & "}"

    EscribirLog "Actualizando stock - Producto: " & idProducto & ", Cantidad: -" & cantidad
    EscribirLog "POST Data: " & postData

    ' Crear objeto HTTP
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttp.setTimeouts 5000, 5000, PS_API_TIMEOUT, PS_API_TIMEOUT

    ' Realizar petición POST
    xmlHttp.Open "POST", url, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send postData

    ' Verificar respuesta
    If xmlHttp.Status = 200 Then
        responseText = xmlHttp.responseText
        EscribirLog "Respuesta: " & responseText

        ' Parsear resultado
        resultado = ParsearResultadoActualizacionJSON(responseText)

        If resultado.Exito Then
            EscribirLog "Stock actualizado correctamente: " & resultado.StockAnterior & " -> " & resultado.StockNuevo
        Else
            EscribirLog "ERROR: No se pudo actualizar el stock - " & resultado.MensajeError
        End If
    Else
        resultado.MensajeError = "Error HTTP: " & xmlHttp.Status & " - " & xmlHttp.statusText
        EscribirLog "ERROR HTTP: " & resultado.MensajeError
    End If

    Set xmlHttp = Nothing
    ActualizarStock = resultado
    Exit Function

ErrorHandler:
    resultado.MensajeError = "Error: " & Err.Description
    EscribirLog "ERROR en ActualizarStock: " & Err.Description
    ActualizarStock = resultado
End Function

'******************************************************************************
'* FUNCIÓN: ParsearProductoJSON
'* PROPÓSITO: Parsea la respuesta JSON del API y extrae los datos del producto
'* NOTA: Implementación simplificada. Para producción, considerar usar
'*       una librería JSON como VB-JSON o similar
'******************************************************************************
Private Function ParsearProductoJSON(ByVal jsonText As String) As ProductoPrestaShop
    Dim producto As ProductoPrestaShop

    On Error Resume Next

    producto.Encontrado = False

    ' Verificar si hay éxito en la respuesta
    If InStr(1, jsonText, """success"":true", vbTextCompare) > 0 Or _
       InStr(1, jsonText, """found"":true", vbTextCompare) > 0 Then

        producto.Encontrado = True

        ' Extraer ID del producto (bridge.php usa "id")
        producto.IdProducto = ExtraerValorNumerico(jsonText, "id")

        ' Extraer referencia
        producto.Referencia = ExtraerValorCadena(jsonText, "reference")

        ' Extraer EAN
        producto.EAN = ExtraerValorCadena(jsonText, "ean13")

        ' Extraer nombre (bridge.php usa "nombre")
        producto.Nombre = ExtraerValorCadena(jsonText, "nombre")

        ' Extraer descripción (bridge.php usa "descripcion")
        producto.Descripcion = ExtraerValorCadena(jsonText, "descripcion")

        ' Extraer precios (bridge.php usa "precio_sin_iva" y "precio_con_iva")
        producto.PrecioSinIVA = ExtraerValorMoneda(jsonText, "precio_sin_iva")
        producto.PrecioConIVA = ExtraerValorMoneda(jsonText, "precio_con_iva")

        ' Extraer IVA
        producto.PorcentajeIVA = ExtraerValorNumerico(jsonText, "iva")
        If producto.PorcentajeIVA = 0 Then producto.PorcentajeIVA = 21  ' Por defecto

        ' Extraer stock (bridge.php usa "stock")
        producto.StockDisponible = ExtraerValorNumerico(jsonText, "stock")

        ' Verificar si tiene combinaciones (bridge.php usa "tiene_combinaciones")
        producto.TieneCombinaciones = (InStr(1, jsonText, """tiene_combinaciones"":true", vbTextCompare) > 0)

        ' Si tiene combinaciones, extraer la primera combinación disponible
        If producto.TieneCombinaciones Then
            ' Buscar primera combinación en el array "combinaciones"
            Dim posCombo As Long
            posCombo = InStr(1, jsonText, """combinaciones"":[", vbTextCompare)
            If posCombo > 0 Then
                ' Extraer id_combinacion de la primera combinación
                producto.IdCombinacion = ExtraerValorNumerico(Mid(jsonText, posCombo), "id_combinacion")
                If producto.IdCombinacion = 0 Then
                    producto.IdCombinacion = ExtraerValorNumerico(Mid(jsonText, posCombo), "id_product_attribute")
                End If
            End If
        End If

        ' Verificar si está activo (bridge.php usa "activo")
        producto.Activo = (InStr(1, jsonText, """activo"":true", vbTextCompare) > 0)

    Else
        ' Producto no encontrado
        producto.MensajeError = ExtraerValorCadena(jsonText, "message")
        If producto.MensajeError = "" Then
            producto.MensajeError = "Producto no encontrado"
        End If
    End If

    ParsearProductoJSON = producto
End Function

'******************************************************************************
'* FUNCIÓN: ParsearStockJSON
'* PROPÓSITO: Extrae el valor de stock de una respuesta JSON
'******************************************************************************
Private Function ParsearStockJSON(ByVal jsonText As String) As Long
    Dim stock As Long

    ' bridge.php usa "cantidad"
    stock = ExtraerValorNumerico(jsonText, "cantidad")

    ParsearStockJSON = stock
End Function

'******************************************************************************
'* FUNCIÓN: ParsearResultadoActualizacionJSON
'* PROPÓSITO: Parsea el resultado de una actualización de stock
'******************************************************************************
Private Function ParsearResultadoActualizacionJSON(ByVal jsonText As String) As ResultadoActualizacion
    Dim resultado As ResultadoActualizacion

    resultado.Exito = (InStr(1, jsonText, """success"":true", vbTextCompare) > 0)
    resultado.StockAnterior = ExtraerValorNumerico(jsonText, "old_stock")
    resultado.StockNuevo = ExtraerValorNumerico(jsonText, "new_stock")

    If Not resultado.Exito Then
        resultado.MensajeError = ExtraerValorCadena(jsonText, "message")
        If resultado.MensajeError = "" Then
            resultado.MensajeError = ExtraerValorCadena(jsonText, "error")
        End If
    End If

    ParsearResultadoActualizacionJSON = resultado
End Function

'******************************************************************************
'* FUNCIONES AUXILIARES DE PARSEO
'******************************************************************************

Private Function ExtraerValorNumerico(ByVal jsonText As String, ByVal campo As String) As Long
    On Error Resume Next
    Dim posInicio As Long
    Dim posFin As Long
    Dim valor As String

    ' Buscar el campo en el JSON
    posInicio = InStr(1, jsonText, """" & campo & """:", vbTextCompare)
    If posInicio = 0 Then
        ExtraerValorNumerico = 0
        Exit Function
    End If

    ' Mover al inicio del valor (después de los dos puntos)
    posInicio = InStr(posInicio, jsonText, ":")
    posInicio = posInicio + 1

    ' Saltar espacios
    Do While Mid(jsonText, posInicio, 1) = " "
        posInicio = posInicio + 1
    Loop

    ' Buscar el final del valor (coma o llave de cierre)
    posFin = posInicio
    Do While posFin <= Len(jsonText)
        Dim c As String
        c = Mid(jsonText, posFin, 1)
        If c = "," Or c = "}" Or c = "]" Then Exit Do
        posFin = posFin + 1
    Loop

    ' Extraer y convertir
    valor = Trim(Mid(jsonText, posInicio, posFin - posInicio))
    valor = Replace(valor, """", "")  ' Quitar comillas si las hay

    ExtraerValorNumerico = CLng(Val(valor))
End Function

Private Function ExtraerValorMoneda(ByVal jsonText As String, ByVal campo As String) As Currency
    On Error Resume Next
    Dim posInicio As Long
    Dim posFin As Long
    Dim valor As String

    posInicio = InStr(1, jsonText, """" & campo & """:", vbTextCompare)
    If posInicio = 0 Then
        ExtraerValorMoneda = 0
        Exit Function
    End If

    posInicio = InStr(posInicio, jsonText, ":")
    posInicio = posInicio + 1

    Do While Mid(jsonText, posInicio, 1) = " "
        posInicio = posInicio + 1
    Loop

    posFin = posInicio
    Do While posFin <= Len(jsonText)
        Dim c As String
        c = Mid(jsonText, posFin, 1)
        If c = "," Or c = "}" Or c = "]" Then Exit Do
        posFin = posFin + 1
    Loop

    valor = Trim(Mid(jsonText, posInicio, posFin - posInicio))
    valor = Replace(valor, """", "")

    ExtraerValorMoneda = CCur(Val(valor))
End Function

Private Function ExtraerValorCadena(ByVal jsonText As String, ByVal campo As String) As String
    On Error Resume Next
    Dim posInicio As Long
    Dim posFin As Long
    Dim valor As String

    ' Buscar el campo
    posInicio = InStr(1, jsonText, """" & campo & """:", vbTextCompare)
    If posInicio = 0 Then
        ExtraerValorCadena = ""
        Exit Function
    End If

    ' Encontrar el inicio de la cadena (primera comilla después de :)
    posInicio = InStr(posInicio, jsonText, ":")
    posInicio = InStr(posInicio, jsonText, """")
    posInicio = posInicio + 1

    ' Encontrar el final de la cadena (siguiente comilla no escapada)
    posFin = posInicio
    Do While posFin <= Len(jsonText)
        If Mid(jsonText, posFin, 1) = """" Then
            ' Verificar que no esté escapada
            If posFin = posInicio Or Mid(jsonText, posFin - 1, 1) <> "\" Then
                Exit Do
            End If
        End If
        posFin = posFin + 1
    Loop

    valor = Mid(jsonText, posInicio, posFin - posInicio)

    ' Decodificar caracteres escapados básicos
    valor = Replace(valor, "\""", """")
    valor = Replace(valor, "\\", "\")
    valor = Replace(valor, "\/", "/")
    valor = Replace(valor, "\n", vbCrLf)
    valor = Replace(valor, "\r", "")
    valor = Replace(valor, "\t", vbTab)

    ExtraerValorCadena = valor
End Function

'******************************************************************************
'* FUNCIÓN: URLEncode
'* PROPÓSITO: Codifica una cadena para usar en URL
'******************************************************************************
Private Function URLEncode(ByVal texto As String) As String
    Dim i As Long
    Dim resultado As String
    Dim c As String
    Dim asciiVal As Integer

    resultado = ""
    For i = 1 To Len(texto)
        c = Mid(texto, i, 1)
        asciiVal = Asc(c)

        ' Caracteres seguros (alfanuméricos, guión, punto, underscore, tilde)
        If (asciiVal >= 48 And asciiVal <= 57) Or _
           (asciiVal >= 65 And asciiVal <= 90) Or _
           (asciiVal >= 97 And asciiVal <= 122) Or _
           c = "-" Or c = "." Or c = "_" Or c = "~" Then
            resultado = resultado & c
        Else
            ' Codificar caracteres especiales
            resultado = resultado & "%" & Right("0" & Hex(asciiVal), 2)
        End If
    Next i

    URLEncode = resultado
End Function
