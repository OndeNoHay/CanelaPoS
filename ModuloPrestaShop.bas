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
Private Const PS_API_TIMEOUT As Long = 20000  ' 30 segundos

'--- Tipos de datos ---

' Combinacion (talla) de un producto
Type CombinacionProducto
    idCombinacion As Long
    Talla As String
    stock As Long
    Disponible As Boolean
End Type

' Producto de PrestaShop
Type ProductoPrestaShop
    idProducto As Long
    Referencia As String
    EAN As String
    Nombre As String
    Descripcion As String
    PrecioSinIVA As Currency
    PrecioConIVA As Currency
    PorcentajeIVA As Double
    StockDisponible As Long
    TieneCombinaciones As Boolean
    idCombinacion As Long
    Activo As Boolean
    encontrado As Boolean
    MensajeError As String
    ' Combinaciones (tallas)
    Combinaciones(1 To 50) As CombinacionProducto  ' Maximo 50 tallas por producto
    NumCombinaciones As Integer
End Type

Type ResultadoActualizacion
    exito As Boolean
    stockAnterior As Long
    stockNuevo As Long
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
    producto.encontrado = False
    producto.TieneCombinaciones = False
    producto.idProducto = 0
    producto.idCombinacion = 0

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

        If producto.encontrado Then
            EscribirLog "Producto encontrado: " & producto.Nombre & " (ID: " & producto.idProducto & ")"
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

    resultado.exito = False
    resultado.stockAnterior = 0
    resultado.stockNuevo = 0

    ' Validar parámetros
    If idProducto <= 0 Then
        resultado.MensajeError = "ID de producto inválido"
        ActualizarStock = resultado
        Exit Function
    End If

    ' Construir URL
    url = PS_API_BRIDGE_URL & "bridge.php?action=actualizar_stock"

    ' Construir datos POST en formato application/x-www-form-urlencoded
    postData = "id_producto=" & idProducto & _
               "&cantidad=" & cantidad & _
               "&id_combinacion=" & idCombinacion

    LogInfo "Actualizando stock - Producto: " & idProducto & _
            " | Cantidad: " & cantidad & _
            " | Combinacion: " & idCombinacion

    ' Crear objeto HTTP
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xmlHttp.setTimeouts 5000, 5000, PS_API_TIMEOUT, PS_API_TIMEOUT

    ' Realizar petición POST
    xmlHttp.Open "POST", url, False
    xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlHttp.Send postData

    ' Verificar respuesta HTTP
    If xmlHttp.Status <> 200 Then
        resultado.MensajeError = "Error HTTP: " & xmlHttp.Status
        LogError resultado.MensajeError
        ActualizarStock = resultado
        Set xmlHttp = Nothing
        Exit Function
    End If

    ' Obtener respuesta
    responseText = xmlHttp.responseText
    LogDebug "Respuesta stock update: " & Left(responseText, 200)

    ' Parsear resultado JSON
    resultado = ParsearResultadoActualizacionJSON(responseText)

    If resultado.exito Then
        LogInfo "Stock actualizado OK - Anterior: " & resultado.stockAnterior & _
                " | Nuevo: " & resultado.stockNuevo
    Else
        LogError "Error al actualizar stock: " & resultado.MensajeError
    End If

    Set xmlHttp = Nothing
    ActualizarStock = resultado
    Exit Function

ErrorHandler:
    resultado.MensajeError = "Error: " & Err.Description
    LogError "ERROR en ActualizarStock: " & Err.Description
    ActualizarStock = resultado
    If Not xmlHttp Is Nothing Then Set xmlHttp = Nothing
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

    producto.encontrado = False

    ' Verificar si hay éxito en la respuesta
    ' NOTA: "success" es un booleano (true/false) sin comillas en JSON
    Dim posSuccess As Long
    Dim posColon As Long
    Dim esExitoso As Boolean

    esExitoso = False
    posSuccess = InStr(1, jsonText, """success""", vbTextCompare)
    If posSuccess > 0 Then
        ' Buscar los : después de "success"
        posColon = InStr(posSuccess, jsonText, ":")
        If posColon > 0 Then
            ' Saltar espacios después de :
            posColon = posColon + 1
            Do While posColon <= Len(jsonText) And Mid(jsonText, posColon, 1) = " "
                posColon = posColon + 1
            Loop

            ' Verificar si empieza con "true" (case insensitive)
            If posColon + 3 <= Len(jsonText) Then
                If LCase(Mid(jsonText, posColon, 4)) = "true" Then
                    esExitoso = True
                End If
            End If
        End If
    End If

    ' Si success = false, extraer mensaje de error y salir
    If Not esExitoso Then
        producto.MensajeError = ExtraerValorCadena(jsonText, "mensaje")
        If producto.MensajeError = "" Then
            producto.MensajeError = ExtraerValorCadena(jsonText, "message")
        End If
        If producto.MensajeError = "" Then
            producto.MensajeError = "Producto no encontrado"
        End If
        ParsearProductoJSON = producto
        Exit Function
    End If

    ' Los datos están dentro de "data": {...}
    ' Extraer el contenido de "data"
    Dim dataContent As String
    Dim posDataStart As Long
    Dim posDataEnd As Long

    posDataStart = InStr(1, jsonText, """data""", vbTextCompare)
    If posDataStart > 0 Then
        ' Buscar el { después de "data":
        posDataStart = InStr(posDataStart, jsonText, "{")
        If posDataStart > 0 Then
            ' Encontrar el } correspondiente
            Dim nivel As Integer
            Dim i As Long
            nivel = 1
            posDataEnd = posDataStart
            For i = posDataStart + 1 To Len(jsonText)
                If Mid(jsonText, i, 1) = "{" Then nivel = nivel + 1
                If Mid(jsonText, i, 1) = "}" Then nivel = nivel - 1
                If nivel = 0 Then
                    posDataEnd = i
                    Exit For
                End If
            Next i

            ' Extraer contenido de data
            dataContent = Mid(jsonText, posDataStart, posDataEnd - posDataStart + 1)

            ' Ahora parsear usando dataContent
            producto.encontrado = True
        Else
            ' No hay contenido en data
            producto.MensajeError = "Respuesta sin datos"
            ParsearProductoJSON = producto
            Exit Function
        End If
    Else
        ' No hay campo "data", intentar parsear directamente
        dataContent = jsonText
        producto.encontrado = True
    End If

    ' Parsear campos usando dataContent

    ' Extraer ID del producto (bridge.php usa "id")
    producto.idProducto = ExtraerValorNumerico(dataContent, "id")

    ' Extraer referencia
    producto.Referencia = ExtraerValorCadena(dataContent, "reference")

    ' Extraer EAN
    producto.EAN = ExtraerValorCadena(dataContent, "ean13")

    ' Extraer nombre (bridge.php usa "nombre")
    producto.Nombre = ExtraerValorCadena(dataContent, "nombre")

    ' Extraer descripción (bridge.php usa "descripcion")
    producto.Descripcion = ExtraerValorCadena(dataContent, "descripcion")

    ' Extraer precios (bridge.php usa "precio_sin_iva" y "precio_con_iva")
    producto.PrecioSinIVA = ExtraerValorMoneda(dataContent, "precio_sin_iva")
    producto.PrecioConIVA = ExtraerValorMoneda(dataContent, "precio_con_iva")

    ' Extraer IVA
    producto.PorcentajeIVA = ExtraerValorNumerico(dataContent, "iva")
    If producto.PorcentajeIVA = 0 Then producto.PorcentajeIVA = 21  ' Por defecto

    ' Extraer stock (bridge.php usa "stock")
    producto.StockDisponible = ExtraerValorNumerico(dataContent, "stock")

    ' Verificar si tiene combinaciones (bridge.php usa "tiene_combinaciones")
    producto.TieneCombinaciones = ExtraerValorBooleano(dataContent, "tiene_combinaciones")

    ' Inicializar array de combinaciones
    producto.NumCombinaciones = 0

    ' Si tiene combinaciones, parsear array completo de tallas
    If producto.TieneCombinaciones Then
        producto.NumCombinaciones = ParsearCombinaciones(dataContent, producto.Combinaciones)

        If producto.NumCombinaciones > 0 Then
            ' Usar la primera combinación como referencia temporal
            ' (el usuario deberá seleccionar la talla específica en el formulario)
            producto.idCombinacion = producto.Combinaciones(1).idCombinacion
            LogInfo "Producto con " & producto.NumCombinaciones & " tallas disponibles"
        Else
            LogWarning "Producto tiene combinaciones pero ninguna con stock > 0"
        End If
    End If

    ' Verificar si está activo (bridge.php usa "activo")
    producto.Activo = ExtraerValorBooleano(dataContent, "activo")

    ' DEBUG: Verificar que el producto se ha parseado correctamente
    If ModoDebug Then
        ModuloLog.EscribirLog "PARSER - Producto parseado: ID=" & producto.idProducto & _
            " | Nombre=" & producto.Nombre & " | Precio=" & producto.PrecioConIVA & _
            " | Stock=" & producto.StockDisponible & " | Encontrado=" & producto.encontrado, LOG_DEBUG
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
    On Error Resume Next

    Dim resultado As ResultadoActualizacion
    Dim dataContent As String
    Dim posDataStart As Long
    Dim posDataEnd As Long
    Dim nivel As Integer
    Dim i As Long

    ' Verificar success usando ExtraerValorBooleano
    resultado.exito = ExtraerValorBooleano(jsonText, "success")

    If resultado.exito Then
        ' Extraer contenido de "data" si existe (mismo patrón que búsqueda de productos)
        posDataStart = InStr(1, jsonText, """data""", vbTextCompare)
        If posDataStart > 0 Then
            posDataStart = InStr(posDataStart, jsonText, "{")
            If posDataStart > 0 Then
                nivel = 1
                For i = posDataStart + 1 To Len(jsonText)
                    If Mid(jsonText, i, 1) = "{" Then nivel = nivel + 1
                    If Mid(jsonText, i, 1) = "}" Then nivel = nivel - 1
                    If nivel = 0 Then
                        posDataEnd = i
                        Exit For
                    End If
                Next i
                dataContent = Mid(jsonText, posDataStart, posDataEnd - posDataStart + 1)
            Else
                dataContent = jsonText
            End If
        Else
            dataContent = jsonText
        End If

        ' Extraer valores de stock (bridge.php usa "stock_anterior" y "stock_nuevo")
        resultado.stockAnterior = ExtraerValorNumerico(dataContent, "stock_anterior")
        resultado.stockNuevo = ExtraerValorNumerico(dataContent, "stock_nuevo")
    Else
        ' Si hay error, extraer mensaje
        resultado.MensajeError = ExtraerValorCadena(jsonText, "mensaje")
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
'* FUNCIÓN: ExtraerValorBooleano
'* PROPÓSITO: Extrae un valor booleano (true/false sin comillas) del JSON
'******************************************************************************
Private Function ExtraerValorBooleano(ByVal jsonText As String, ByVal campo As String) As Boolean
    On Error Resume Next
    Dim posInicio As Long
    Dim posColon As Long

    ExtraerValorBooleano = False

    ' Buscar el campo
    posInicio = InStr(1, jsonText, """" & campo & """:", vbTextCompare)
    If posInicio = 0 Then
        ' Buscar también con espacios: "campo" :
        posInicio = InStr(1, jsonText, """" & campo & """", vbTextCompare)
        If posInicio = 0 Then Exit Function
    End If

    ' Buscar los : después del campo
    posColon = InStr(posInicio, jsonText, ":")
    If posColon = 0 Then Exit Function

    ' Saltar espacios después de :
    posColon = posColon + 1
    Do While posColon <= Len(jsonText) And Mid(jsonText, posColon, 1) = " "
        posColon = posColon + 1
    Loop

    ' Verificar si empieza con "true" (case insensitive)
    If posColon + 3 <= Len(jsonText) Then
        If LCase(Mid(jsonText, posColon, 4)) = "true" Then
            ExtraerValorBooleano = True
        End If
    End If
End Function

'******************************************************************************
'* FUNCIÓN: ParsearCombinaciones
'* PROPÓSITO: Extrae array de combinaciones (tallas) del JSON
'* PARÁMETROS:
'*   - jsonText: Texto JSON que contiene "combinaciones": [...]
'*   - combos: Array donde se almacenarán las combinaciones (ByRef)
'* RETORNA: Número de combinaciones parseadas (solo las que tienen stock > 0)
'******************************************************************************
Private Function ParsearCombinaciones(ByVal jsonText As String, ByRef combos() As CombinacionProducto) As Integer
    On Error Resume Next

    Dim posArray As Long
    Dim posStart As Long
    Dim posEnd As Long
    Dim nivel As Integer
    Dim i As Long
    Dim numCombos As Integer
    Dim objetoCombo As String
    Dim combo As CombinacionProducto

    numCombos = 0

    ' Buscar "combinaciones": [
    posArray = InStr(1, jsonText, """combinaciones""", vbTextCompare)
    If posArray = 0 Then
        ParsearCombinaciones = 0
        Exit Function
    End If

    ' Buscar el [ que abre el array
    posArray = InStr(posArray, jsonText, "[")
    If posArray = 0 Then
        ParsearCombinaciones = 0
        Exit Function
    End If

    ' Buscar cada objeto {...} dentro del array
    posStart = posArray + 1

    Do While posStart < Len(jsonText) And numCombos < 50
        ' Saltar espacios y comas
        Do While posStart < Len(jsonText)
            Dim ch As String
            ch = Mid(jsonText, posStart, 1)
            If ch <> " " And ch <> vbCrLf And ch <> vbLf And ch <> vbTab And ch <> "," Then
                Exit Do
            End If
            posStart = posStart + 1
        Loop

        ' Si encontramos ], terminamos
        If Mid(jsonText, posStart, 1) = "]" Then Exit Do

        ' Si no es {, saltar
        If Mid(jsonText, posStart, 1) <> "{" Then Exit Do

        ' Encontrar el } correspondiente
        nivel = 1
        posEnd = posStart
        For i = posStart + 1 To Len(jsonText)
            If Mid(jsonText, i, 1) = "{" Then nivel = nivel + 1
            If Mid(jsonText, i, 1) = "}" Then nivel = nivel - 1
            If nivel = 0 Then
                posEnd = i
                Exit For
            End If
        Next i

        If nivel <> 0 Then Exit Do ' No se encontró el cierre

        ' Extraer objeto completo
        objetoCombo = Mid(jsonText, posStart, posEnd - posStart + 1)

        ' Parsear campos de la combinación
        combo.idCombinacion = ExtraerValorNumerico(objetoCombo, "id_combinacion")
        If combo.idCombinacion = 0 Then
            combo.idCombinacion = ExtraerValorNumerico(objetoCombo, "id_product_attribute")
        End If
        combo.Talla = ExtraerValorCadena(objetoCombo, "talla")
        combo.stock = ExtraerValorNumerico(objetoCombo, "stock")
        combo.Disponible = ExtraerValorBooleano(objetoCombo, "disponible")

        ' Solo agregar si tiene stock > 0 (requisito del usuario)
        If combo.stock > 0 Then
            numCombos = numCombos + 1
            combos(numCombos) = combo

            If ModoDebug Then
                LogDebug "Combinacion parseada: " & combo.Talla & " (ID: " & combo.idCombinacion & ", Stock: " & combo.stock & ")"
            End If
        End If

        ' Mover al siguiente objeto
        posStart = posEnd + 1
    Loop

    ParsearCombinaciones = numCombos
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

'******************************************************************************
'* FUNCIÓN: BuscarProductosPorRangoID
'* PROPÓSITO: Busca productos en PrestaShop por rango de IDs
'* PARÁMETROS:
'*   - idInicio: ID inicial del rango
'*   - idFin: ID final del rango
'*   - productos(): Array ByRef donde se almacenarán los productos encontrados
'* RETORNA: Número de productos encontrados (0 si hay error o no se encuentran)
'* NOTA: Solo retorna productos activos y con stock > 0
'******************************************************************************
Public Function BuscarProductosPorRangoID(ByVal idInicio As Long, ByVal idFin As Long, _
    ByRef productos() As ProductoPrestaShop) As Integer
    On Error GoTo ErrorHandler

    Dim xmlHttp As Object
    Dim url As String
    Dim responseText As String
    Dim numProductos As Integer

    BuscarProductosPorRangoID = 0
    numProductos = 0

    ' Validar rango
    If idInicio < 1 Or idFin < idInicio Then
        LogError "Rango de IDs inválido: " & idInicio & " - " & idFin
        Exit Function
    End If

    ' Limitar el rango a 500 productos
    If (idFin - idInicio) > 500 Then
        LogError "Rango demasiado grande: máximo 500 productos"
        Exit Function
    End If

    ' Construir URL para búsqueda por rango
    url = PS_API_BRIDGE_URL & "bridge.php?action=buscar_productos_rango&id_inicio=" & idInicio & "&id_fin=" & idFin

    LogInfo "Buscando productos del ID " & idInicio & " al " & idFin

    ' Crear objeto HTTP
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configurar timeout largo (puede tardar con muchos productos)
    xmlHttp.setTimeouts 10000, 10000, 60000, 60000

    ' Realizar petición GET
    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.setRequestHeader "Accept", "application/json"
    xmlHttp.Send

    ' Verificar respuesta
    If xmlHttp.Status = 200 Then
        responseText = xmlHttp.responseText
        LogDebug "Respuesta recibida: " & Left(responseText, 200)

        ' Parsear JSON response con array de productos
        numProductos = ParsearProductosRangoJSON(responseText, productos)

        If numProductos > 0 Then
            LogInfo "Productos encontrados en rango: " & numProductos
        Else
            LogInfo "No se encontraron productos activos con stock en el rango"
        End If
    Else
        LogError "Error HTTP: " & xmlHttp.Status & " - " & xmlHttp.statusText
    End If

    Set xmlHttp = Nothing
    BuscarProductosPorRangoID = numProductos
    Exit Function

ErrorHandler:
    LogError "Error en BuscarProductosPorRangoID: " & Err.Description
    BuscarProductosPorRangoID = 0
End Function

'******************************************************************************
'* FUNCIÓN: ParsearProductosRangoJSON
'* PROPÓSITO: Parsea la respuesta JSON del endpoint de búsqueda por rango
'* RETORNA: Número de productos parseados
'******************************************************************************
Private Function ParsearProductosRangoJSON(ByVal jsonText As String, _
    ByRef productos() As ProductoPrestaShop) As Integer
    On Error Resume Next

    Dim numProductos As Integer
    Dim posArray As Long
    Dim posStart As Long
    Dim posEnd As Long
    Dim nivel As Integer
    Dim i As Long
    Dim objetoProducto As String
    Dim producto As ProductoPrestaShop

    numProductos = 0

    ' Verificar success
    If Not ExtraerValorBooleano(jsonText, "success") Then
        LogError "Error en respuesta del servidor: " & ExtraerValorCadena(jsonText, "mensaje")
        ParsearProductosRangoJSON = 0
        Exit Function
    End If

    ' Buscar el array "productos" dentro de "data"
    ' Patrón: "data":{"productos":[...]}
    posArray = InStr(1, jsonText, """productos""", vbTextCompare)
    If posArray = 0 Then
        LogWarning "No se encontró array de productos en respuesta"
        ParsearProductosRangoJSON = 0
        Exit Function
    End If

    ' Buscar el [ que abre el array
    posArray = InStr(posArray, jsonText, "[")
    If posArray = 0 Then
        ParsearProductosRangoJSON = 0
        Exit Function
    End If

    ' Inicializar array de productos (máximo 500)
    ReDim productos(1 To 500)

    ' Buscar cada objeto {...} dentro del array
    posStart = posArray + 1

    Do While posStart < Len(jsonText) And numProductos < 500
        ' Saltar espacios y comas
        Do While posStart < Len(jsonText)
            Dim ch As String
            ch = Mid(jsonText, posStart, 1)
            If ch <> " " And ch <> vbCrLf And ch <> vbLf And ch <> vbTab And ch <> "," Then
                Exit Do
            End If
            posStart = posStart + 1
        Loop

        ' Si encontramos ], terminamos
        If Mid(jsonText, posStart, 1) = "]" Then Exit Do

        ' Si no es {, saltar
        If Mid(jsonText, posStart, 1) <> "{" Then Exit Do

        ' Encontrar el } correspondiente
        nivel = 1
        posEnd = posStart
        For i = posStart + 1 To Len(jsonText)
            If Mid(jsonText, i, 1) = "{" Then nivel = nivel + 1
            If Mid(jsonText, i, 1) = "}" Then nivel = nivel - 1
            If nivel = 0 Then
                posEnd = i
                Exit For
            End If
        Next i

        If nivel <> 0 Then Exit Do ' No se encontró el cierre

        ' Extraer objeto completo
        objetoProducto = Mid(jsonText, posStart, posEnd - posStart + 1)

        ' Parsear el producto (reutilizamos ParsearProductoJSON pero con el objeto individual)
        ' Como ParsearProductoJSON espera el JSON completo con "success" y "data",
        ' vamos a parsear manualmente este objeto
        producto = ParsearObjetoProducto(objetoProducto)

        If producto.encontrado Then
            numProductos = numProductos + 1
            productos(numProductos) = producto
        End If

        ' Mover al siguiente objeto
        posStart = posEnd + 1
    Loop

    ' Redimensionar array al tamaño real
    If numProductos > 0 Then
        ReDim Preserve productos(1 To numProductos)
    Else
        Erase productos
    End If

    ParsearProductosRangoJSON = numProductos
End Function

'******************************************************************************
'* FUNCIÓN: ParsearObjetoProducto
'* PROPÓSITO: Parsea un objeto JSON individual de producto
'******************************************************************************
Private Function ParsearObjetoProducto(ByVal jsonText As String) As ProductoPrestaShop
    On Error Resume Next

    Dim producto As ProductoPrestaShop

    producto.encontrado = True
    producto.TieneCombinaciones = False

    ' Parsear campos usando funciones existentes
    producto.idProducto = ExtraerValorNumerico(jsonText, "id")
    producto.Referencia = ExtraerValorCadena(jsonText, "reference")
    producto.EAN = ExtraerValorCadena(jsonText, "ean13")
    producto.Nombre = ExtraerValorCadena(jsonText, "nombre")
    producto.Descripcion = ExtraerValorCadena(jsonText, "descripcion")
    producto.PrecioSinIVA = ExtraerValorMoneda(jsonText, "precio_sin_iva")
    producto.PrecioConIVA = ExtraerValorMoneda(jsonText, "precio_con_iva")
    producto.PorcentajeIVA = ExtraerValorNumerico(jsonText, "iva")
    producto.StockDisponible = ExtraerValorNumerico(jsonText, "stock")
    producto.Activo = ExtraerValorBooleano(jsonText, "activo")
    producto.TieneCombinaciones = ExtraerValorBooleano(jsonText, "tiene_combinaciones")

    ' Parsear combinaciones si existen
    producto.NumCombinaciones = 0
    If producto.TieneCombinaciones Then
        producto.NumCombinaciones = ParsearCombinaciones(jsonText, producto.Combinaciones)
        If producto.NumCombinaciones > 0 Then
            producto.idCombinacion = producto.Combinaciones(1).idCombinacion
        End If
    End If

    ParsearObjetoProducto = producto
End Function
