Attribute VB_Name = "ModuloPrestaShop"
'========================================================================
' MÓDULO DE INTEGRACIÓN CON PRESTASHOP 8.1
'========================================================================
' Propósito: Comunicación HTTP con API Bridge (PHP) para consultar
'            productos y stock de PrestaShop
'
' Fase 1: SOLO LECTURA
' - Buscar producto por código
' - Obtener stock de producto
' - Obtener información completa de producto
'
' Dependencias: WinHTTP (WinHttp.WinHttpRequest.5.1)
' Autor: Claude Code
' Fecha: 19/12/2025
'========================================================================

Option Explicit

' Constantes de configuración (se cargan desde BD Access)
Private API_BRIDGE_URL As String
Private API_TIMEOUT As Integer
Private DEBUG_MODE As Boolean

' Estado de inicialización
Private Inicializado As Boolean

'========================================================================
' TIPO DE DATOS: PRODUCTO PRESTASHOP
'========================================================================
Public Type ProductoPS
    ID As Long
    Reference As String
    EAN13 As String
    Nombre As String
    Descripcion As String
    PrecioSinIVA As Currency
    PrecioConIVA As Currency
    IVA As Integer
    Stock As Long
    Activo As Boolean
    URLImagen As String
    FechaConsulta As Date
    Encontrado As Boolean
End Type

'========================================================================
' INICIALIZACIÓN
'========================================================================

Public Function InicializarModuloPS() As Boolean
    On Error GoTo ErrorHandler

    ' Cargar configuración desde tabla ConfigAPI
    API_BRIDGE_URL = GetConfigValue("API_BRIDGE_URL")
    API_TIMEOUT = Val(GetConfigValue("API_TIMEOUT"))
    DEBUG_MODE = (GetConfigValue("DEBUG_MODE") = "True")

    ' Validar configuración
    If API_BRIDGE_URL = "" Then
        MsgBox "Error: API_BRIDGE_URL no configurada en tabla ConfigAPI", vbCritical
        InicializarModuloPS = False
        Exit Function
    End If

    ' Verificar conectividad con API Bridge
    If Not TestConexionAPIBridge() Then
        MsgBox "Advertencia: No se pudo conectar con API Bridge" & vbCrLf & _
               "URL: " & API_BRIDGE_URL & vbCrLf & _
               "El sistema funcionará en modo OFFLINE", vbExclamation
        InicializarModuloPS = False
        Exit Function
    End If

    Inicializado = True
    InicializarModuloPS = True

    If DEBUG_MODE Then
        Debug.Print "ModuloPrestaShop inicializado correctamente"
        Debug.Print "API Bridge URL: " & API_BRIDGE_URL
    End If

    Exit Function

ErrorHandler:
    MsgBox "Error al inicializar módulo PrestaShop: " & Err.Description, vbCritical
    InicializarModuloPS = False
End Function

'========================================================================
' FUNCIÓN PRINCIPAL: BUSCAR PRODUCTO POR CÓDIGO
'========================================================================

Public Function BuscarProductoPorCodigo(Codigo As String) As ProductoPS
    On Error GoTo ErrorHandler

    Dim resultado As ProductoPS
    resultado.Encontrado = False

    ' Verificar inicialización
    If Not Inicializado Then
        If Not InicializarModuloPS() Then
            BuscarProductoPorCodigo = resultado
            Exit Function
        End If
    End If

    ' Intentar buscar en caché local primero
    resultado = BuscarEnCache(Codigo)
    If resultado.Encontrado Then
        If DEBUG_MODE Then Debug.Print "Producto encontrado en caché: " & Codigo
        BuscarProductoPorCodigo = resultado
        Exit Function
    End If

    ' Si no está en caché, consultar API Bridge
    Dim url As String
    url = API_BRIDGE_URL & "?action=buscar_producto&codigo=" & URLEncode(Codigo)

    Dim respuestaJSON As String
    respuestaJSON = HacerPeticionHTTP(url)

    If respuestaJSON = "" Then
        ' Error de conexión, devolver resultado vacío
        BuscarProductoPorCodigo = resultado
        Exit Function
    End If

    ' Parsear respuesta JSON
    resultado = ParsearProductoJSON(respuestaJSON)

    ' Si se encontró, guardar en caché
    If resultado.Encontrado Then
        GuardarEnCache resultado
        RegistrarLogSync "BUSQUEDA", resultado.ID, resultado.Reference, _
                        "Producto encontrado: " & resultado.Nombre, respuestaJSON
    Else
        RegistrarLogSync "BUSQUEDA", 0, Codigo, "Producto no encontrado", respuestaJSON
    End If

    BuscarProductoPorCodigo = resultado
    Exit Function

ErrorHandler:
    MsgBox "Error al buscar producto: " & Err.Description, vbCritical
    resultado.Encontrado = False
    BuscarProductoPorCodigo = resultado
End Function

'========================================================================
' FUNCIÓN: OBTENER STOCK DE PRODUCTO
'========================================================================

Public Function ObtenerStockProducto(IdProducto As Long) As Long
    On Error GoTo ErrorHandler

    ObtenerStockProducto = 0

    ' Verificar inicialización
    If Not Inicializado Then
        If Not InicializarModuloPS() Then
            Exit Function
        End If
    End If

    ' Construir URL
    Dim url As String
    url = API_BRIDGE_URL & "?action=obtener_stock&id=" & IdProducto

    ' Hacer petición
    Dim respuestaJSON As String
    respuestaJSON = HacerPeticionHTTP(url)

    If respuestaJSON = "" Then
        Exit Function
    End If

    ' Parsear JSON para extraer cantidad
    Dim stock As Long
    stock = ParsearStockJSON(respuestaJSON)

    ObtenerStockProducto = stock

    RegistrarLogSync "STOCK", IdProducto, "", "Stock obtenido: " & stock, respuestaJSON

    Exit Function

ErrorHandler:
    MsgBox "Error al obtener stock: " & Err.Description, vbCritical
    ObtenerStockProducto = 0
End Function

'========================================================================
' FUNCIÓN: BUSCAR EN CACHÉ LOCAL
'========================================================================

Private Function BuscarEnCache(Codigo As String) As ProductoPS
    On Error GoTo ErrorHandler

    Dim rs As Recordset
    Dim producto As ProductoPS
    producto.Encontrado = False

    ' Buscar en tabla ProductosPS
    Dim sql As String
    sql = "SELECT * FROM ProductosPS WHERE Referencia='" & Codigo & "'"

    Set rs = bdtienda.OpenRecordset(sql)

    If Not rs.EOF Then
        ' Verificar si el caché es válido (menos de 60 minutos)
        Dim minutosCache As Long
        minutosCache = Val(GetConfigValue("CACHE_EXPIRATION_MINUTES"))

        If Not IsNull(rs!UltimaConsulta) Then
            If DateDiff("n", rs!UltimaConsulta, Now) < minutosCache Then
                ' Caché válido, cargar datos
                producto.ID = rs!IDProductoPS
                producto.Reference = rs!Referencia & ""
                producto.EAN13 = rs!EAN13 & ""
                producto.Nombre = rs!Nombre & ""
                producto.Descripcion = rs!Descripcion & ""
                producto.PrecioSinIVA = rs!PrecioSinIVA
                producto.PrecioConIVA = rs!PrecioConIVA
                producto.IVA = rs!IVA
                producto.Stock = rs!StockPS
                producto.Activo = rs!Activo
                producto.URLImagen = rs!URLImagen & ""
                producto.FechaConsulta = rs!UltimaConsulta
                producto.Encontrado = True
            End If
        End If
    End If

    rs.Close
    Set rs = Nothing

    BuscarEnCache = producto
    Exit Function

ErrorHandler:
    producto.Encontrado = False
    BuscarEnCache = producto
End Function

'========================================================================
' FUNCIÓN: GUARDAR EN CACHÉ LOCAL
'========================================================================

Private Sub GuardarEnCache(producto As ProductoPS)
    On Error GoTo ErrorHandler

    Dim rs As Recordset
    Dim sql As String

    ' Verificar si ya existe
    sql = "SELECT * FROM ProductosPS WHERE IDProductoPS=" & producto.ID
    Set rs = bdtienda.OpenRecordset(sql)

    If rs.EOF Then
        ' Insertar nuevo
        rs.AddNew
        rs!IDProductoPS = producto.ID
    Else
        ' Actualizar existente
        rs.Edit
    End If

    ' Actualizar campos
    rs!Referencia = producto.Reference
    rs!EAN13 = producto.EAN13
    rs!Nombre = producto.Nombre
    rs!Descripcion = producto.Descripcion
    rs!PrecioSinIVA = producto.PrecioSinIVA
    rs!PrecioConIVA = producto.PrecioConIVA
    rs!IVA = producto.IVA
    rs!StockPS = producto.Stock
    rs!Activo = producto.Activo
    rs!URLImagen = producto.URLImagen
    rs!UltimaConsulta = Now
    rs!EstadoSync = "OK"

    rs.Update
    rs.Close
    Set rs = Nothing

    Exit Sub

ErrorHandler:
    If DEBUG_MODE Then Debug.Print "Error al guardar en caché: " & Err.Description
End Sub

'========================================================================
' FUNCIÓN: HACER PETICIÓN HTTP
'========================================================================

Private Function HacerPeticionHTTP(url As String) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim respuesta As String
    Dim tiempoInicio As Double

    tiempoInicio = Timer

    ' Crear objeto WinHTTP
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Configurar petición
    http.Open "GET", url, False
    http.SetTimeouts API_TIMEOUT * 1000, API_TIMEOUT * 1000, API_TIMEOUT * 1000, API_TIMEOUT * 1000

    ' Enviar petición
    http.Send

    ' Obtener respuesta
    respuesta = http.ResponseText

    If DEBUG_MODE Then
        Debug.Print "Petición HTTP: " & url
        Debug.Print "Tiempo: " & Format((Timer - tiempoInicio) * 1000, "0") & "ms"
        Debug.Print "Código: " & http.Status
    End If

    ' Verificar código HTTP
    If http.Status <> 200 Then
        MsgBox "Error HTTP " & http.Status & ": " & http.StatusText, vbExclamation
        HacerPeticionHTTP = ""
        Exit Function
    End If

    HacerPeticionHTTP = respuesta

    Set http = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Error de conexión: " & Err.Description & vbCrLf & _
           "Verificar conexión a internet y que API Bridge esté activo", vbCritical
    HacerPeticionHTTP = ""
End Function

'========================================================================
' FUNCIÓN: PARSEAR JSON DE PRODUCTO
'========================================================================

Private Function ParsearProductoJSON(jsonStr As String) As ProductoPS
    On Error GoTo ErrorHandler

    Dim producto As ProductoPS
    producto.Encontrado = False

    ' Verificar si la respuesta es exitosa
    If InStr(jsonStr, """success"": true") = 0 And InStr(jsonStr, """success"":true") = 0 Then
        ' Error en la respuesta
        ParsearProductoJSON = producto
        Exit Function
    End If

    ' Parsear campos manualmente (VB6 no tiene JSON nativo)
    producto.ID = ConvertirALong(ExtraerValorJSON(jsonStr, "id", "number"))
    producto.Reference = ExtraerValorJSON(jsonStr, "reference", "string")
    producto.EAN13 = ExtraerValorJSON(jsonStr, "ean13", "string")
    producto.Nombre = ExtraerValorJSON(jsonStr, "nombre", "string")
    producto.Descripcion = ExtraerValorJSON(jsonStr, "descripcion", "string")

    ' Convertir precios con manejo de errores
    producto.PrecioSinIVA = ConvertirACurrency(ExtraerValorJSON(jsonStr, "precio_sin_iva", "number"))
    producto.PrecioConIVA = ConvertirACurrency(ExtraerValorJSON(jsonStr, "precio_con_iva", "number"))

    producto.IVA = ConvertirAInteger(ExtraerValorJSON(jsonStr, "iva", "number"))
    producto.Stock = ConvertirALong(ExtraerValorJSON(jsonStr, "stock", "number"))
    producto.Activo = (ExtraerValorJSON(jsonStr, "activo", "boolean") = "true")
    producto.URLImagen = ExtraerValorJSON(jsonStr, "url_imagen", "string")
    producto.FechaConsulta = Now
    producto.Encontrado = True

    ParsearProductoJSON = producto
    Exit Function

ErrorHandler:
    If DEBUG_MODE Then Debug.Print "Error al parsear JSON: " & Err.Description
    producto.Encontrado = False
    ParsearProductoJSON = producto
End Function

'========================================================================
' FUNCIÓN: PARSEAR JSON DE STOCK
'========================================================================

Private Function ParsearStockJSON(jsonStr As String) As Long
    On Error GoTo ErrorHandler

    ParsearStockJSON = 0

    ' Verificar éxito
    If InStr(jsonStr, """success"": true") = 0 And InStr(jsonStr, """success"":true") = 0 Then
        Exit Function
    End If

    ' Extraer cantidad
    Dim cantidad As String
    cantidad = ExtraerValorJSON(jsonStr, "cantidad", "number")

    ParsearStockJSON = CLng(cantidad)
    Exit Function

ErrorHandler:
    ParsearStockJSON = 0
End Function

'========================================================================
' FUNCIÓN: EXTRAER VALOR DE JSON (PARSER SIMPLE)
'========================================================================

Private Function ExtraerValorJSON(jsonStr As String, clave As String, tipo As String) As String
    On Error GoTo ErrorHandler

    Dim patron As String
    Dim posInicio As Long
    Dim posFin As Long
    Dim valor As String

    ' Construir patrón de búsqueda según tipo
    Select Case tipo
        Case "string"
            patron = """" & clave & """: """
        Case "number", "boolean"
            patron = """" & clave & """: "
        Case Else
            patron = """" & clave & """: "
    End Select

    ' Buscar inicio
    posInicio = InStr(jsonStr, patron)
    If posInicio = 0 Then
        ExtraerValorJSON = ""
        Exit Function
    End If

    posInicio = posInicio + Len(patron)

    ' Buscar fin según tipo
    Select Case tipo
        Case "string"
            posFin = InStr(posInicio, jsonStr, """")
        Case "number"
            posFin = InStr(posInicio, jsonStr, ",")
            If posFin = 0 Then posFin = InStr(posInicio, jsonStr, "}")
        Case "boolean"
            posFin = InStr(posInicio, jsonStr, ",")
            If posFin = 0 Then posFin = InStr(posInicio, jsonStr, "}")
    End Select

    If posFin = 0 Then
        ExtraerValorJSON = ""
        Exit Function
    End If

    ' Extraer valor
    valor = Mid(jsonStr, posInicio, posFin - posInicio)
    valor = Trim(valor)

    ExtraerValorJSON = valor
    Exit Function

ErrorHandler:
    ExtraerValorJSON = ""
End Function

'========================================================================
' FUNCIÓN: TEST DE CONEXIÓN CON API BRIDGE
'========================================================================

Public Function TestConexionAPIBridge() As Boolean
    On Error GoTo ErrorHandler

    Dim url As String
    url = API_BRIDGE_URL & "?action=test"

    Dim respuesta As String
    respuesta = HacerPeticionHTTP(url)

    If respuesta = "" Then
        TestConexionAPIBridge = False
        Exit Function
    End If

    ' Verificar que la respuesta contiene "success": true
    If InStr(respuesta, """success"": true") > 0 Or InStr(respuesta, """success"":true") > 0 Then
        TestConexionAPIBridge = True
        If DEBUG_MODE Then Debug.Print "Test de conexión OK"
    Else
        TestConexionAPIBridge = False
    End If

    Exit Function

ErrorHandler:
    TestConexionAPIBridge = False
End Function

'========================================================================
' FUNCIÓN: REGISTRAR LOG DE SINCRONIZACIÓN
'========================================================================

Private Sub RegistrarLogSync(tipoOp As String, idProducto As Long, _
                             referencia As String, descripcion As String, _
                             respuestaAPI As String)
    On Error GoTo ErrorHandler

    Dim rs As Recordset
    Set rs = bdtienda.OpenRecordset("LogSincronizacion")

    With rs
        .AddNew
        !TipoOperacion = tipoOp
        If idProducto > 0 Then !IDProductoPS = idProducto
        If referencia <> "" Then !referencia = referencia
        !descripcion = descripcion
        !RespuestaAPI = Left(respuestaAPI, 65000) ' Límite de Memo
        !CodigoHTTP = 200
        !UsuarioVB = Environ$("USERNAME")
        .Update
    End With

    rs.Close
    Set rs = Nothing

    Exit Sub

ErrorHandler:
    ' No detener ejecución si falla el log
    If DEBUG_MODE Then Debug.Print "Error al registrar log: " & Err.Description
End Sub

'========================================================================
' FUNCIÓN: OBTENER VALOR DE CONFIGURACIÓN
'========================================================================

Private Function GetConfigValue(clave As String) As String
    On Error GoTo ErrorHandler

    Dim rs As Recordset
    Dim sql As String

    sql = "SELECT Valor FROM ConfigAPI WHERE Clave='" & clave & "'"
    Set rs = bdtienda.OpenRecordset(sql)

    If Not rs.EOF Then
        GetConfigValue = rs!Valor & ""
    Else
        GetConfigValue = ""
    End If

    rs.Close
    Set rs = Nothing

    Exit Function

ErrorHandler:
    GetConfigValue = ""
End Function

'========================================================================
' FUNCIÓN: URL ENCODE (CODIFICAR CARACTERES ESPECIALES)
'========================================================================

Private Function URLEncode(texto As String) As String
    Dim i As Integer
    Dim resultado As String
    Dim caracter As String

    resultado = ""

    For i = 1 To Len(texto)
        caracter = Mid(texto, i, 1)

        ' Codificar caracteres especiales
        Select Case caracter
            Case " "
                resultado = resultado & "%20"
            Case "-", "_", ".", "~"
                resultado = resultado & caracter
            Case "A" To "Z", "a" To "z", "0" To "9"
                resultado = resultado & caracter
            Case Else
                ' Codificar otros caracteres
                resultado = resultado & "%" & Hex(Asc(caracter))
        End Select
    Next i

    URLEncode = resultado
End Function

'========================================================================
' FUNCIONES AUXILIARES DE CONVERSIÓN
'========================================================================

Private Function ConvertirACurrency(valor As String) As Currency
    On Error GoTo ErrorHandler

    If valor = "" Or IsNull(valor) Then
        ConvertirACurrency = 0
        Exit Function
    End If

    ' Reemplazar punto por coma (formato español)
    valor = Replace(valor, ".", ",")

    ConvertirACurrency = CCur(valor)
    Exit Function

ErrorHandler:
    If DEBUG_MODE Then Debug.Print "Error al convertir a Currency: " & valor & " - " & Err.Description
    ConvertirACurrency = 0
End Function

Private Function ConvertirALong(valor As String) As Long
    On Error GoTo ErrorHandler

    If valor = "" Or IsNull(valor) Then
        ConvertirALong = 0
        Exit Function
    End If

    ConvertirALong = CLng(valor)
    Exit Function

ErrorHandler:
    If DEBUG_MODE Then Debug.Print "Error al convertir a Long: " & valor & " - " & Err.Description
    ConvertirALong = 0
End Function

Private Function ConvertirAInteger(valor As String) As Integer
    On Error GoTo ErrorHandler

    If valor = "" Or IsNull(valor) Then
        ConvertirAInteger = 0
        Exit Function
    End If

    ConvertirAInteger = CInt(valor)
    Exit Function

ErrorHandler:
    If DEBUG_MODE Then Debug.Print "Error al convertir a Integer: " & valor & " - " & Err.Description
    ConvertirAInteger = 0
End Function
