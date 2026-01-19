Attribute VB_Name = "ModuloLog"
'******************************************************************************
'* MÓDULO: ModuloLog.bas
'* PROPÓSITO: Sistema de logging para depuración y auditoría
'* AUTOR: Claude Code
'* FECHA: 2025-12-29
'******************************************************************************

Option Explicit

'--- Variables globales ---
Private logHabilitado As Boolean
Private logFilePath As String
Private logFileHandle As Integer

'--- Niveles de log ---
Public Enum NivelLog
    LOG_INFO = 1
    LOG_WARNING = 2
    LOG_ERROR = 3
    LOG_DEBUG = 4
End Enum

'******************************************************************************
'* FUNCIÓN: InicializarLog
'* PROPÓSITO: Inicializa el sistema de logging
'* PARÁMETROS:
'*   - rutaArchivo: Ruta del archivo de log (opcional)
'*   - habilitado: True para habilitar el log (opcional, por defecto True)
'******************************************************************************
Public Sub InicializarLog(Optional ByVal rutaArchivo As String = "", _
                         Optional ByVal habilitado As Boolean = True)
    On Error Resume Next

    logHabilitado = habilitado

    If Not logHabilitado Then Exit Sub

    ' Si no se especifica ruta, usar carpeta de la aplicación
    If rutaArchivo = "" Then
        rutaArchivo = App.Path & "\logs\prestashop_" & Format(Now, "yyyymmdd") & ".log"
    End If

    logFilePath = rutaArchivo

    ' Crear carpeta de logs si no existe
    Dim logFolder As String
    logFolder = Left(logFilePath, InStrRev(logFilePath, "\"))

    ' Crear directorio si no existe
    If Len(Dir(logFolder, vbDirectory)) = 0 Then
        MkDir logFolder
    End If

    ' Escribir encabezado inicial
    EscribirLog String(80, "=")
    EscribirLog "INICIO DE SESIÓN: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    EscribirLog "Aplicación: CanelaPoS - Integración PrestaShop"
    EscribirLog String(80, "=")
End Sub

'******************************************************************************
'* FUNCIÓN: EscribirLog
'* PROPÓSITO: Escribe una entrada en el log
'* PARÁMETROS:
'*   - mensaje: Mensaje a escribir
'*   - nivel: Nivel de log (INFO, WARNING, ERROR, DEBUG)
'******************************************************************************
Public Sub EscribirLog(ByVal mensaje As String, Optional ByVal nivel As NivelLog = LOG_INFO)
    On Error Resume Next

    If Not logHabilitado Then Exit Sub

    Dim timestamp As String
    Dim nivelTexto As String
    Dim lineaCompleta As String

    ' Formatear timestamp
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    ' Determinar texto del nivel
    Select Case nivel
        Case LOG_INFO
            nivelTexto = "INFO"
        Case LOG_WARNING
            nivelTexto = "WARNING"
        Case LOG_ERROR
            nivelTexto = "ERROR"
        Case LOG_DEBUG
            nivelTexto = "DEBUG"
        Case Else
            nivelTexto = "INFO"
    End Select

    ' Construir línea completa
    lineaCompleta = "[" & timestamp & "] [" & nivelTexto & "] " & mensaje

    ' Escribir en archivo
    logFileHandle = FreeFile
    Open logFilePath For Append As #logFileHandle
    Print #logFileHandle, lineaCompleta
    Close #logFileHandle

    ' Si es error, también escribir en Debug
    If nivel = LOG_ERROR Then
        Debug.Print lineaCompleta
    End If
End Sub

'******************************************************************************
'* FUNCIÓN: LogError
'* PROPÓSITO: Escribe un error en el log
'******************************************************************************
Public Sub LogError(ByVal mensaje As String)
    EscribirLog mensaje, LOG_ERROR
End Sub

'******************************************************************************
'* FUNCIÓN: LogWarning
'* PROPÓSITO: Escribe un warning en el log
'******************************************************************************
Public Sub LogWarning(ByVal mensaje As String)
    EscribirLog mensaje, LOG_WARNING
End Sub

'******************************************************************************
'* FUNCIÓN: LogDebug
'* PROPÓSITO: Escribe información de debug en el log
'******************************************************************************
Public Sub LogDebug(ByVal mensaje As String)
    EscribirLog mensaje, LOG_DEBUG
End Sub

'******************************************************************************
'* FUNCIÓN: LogInfo
'* PROPÓSITO: Escribe información general en el log
'******************************************************************************
Public Sub LogInfo(ByVal mensaje As String)
    EscribirLog mensaje, LOG_INFO
End Sub

'******************************************************************************
'* FUNCIÓN: LogOperacionVenta
'* PROPÓSITO: Registra una operación de venta con PrestaShop
'******************************************************************************
Public Sub LogOperacionVenta(ByVal idArticulo As Long, ByVal codigo As String, _
                            ByVal cantidad As Long, ByVal exito As Boolean, _
                            Optional ByVal detalles As String = "")
    Dim mensaje As String

    mensaje = "VENTA - Artículo: " & idArticulo & " | Código: " & codigo & _
              " | Cantidad: " & cantidad & " | Éxito: " & IIf(exito, "SÍ", "NO")

    If detalles <> "" Then
        mensaje = mensaje & " | " & detalles
    End If

    If exito Then
        EscribirLog mensaje, LOG_INFO
    Else
        EscribirLog mensaje, LOG_WARNING
    End If
End Sub

'******************************************************************************
'* FUNCIÓN: LogSincronizacionStock
'* PROPÓSITO: Registra una operación de sincronización de stock
'******************************************************************************
Public Sub LogSincronizacionStock(ByVal idProducto As Long, ByVal stockAnterior As Long, _
                                 ByVal stockNuevo As Long, ByVal exito As Boolean, _
                                 Optional ByVal error As String = "")
    Dim mensaje As String

    mensaje = "SYNC STOCK - Producto PS ID: " & idProducto & _
              " | Stock anterior: " & stockAnterior & _
              " | Stock nuevo: " & stockNuevo & _
              " | Éxito: " & IIf(exito, "SÍ", "NO")

    If error <> "" Then
        mensaje = mensaje & " | Error: " & error
    End If

    If exito Then
        EscribirLog mensaje, LOG_INFO
    Else
        EscribirLog mensaje, LOG_ERROR
    End If
End Sub

'******************************************************************************
'* FUNCIÓN: LogBusquedaProducto
'* PROPÓSITO: Registra una búsqueda de producto
'******************************************************************************
Public Sub LogBusquedaProducto(ByVal codigo As String, ByVal encontrado As Boolean, _
                              Optional ByVal idProducto As Long = 0, _
                              Optional ByVal nombreProducto As String = "")
    Dim mensaje As String

    mensaje = "BÚSQUEDA - Código: " & codigo & _
              " | Encontrado: " & IIf(encontrado, "SÍ", "NO")

    If encontrado Then
        mensaje = mensaje & " | ID PS: " & idProducto
        If nombreProducto <> "" Then
            mensaje = mensaje & " | Nombre: " & nombreProducto
        End If
        EscribirLog mensaje, LOG_INFO
    Else
        EscribirLog mensaje, LOG_DEBUG
    End If
End Sub

'******************************************************************************
'* FUNCIÓN: CerrarLog
'* PROPÓSITO: Cierra el log y escribe mensaje de finalización
'******************************************************************************
Public Sub CerrarLog()
    On Error Resume Next

    If Not logHabilitado Then Exit Sub

    EscribirLog String(80, "=")
    EscribirLog "FIN DE SESIÓN: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    EscribirLog String(80, "=")
End Sub

'******************************************************************************
'* FUNCIÓN: HabilitarLog
'* PROPÓSITO: Habilita o deshabilita el logging
'******************************************************************************
Public Sub HabilitarLog(ByVal habilitar As Boolean)
    logHabilitado = habilitar
End Sub

'******************************************************************************
'* FUNCIÓN: EstaHabilitado
'* PROPÓSITO: Devuelve si el logging está habilitado
'******************************************************************************
Public Function EstaHabilitado() As Boolean
    EstaHabilitado = logHabilitado
End Function

'******************************************************************************
'* FUNCIÓN: ObtenerRutaLog
'* PROPÓSITO: Devuelve la ruta del archivo de log actual
'******************************************************************************
Public Function ObtenerRutaLog() As String
    ObtenerRutaLog = logFilePath
End Function

'******************************************************************************
'* FUNCIÓN: LimpiarLogsAntiguos
'* PROPÓSITO: Elimina logs antiguos (más de X días)
'******************************************************************************
Public Sub LimpiarLogsAntiguos(Optional ByVal diasAMantener As Integer = 30)
    On Error Resume Next

    Dim logFolder As String
    Dim archivo As String
    Dim fechaArchivo As Date
    Dim fechaLimite As Date

    logFolder = App.Path & "\logs\"

    ' Verificar que existe la carpeta
    If Len(Dir(logFolder, vbDirectory)) = 0 Then Exit Sub

    fechaLimite = Date - diasAMantener

    ' Buscar archivos de log
    archivo = Dir(logFolder & "prestashop_*.log")

    Do While archivo <> ""
        ' Obtener fecha del archivo
        fechaArchivo = FileDateTime(logFolder & archivo)

        ' Si es más antiguo que el límite, eliminar
        If fechaArchivo < fechaLimite Then
            Kill logFolder & archivo
            Debug.Print "Log eliminado: " & archivo
        End If

        archivo = Dir
    Loop
End Sub

'******************************************************************************
'* FUNCIÓN: MostrarLog
'* PROPÓSITO: Abre el archivo de log en el Notepad
'******************************************************************************
Public Sub MostrarLog()
    On Error Resume Next

    If logFilePath = "" Then
        MsgBox "No hay archivo de log activo", vbInformation
        Exit Sub
    End If

    ' Verificar que existe el archivo
    If Len(Dir(logFilePath)) = 0 Then
        MsgBox "El archivo de log no existe: " & logFilePath, vbInformation
        Exit Sub
    End If

    ' Abrir en Notepad
    Shell "notepad.exe """ & logFilePath & """", vbNormalFocus
End Sub
