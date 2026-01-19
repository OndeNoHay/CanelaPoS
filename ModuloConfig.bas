Attribute VB_Name = "ModuloConfig"
'******************************************************************************
'* MÓDULO: ModuloConfig.bas
'* PROPÓSITO: Gestión de configuración de integración PrestaShop
'* AUTOR: Claude Code
'* FECHA: 2025-12-29
'******************************************************************************
'--- Declaraciones API de Windows ---
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
     ByVal lpFileName As String) As Long
     
     
Option Explicit

'--- Estructura de configuración ---
Type ConfigPrestaShop
    IntegracionHabilitada As Boolean
    BuscarEnPrestaShop As Boolean
    ActualizarStockAutomatico As Boolean
    MostrarMensajesError As Boolean
    TimeoutSegundos As Integer
    logHabilitado As Boolean
    ModoDebug As Boolean
    URLAPIBridge As String
End Type

'--- Variables globales ---
Private config As ConfigPrestaShop
Private configCargada As Boolean
Private archivoINI As String

'******************************************************************************
'* FUNCIÓN: CargarConfiguracion
'* PROPÓSITO: Carga la configuración desde archivo INI
'* RETORNA: True si se cargó correctamente
'******************************************************************************
Public Function CargarConfiguracion() As Boolean
    On Error GoTo ErrorHandler

    Dim valor As String

    ' Determinar ruta del archivo INI
    archivoINI = App.Path & "\config\prestashop.ini"

    ' Valores por defecto
    config.IntegracionHabilitada = True
    config.BuscarEnPrestaShop = True
    config.ActualizarStockAutomatico = True
    config.MostrarMensajesError = False  ' No molestar al usuario con errores de API
    config.TimeoutSegundos = 30
    config.logHabilitado = True
    config.ModoDebug = False
    config.URLAPIBridge = "https://www.canelamoda.es/api_bridge/"

    ' Verificar si existe el archivo
    If Len(Dir(archivoINI)) = 0 Then
        ' Si no existe, crear con valores por defecto
        GuardarConfiguracion
        CargarConfiguracion = True
        configCargada = True
        Exit Function
    End If

    ' Cargar valores del archivo INI
    config.IntegracionHabilitada = (LeerINI("General", "IntegracionHabilitada", "1") = "1")
    config.BuscarEnPrestaShop = (LeerINI("General", "BuscarEnPrestaShop", "1") = "1")
    config.ActualizarStockAutomatico = (LeerINI("General", "ActualizarStockAutomatico", "1") = "1")
    config.MostrarMensajesError = (LeerINI("General", "MostrarMensajesError", "0") = "1")
    config.TimeoutSegundos = CInt(Val(LeerINI("General", "TimeoutSegundos", "30")))
    config.logHabilitado = (LeerINI("General", "LogHabilitado", "1") = "1")
    config.ModoDebug = (LeerINI("General", "ModoDebug", "0") = "1")
    config.URLAPIBridge = LeerINI("API", "URLBridge", "https://www.canelamoda.es/api_bridge/")

    configCargada = True
    CargarConfiguracion = True
    Exit Function

ErrorHandler:
    LogError "Error al cargar configuración: " & Err.Description
    ' Usar valores por defecto
    configCargada = True
    CargarConfiguracion = False
End Function

'******************************************************************************
'* FUNCIÓN: GuardarConfiguracion
'* PROPÓSITO: Guarda la configuración actual en archivo INI
'******************************************************************************
Public Sub GuardarConfiguracion()
    On Error Resume Next

    Dim configFolder As String

    ' Crear carpeta de configuración si no existe
    configFolder = App.Path & "\config\"
    If Len(Dir(configFolder, vbDirectory)) = 0 Then
        MkDir configFolder
    End If

    ' Guardar valores
    EscribirINI "General", "IntegracionHabilitada", IIf(config.IntegracionHabilitada, "1", "0")
    EscribirINI "General", "BuscarEnPrestaShop", IIf(config.BuscarEnPrestaShop, "1", "0")
    EscribirINI "General", "ActualizarStockAutomatico", IIf(config.ActualizarStockAutomatico, "1", "0")
    EscribirINI "General", "MostrarMensajesError", IIf(config.MostrarMensajesError, "1", "0")
    EscribirINI "General", "TimeoutSegundos", CStr(config.TimeoutSegundos)
    EscribirINI "General", "LogHabilitado", IIf(config.logHabilitado, "1", "0")
    EscribirINI "General", "ModoDebug", IIf(config.ModoDebug, "1", "0")
    EscribirINI "API", "URLBridge", config.URLAPIBridge

    LogInfo "Configuración guardada correctamente"
End Sub

'******************************************************************************
'* FUNCIÓN: ObtenerConfig
'* PROPÓSITO: Devuelve la configuración actual
'******************************************************************************
Public Function ObtenerConfig() As ConfigPrestaShop
    ' Si no está cargada, cargarla
    If Not configCargada Then
        CargarConfiguracion
    End If

    ObtenerConfig = config
End Function

'******************************************************************************
'* FUNCIÓN: EstablecerConfig
'* PROPÓSITO: Establece una nueva configuración
'******************************************************************************
Public Sub EstablecerConfig(nuevaConfig As ConfigPrestaShop)
    config = nuevaConfig
    configCargada = True
    GuardarConfiguracion
End Sub

'******************************************************************************
'* PROPIEDADES DE ACCESO RÁPIDO
'******************************************************************************

Public Function IntegracionHabilitada() As Boolean
    If Not configCargada Then CargarConfiguracion
    IntegracionHabilitada = config.IntegracionHabilitada
End Function

Public Sub HabilitarIntegracion(ByVal habilitar As Boolean)
    If Not configCargada Then CargarConfiguracion
    config.IntegracionHabilitada = habilitar
    GuardarConfiguracion
End Sub

Public Function BuscarEnPrestaShop() As Boolean
    If Not configCargada Then CargarConfiguracion
    BuscarEnPrestaShop = config.BuscarEnPrestaShop
End Function

Public Function ActualizarStockAutomatico() As Boolean
    If Not configCargada Then CargarConfiguracion
    ActualizarStockAutomatico = config.ActualizarStockAutomatico
End Function

Public Function logHabilitado() As Boolean
    If Not configCargada Then CargarConfiguracion
    logHabilitado = config.logHabilitado
End Function

Public Function ModoDebug() As Boolean
    If Not configCargada Then CargarConfiguracion
    ModoDebug = config.ModoDebug
End Function

Public Function ObtenerURLAPIBridge() As String
    If Not configCargada Then CargarConfiguracion
    ObtenerURLAPIBridge = config.URLAPIBridge
End Function

'******************************************************************************
'* FUNCIONES DE LECTURA/ESCRITURA DE ARCHIVO INI
'******************************************************************************



'******************************************************************************
'* FUNCIÓN: LeerINI
'* PROPÓSITO: Lee un valor del archivo INI
'******************************************************************************
Private Function LeerINI(ByVal seccion As String, ByVal clave As String, _
                        ByVal valorPorDefecto As String) As String
    Dim buffer As String
    Dim resultado As Long

    buffer = String(255, 0)
    resultado = GetPrivateProfileString(seccion, clave, valorPorDefecto, buffer, Len(buffer), archivoINI)

    If resultado > 0 Then
        LeerINI = Left(buffer, resultado)
    Else
        LeerINI = valorPorDefecto
    End If
End Function

'******************************************************************************
'* FUNCIÓN: EscribirINI
'* PROPÓSITO: Escribe un valor en el archivo INI
'******************************************************************************
Private Sub EscribirINI(ByVal seccion As String, ByVal clave As String, ByVal valor As String)
    WritePrivateProfileString seccion, clave, valor, archivoINI
End Sub

'******************************************************************************
'* FUNCIÓN: CrearConfiguracionPorDefecto
'* PROPÓSITO: Crea un archivo de configuración con valores por defecto
'******************************************************************************
Public Sub CrearConfiguracionPorDefecto()
    On Error Resume Next

    Dim configFolder As String
    Dim fileNum As Integer

    configFolder = App.Path & "\config\"

    ' Crear carpeta si no existe
    If Len(Dir(configFolder, vbDirectory)) = 0 Then
        MkDir configFolder
    End If

    ' Crear archivo con comentarios
    fileNum = FreeFile
    Open archivoINI For Output As #fileNum

    Print #fileNum, "; ======================================================================"
    Print #fileNum, "; CONFIGURACIÓN DE INTEGRACIÓN PRESTASHOP"
    Print #fileNum, "; CanelaPoS - Archivo generado automáticamente"
    Print #fileNum, "; Fecha: " & Format(Now, "dd/mm/yyyy hh:nn:ss")
    Print #fileNum, "; ======================================================================"
    Print #fileNum, ""
    Print #fileNum, "[General]"
    Print #fileNum, "; Habilita/deshabilita toda la integración con PrestaShop (1=Sí, 0=No)"
    Print #fileNum, "IntegracionHabilitada=1"
    Print #fileNum, ""
    Print #fileNum, "; Buscar productos en PrestaShop al escanear código (1=Sí, 0=No)"
    Print #fileNum, "BuscarEnPrestaShop=1"
    Print #fileNum, ""
    Print #fileNum, "; Actualizar stock automáticamente después de venta (1=Sí, 0=No)"
    Print #fileNum, "ActualizarStockAutomatico=1"
    Print #fileNum, ""
    Print #fileNum, "; Mostrar mensajes de error al usuario (1=Sí, 0=No)"
    Print #fileNum, "; Recomendado: 0 (los errores se registran en el log)"
    Print #fileNum, "MostrarMensajesError=0"
    Print #fileNum, ""
    Print #fileNum, "; Timeout en segundos para llamadas API"
    Print #fileNum, "TimeoutSegundos=30"
    Print #fileNum, ""
    Print #fileNum, "; Habilitar logging de operaciones (1=Sí, 0=No)"
    Print #fileNum, "LogHabilitado=1"
    Print #fileNum, ""
    Print #fileNum, "; Modo debug - registra información detallada (1=Sí, 0=No)"
    Print #fileNum, "ModoDebug=0"
    Print #fileNum, ""
    Print #fileNum, "[API]"
    Print #fileNum, "; URL del API Bridge (NO CAMBIAR sin autorización)"
    Print #fileNum, "URLBridge=https://www.canelamoda.es/api_bridge/"
    Print #fileNum, ""
    Print #fileNum, "; ======================================================================"
    Print #fileNum, "; NOTAS:"
    Print #fileNum, "; - La API Key está configurada en el servidor (api_bridge.php)"
    Print #fileNum, "; - Los logs se guardan en la carpeta 'logs' de la aplicación"
    Print #fileNum, "; - Si hay problemas, desactivar IntegracionHabilitada temporalmente"
    Print #fileNum, "; ======================================================================"

    Close #fileNum

    MsgBox "Archivo de configuración creado en:" & vbCrLf & archivoINI, vbInformation
End Sub

'******************************************************************************
'* FUNCIÓN: MostrarConfiguracion
'* PROPÓSITO: Muestra la configuración actual en un mensaje
'******************************************************************************
Public Sub MostrarConfiguracion()
    Dim msg As String

    If Not configCargada Then CargarConfiguracion

    msg = "CONFIGURACIÓN PRESTASHOP" & vbCrLf & String(50, "=") & vbCrLf & vbCrLf
    msg = msg & "Integración habilitada: " & IIf(config.IntegracionHabilitada, "SÍ", "NO") & vbCrLf
    msg = msg & "Buscar en PrestaShop: " & IIf(config.BuscarEnPrestaShop, "SÍ", "NO") & vbCrLf
    msg = msg & "Actualizar stock auto: " & IIf(config.ActualizarStockAutomatico, "SÍ", "NO") & vbCrLf
    msg = msg & "Mostrar errores: " & IIf(config.MostrarMensajesError, "SÍ", "NO") & vbCrLf
    msg = msg & "Timeout: " & config.TimeoutSegundos & " seg" & vbCrLf
    msg = msg & "Log habilitado: " & IIf(config.logHabilitado, "SÍ", "NO") & vbCrLf
    msg = msg & "Modo debug: " & IIf(config.ModoDebug, "SÍ", "NO") & vbCrLf
    msg = msg & vbCrLf & "URL API Bridge:" & vbCrLf & config.URLAPIBridge & vbCrLf
    msg = msg & vbCrLf & "Archivo config:" & vbCrLf & archivoINI

    MsgBox msg, vbInformation, "Configuración PrestaShop"
End Sub

'******************************************************************************
'* FUNCIÓN: EditarConfiguracion
'* PROPÓSITO: Abre el archivo INI en Notepad para edición
'******************************************************************************
Public Sub EditarConfiguracion()
    On Error Resume Next

    If Not configCargada Then CargarConfiguracion

    ' Si no existe, crear
    If Len(Dir(archivoINI)) = 0 Then
        CrearConfiguracionPorDefecto
    End If

    ' Abrir en Notepad
    Shell "notepad.exe """ & archivoINI & """", vbNormalFocus
End Sub
