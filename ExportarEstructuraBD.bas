'==============================================================================
' MÓDULO: ExportarEstructuraBD
' DESCRIPCIÓN: Exporta la estructura completa de una base de datos Access
'              a un archivo de texto para análisis por Claude Code
' 
' INSTRUCCIONES DE USO:
' 1. Abre tu base de datos Access
' 2. Pulsa Alt+F11 para abrir el Editor de VBA
' 3. Menú Insertar → Módulo
' 4. Copia y pega todo este código
' 5. Pulsa F5 o ejecuta la macro "ExportarEstructuraCompleta"
' 6. El archivo se guardará en la misma carpeta que la BD
'==============================================================================

Option Compare Database
Option Explicit

' Constantes para tipos de datos de Access
Private Const adBoolean = 11
Private Const adTinyInt = 16
Private Const adSmallInt = 2
Private Const adInteger = 3
Private Const adBigInt = 20
Private Const adSingle = 4
Private Const adDouble = 5
Private Const adCurrency = 6
Private Const adDecimal = 14
Private Const adNumeric = 131
Private Const adDate = 7
Private Const adDBDate = 133
Private Const adDBTime = 134
Private Const adDBTimeStamp = 135
Private Const adChar = 129
Private Const adVarChar = 200
Private Const adLongVarChar = 201
Private Const adWChar = 130
Private Const adVarWChar = 202
Private Const adLongVarWChar = 203
Private Const adBinary = 128
Private Const adVarBinary = 204
Private Const adLongVarBinary = 205
Private Const adGUID = 72

'==============================================================================
' PROCEDIMIENTO PRINCIPAL
'==============================================================================
Public Sub ExportarEstructuraCompleta()
    Dim strRutaArchivo As String
    Dim intArchivo As Integer
    Dim strFecha As String
    
    On Error GoTo ErrorHandler
    
    ' Generar nombre de archivo con fecha
    strFecha = Format(Now, "yyyymmdd_hhmmss")
    strRutaArchivo = CurrentProject.Path & "\estructura_bd_" & strFecha & ".md"
    
    ' Abrir archivo para escritura
    intArchivo = FreeFile
    Open strRutaArchivo For Output As #intArchivo
    
    ' Escribir encabezado
    EscribirEncabezado intArchivo
    
    ' Exportar información general
    ExportarInfoGeneral intArchivo
    
    ' Exportar tablas
    ExportarTablas intArchivo
    
    ' Exportar relaciones
    ExportarRelaciones intArchivo
    
    ' Exportar consultas
    ExportarConsultas intArchivo
    
    ' Exportar formularios (solo nombres y controles principales)
    ExportarFormularios intArchivo
    
    ' Exportar módulos (solo nombres)
    ExportarModulos intArchivo
    
    ' Cerrar archivo
    Close #intArchivo
    
    MsgBox "Estructura exportada correctamente a:" & vbCrLf & vbCrLf & _
           strRutaArchivo, vbInformation, "Exportación Completada"
    
    Exit Sub
    
ErrorHandler:
    If intArchivo > 0 Then Close #intArchivo
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

'==============================================================================
' FUNCIONES DE EXPORTACIÓN
'==============================================================================

Private Sub EscribirEncabezado(intArchivo As Integer)
    Print #intArchivo, "# Estructura de Base de Datos Access"
    Print #intArchivo, ""
    Print #intArchivo, "Archivo generado automáticamente para análisis por Claude Code"
    Print #intArchivo, ""
    Print #intArchivo, "---"
    Print #intArchivo, ""
End Sub

Private Sub ExportarInfoGeneral(intArchivo As Integer)
    Print #intArchivo, "## Información General"
    Print #intArchivo, ""
    Print #intArchivo, "| Propiedad | Valor |"
    Print #intArchivo, "|-----------|-------|"
    Print #intArchivo, "| Nombre del archivo | " & CurrentProject.Name & " |"
    Print #intArchivo, "| Ruta | " & CurrentProject.Path & " |"
    Print #intArchivo, "| Fecha de exportación | " & Format(Now, "dd/mm/yyyy hh:mm:ss") & " |"
    Print #intArchivo, "| Versión de Access | " & Application.Version & " |"
    Print #intArchivo, ""
End Sub

Private Sub ExportarTablas(intArchivo As Integer)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim idxFld As DAO.Field
    Dim strTipo As String
    Dim strAtributos As String
    Dim intContadorTablas As Integer
    
    Set db = CurrentDb
    
    Print #intArchivo, "## Tablas"
    Print #intArchivo, ""
    
    intContadorTablas = 0
    
    For Each tdf In db.TableDefs
        ' Omitir tablas del sistema
        If Left(tdf.Name, 4) <> "MSys" And Left(tdf.Name, 1) <> "~" Then
            intContadorTablas = intContadorTablas + 1
            
            Print #intArchivo, "### " & intContadorTablas & ". " & tdf.Name
            Print #intArchivo, ""
            
            ' Información de la tabla
            Print #intArchivo, "**Registros aproximados:** " & DCount("*", tdf.Name)
            Print #intArchivo, ""
            
            ' Campos
            Print #intArchivo, "#### Campos"
            Print #intArchivo, ""
            Print #intArchivo, "| # | Campo | Tipo | Tamaño | Requerido | Valor Predeterminado |"
            Print #intArchivo, "|---|-------|------|--------|-----------|---------------------|"
            
            Dim intCampo As Integer
            intCampo = 0
            
            For Each fld In tdf.Fields
                intCampo = intCampo + 1
                strTipo = ObtenerNombreTipo(fld.Type)
                strAtributos = ""
                
                ' Verificar si es requerido
                If fld.Required Then
                    strAtributos = "Sí"
                Else
                    strAtributos = "No"
                End If
                
                Print #intArchivo, "| " & intCampo & " | " & fld.Name & " | " & strTipo & _
                      " | " & fld.Size & " | " & strAtributos & _
                      " | " & Nz(fld.DefaultValue, "-") & " |"
            Next fld
            
            Print #intArchivo, ""
            
            ' Índices
            If tdf.Indexes.Count > 0 Then
                Print #intArchivo, "#### Índices"
                Print #intArchivo, ""
                Print #intArchivo, "| Índice | Campos | Primario | Único |"
                Print #intArchivo, "|--------|--------|----------|-------|"
                
                For Each idx In tdf.Indexes
                    Dim strCamposIdx As String
                    strCamposIdx = ""
                    
                    For Each idxFld In idx.Fields
                        If strCamposIdx <> "" Then strCamposIdx = strCamposIdx & ", "
                        strCamposIdx = strCamposIdx & idxFld.Name
                    Next idxFld
                    
                    Print #intArchivo, "| " & idx.Name & " | " & strCamposIdx & _
                          " | " & IIf(idx.Primary, "Sí", "No") & _
                          " | " & IIf(idx.Unique, "Sí", "No") & " |"
                Next idx
                
                Print #intArchivo, ""
            End If
            
            Print #intArchivo, "---"
            Print #intArchivo, ""
        End If
    Next tdf
    
    Set db = Nothing
End Sub

Private Sub ExportarRelaciones(intArchivo As Integer)
    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim strAtributos As String
    
    Set db = CurrentDb
    
    Print #intArchivo, "## Relaciones"
    Print #intArchivo, ""
    
    If db.Relations.Count = 0 Then
        Print #intArchivo, "*No hay relaciones definidas*"
        Print #intArchivo, ""
    Else
        Print #intArchivo, "| Nombre | Tabla Principal | Tabla Relacionada | Campo Principal | Campo Foráneo | Integridad |"
        Print #intArchivo, "|--------|-----------------|-------------------|-----------------|---------------|------------|"
        
        For Each rel In db.Relations
            strAtributos = ""
            
            ' Verificar atributos de integridad referencial
            If (rel.Attributes And dbRelationUpdateCascade) Then
                strAtributos = "Actualizar cascada"
            End If
            If (rel.Attributes And dbRelationDeleteCascade) Then
                If strAtributos <> "" Then strAtributos = strAtributos & ", "
                strAtributos = strAtributos & "Eliminar cascada"
            End If
            If strAtributos = "" Then strAtributos = "Ninguna"
            
            For Each fld In rel.Fields
                Print #intArchivo, "| " & rel.Name & " | " & rel.Table & _
                      " | " & rel.ForeignTable & " | " & fld.Name & _
                      " | " & fld.ForeignName & " | " & strAtributos & " |"
            Next fld
        Next rel
        
        Print #intArchivo, ""
    End If
    
    Set db = Nothing
End Sub

Private Sub ExportarConsultas(intArchivo As Integer)
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim intContador As Integer
    
    Set db = CurrentDb
    
    Print #intArchivo, "## Consultas"
    Print #intArchivo, ""
    
    intContador = 0
    
    For Each qdf In db.QueryDefs
        ' Omitir consultas del sistema
        If Left(qdf.Name, 1) <> "~" Then
            intContador = intContador + 1
            
            Print #intArchivo, "### " & intContador & ". " & qdf.Name
            Print #intArchivo, ""
            Print #intArchivo, "**Tipo:** " & ObtenerTipoConsulta(qdf.Type)
            Print #intArchivo, ""
            Print #intArchivo, "```sql"
            Print #intArchivo, qdf.SQL
            Print #intArchivo, "```"
            Print #intArchivo, ""
        End If
    Next qdf
    
    If intContador = 0 Then
        Print #intArchivo, "*No hay consultas definidas*"
        Print #intArchivo, ""
    End If
    
    Set db = Nothing
End Sub

Private Sub ExportarFormularios(intArchivo As Integer)
    Dim obj As AccessObject
    Dim frm As Form
    Dim ctl As Control
    Dim intContador As Integer
    Dim strControles As String
    
    Print #intArchivo, "## Formularios"
    Print #intArchivo, ""
    
    intContador = 0
    
    For Each obj In CurrentProject.AllForms
        intContador = intContador + 1
        
        Print #intArchivo, "### " & intContador & ". " & obj.Name
        Print #intArchivo, ""
        
        ' Intentar abrir el formulario en modo diseño para obtener controles
        On Error Resume Next
        DoCmd.OpenForm obj.Name, acDesign, , , , acHidden
        
        If Err.Number = 0 Then
            Set frm = Forms(obj.Name)
            
            Print #intArchivo, "**Origen de registro:** " & Nz(frm.RecordSource, "*Sin origen*")
            Print #intArchivo, ""
            
            ' Listar controles importantes
            strControles = ""
            For Each ctl In frm.Controls
                Select Case ctl.ControlType
                    Case acTextBox, acComboBox, acListBox, acCommandButton, acSubform
                        If strControles <> "" Then strControles = strControles & ", "
                        strControles = strControles & ctl.Name & " (" & ObtenerTipoControl(ctl.ControlType) & ")"
                End Select
            Next ctl
            
            If strControles <> "" Then
                Print #intArchivo, "**Controles principales:** " & strControles
                Print #intArchivo, ""
            End If
            
            DoCmd.Close acForm, obj.Name, acSaveNo
        Else
            Print #intArchivo, "*No se pudo abrir para análisis*"
            Print #intArchivo, ""
            Err.Clear
        End If
        On Error GoTo 0
    Next obj
    
    If intContador = 0 Then
        Print #intArchivo, "*No hay formularios*"
        Print #intArchivo, ""
    End If
End Sub

Private Sub ExportarModulos(intArchivo As Integer)
    Dim obj As AccessObject
    Dim intContador As Integer
    
    Print #intArchivo, "## Módulos de Código"
    Print #intArchivo, ""
    Print #intArchivo, "| # | Nombre | Tipo |"
    Print #intArchivo, "|---|--------|------|"
    
    intContador = 0
    
    ' Módulos estándar
    For Each obj In CurrentProject.AllModules
        intContador = intContador + 1
        Print #intArchivo, "| " & intContador & " | " & obj.Name & " | Módulo estándar |"
    Next obj
    
    ' Módulos de clase
    On Error Resume Next
    For Each obj In CurrentDb.Containers("Modules").Documents
        ' Solo si no está ya listado
    Next obj
    On Error GoTo 0
    
    If intContador = 0 Then
        Print #intArchivo, "| - | *No hay módulos* | - |"
    End If
    
    Print #intArchivo, ""
    
    ' Nota sobre código en formularios
    Print #intArchivo, "**Nota:** Los formularios pueden contener código adicional en sus módulos de clase."
    Print #intArchivo, ""
End Sub

'==============================================================================
' FUNCIONES AUXILIARES
'==============================================================================

Private Function ObtenerNombreTipo(intTipo As Integer) As String
    Select Case intTipo
        Case dbBoolean: ObtenerNombreTipo = "Sí/No"
        Case dbByte: ObtenerNombreTipo = "Byte"
        Case dbInteger: ObtenerNombreTipo = "Entero"
        Case dbLong: ObtenerNombreTipo = "Entero largo"
        Case dbCurrency: ObtenerNombreTipo = "Moneda"
        Case dbSingle: ObtenerNombreTipo = "Simple"
        Case dbDouble: ObtenerNombreTipo = "Doble"
        Case dbDate: ObtenerNombreTipo = "Fecha/Hora"
        Case dbText: ObtenerNombreTipo = "Texto"
        Case dbLongBinary: ObtenerNombreTipo = "Objeto OLE"
        Case dbMemo: ObtenerNombreTipo = "Memo"
        Case dbGUID: ObtenerNombreTipo = "GUID"
        Case dbBigInt: ObtenerNombreTipo = "Entero grande"
        Case dbVarBinary: ObtenerNombreTipo = "Binario"
        Case dbChar: ObtenerNombreTipo = "Carácter"
        Case dbNumeric: ObtenerNombreTipo = "Numérico"
        Case dbDecimal: ObtenerNombreTipo = "Decimal"
        Case dbFloat: ObtenerNombreTipo = "Flotante"
        Case dbTime: ObtenerNombreTipo = "Hora"
        Case dbTimeStamp: ObtenerNombreTipo = "Marca de tiempo"
        Case Else: ObtenerNombreTipo = "Tipo " & intTipo
    End Select
End Function

Private Function ObtenerTipoConsulta(intTipo As Integer) As String
    Select Case intTipo
        Case 0: ObtenerTipoConsulta = "Selección"
        Case 16: ObtenerTipoConsulta = "Tabla de referencias cruzadas"
        Case 32: ObtenerTipoConsulta = "Eliminación"
        Case 48: ObtenerTipoConsulta = "Actualización"
        Case 64: ObtenerTipoConsulta = "Anexar"
        Case 80: ObtenerTipoConsulta = "Creación de tabla"
        Case 96: ObtenerTipoConsulta = "Definición de datos"
        Case 112: ObtenerTipoConsulta = "Paso a través SQL"
        Case 128: ObtenerTipoConsulta = "Unión"
        Case 144: ObtenerTipoConsulta = "Paso a través SPT a granel"
        Case 224: ObtenerTipoConsulta = "Compuesta"
        Case 240: ObtenerTipoConsulta = "Procedimiento"
        Case Else: ObtenerTipoConsulta = "Tipo " & intTipo
    End Select
End Function

Private Function ObtenerTipoControl(intTipo As Integer) As String
    Select Case intTipo
        Case acTextBox: ObtenerTipoControl = "TextBox"
        Case acComboBox: ObtenerTipoControl = "ComboBox"
        Case acListBox: ObtenerTipoControl = "ListBox"
        Case acCommandButton: ObtenerTipoControl = "Botón"
        Case acSubform: ObtenerTipoControl = "Subformulario"
        Case acLabel: ObtenerTipoControl = "Etiqueta"
        Case acCheckBox: ObtenerTipoControl = "CheckBox"
        Case acOptionButton: ObtenerTipoControl = "OptionButton"
        Case acOptionGroup: ObtenerTipoControl = "Grupo de opciones"
        Case acImage: ObtenerTipoControl = "Imagen"
        Case Else: ObtenerTipoControl = "Control"
    End Select
End Function
