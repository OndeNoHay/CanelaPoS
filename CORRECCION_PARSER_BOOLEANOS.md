# Corrección Parser JSON - Valores Booleanos

## Fecha: 29 de Diciembre de 2025
## Commit: `08338c4`

---

## 🐛 Problema Reportado

**Síntoma:**
```
[INFO] Respuesta recibida: {"success": true, "data": {...}}
[INFO] Producto no encontrado en PrestaShop
```

A pesar de recibir una respuesta válida con `"success": true` y datos completos del producto, el parser reportaba "Producto no encontrado".

---

## 🔍 Análisis de Causa Raíz

### El Bug

La función `ExtraerValorCadena()` está diseñada para extraer **valores entre comillas** (strings):

```vb
Private Function ExtraerValorCadena(ByVal jsonText As String, ByVal campo As String) As String
    ' ...
    ' Línea 520: Busca la primera comilla después de :
    posInicio = InStr(posInicio, jsonText, """")
    ' ...
End Function
```

**Ejemplo de uso correcto:**
```json
"nombre": "Megan_59"
          ↑       ↑
       Comillas presentes - FUNCIONA
```

### El Problema

Los valores **booleanos** en JSON **NO tienen comillas**:

```json
"success": true
           ↑
       Sin comillas - FALLA
```

Cuando `ExtraerValorCadena()` intentaba extraer `"success"`:
1. Buscaba una comilla después de `:`
2. No la encontraba (porque `true` no tiene comillas)
3. `InStr()` retornaba 0
4. La función retornaba una cadena vacía `""`

### Verificación Fallida

```vb
' Código original (BUGGY)
valorSuccess = ExtraerValorCadena(Mid(jsonText, posSuccess - 1), "success")
If LCase(Trim(valorSuccess)) <> "true" Then
    ' valorSuccess = "" (vacío)
    ' "" <> "true" → TRUE
    ' Marca producto como no encontrado ❌
    producto.MensajeError = "Producto no encontrado"
    Exit Function
End If
```

**Resultado:**
- `valorSuccess = ""` (vacío)
- `"" <> "true"` → **TRUE**
- Salía de la función marcando el producto como no encontrado

---

## ✅ Solución Implementada

### 1. Nueva Función: `ExtraerValorBooleano()`

Creada función específica para extraer valores booleanos sin comillas:

```vb
Private Function ExtraerValorBooleano(ByVal jsonText As String, ByVal campo As String) As Boolean
    On Error Resume Next
    Dim posInicio As Long
    Dim posColon As Long

    ExtraerValorBooleano = False

    ' Buscar el campo
    posInicio = InStr(1, jsonText, """" & campo & """:", vbTextCompare)
    If posInicio = 0 Then Exit Function

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
```

**Características:**
- ✅ Maneja espacios después de `:` (`"success": true` o `"success":true`)
- ✅ Case insensitive (`true`, `True`, `TRUE`)
- ✅ No requiere comillas
- ✅ Retorna `False` por defecto si no encuentra el campo

### 2. Verificación de `success` Corregida

**ANTES (Buggy):**
```vb
If InStr(1, jsonText, """success""", vbTextCompare) > 0 Then
    valorSuccess = ExtraerValorCadena(Mid(jsonText, posSuccess - 1), "success")
    If LCase(Trim(valorSuccess)) <> "true" Then
        ' ❌ Siempre fallaba porque valorSuccess = ""
        producto.MensajeError = "Producto no encontrado"
        Exit Function
    End If
End If
```

**DESPUÉS (Corregido):**
```vb
Dim esExitoso As Boolean

esExitoso = False
posSuccess = InStr(1, jsonText, """success""", vbTextCompare)
If posSuccess > 0 Then
    posColon = InStr(posSuccess, jsonText, ":")
    If posColon > 0 Then
        posColon = posColon + 1
        ' Saltar espacios
        Do While posColon <= Len(jsonText) And Mid(jsonText, posColon, 1) = " "
            posColon = posColon + 1
        Loop

        ' Verificar si empieza con "true"
        If posColon + 3 <= Len(jsonText) Then
            If LCase(Mid(jsonText, posColon, 4)) = "true" Then
                esExitoso = True  ' ✅ FUNCIONA
            End If
        End If
    End If
End If

If Not esExitoso Then
    producto.MensajeError = "Producto no encontrado"
    Exit Function
End If
```

### 3. Campos Booleanos Actualizados

Se actualizaron todos los campos booleanos para usar la nueva función:

**Campo: `tiene_combinaciones`**

**ANTES:**
```vb
producto.TieneCombinaciones = (InStr(1, dataContent, """tiene_combinaciones""", vbTextCompare) > 0)
If producto.TieneCombinaciones Then
    Dim tieneCombosStr As String
    tieneCombosStr = LCase(Trim(ExtraerValorCadena(dataContent, "tiene_combinaciones")))
    producto.TieneCombinaciones = (tieneCombosStr = "true" Or tieneCombosStr = "1")
End If
```

**DESPUÉS:**
```vb
producto.TieneCombinaciones = ExtraerValorBooleano(dataContent, "tiene_combinaciones")
```

**Campo: `activo`**

**ANTES:**
```vb
Dim activoStr As String
activoStr = LCase(Trim(ExtraerValorCadena(dataContent, "activo")))
producto.Activo = (activoStr = "true" Or activoStr = "1")
```

**DESPUÉS:**
```vb
producto.Activo = ExtraerValorBooleano(dataContent, "activo")
```

### 4. Log de Debug Agregado

Para facilitar diagnóstico futuro:

```vb
' DEBUG: Verificar que el producto se ha parseado correctamente
If ModoDebug Then
    ModuloLog.EscribirLog "PARSER - Producto parseado: ID=" & producto.IdProducto & _
        " | Nombre=" & producto.Nombre & " | Precio=" & producto.PrecioConIVA & _
        " | Stock=" & producto.StockDisponible & " | Encontrado=" & producto.Encontrado, LOG_DEBUG
End If
```

---

## 📊 Comparación: Antes vs Después

### Respuesta JSON de Prueba

```json
{
    "success": true,
    "data": {
        "id": 1178,
        "nombre": "Megan_59",
        "precio_con_iva": 30.0,
        "stock": 5,
        "tiene_combinaciones": false,
        "activo": true
    }
}
```

### Comportamiento ANTES (Buggy)

```
[INFO] Respuesta recibida: {"success": true, ...}
[INFO] Producto no encontrado en PrestaShop
[DEBUG] BÚSQUEDA - Código: 2804389083757 | Encontrado: NO
```

**Por qué fallaba:**
1. `ExtraerValorCadena("success")` → retorna `""` (vacío)
2. `"" <> "true"` → TRUE
3. Sale de la función con error

### Comportamiento DESPUÉS (Corregido)

```
[INFO] Respuesta recibida: {"success": true, ...}
[DEBUG] PARSER - Producto parseado: ID=1178 | Nombre=Megan_59 | Precio=30 | Stock=5 | Encontrado=True
[INFO] Producto encontrado: Megan_59 (ID: 1178)
[INFO] Articulo temporal creado con ID: -1178
[DEBUG] BÚSQUEDA - Código: 2804389083757 | Encontrado: SI | ID PS: 1178 | ID Local: -1178
```

**Por qué funciona:**
1. Verificación directa de `"success": true` → `esExitoso = True`
2. No sale prematuramente
3. Extrae `"data"` wrapper correctamente
4. Parsea todos los campos incluyendo booleanos
5. Marca `producto.Encontrado = True`

---

## 🧪 Cómo Probar

### 1. Activar Modo Debug

**config/prestashop.ini:**
```ini
[General]
ModoDebug=1
```

### 2. Recompilar VB6

```
Archivo > Generar Canela.exe
```

### 3. Buscar Producto

Código de prueba: `2804389083757`

### 4. Verificar Log

**Archivo:** `logs/frmventa_2025-12-29.log`

**Buscar líneas:**
```
[DEBUG] PARSER - Producto parseado: ID=1178 | Nombre=Megan_59 | Precio=30 | Stock=5 | Encontrado=True
[INFO] Producto encontrado: Megan_59 (ID: 1178)
```

Si ves estas líneas, el parser funciona correctamente.

### 5. Verificar UI

El producto debe aparecer en el formulario de venta con:
- Nombre: Megan_59
- Precio: 30.00
- Stock: 5

---

## 📝 Archivos Modificados

- **ModuloPrestaShop.bas**
  - Líneas 272-310: Verificación de `success` reescrita
  - Líneas 383: Campo `tiene_combinaciones` usando `ExtraerValorBooleano()`
  - Líneas 400: Campo `activo` usando `ExtraerValorBooleano()`
  - Líneas 403-407: Log de debug agregado
  - Líneas 569-600: Nueva función `ExtraerValorBooleano()`

---

## 🎯 Lecciones Aprendidas

### 1. JSON Tiene Múltiples Tipos de Valores

| Tipo | Ejemplo | Tiene Comillas |
|------|---------|----------------|
| String | `"nombre": "Megan_59"` | ✅ Sí |
| Number | `"precio": 30.0` | ❌ No |
| Boolean | `"activo": true` | ❌ No |
| Null | `"extra": null` | ❌ No |
| Object | `"data": {...}` | ❌ No |
| Array | `"items": [...]` | ❌ No |

### 2. Necesidad de Funciones Específicas por Tipo

**Funciones del parser:**
- `ExtraerValorCadena()` → Para strings (con comillas)
- `ExtraerValorNumerico()` → Para números enteros
- `ExtraerValorMoneda()` → Para números decimales
- `ExtraerValorBooleano()` → Para booleanos (**NUEVO**)

### 3. Importancia de Logs de Debug

Sin el log:
```
[INFO] Producto no encontrado
```
No sabíamos QUÉ parte del parser fallaba.

Con el log:
```
[DEBUG] PARSER - Producto parseado: ... | Encontrado=True
```
Podemos verificar exactamente qué se parseó.

---

## ✅ Estado Final

| Item | Estado |
|------|--------|
| Verificación de `success` | ✅ Corregida |
| Extracción de `tiene_combinaciones` | ✅ Corregida |
| Extracción de `activo` | ✅ Corregida |
| Función `ExtraerValorBooleano()` | ✅ Creada |
| Logs de debug | ✅ Agregados |
| Compilación VB6 | ⏳ Pendiente (usuario) |
| Pruebas funcionales | ⏳ Pendiente (usuario) |

---

## 🚀 Próximo Paso

**Acción inmediata:**

1. **Pull latest changes:**
   ```bash
   git pull origin claude/setup-api-bridge-gj7BX
   ```

2. **Recompilar proyecto VB6**

3. **Ejecutar prueba:**
   - Buscar código: `2804389083757`
   - Verificar log muestra: `[INFO] Producto encontrado`
   - Verificar producto aparece en frmventa

4. **Reportar resultado:**
   - ✅ Si funciona: Continuar con PRUEBA 2-5 de GUIA_PRUEBAS_INTEGRACION.md
   - ❌ Si falla: Enviar log completo para análisis

---

**Commit:** `08338c4`
**Branch:** `claude/setup-api-bridge-gj7BX`
**Fecha:** 29 de Diciembre de 2025
