# 🔍 INSTRUCCIONES DE DEBUG Y PRUEBAS

**Fecha:** 25/12/2025
**Commit:** 61fc70b
**Rama:** `claude/vb6-prestashop-integration-i575X`

---

## ⚠️ PROBLEMA IDENTIFICADO Y CORREGIDO

### El Problema

El usuario reportó que al buscar un producto con combinaciones (tallas), **solo veía el mensaje básico sin las tallas**.

**Mensaje que veía:**
```
Producto encontrado en PrestaShop:

Nombre: Megan 121
Precio: 29,99€
Stock disponible: 3

Codigo: FAC-10064076
```

**Lo que DEBÍA ver (si tiene tallas):**
```
Producto encontrado en PrestaShop:

Nombre: Megan 121
Precio: 29,99€
Stock disponible: 3

Codigo: FAC-10064076

TALLAS DISPONIBLES:
1. S (Stock: 2) ***DISPONIBLE***
2. M (Stock: 1) ***DISPONIBLE***
3. L (Stock: 0) [AGOTADA]
```

### Causas Raíz

1. **Conflictos de merge sin resolver** en `ModuloPrestaShop.bas`
   - Marcadores `<<<<<<< Updated upstream` y `>>>>>>> Stashed changes` en el código
   - Causaban errores de sintaxis

2. **Estructura JSON mal parseada**
   - El API Bridge responde con: `{"success": true, "data": {...}}`
   - El código VB6 buscaba `"tiene_combinaciones"` en el JSON root
   - Pero está anidado dentro de `"data"`
   - Solución: Extraer primero el objeto `"data"` y luego parsear

3. **Problemas de codificación**
   - Símbolo € se mostraba como �
   - Acentos (á, é, ó) causaban problemas en VB6
   - Solución: Usar `Chr(128)` para € y eliminar acentos

---

## 🛠️ CORRECCIONES APLICADAS

### ModuloPrestaShop.bas

**Función `ParsearProductoJSON()` - COMPLETAMENTE REESCRITA:**

```vb
' ANTES (INCORRECTO):
producto.ID = ConvertirALong(ExtraerValorJSON(jsonStr, "id", "number"))
...
producto.TieneCombinaciones = (ExtraerValorJSON(jsonStr, "tiene_combinaciones", "boolean") = "true")
' ^ Buscaba en jsonStr completo, pero "id" y "tiene_combinaciones" están dentro de "data"

' AHORA (CORRECTO):
' 1. Extraer objeto "data" del JSON
Dim dataJSON As String
posDataInicio = InStr(jsonStr, """data"":")
' ... extraer el objeto completo entre { }
dataJSON = Mid(jsonStr, posDataInicio, posDataFin - posDataInicio + 1)

' 2. Parsear campos desde dataJSON
producto.ID = ConvertirALong(ExtraerValorJSON(dataJSON, "id", "number"))
...
producto.TieneCombinaciones = (ExtraerValorJSON(dataJSON, "tiene_combinaciones", "boolean") = "true")

' 3. Pasar dataJSON a ParsearCombinacionesJSON
ParsearCombinacionesJSON dataJSON, producto  ' NO jsonStr completo
```

**Debug Logging Añadido:**
- Muestra primeros 2000 caracteres del JSON recibido
- Indica si se extrajo correctamente el objeto "data"
- Muestra valor raw de `tiene_combinaciones`
- Log completo del proceso de parseo de combinaciones

### frmventa.frm

**Mensaje adaptado al formato del usuario:**
```vb
mensaje = "Producto encontrado en PrestaShop:" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
mensaje = mensaje & "Nombre: " & productoPS.Nombre & vbCrLf & vbCrLf
mensaje = mensaje & "Precio: " & Format(productoPS.PrecioConIVA, "0,00") & Chr(128) & vbCrLf & vbCrLf
mensaje = mensaje & "Stock disponible: " & productoPS.Stock & vbCrLf & vbCrLf & vbCrLf & vbCrLf
mensaje = mensaje & "Codigo: " & productoPS.Reference & vbCrLf
```

**Cambios de codificación:**
- `Chr(128)` para el símbolo € (en lugar de "€" directo)
- `"***DISPONIBLE***"` en lugar de "✓✓✓DISPONIBLE"
- Eliminados todos los acentos: "Seleccion" en lugar de "Selección"

**Debug añadido:**
```vb
Debug.Print ">>> FRMVENTA: TieneCombinaciones=" & productoPS.TieneCombinaciones
Debug.Print ">>> FRMVENTA: NumCombinaciones=" & productoPS.NumCombinaciones
Debug.Print ">>> Talla " & i & ": " & productoPS.Combinaciones(i).Talla & " - Stock: " & productoPS.Combinaciones(i).Stock
```

---

## 🧪 CÓMO PROBAR LAS CORRECCIONES

### Paso 1: Preparación

1. **Activar modo DEBUG en Access:**
   ```
   - Abrir canela.mdb
   - Ir a tabla ConfigAPI
   - Buscar registro con Clave = "DEBUG_MODE"
   - Cambiar Valor a "True"
   - Guardar
   ```

2. **Abrir proyecto VB6:**
   ```
   - Abrir el archivo .vbp del proyecto en Visual Basic 6
   - Presionar F5 para ejecutar en modo DEBUG (NO compilar .exe todavía)
   ```

3. **Abrir ventana Immediate:**
   ```
   - En VB6, presionar Ctrl+G
   - Aparece la ventana "Immediate"
   - Esta ventana mostrará todo el debug logging
   ```

### Paso 2: Probar Producto CON Combinaciones

1. **Identificar un producto con tallas en PrestaShop:**
   - Ir a Admin PrestaShop → Catálogo → Productos
   - Buscar un producto que tenga **Combinaciones** configuradas
   - Anotar la **Referencia** (ej: "FAC-10064076")

2. **Hacer búsqueda en el POS:**
   - Ejecutar la aplicación VB6 (F5)
   - Abrir formulario de venta
   - Buscar el producto por su referencia

3. **Verificar en ventana Immediate:**

   Deberías ver algo como esto:

   ```
   =========================================
   JSON RECIBIDO (primeros 2000 chars):
   {"success":true,"data":{"id":1234,"reference":"FAC-10064076","nombre":"Megan 121","precio_sin_iva":24.79,"precio_con_iva":29.99,"stock":3,"tiene_combinaciones":true,"combinaciones":[{"id_combinacion":456,"id_product_attribute":456,"talla":"S","stock":2,"disponible":true},{"id_combinacion":457,"id_product_attribute":457,"talla":"M","stock":1,"disponible":true},{"id_combinacion":458,"id_product_attribute":458,"talla":"L","stock":0,"disponible":false}]}}
   =========================================
   Objeto 'data' extraido correctamente (425 caracteres)
   Producto parseado OK:
     ID: 1234
     Reference: FAC-10064076
     Nombre: Megan 121
     Precio: 29,99
     Stock: 3
     tiene_combinaciones (raw): 'true'
     TieneCombinaciones (boolean): True
   >>> Intentando parsear combinaciones desde objeto data...
   >>> ParsearCombinacionesJSON: INICIO
   >>> Array combinaciones encontrado en posicion: 156
   Combinaciones encontradas: 3
     Talla 1: S (Stock: 2)
     Talla 2: M (Stock: 1)
     Talla 3: L (Stock: 0)
   >>> Combinaciones parseadas: 3
   >>> FRMVENTA: TieneCombinaciones=True
   >>> FRMVENTA: NumCombinaciones=3
   >>> FRMVENTA: Entrando a rama de combinaciones...
   >>> Talla 1: S - Stock: 2
   >>> Talla 2: M - Stock: 1
   >>> Talla 3: L - Stock: 0
   ```

4. **Verificar mensaje en pantalla:**

   El MsgBox debería mostrar:
   ```
   Producto encontrado en PrestaShop:

   Nombre: Megan 121
   Precio: 29,99€
   Stock disponible: 3

   Codigo: FAC-10064076

   TALLAS DISPONIBLES:
   1. S (Stock: 2) ***DISPONIBLE***
   2. M (Stock: 1) ***DISPONIBLE***
   3. L (Stock: 0) [AGOTADA]
   ```

5. **Seleccionar talla:**
   - Debería aparecer un InputBox pidiendo número de talla
   - Ingresar "1" (para talla S) o "2" (para talla M)
   - El sistema debería buscar en Access: `WHERE codigo='FAC-10064076' AND talla='S'`

### Paso 3: Probar Producto SIN Combinaciones

1. **Identificar producto sin tallas:**
   - Producto que NO tenga combinaciones en PrestaShop

2. **Buscar en POS**

3. **Verificar Immediate Window:**
   ```
   =========================================
   JSON RECIBIDO (primeros 2000 chars):
   {"success":true,"data":{"id":5678,"reference":"BOLSO-001",...,"tiene_combinaciones":false,"combinaciones":[]}}
   =========================================
   Objeto 'data' extraido correctamente (220 caracteres)
   Producto parseado OK:
     ...
     tiene_combinaciones (raw): 'false'
     TieneCombinaciones (boolean): False
   Producto SIN combinaciones (estandar)
   >>> FRMVENTA: TieneCombinaciones=False
   >>> FRMVENTA: NumCombinaciones=0
   ```

4. **Mensaje esperado:**
   ```
   Producto encontrado en PrestaShop:

   Nombre: Bolso de Mano
   Precio: 35,00€
   Stock disponible: 8

   Codigo: BOLSO-001
   ```

   **NO** debería mostrar lista de tallas.
   **SÍ** debería proceder directamente a buscar en Access por código.

---

## 🐛 DIAGNÓSTICO DE PROBLEMAS

### Problema: "No veo nada en Immediate Window"

**Causa:** DEBUG_MODE no está activado

**Solución:**
1. Cerrar aplicación VB6
2. Abrir Access → canela.mdb
3. Tabla ConfigAPI → registro DEBUG_MODE → cambiar a "True"
4. Cerrar Access
5. Volver a ejecutar VB6 (F5)

---

### Problema: "Immediate Window muestra: tiene_combinaciones (raw): ''"

**Causa:** El campo `tiene_combinaciones` no está en el JSON, o no se extrajo correctamente el objeto "data"

**Diagnóstico:**
1. Mirar el JSON recibido (primeras líneas en Immediate)
2. Buscar manualmente `"tiene_combinaciones"` en el texto
3. Verificar que esté dentro de `"data": { ... }`

**Posibles causas:**
- API Bridge desactualizado (no tiene la función `obtenerCombinaciones`)
- PrestaShop API devolvió error
- Producto realmente no tiene combinaciones

**Solución:**
1. Verificar que subiste el `bridge.php` actualizado al servidor
2. Probar URL directa: `https://canelamoda.es/api_bridge/bridge.php?action=buscar_producto&codigo=FAC-10064076`
3. Ver la respuesta JSON en el navegador

---

### Problema: "TieneCombinaciones=True pero NumCombinaciones=0"

**Causa:** El array `"combinaciones"` está vacío o no se pudo parsear

**Diagnóstico:**
Mirar el log:
```
>>> ParsearCombinacionesJSON: INICIO
>>> Array combinaciones encontrado en posicion: 0
>>> ERROR: No se encontro array 'combinaciones' en JSON
```

Si posición es 0, significa que no encontró `"combinaciones": [`

**Solución:**
- Verificar JSON en Immediate Window
- Buscar manualmente la palabra `"combinaciones"`
- Si está vacío: `"combinaciones": []`, el producto no tiene tallas en PrestaShop
- Si está lleno pero posición=0, hay un problema de formato JSON

---

### Problema: "Error al convertir a Currency"

**Ejemplo en Immediate:**
```
Error al convertir a Currency: 29.99
```

**Causa:** El valor tiene punto decimal pero VB6 espera coma

**Estado:** **CORREGIDO** con la función `ConvertirACurrency()` que hace `Replace(".", ",")`

Si sigues viendo este error, verifica que estés usando la versión corregida del código.

---

### Problema: "Símbolo € se ve como �"

**Causa:** Problemas de codificación Windows-1252 vs UTF-8

**Solución:** Usar `Chr(128)` en lugar de "€"

**Estado:** **CORREGIDO** en el commit actual

---

## 📋 CHECKLIST DE VERIFICACIÓN

Antes de reportar un problema, verifica:

- [ ] DEBUG_MODE está en "True" en tabla ConfigAPI
- [ ] Ventana Immediate está abierta (Ctrl+G en VB6)
- [ ] Ejecutando en modo DEBUG (F5), no .exe compilado
- [ ] El producto que pruebas SÍ tiene combinaciones en PrestaShop Admin
- [ ] API Bridge actualizado en servidor (bridge.php con fecha 25/12/2025)
- [ ] URL API sin "www" (https://canelamoda.es/api/, NO www.canelamoda.es)

---

## 📞 INFORMACIÓN PARA REPORTAR PROBLEMAS

Si después de seguir estos pasos el problema persiste, proporciona:

1. **Output completo de Immediate Window** (copiar y pegar todo)
2. **Captura de pantalla** del mensaje que aparece
3. **Código del producto** que estás probando
4. **Verificación en PrestaShop Admin:**
   - ¿El producto tiene combinaciones?
   - ¿Cuántas combinaciones tiene?
   - Captura de pantalla de la pestaña "Combinaciones" del producto

---

## ✅ PRÓXIMOS PASOS SI TODO FUNCIONA

Si las pruebas son exitosas:

1. **Compilar versión final:**
   - En VB6: Archivo → Generar .exe
   - Probar el .exe compilado (debería funcionar igual)

2. **Desactivar DEBUG_MODE:**
   - En Access: ConfigAPI → DEBUG_MODE → cambiar a "False"
   - Esto evita llenar logs innecesariamente en producción

3. **Probar en entorno real:**
   - Realizar ventas reales
   - Verificar que el stock se actualiza correctamente
   - Probar con varios productos diferentes

4. **Documentar productos problema:**
   - Si algún producto específico falla, anotarlo
   - Puede haber casos edge que necesiten atención

---

**¡La detección y visualización de combinaciones debería funcionar correctamente ahora!** 🎉

Con el debug extensivo, podemos identificar exactamente dónde está el problema si algo sigue fallando.
