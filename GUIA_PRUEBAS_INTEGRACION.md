# Guía de Pruebas - Integración PrestaShop

## Fecha: 29 de Diciembre de 2025

---

## Estado Actual del Proyecto

✅ **COMPLETADO:**
- Módulos de integración creados (4 archivos .bas)
- Integración en frmventa.frm
- Error 400 corregido (endpoints y parámetros)
- Parser JSON corregido (wrapper "data" y espacios)
- Sistema de logging implementado
- Configuración INI funcional

⏳ **PENDIENTE:**
- Pruebas funcionales completas
- Verificación de visualización en UI
- Actualización de stock (Fase 2)

---

## Pre-requisitos de Pruebas

### 1. Verificar Archivos Actualizados

Asegurar que tienes la versión más reciente:

```bash
git pull origin claude/setup-api-bridge-gj7BX
```

**Commits requeridos:**
- `6c89a52` - Parser JSON corregido
- `294924d` - Documentación del parser
- `9cf1b3a` - Corrección Error 400

### 2. Compilar Proyecto VB6

1. Abrir **Canela.vbp** en Visual Basic 6
2. Menú: **Archivo > Generar Canela.exe**
3. Verificar que NO hay errores de compilación
4. Cerrar y reabrir VB6 si es necesario

### 3. Verificar Configuración

**Archivo:** `config/prestashop.ini`

```ini
[General]
IntegracionHabilitada=1
BuscarEnPrestaShop=1
LogHabilitado=1
ModoDebug=1  ← IMPORTANTE: Activar debug para pruebas

[API]
URLBridge=https://www.canelamoda.es/api_bridge/
```

Si no existe, el sistema lo creará automáticamente con valores por defecto.

### 4. Preparar Entorno de Logs

```bash
# Crear carpeta de logs si no existe
mkdir logs

# Limpiar logs anteriores (opcional)
rm logs/frmventa_*.log
```

---

## Plan de Pruebas

### PRUEBA 1: Producto Existente en PrestaShop

**Objetivo:** Verificar búsqueda exitosa y visualización en UI

**Código de prueba:** `2804389083757` (Producto: Megan_59, ID: 1178)

#### Pasos:

1. **Ejecutar aplicación VB6**
   - Abrir frmventa

2. **Buscar producto**
   - En campo de búsqueda, introducir: `2804389083757`
   - Presionar Enter o hacer clic en botón de búsqueda

3. **Verificar log en tiempo real**

   Abrir archivo: `logs/frmventa_YYYY-MM-DD.log`

   **Resultado esperado:**
   ```
   [INFO] ===== INICIO DE SESION =====
   [INFO] Configuración cargada correctamente
   [INFO] BUSQUEDA - Codigo solicitado: 2804389083757
   [INFO] Buscando producto en PrestaShop: 2804389083757
   [INFO] URL: https://www.canelamoda.es/api_bridge/bridge.php?action=buscar_producto&codigo=2804389083757
   [INFO] Respuesta recibida: {"success": true, "data": {...}}
   [INFO] Producto encontrado: Megan_59 (ID: 1178)
   [DEBUG] PRODUCTO PS - ID: 1178 | Nombre: Megan_59 | Precio: 30.00 | Stock: 5
   [INFO] Articulo temporal creado con ID: -1178
   [DEBUG] BUSQUEDA - Codigo: 2804389083757 | Encontrado: SI | ID PS: 1178 | ID Local: -1178
   ```

4. **Verificar visualización en frmventa**

   **Campos que deben mostrarse:**
   - **Nombre producto:** Megan_59
   - **Precio:** 30.00
   - **Stock disponible:** 5
   - **Código:** 2804389083757

5. **Verificar BD temporal**

   Abrir `canela_true.mdb` en Access:
   ```sql
   SELECT * FROM articulos WHERE idart < 0
   ```

   **Debe existir:**
   - `idart = -1178`
   - `codigo = "2804389083757"`
   - `tipo = "Megan_59"`
   - `precioventa = 30.00`
   - `vendido = False`

#### Resultado de la Prueba:

| ✓ | Criterio | Estado | Observaciones |
|---|----------|--------|---------------|
| ☐ | Log muestra "Producto encontrado" | | |
| ☐ | Producto visible en frmventa | | |
| ☐ | Precio correcto (30.00) | | |
| ☐ | Stock correcto (5) | | |
| ☐ | Registro temporal creado en BD | | |

---

### PRUEBA 2: Producto NO Existente

**Objetivo:** Verificar fallback a búsqueda local

**Código de prueba:** `9999999999999` (No existe)

#### Pasos:

1. Buscar código: `9999999999999`

2. **Verificar log:**
   ```
   [INFO] Buscando producto en PrestaShop: 9999999999999
   [INFO] URL: https://www.canelamoda.es/api_bridge/bridge.php?action=buscar_producto&codigo=9999999999999
   [INFO] Respuesta recibida: {"success": false, "mensaje": "Producto no encontrado"}
   [INFO] Producto no encontrado en PrestaShop
   [DEBUG] BUSQUEDA - Codigo: 9999999999999 | Encontrado: NO
   [INFO] Buscando en base de datos local...
   ```

3. **Verificar comportamiento:**
   - NO debe mostrar error al usuario
   - Debe buscar automáticamente en BD local
   - Si tampoco existe localmente, mostrar mensaje estándar

#### Resultado de la Prueba:

| ✓ | Criterio | Estado | Observaciones |
|---|----------|--------|---------------|
| ☐ | Log muestra "Producto no encontrado" | | |
| ☐ | NO hay error HTTP 400 | | |
| ☐ | Busca automáticamente en BD local | | |
| ☐ | No bloquea la aplicación | | |

---

### PRUEBA 3: Venta Completa (Sincronización Stock)

**Objetivo:** Verificar flujo completo de venta

**Código de prueba:** `2804389083757`

#### Pasos:

1. **Agregar producto a venta**
   - Buscar código: `2804389083757`
   - Verificar que aparece en lista de venta

2. **Completar venta**
   - Hacer clic en botón "Finalizar Venta" o equivalente
   - Verificar que venta se registra correctamente

3. **Verificar log de sincronización:**

   **COMPORTAMIENTO ACTUAL (Fase 1 - Solo lectura):**
   ```
   [INFO] SYNC STOCK - Iniciando sincronizacion post-venta
   [INFO] SYNC STOCK - Articulos a sincronizar: 1
   [WARNING] Actualizacion de stock aun no implementada en bridge.php (Fase 2)
   [INFO] SYNC STOCK - Articulo temporal eliminado: -1178
   ```

   **COMPORTAMIENTO FUTURO (Fase 2 - Con actualización):**
   ```
   [INFO] SYNC STOCK - Producto PS ID: 1178 | Stock anterior: 5 | Stock nuevo: 4 | Exito: SI
   ```

4. **Verificar limpieza de BD:**
   ```sql
   SELECT * FROM articulos WHERE idart = -1178
   ```
   **Resultado esperado:** 0 registros (eliminado después de venta)

#### Resultado de la Prueba:

| ✓ | Criterio | Estado | Observaciones |
|---|----------|--------|---------------|
| ☐ | Producto se agrega a venta | | |
| ☐ | Venta se completa sin errores | | |
| ☐ | Log registra intento de sync | | |
| ☐ | Artículo temporal eliminado | | |
| ☐ | Venta NO se bloquea por sync | | |

---

### PRUEBA 4: Cancelar Venta

**Objetivo:** Verificar limpieza al cancelar

#### Pasos:

1. Buscar producto: `2804389083757`
2. Agregarlo a venta
3. Hacer clic en botón "Cancelar" o "Borrar"

4. **Verificar log:**
   ```
   [INFO] CANCELAR VENTA - Limpiando articulos temporales
   [INFO] Articulo temporal eliminado: -1178
   ```

5. **Verificar BD:**
   - Artículo temporal debe estar eliminado

#### Resultado de la Prueba:

| ✓ | Criterio | Estado | Observaciones |
|---|----------|--------|---------------|
| ☐ | Venta se cancela correctamente | | |
| ☐ | Artículos temporales eliminados | | |
| ☐ | No quedan registros basura | | |

---

### PRUEBA 5: Error de Conectividad

**Objetivo:** Verificar manejo de errores de red

#### Pasos:

1. **Simular error de red:**
   - Opción 1: Desconectar internet temporalmente
   - Opción 2: Modificar URL en config/prestashop.ini a URL inválida

2. Buscar producto: `2804389083757`

3. **Verificar log:**
   ```
   [ERROR] Error al buscar producto en PrestaShop: [Descripcion del error]
   [INFO] Buscando en base de datos local...
   ```

4. **Verificar comportamiento:**
   - NO debe cerrar la aplicación
   - Debe buscar automáticamente en BD local
   - Usuario puede seguir trabajando normalmente

#### Resultado de la Prueba:

| ✓ | Criterio | Estado | Observaciones |
|---|----------|--------|---------------|
| ☐ | Error registrado en log | | |
| ☐ | NO cierra aplicación | | |
| ☐ | Fallback a BD local funciona | | |
| ☐ | Usuario puede continuar | | |

---

## Diagnóstico de Problemas

### Problema: "Producto no encontrado" (pero debería existir)

**Verificar:**

1. **URL correcta en log:**
   ```
   Debe ser: .../bridge.php?action=buscar_producto&codigo=...
   NO: .../bridge.php?action=search&code=...
   ```

2. **Respuesta del servidor:**
   - Buscar en log: `[INFO] Respuesta recibida:`
   - Copiar JSON completo
   - Verificar que `"success": true`
   - Verificar que existe campo `"data": {...}`

3. **Probar API directamente:**

   Abrir en navegador:
   ```
   https://www.canelamoda.es/api_bridge/bridge.php?action=buscar_producto&codigo=2804389083757
   ```

   Debe retornar JSON válido.

4. **Verificar parser:**
   - Línea en log debe mostrar: `[INFO] Producto encontrado: ...`
   - Si dice "Producto no encontrado" pero JSON tiene datos, hay problema en parser

### Problema: Error HTTP 400

**Solución:**
- Verificar commit `9cf1b3a` está aplicado
- URL debe usar `action=buscar_producto` y parámetro `codigo`

### Problema: Caracteres extraños en log

**Causa:** Problema de codificación

**Solución:**
```bash
# Verificar codificación de archivos .bas
file -i ModuloPrestaShop.bas

# Debe ser: charset=us-ascii o charset=iso-8859-1
# Si es charset=utf-8, convertir:
iconv -f UTF-8 -t WINDOWS-1252 ModuloPrestaShop.bas -o ModuloPrestaShop_fixed.bas
```

### Problema: No aparece en frmventa (pero log dice "encontrado")

**Verificar:**

1. **CrearArticuloDesdePrestaShop se ejecutó:**
   ```
   [INFO] Articulo temporal creado con ID: -XXXX
   ```

2. **SQL de búsqueda incluye ID negativo:**
   - Debug en VB6: Ver valor de variable `SqlArticulos`
   - Debe ser: `... where idart = -1178 ...`

3. **Controles de formulario:**
   - Verificar que campos están habilitados (Enabled=True)
   - Verificar que se llama a `PoneArticulos` después de búsqueda

---

## Checklist Final

### ✅ Antes de Reportar Éxito

- [ ] Prueba 1 completa (producto existente)
- [ ] Prueba 2 completa (producto no existente)
- [ ] Prueba 3 completa (venta completa)
- [ ] Prueba 4 completa (cancelar venta)
- [ ] Prueba 5 completa (error conectividad)
- [ ] Logs sin errores críticos
- [ ] No quedan artículos temporales basura en BD
- [ ] Aplicación NO se cierra inesperadamente

### ✅ Información para Reportar

Al reportar resultados, incluir:

1. **Estado de cada prueba** (tabla con ✓)
2. **Extracto del log** (últimas 50 líneas)
3. **Capturas de pantalla:**
   - frmventa con producto cargado
   - Mensajes de error (si los hay)
4. **Versión del commit:**
   ```bash
   git log --oneline -1
   ```

---

## Próximos Pasos (Fase 2)

Una vez completadas todas las pruebas con éxito:

1. **Implementar actualización de stock en bridge.php**
   - Endpoint: `POST /bridge.php?action=actualizar_stock`
   - Parámetros: `id`, `cantidad`, `operacion`

2. **Habilitar sync en VB6**
   - Descomentar código en ModuloPrestaShop.bas línea 215-216
   - Actualizar formato JSON de petición

3. **Probar con productos con combinaciones**
   - Productos con tallas/colores
   - Verificar actualización de stock por combinación

---

## Ayuda Adicional

### Ver Log en Tiempo Real (Windows)

```cmd
# PowerShell
Get-Content logs\frmventa_2025-12-29.log -Wait -Tail 20

# CMD (actualizar manualmente)
type logs\frmventa_2025-12-29.log
```

### Exportar Log para Análisis

```bash
# Copiar log del día
cp logs/frmventa_$(date +%Y-%m-%d).log log_pruebas.txt
```

### Limpiar Artículos Temporales Manualmente

Si quedan artículos basura en BD:

```sql
-- En Access
DELETE FROM articulos WHERE idart < 0;
```

---

**Última actualización:** 29/12/2025
**Versión de la guía:** 1.0
**Branch:** `claude/setup-api-bridge-gj7BX`
