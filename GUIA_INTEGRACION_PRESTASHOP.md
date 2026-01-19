# Guía de Integración PrestaShop - CanelaPoS

## Fecha: 29 de Diciembre de 2025

---

## 📋 Resumen

Esta guía documenta la integración del sistema POS local (VB6 + Access) con PrestaShop mediante API Bridge.

## 🎯 Objetivos Cumplidos

1. ✅ Búsqueda de productos en PrestaShop por código/EAN
2. ✅ Detección automática de productos con combinaciones (tallas)
3. ✅ Mapeo automático de datos PrestaShop → BD Local
4. ✅ Actualización automática de stock después de venta
5. ✅ Sistema de logging para depuración
6. ✅ Configuración flexible mediante archivo INI

---

## 📁 Archivos Creados

### Módulos VB6

1. **ModuloPrestaShop.bas**
   - Comunicación con API Bridge
   - Funciones de búsqueda y actualización de stock
   - Parseo de JSON (implementación simplificada)

2. **ModuloLog.bas**
   - Sistema de logging a archivo
   - Logs rotativos por fecha
   - Diferentes niveles: INFO, WARNING, ERROR, DEBUG

3. **ModuloConfig.bas**
   - Gestión de configuración mediante archivo INI
   - Lectura/escritura de parámetros
   - Interfaz de Windows API para archivos INI

4. **ModuloIntegracion.bas**
   - Orquestación entre POS local y PrestaShop
   - Gestión de artículos temporales
   - Sincronización de stock post-venta

### Archivos de Configuración

- **config/prestashop.ini** - Configuración de la integración (se crea automáticamente)
- **logs/prestashop_YYYYMMDD.log** - Logs diarios (se crean automáticamente)

---

## 🔧 Modificaciones Necesarias en frmventa.frm

### 1. Declaraciones en la sección General

```vb
Option Explicit

' NOTA: Agregar al inicio del módulo de frmventa.frm
' Variable para tracking de productos de PrestaShop
Private productoPrestaShop As Boolean
```

### 2. Modificar Form_Load

```vb
Private Sub Form_Load()
    ' ... código existente ...

    ' NUEVO: Inicializar integración PrestaShop
    InicializarIntegracion

    ' Resto del código existente
End Sub
```

### 3. Modificar cmdarticulo_Click

**UBICACIÓN:** Línea 1185 aproximadamente

**MODIFICACIÓN:** Agregar búsqueda en PrestaShop antes de buscar en BD local

```vb
Private Sub cmdarticulo_Click()
    On Error GoTo sehodio

    Dim idArtPrestaShop As Long

    ModoBusca = "articulos"
    If CodigoBusca = "" Then CodigoBusca = InputBox("Escriba el código")

    If CodigoBusca <> "" Then
        ' ===== NUEVO: Intentar buscar en PrestaShop primero =====
        idArtPrestaShop = BuscarProductoPrestaShop(CodigoBusca)

        If idArtPrestaShop <> 0 Then
            ' Producto encontrado en PrestaShop y agregado a BD local
            ' Buscar el artículo recién creado
            SqlArticulos = "Select idart, codigo, tipo, precioventa, " _
            & " color, talla, extra from articulos where " _
            & " idart = " & idArtPrestaShop & " order by codigo"
        Else
            ' Si no está en PrestaShop, buscar en BD local (comportamiento original)
            SqlArticulos = "Select idart, codigo, tipo, precioventa, " _
            & " color, talla, extra from articulos where vendido = false and apartado = false and" _
            & " idart = " & CodigoBusca & " order by codigo"
        End If
        ' ===== FIN NUEVO =====
    Else
        CodigoBusca = InputBox("Escriba algún dato para buscar")
        SqlArticulos = "Select idart, codigo, tipo, precioventa, color, talla, extra " _
        & "from articulos where vendido = false and apartado = false and(codigo " _
        & "like '*" & CodigoBusca & "*' or precioventa like '*" & CodigoBusca & "*' or " _
        & "talla like '*" & CodigoBusca & "*' or tipo like '*" & CodigoBusca & "*') order by codigo"
    End If

    Set RsArticulo = bdtienda.OpenRecordset(SqlArticulos)
    If RsArticulo.EOF Then
        CodigoBusca = ""
        Exit Sub
    End If
    RsArticulo.MoveLast
    If RsArticulo.RecordCount > 1 Then
        frmarticulos.Show
    Else
        NumArtVend = NumArtVend + 1
        PoneArticulos
    End If
    CodigoBusca = ""
    Exit Sub

sehodio:
    MsgBox ("No se han encontrado datos")
End Sub
```

### 4. Modificar MarcaVenta (para sincronización de stock)

**UBICACIÓN:** Después de la línea que llama a `MarcaVendido` (aprox. línea 1771)

**MODIFICACIÓN:** Agregar sincronización de stock después de marcar vendido

```vb
Private Sub MarcaVenta()
    ' ... todo el código existente hasta ...

    MarcaVendido

    ' ===== NUEVO: Sincronizar stock con PrestaShop =====
    SincronizarStockVendido
    ' ===== FIN NUEVO =====

    CmbBorraArt_Click
    cmdBorrar_Click
    ' ... resto del código existente ...
End Sub
```

### 5. Modificar cmdBorrar_Click (para cancelación de venta)

**UBICACIÓN:** Línea 1229 aproximadamente

**MODIFICACIÓN:** Cancelar sincronización si se borran datos

```vb
Private Sub cmdBorrar_Click()
    ' ... código existente ...

    ' ===== NUEVO: Cancelar venta en PrestaShop si había artículos =====
    CancelarVenta
    ' ===== FIN NUEVO =====

    ' ... resto del código existente ...
End Sub
```

### 6. Modificar Form_Unload (opcional - limpieza)

```vb
Private Sub Form_Unload(Cancel As Integer)
    ' Código existente (si hay)

    ' NUEVO: Finalizar integración
    FinalizarIntegracion
End Sub
```

---

## 🎯 Cómo Funciona la Integración

### Flujo de Búsqueda de Producto

```
Usuario escanea código
        ↓
TxtBusca_KeyPress (Enter)
        ↓
cmdarticulo_Click
        ↓
BuscarProductoPrestaShop(codigo)
        ↓
    ¿Encontrado en PS?
    ├─ SÍ → Crear artículo temporal en BD local
    │        ID negativo para identificarlo
    │        Registrar para sincronización
    │        Mostrar producto
    │
    └─ NO → Buscar en BD local (comportamiento normal)
              Continuar venta sin sincronización
```

### Flujo de Venta Completada

```
Usuario completa venta
        ↓
MarcaVenta
        ↓
MarcaVendido (marca vendido en BD local)
        ↓
SincronizarStockVendido
        ↓
    Para cada artículo de PrestaShop:
        - Llamar API Bridge para decrementar stock
        - Registrar en log
        - Eliminar artículo temporal (ID negativo)
```

---

## ⚙️ Configuración

### Archivo: config/prestashop.ini

```ini
[General]
; Habilita/deshabilita toda la integración (1=Sí, 0=No)
IntegracionHabilitada=1

; Buscar productos en PrestaShop al escanear código (1=Sí, 0=No)
BuscarEnPrestaShop=1

; Actualizar stock automáticamente después de venta (1=Sí, 0=No)
ActualizarStockAutomatico=1

; Mostrar mensajes de error al usuario (1=Sí, 0=No)
; Recomendado: 0 (los errores se registran en el log)
MostrarMensajesError=0

; Timeout en segundos para llamadas API
TimeoutSegundos=30

; Habilitar logging de operaciones (1=Sí, 0=No)
LogHabilitado=1

; Modo debug - registra información detallada (1=Sí, 0=No)
ModoDebug=0

[API]
; URL del API Bridge (NO CAMBIAR sin autorización)
URLBridge=https://www.canelamoda.es/api_bridge/
```

### Editar Configuración

Desde VB6:
```vb
ModuloConfig.EditarConfiguracion  ' Abre el INI en Notepad
ModuloConfig.MostrarConfiguracion ' Muestra config actual
```

---

## 📊 Sistema de Logging

### Ubicación de Logs

- Carpeta: `[App.Path]\logs\`
- Formato: `prestashop_YYYYMMDD.log`
- Rotación: Diaria (se crea un archivo nuevo cada día)
- Retención: 30 días (los logs más antiguos se eliminan automáticamente)

### Ver Logs

Desde VB6:
```vb
ModuloLog.MostrarLog  ' Abre el log actual en Notepad
```

### Ejemplo de Log

```
[2025-12-29 14:23:15] [INFO] Sistema de integración PrestaShop iniciado
[2025-12-29 14:23:45] [INFO] BÚSQUEDA - Código: 12345 | Encontrado: SÍ | ID PS: 789
[2025-12-29 14:24:10] [INFO] Artículo creado desde PrestaShop - ID Local: -7890001
[2025-12-29 14:25:30] [INFO] SYNC STOCK - Producto PS ID: 789 | Stock anterior: 5 | Stock nuevo: 4 | Éxito: SÍ
```

---

## 🔍 Detalles Técnicos

### Productos con Combinaciones

PrestaShop maneja dos tipos de productos:
1. **Simples:** Stock único para el producto
2. **Con combinaciones:** Stock separado por cada combinación (ej: tallas)

La integración detecta automáticamente el tipo y maneja correctamente ambos casos.

### Identificación de Artículos Temporales

Los artículos creados desde PrestaShop tienen **ID negativos**:
- Cálculo: `-(IdProductoPS * 10000 + IdCombinacion)`
- Ejemplo: Producto PS #789, Combinación #12 → ID local: -7890012
- Esto evita conflictos con IDs reales de la BD local
- Se eliminan automáticamente después de sincronizar stock

### Manejo de Errores

La integración está diseñada para **no interrumpir** el flujo normal de venta:
- Si la API falla → Se continúa con venta local sin sincronización
- Si timeout → Se registra en log pero no se muestra error al usuario
- Si producto no existe → Se busca en BD local normalmente

---

## 🧪 Pruebas Recomendadas

### 1. Prueba de Búsqueda
- Escanear un código que SÍ exista en PrestaShop
- Verificar que el producto se muestra correctamente
- Verificar precio, nombre, stock en pantalla

### 2. Prueba de Venta Completa
- Escanear producto de PrestaShop
- Completar venta normalmente
- Verificar en log que stock se actualizó
- Verificar en PrestaShop admin que stock decrementó

### 3. Prueba de Producto No Encontrado
- Escanear código que NO existe en PrestaShop
- Verificar que continúa búsqueda normal en BD local
- No debe mostrar errores al usuario

### 4. Prueba de Conexión Fallida
- Desactivar internet temporalmente
- Intentar escanear producto
- Verificar que venta local funciona normalmente
- Verificar registro en log del error de conexión

### 5. Prueba de Configuración
- Desactivar IntegracionHabilitada en INI
- Verificar que sistema funciona 100% local
- Reactivar y verificar que vuelve a funcionar

---

## 📝 API Bridge - Endpoints Esperados

### Búsqueda de Producto
```
GET /bridge.php?action=search&code={codigo}

Respuesta esperada:
{
  "success": true,
  "found": true,
  "id_product": 123,
  "reference": "ABC123",
  "ean13": "1234567890123",
  "name": "Nombre del producto",
  "price": 25.00,
  "price_with_tax": 30.25,
  "tax_rate": 21,
  "quantity": 10,
  "has_combinations": false,
  "active": "1"
}
```

### Obtener Stock
```
GET /bridge.php?action=stock&product_id={id}&combination_id={id}

Respuesta esperada:
{
  "success": true,
  "quantity": 10
}
```

### Actualizar Stock
```
POST /bridge.php?action=update_stock
Content-Type: application/json

{
  "product_id": 123,
  "quantity": 1,
  "operation": "decrease",
  "combination_id": 0
}

Respuesta esperada:
{
  "success": true,
  "old_stock": 10,
  "new_stock": 9
}
```

---

## 🐛 Resolución de Problemas

### Problema: No encuentra productos en PrestaShop

**Soluciones:**
1. Verificar que `IntegracionHabilitada=1` en el INI
2. Verificar que `BuscarEnPrestaShop=1` en el INI
3. Revisar el log para ver si hay errores de conexión
4. Verificar URL del API Bridge en configuración
5. Probar el test_bridge.html en navegador

### Problema: Stock no se actualiza en PrestaShop

**Soluciones:**
1. Verificar que `ActualizarStockAutomatico=1` en el INI
2. Revisar el log - buscar líneas "SYNC STOCK"
3. Verificar que el producto tenga ID válido en PrestaShop
4. Comprobar permisos de la API Key en PrestaShop

### Problema: Errores de timeout

**Soluciones:**
1. Aumentar `TimeoutSegundos` en el INI (probar con 60)
2. Verificar velocidad de conexión a internet
3. Verificar que el servidor PrestaShop responde rápido

### Problema: Logs no se crean

**Soluciones:**
1. Verificar que `LogHabilitado=1` en el INI
2. Verificar permisos de escritura en carpeta de aplicación
3. Crear carpeta `logs` manualmente si no existe

---

## 📞 Soporte

Para problemas con la integración:
1. Revisar siempre el archivo de log primero
2. Habilitar `ModoDebug=1` para información detallada
3. Verificar configuración en prestashop.ini
4. Probar el API Bridge directamente en test_bridge.html

---

## 📌 Notas Importantes

- ⚠️ La API Key debe estar configurada en el servidor (api_bridge.php)
- ⚠️ No compartir archivos de log (pueden contener información sensible)
- ⚠️ Mantener backups de la BD local antes de grandes cambios
- ⚠️ Los artículos temporales (ID negativo) no deben editarse manualmente
- ✅ La integración funciona en modo "fail-safe" - nunca bloquea ventas

---

**Fin del documento**
