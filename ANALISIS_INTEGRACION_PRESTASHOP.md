# Análisis de Integración VB6 - PrestaShop 8.1
## Sistema POS Canela Moda

**Fecha:** 19 de diciembre de 2025
**Analista:** Claude Code
**Objetivo:** Integrar POS VB6 con PrestaShop 8.1 para sincronización de inventario

---

## RESPUESTAS A TUS PREGUNTAS

### 1️⃣ ¿Qué archivos del proyecto son críticos para esta integración?

#### **Archivos CRÍTICOS (Modificación requerida):**

| Archivo | Líneas | Función | Modificaciones necesarias |
|---------|--------|---------|---------------------------|
| **frmventa.frm** | 2220 | TPV principal | Búsqueda de productos (línea 1185-1199), Inserción de ventas (línea 1680), Actualización de stock post-venta |
| **Module1.bas** | 597 | Variables globales y funciones core | Agregar funciones de API PrestaShop, tipos de datos para productos PS |
| **frmelige.frm** | 939+ | Menú principal | Inicialización de conexión API en Form_Load (línea 939) |
| **frmaddart.frm** | 724+ | Alta de artículos | Sincronizar nuevos productos con PrestaShop (línea 710-724) |
| **frmaddart2.frm** | 748+ | Alta de artículos v2 | Sincronizar nuevos productos con PrestaShop (línea 748) |
| **FrmInventario.frm** | - | Gestión de stock | Actualización masiva de stock a PrestaShop |

#### **Archivos de SOPORTE (Lectura necesaria):**

- `canela.mdb` - Base de datos Access (54,036 productos)
- `estructura_bd_20251219_181525.md` - Esquema de BD
- `claude.md` - Notas sobre PrestaShop (reference, ean-13)

#### **Archivos NUEVOS a crear:**

- `ModuloPrestaShop.bas` - Módulo de comunicación con API PrestaShop
- `ClsProductoPS.cls` - Clase para mapear productos PrestaShop
- `ConfigAPI.ini` - Archivo de configuración (API Key, URL)

---

### 2️⃣ ¿Hay limitaciones técnicas en VB6 para llamadas HTTPS?

#### **SÍ, hay limitaciones importantes:**

##### **Problema 1: Certificados SSL/TLS modernos**

VB6 fue diseñado en 1998, antes de TLS 1.2/1.3:

- **MSXML2.XMLHTTP**: Solo soporta SSL 3.0/TLS 1.0 por defecto
- **WinHTTP.WinHttpRequest.5.1**: ✅ **RECOMENDADO** - Mejor soporte TLS 1.2
- PrestaShop moderno requiere TLS 1.2+ para HTTPS

**Solución:**
```vb
' En Module1.bas o ModuloPrestaShop.bas
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

' WinHTTP soporta TLS 1.2 en Windows 7+ con actualizaciones
http.SetTimeouts 5000, 10000, 30000, 60000  ' Timeouts personalizados
```

##### **Problema 2: Sistema Operativo**

- **Windows XP/Vista**: NO soportan TLS 1.2 nativamente
- **Windows 7+**: Requiere KB3140245 (parche TLS 1.2)
- **Windows 10/11**: ✅ Sin problemas

**Tu entorno (Windows 11)**: ✅ **Compatible sin problemas**

##### **Problema 3: Parseo de XML**

PrestaShop devuelve XML complejo:
```xml
<prestashop>
  <product>
    <id>123</id>
    <reference><![CDATA[ABC-12345678]]></reference>
    <name><language id="1"><![CDATA[Vestido Rojo]]></language></name>
    <price>45.00</price>
    <stock_available>
      <id>456</id>
      <quantity>5</quantity>
    </stock_available>
  </product>
</prestashop>
```

VB6 puede parsearlo con `MSXML2.DOMDocument`, pero es verboso y propenso a errores.

##### **Problema 4: Autenticación Basic Auth**

PrestaShop usa Basic Auth con API Key de 32 caracteres:
```
Authorization: Basic BASE64(API_KEY:)
```

VB6 NO tiene función Base64 nativa, hay que implementarla o usar servicio intermediario.

#### **📊 Comparación de opciones para HTTPS en VB6:**

| Método | TLS 1.2 | Base64 | XML Parse | Dificultad | Recomendación |
|--------|---------|--------|-----------|------------|---------------|
| **WinHTTP** | ✅ Sí | ❌ No | ⚠️ Complejo | Alta | ⭐⭐⭐ Viable con trabajo |
| **MSXML2** | ⚠️ Limitado | ❌ No | ⚠️ Complejo | Alta | ⭐⭐ No recomendado |
| **API Bridge PHP/Python** | ✅ Sí | ✅ Sí | ✅ Automático | Baja | ⭐⭐⭐⭐⭐ **ÓPTIMO** |

---

### 3️⃣ ¿Recomiendas crear un servicio intermediario o conectar directamente desde VB6?

## 🏆 **RECOMENDACIÓN: Servicio Intermediario (API Bridge)**

### **Arquitectura Propuesta:**

```
┌──────────────────────────────────────────────────────────────┐
│                     POS VB6 (Cliente)                        │
│  ┌────────────┐  ┌──────────────┐  ┌──────────────┐         │
│  │ frmventa   │  │ frmaddart    │  │ FrmInventario│         │
│  │ (TPV)      │  │ (Productos)  │  │ (Stock)      │         │
│  └──────┬─────┘  └──────┬───────┘  └──────┬───────┘         │
│         │                │                  │                 │
│         └────────────────┼──────────────────┘                 │
│                          ▼                                    │
│              ┌────────────────────────┐                       │
│              │ ModuloPrestaShop.bas   │                       │
│              │ (HTTP Client VB6)      │                       │
│              └───────────┬────────────┘                       │
│                          │                                    │
│                          │ HTTP JSON                          │
│                          │ (Simple)                           │
└──────────────────────────┼────────────────────────────────────┘
                           │
                           ▼
            ┌──────────────────────────────┐
            │   API BRIDGE (Intermediario) │
            │   - PHP 7.4+ ó Python 3.x    │
            │   - Localhost o servidor     │
            │   - Puerto 8080              │
            │                              │
            │   Endpoints:                 │
            │   GET  /producto/{codigo}    │
            │   POST /stock/actualizar     │
            │   GET  /stock/{id}           │
            └──────────────┬───────────────┘
                           │
                           │ HTTPS + XML
                           │ Basic Auth
                           │ TLS 1.2+
                           ▼
            ┌──────────────────────────────┐
            │   PRESTASHOP 8.1 (API)       │
            │   https://canelamoda.es/api/ │
            │                              │
            │   Recursos:                  │
            │   - /products                │
            │   - /stock_availables        │
            └──────────────────────────────┘
```

### **✅ Ventajas del API Bridge:**

1. **Simplicidad en VB6:**
   - VB6 hace peticiones HTTP simples a `localhost:8080`
   - Recibe JSON simple: `{"nombre":"Vestido","precio":45.50,"stock":5}`
   - No necesita parsear XML complejo de PrestaShop
   - No necesita implementar Base64 para autenticación

2. **Caché local:**
   - El bridge puede cachear productos consultados frecuentemente
   - Reduce latencia: 50ms vs 500ms (directo a PrestaShop)
   - Funciona offline: devuelve datos cacheados si PrestaShop no responde

3. **Logging y debugging:**
   - Todas las peticiones se registran en archivo
   - Fácil identificar errores de API
   - Monitoreo de sincronizaciones

4. **Seguridad:**
   - API Key de PrestaShop NO está en código VB6
   - Bridge maneja autenticación
   - Puede agregar capa de autenticación adicional

5. **Mantenimiento:**
   - Cambios en API de PrestaShop: solo actualizar bridge
   - No recompilar VB6
   - Testeable independientemente

6. **Cola de sincronización:**
   - Si PrestaShop cae, bridge encola actualizaciones
   - Reintenta automáticamente
   - VB6 no se bloquea esperando respuesta

### **❌ Desventajas (mínimas):**

1. Componente adicional a desplegar (pero simple: 1 archivo PHP/Python)
2. Requiere servidor web local (IIS, Apache, o Python simple HTTP server)

### **Implementación rápida del Bridge (PHP):**

```php
<?php
// api_bridge.php - Servidor intermediario

header('Content-Type: application/json');

$PRESTASHOP_URL = 'https://www.canelamoda.es/api/';
$API_KEY = file_get_contents('api_key.txt'); // API Key en archivo separado

function callPrestaShop($resource, $method = 'GET', $data = null) {
    global $PRESTASHOP_URL, $API_KEY;

    $ch = curl_init();
    $url = $PRESTASHOP_URL . $resource;

    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_USERPWD, $API_KEY . ':');
    curl_setopt($ch, CURLOPT_TIMEOUT, 30);

    if ($method == 'POST' || $method == 'PUT') {
        curl_setopt($ch, CURLOPT_CUSTOMREQUEST, $method);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $data);
    }

    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    return ['code' => $httpCode, 'data' => $response];
}

// Endpoint: Buscar producto por código
if ($_GET['action'] == 'buscar_producto') {
    $codigo = $_GET['codigo'];
    $result = callPrestaShop("products?filter[reference]=$codigo&display=full");

    if ($result['code'] == 200) {
        $xml = simplexml_load_string($result['data']);
        $product = $xml->products->product;

        echo json_encode([
            'success' => true,
            'id' => (int)$product->id,
            'nombre' => (string)$product->name->language,
            'precio' => (float)$product->price,
            'stock' => (int)$product->stock_available->quantity
        ]);
    } else {
        echo json_encode(['success' => false, 'error' => 'Producto no encontrado']);
    }
}

// Endpoint: Actualizar stock
if ($_GET['action'] == 'actualizar_stock') {
    $id_producto = $_POST['id'];
    $cantidad = $_POST['cantidad'];

    // Lógica de actualización...
    echo json_encode(['success' => true]);
}
?>
```

---

### 4️⃣ ¿Qué estructura de tabla adicional necesitaría en Access para caché/sincronización?

## 📊 **Tablas Nuevas Necesarias:**

### **Tabla 1: ConfigAPI**
Almacena configuración de la API (sin hardcodear en código).

```sql
CREATE TABLE ConfigAPI (
    Clave VARCHAR(50) PRIMARY KEY,
    Valor TEXT,
    FechaModificacion DATETIME DEFAULT Now()
);

-- Datos iniciales
INSERT INTO ConfigAPI (Clave, Valor) VALUES
    ('API_URL', 'http://localhost:8080'),
    ('API_TIMEOUT', '30'),
    ('SYNC_ENABLED', 'True'),
    ('LAST_SYNC', ''),
    ('DEBUG_MODE', 'False');
```

**Acceso desde VB6:**
```vb
Function GetConfig(Clave As String) As String
    Dim rs As Recordset
    Set rs = bdtienda.OpenRecordset("SELECT Valor FROM ConfigAPI WHERE Clave='" & Clave & "'")
    If Not rs.EOF Then GetConfig = rs!Valor
    rs.Close
End Function
```

---

### **Tabla 2: ProductosPS (Caché de productos)**
Almacena datos de PrestaShop localmente para funcionamiento offline.

```sql
CREATE TABLE ProductosPS (
    IDProductoPS LONG PRIMARY KEY,           -- ID de PrestaShop
    Referencia VARCHAR(50) UNIQUE NOT NULL,  -- Código del producto (reference)
    EAN13 VARCHAR(13),                       -- Código de barras
    Nombre VARCHAR(255),                     -- Nombre del producto
    Descripcion TEXT,                        -- Descripción corta
    PrecioSinIVA CURRENCY,                   -- Precio base
    PrecioConIVA CURRENCY,                   -- Precio de venta
    IVA INTEGER DEFAULT 21,                  -- % IVA aplicado
    StockPS LONG DEFAULT 0,                  -- Stock en PrestaShop
    StockLocal LONG DEFAULT 0,               -- Stock en Access (calculado)
    DiferenciaStock LONG DEFAULT 0,          -- Diferencia a sincronizar
    UltimaConsulta DATETIME,                 -- Última vez que se consultó PS
    UltimaActualizacion DATETIME,            -- Última actualización enviada a PS
    EstadoSync VARCHAR(20) DEFAULT 'OK',     -- 'OK', 'PENDIENTE', 'ERROR', 'CONFLICTO'
    URLImagen VARCHAR(255),                  -- URL de la imagen web
    Activo BIT DEFAULT 1                     -- Si está activo en PrestaShop
);

CREATE INDEX idx_referencia ON ProductosPS(Referencia);
CREATE INDEX idx_estado ON ProductosPS(EstadoSync);
```

**Relación con tabla `articulos`:**
- El campo `articulos.codigo` se mapea a `ProductosPS.Referencia`
- El campo `articulos.idarticulo` puede mapearse a `ProductosPS.EAN13`

---

### **Tabla 3: LogSincronizacion (Auditoría)**
Registra todas las operaciones con PrestaShop para debugging.

```sql
CREATE TABLE LogSincronizacion (
    ID AUTOINCREMENT PRIMARY KEY,
    FechaHora DATETIME DEFAULT Now(),
    TipoOperacion VARCHAR(50),               -- 'BUSQUEDA', 'UPDATE_STOCK', 'GET_STOCK', 'ERROR'
    IDProductoPS LONG,                       -- ID del producto afectado
    Referencia VARCHAR(50),                  -- Código del producto
    Descripcion TEXT,                        -- Descripción de la operación
    RespuestaAPI TEXT,                       -- Respuesta completa de la API
    CodigoHTTP INTEGER,                      -- 200, 404, 500, etc.
    TiempoRespuesta INTEGER,                 -- Milisegundos
    UsuarioVB VARCHAR(50) DEFAULT Environ$('USERNAME')
);

CREATE INDEX idx_fecha ON LogSincronizacion(FechaHora);
CREATE INDEX idx_tipo ON LogSincronizacion(TipoOperacion);
```

**Ejemplo de log:**
```vb
Sub LogAPI(TipoOp As String, IDProd As Long, Descrip As String, _
           RespAPI As String, HttpCode As Integer, TiempoMs As Integer)
    Dim rs As Recordset
    Set rs = bdtienda.OpenRecordset("LogSincronizacion")
    With rs
        .AddNew
        !TipoOperacion = TipoOp
        !IDProductoPS = IDProd
        !Descripcion = Descrip
        !RespuestaAPI = Left(RespAPI, 65000)  ' Memo max size
        !CodigoHTTP = HttpCode
        !TiempoRespuesta = TiempoMs
        .Update
    End With
End Sub
```

---

### **Tabla 4: ColaSyncStock (Cola de sincronización offline)**
Cuando no hay conexión, las ventas se encolan para sincronizar después.

```sql
CREATE TABLE ColaSyncStock (
    ID AUTOINCREMENT PRIMARY KEY,
    IDVenta LONG,                            -- Relacionado con tabla venta
    IDProductoPS LONG,                       -- ID en PrestaShop
    Referencia VARCHAR(50),                  -- Código del producto
    CantidadVendida INTEGER DEFAULT 1,       -- Unidades vendidas
    FechaVenta DATETIME,                     -- Fecha de la venta
    Procesado BIT DEFAULT 0,                 -- ¿Ya se sincronizó?
    FechaProcesado DATETIME,                 -- Cuándo se sincronizó
    Reintentos INTEGER DEFAULT 0,            -- Número de intentos fallidos
    ErrorMensaje TEXT                        -- Mensaje de error si falló
);

CREATE INDEX idx_procesado ON ColaSyncStock(Procesado);
CREATE INDEX idx_fecha ON ColaSyncStock(FechaVenta);
```

**Flujo de uso:**
1. **Venta completada** → Se intenta actualizar stock en PrestaShop
2. **Si falla** (sin conexión) → Se inserta en `ColaSyncStock`
3. **Proceso batch** → Cada X minutos, procesa cola pendiente
4. **Al completarse** → `Procesado = True`, `FechaProcesado = Now()`

---

### **Tabla 5: MapeoArticulosPS (Mapeo entre tablas)**
Relaciona artículos locales con productos de PrestaShop.

```sql
CREATE TABLE MapeoArticulosPS (
    IDArticuloLocal LONG PRIMARY KEY,        -- articulos.idart
    IDProductoPS LONG,                       -- ProductosPS.IDProductoPS
    Referencia VARCHAR(50),                  -- Código común
    FechaMapeo DATETIME DEFAULT Now(),
    MapeadoPor VARCHAR(50) DEFAULT Environ$('USERNAME')
);
```

**Uso:**
- Al buscar un producto por código, se registra la asociación
- Permite consultas rápidas sin reconsultar PrestaShop

---

## 📋 **Script SQL Completo para Crear Tablas:**

```sql
-- ConfigAPI
CREATE TABLE ConfigAPI (
    Clave VARCHAR(50) PRIMARY KEY,
    Valor TEXT,
    FechaModificacion DATETIME DEFAULT Now()
);

-- ProductosPS
CREATE TABLE ProductosPS (
    IDProductoPS LONG PRIMARY KEY,
    Referencia VARCHAR(50) UNIQUE NOT NULL,
    EAN13 VARCHAR(13),
    Nombre VARCHAR(255),
    Descripcion TEXT,
    PrecioSinIVA CURRENCY,
    PrecioConIVA CURRENCY,
    IVA INTEGER DEFAULT 21,
    StockPS LONG DEFAULT 0,
    StockLocal LONG DEFAULT 0,
    DiferenciaStock LONG DEFAULT 0,
    UltimaConsulta DATETIME,
    UltimaActualizacion DATETIME,
    EstadoSync VARCHAR(20) DEFAULT 'OK',
    URLImagen VARCHAR(255),
    Activo BIT DEFAULT 1
);

-- Índices ProductosPS
CREATE INDEX idx_referencia ON ProductosPS(Referencia);
CREATE INDEX idx_estado ON ProductosPS(EstadoSync);

-- LogSincronizacion
CREATE TABLE LogSincronizacion (
    ID AUTOINCREMENT PRIMARY KEY,
    FechaHora DATETIME DEFAULT Now(),
    TipoOperacion VARCHAR(50),
    IDProductoPS LONG,
    Referencia VARCHAR(50),
    Descripcion TEXT,
    RespuestaAPI TEXT,
    CodigoHTTP INTEGER,
    TiempoRespuesta INTEGER,
    UsuarioVB VARCHAR(50)
);

-- Índices LogSincronizacion
CREATE INDEX idx_fecha ON LogSincronizacion(FechaHora);
CREATE INDEX idx_tipo ON LogSincronizacion(TipoOperacion);

-- ColaSyncStock
CREATE TABLE ColaSyncStock (
    ID AUTOINCREMENT PRIMARY KEY,
    IDVenta LONG,
    IDProductoPS LONG,
    Referencia VARCHAR(50),
    CantidadVendida INTEGER DEFAULT 1,
    FechaVenta DATETIME,
    Procesado BIT DEFAULT 0,
    FechaProcesado DATETIME,
    Reintentos INTEGER DEFAULT 0,
    ErrorMensaje TEXT
);

-- Índices ColaSyncStock
CREATE INDEX idx_procesado ON ColaSyncStock(Procesado);
CREATE INDEX idx_fecha_venta ON ColaSyncStock(FechaVenta);

-- MapeoArticulosPS
CREATE TABLE MapeoArticulosPS (
    IDArticuloLocal LONG PRIMARY KEY,
    IDProductoPS LONG,
    Referencia VARCHAR(50),
    FechaMapeo DATETIME DEFAULT Now(),
    MapeadoPor VARCHAR(50)
);

-- Datos iniciales ConfigAPI
INSERT INTO ConfigAPI (Clave, Valor) VALUES ('API_URL', 'http://localhost:8080');
INSERT INTO ConfigAPI (Clave, Valor) VALUES ('API_TIMEOUT', '30');
INSERT INTO ConfigAPI (Clave, Valor) VALUES ('SYNC_ENABLED', 'True');
INSERT INTO ConfigAPI (Clave, Valor) VALUES ('DEBUG_MODE', 'False');
```

---

## 🔄 **Cómo Usar el Caché en VB6:**

### **Ejemplo 1: Buscar producto (con caché)**
```vb
Function BuscarProductoPorCodigo(codigo As String) As ProductoPS
    Dim rs As Recordset
    Dim producto As ProductoPS

    ' Primero buscar en caché local
    Set rs = bdtienda.OpenRecordset("SELECT * FROM ProductosPS WHERE Referencia='" & codigo & "'")

    If Not rs.EOF Then
        ' Verificar si caché es reciente (menos de 1 hora)
        If DateDiff("n", rs!UltimaConsulta, Now) < 60 Then
            ' Usar datos del caché
            producto.ID = rs!IDProductoPS
            producto.Nombre = rs!Nombre
            producto.Precio = rs!PrecioConIVA
            producto.Stock = rs!StockPS
            BuscarProductoPorCodigo = producto
            Exit Function
        End If
    End If

    ' Si no está en caché o está desactualizado, consultar PrestaShop
    producto = ConsultarPrestaShop(codigo)

    ' Actualizar caché
    If rs.EOF Then
        rs.AddNew
    Else
        rs.Edit
    End If
    rs!IDProductoPS = producto.ID
    rs!Referencia = codigo
    rs!Nombre = producto.Nombre
    rs!PrecioConIVA = producto.Precio
    rs!StockPS = producto.Stock
    rs!UltimaConsulta = Now
    rs.Update

    BuscarProductoPorCodigo = producto
End Function
```

---

## 📈 **Beneficios del Sistema de Caché:**

1. **Performance:** Consultas locales < 10ms vs 500ms+ a PrestaShop
2. **Offline:** POS sigue funcionando sin internet
3. **Auditabilidad:** Todo queda registrado en `LogSincronizacion`
4. **Resiliencia:** Cola de sincronización procesa cuando hay conexión
5. **Consistencia:** `DiferenciaStock` permite detectar conflictos

---

## 🎯 **RESUMEN EJECUTIVO**

### **Archivos Críticos:**
- `frmventa.frm`, `Module1.bas`, `frmelige.frm`, `frmaddart.frm`, `FrmInventario.frm`

### **Limitaciones VB6:**
- ✅ **Superables** con WinHTTP en Windows 11
- 🏆 **Mejor opción:** API Bridge (PHP/Python) intermediario

### **Arquitectura Recomendada:**
- VB6 → API Bridge (localhost:8080) → PrestaShop API
- Caché local en Access con 5 tablas nuevas
- Sistema de cola para sincronización offline

### **Tablas Nuevas:**
1. `ConfigAPI` - Configuración
2. `ProductosPS` - Caché de productos
3. `LogSincronizacion` - Auditoría
4. `ColaSyncStock` - Cola offline
5. `MapeoArticulosPS` - Relaciones

---

## 🚀 **PRÓXIMOS PASOS**

1. ✅ **APROBACIÓN DE ARQUITECTURA** ← **Estás aquí**
2. Crear tablas en Access (`canela.mdb`)
3. Desarrollar API Bridge (PHP/Python)
4. Crear `ModuloPrestaShop.bas` en VB6
5. Modificar `frmventa.frm` para búsqueda con PS
6. Implementar actualización de stock post-venta
7. Testing en entorno local
8. Despliegue en producción

---

## 📞 **PREGUNTAS DE VALIDACIÓN**

Antes de continuar con la implementación, necesito confirmar:

1. **¿Apruebas la arquitectura con API Bridge?** (Alternativa: conexión directa VB6-PrestaShop)
2. **¿Prefieres API Bridge en PHP o Python?** (PHP es más común en servidores web)
3. **¿Dónde ejecutarás el API Bridge?** (localhost en el mismo PC del POS, o servidor remoto)
4. **¿Tienes ya la API Key de PrestaShop?** (Se genera desde el panel de administración)
5. **¿Quieres que implemente primero un prototipo de solo lectura?** (buscar productos, ver stock) antes de actualizar stock

---

**Una vez confirmes estos puntos, procederé a la implementación completa.**
