# 📦 GUÍA DE INSTALACIÓN - API BRIDGE PRESTASHOP

**Proyecto:** Integración POS VB6 con PrestaShop 8.1
**Fase:** 1 - Solo Lectura
**Fecha:** 19/12/2025

---

## 📋 ÍNDICE

1. [Requisitos previos](#requisitos-previos)
2. [Instalación en Access (Base de datos)](#instalación-en-access)
3. [Instalación en Servidor (API Bridge PHP)](#instalación-en-servidor)
4. [Configuración de PrestaShop](#configuración-de-prestashop)
5. [Configuración del POS VB6](#configuración-del-pos-vb6)
6. [Testing y Verificación](#testing-y-verificación)
7. [Troubleshooting](#troubleshooting)

---

## 🔧 REQUISITOS PREVIOS

### Servidor Web
- **PHP:** 7.4 o superior (recomendado 8.0+)
- **Extensiones PHP requeridas:**
  - `curl` (para peticiones HTTPS)
  - `simplexml` (para parsear XML de PrestaShop)
  - `json` (incluido por defecto)
- **Servidor:** Apache, Nginx, o cualquier servidor compatible con PHP
- **Acceso FTP:** Para subir archivos al servidor

### PrestaShop
- **Versión:** 8.1 (verificada, pero compatible con 1.7+)
- **Webservice API:** Habilitado
- **API Key:** Generada con permisos de lectura en `products` y `stock_availables`

### POS VB6
- **Windows:** 7 o superior (Windows 11 confirmado)
- **Microsoft Access:** Base de datos `canela.mdb`
- **Conexión a Internet:** Para comunicarse con API Bridge

---

## 💾 INSTALACIÓN EN ACCESS

### Paso 1: Abrir Base de Datos

1. Abrir `canela.mdb` en Microsoft Access
2. Hacer backup de la base de datos antes de modificar:
   ```
   Copiar: canela.mdb → canela_backup_20251219.mdb
   ```

### Paso 2: Crear Tablas

1. En Access, ir a: **Crear → Diseño de Consulta**
2. Cerrar la ventana "Mostrar tabla"
3. Ir a: **Ver → Vista SQL**
4. Copiar y ejecutar **CADA BLOQUE** del archivo `crear_tablas_prestashop.sql`

**IMPORTANTE:** Ejecutar los bloques en este orden:

```sql
-- 1. Tabla ConfigAPI
CREATE TABLE ConfigAPI (...);
-- Ejecutar el INSERT de datos iniciales

-- 2. Tabla ProductosPS
CREATE TABLE ProductosPS (...);
CREATE UNIQUE INDEX idx_referencia ON ProductosPS(Referencia);
CREATE INDEX idx_estado ON ProductosPS(EstadoSync);

-- 3. Tabla LogSincronizacion
CREATE TABLE LogSincronizacion (...);
CREATE INDEX idx_fecha ON LogSincronizacion(FechaHora);
CREATE INDEX idx_tipo ON LogSincronizacion(TipoOperacion);

-- 4. Tabla MapeoArticulosPS
CREATE TABLE MapeoArticulosPS (...);
CREATE INDEX idx_idproductops ON MapeoArticulosPS(IDProductoPS);

-- 5. Tabla ColaSyncStock
CREATE TABLE ColaSyncStock (...);
CREATE INDEX idx_procesado ON ColaSyncStock(Procesado);
CREATE INDEX idx_fecha_venta ON ColaSyncStock(FechaVenta);
```

### Paso 3: Verificar Tablas Creadas

Ejecutar esta consulta para verificar:

```sql
SELECT MSysObjects.Name
FROM MSysObjects
WHERE MSysObjects.Type=1
  AND MSysObjects.Name IN ('ConfigAPI','ProductosPS','LogSincronizacion','MapeoArticulosPS','ColaSyncStock')
ORDER BY MSysObjects.Name;
```

**Resultado esperado:** 5 tablas

### Paso 4: Configurar URL del API Bridge

1. Abrir tabla `ConfigAPI` en Access
2. Localizar el registro con `Clave = 'API_BRIDGE_URL'`
3. **TEMPORALMENTE** dejarlo como está:
   ```
   API_BRIDGE_URL = https://www.canelamoda.es/api_bridge/bridge.php
   ```
4. Lo actualizaremos después de subir el bridge al servidor

---

## 🌐 INSTALACIÓN EN SERVIDOR

### Paso 1: Preparar Archivos

En tu PC, tienes la carpeta `api_bridge/` con estos archivos:

```
api_bridge/
├── bridge.php                    (API Bridge principal)
├── api_config.php.example        (Plantilla de configuración)
├── .htaccess                     (Seguridad)
└── cache/                        (se creará automáticamente)
```

### Paso 2: Configurar API Key

1. **Renombrar archivo:**
   ```
   api_config.php.example → api_config.php
   ```

2. **Editar `api_config.php`** con un editor de texto:

   ```php
   // Línea 14: URL de tu API PrestaShop
   define('PRESTASHOP_API_URL', 'https://www.canelamoda.es/api/');

   // Línea 18: TU API KEY (32 caracteres)
   define('PRESTASHOP_API_KEY', 'AQUI_TU_API_KEY_DE_32_CARACTERES');

   // Línea 21: Idioma (1 = Español)
   define('PRESTASHOP_LANGUAGE_ID', 1);

   // Línea 30: Activar DEBUG inicialmente
   define('DEBUG_MODE', true);
   ```

3. **Guardar cambios**

### Paso 3: Subir por FTP

**Conectar a tu servidor por FTP:**

- **Host:** ftp.canelamoda.es (o tu servidor FTP)
- **Usuario:** [tu usuario FTP]
- **Contraseña:** [tu contraseña]

**Subir archivos:**

1. Navegar a la raíz de tu sitio web (generalmente `/public_html/` o `/www/`)
2. Crear carpeta nueva: `api_bridge`
3. Subir TODOS los archivos:
   ```
   api_bridge/bridge.php
   api_bridge/api_config.php
   api_bridge/.htaccess
   ```

**Resultado final en servidor:**
```
https://www.canelamoda.es/
└── api_bridge/
    ├── bridge.php
    ├── api_config.php
    └── .htaccess
```

### Paso 4: Configurar Permisos

**Vía FTP, establecer permisos:**

```
api_bridge/                 → 755 (rwxr-xr-x)
api_bridge/bridge.php       → 644 (rw-r--r--)
api_bridge/api_config.php   → 600 (rw-------) ← IMPORTANTE: Solo lectura del servidor
api_bridge/.htaccess        → 644 (rw-r--r--)
```

**Crear directorio de caché:**

- Crear carpeta: `api_bridge/cache/`
- Permisos: `777` (rwxrwxrwx) - Para que PHP pueda escribir

### Paso 5: Verificar Instalación

**Abrir en navegador:**

```
https://www.canelamoda.es/api_bridge/bridge.php?action=test
```

**Respuesta esperada (JSON):**

```json
{
    "success": true,
    "data": {
        "mensaje": "Conexión exitosa con PrestaShop",
        "prestashop_url": "https://www.canelamoda.es/api/",
        "api_key_configurada": true,
        "tiempo_respuesta_ms": 250,
        "debug_mode": true,
        "cache_enabled": true,
        "php_version": "8.0.28",
        "curl_disponible": true,
        "timestamp": "2025-12-19 14:30:00"
    }
}
```

**Si ves esto:** ✅ **¡Instalación correcta!**

**Si ves error:** 🔍 Ver sección [Troubleshooting](#troubleshooting)

---

## 🔑 CONFIGURACIÓN DE PRESTASHOP

### Generar API Key (si no la tienes)

1. **Acceder al panel de administración de PrestaShop:**
   ```
   https://www.canelamoda.es/admin12345/  (tu URL de admin)
   ```

2. **Navegar a:**
   ```
   Parámetros Avanzados → Webservice
   ```

3. **Activar Webservice:**
   - Si no está activado, activar: **"Habilitar el servicio web de PrestaShop"**
   - Guardar

4. **Agregar nueva clave:**
   - Clic en: **"Añadir una nueva clave"**
   - **Clave:** Se generará automáticamente (32 caracteres)
   - **Descripción:** "POS VB6 - Solo Lectura"
   - **Estado:** Activado ✅

5. **Configurar permisos:**

   Marcar **SOLO ESTAS OPCIONES** (solo lectura):

   ```
   ☑ products                    GET: ✅   PUT: ⬜   POST: ⬜   DELETE: ⬜
   ☑ stock_availables            GET: ✅   PUT: ⬜   POST: ⬜   DELETE: ⬜
   ☑ images                      GET: ✅   PUT: ⬜   POST: ⬜   DELETE: ⬜
   ☑ product_options_values      GET: ✅   PUT: ⬜   POST: ⬜   DELETE: ⬜
   ```

6. **Guardar**

7. **Copiar la API Key** (32 caracteres)

   Ejemplo: `A1B2C3D4E5F6G7H8I9J0K1L2M3N4O5P6`

8. **Actualizar `api_config.php`** en el servidor con esta clave

---

## 🖥️ CONFIGURACIÓN DEL POS VB6

### Paso 1: Agregar Módulo al Proyecto

1. **Abrir tu proyecto VB6** (archivo `.vbp`)

2. **Agregar módulo:**
   - Menú: **Proyecto → Agregar Módulo → Módulo existente**
   - Seleccionar: `ModuloPrestaShop.bas`

3. **Verificar referencias COM:**
   - Menú: **Proyecto → Referencias**
   - Buscar y marcar:
     ```
     ☑ Microsoft Scripting Runtime
     ```
   - Nota: WinHTTP se carga dinámicamente, no necesita referencia

### Paso 2: Inicializar en Form_Load del Menú Principal

**Archivo:** `frmelige.frm`

**Modificar la función `Form_Load` (línea 939):**

```vb
Private Sub Form_Load()
    ' Código existente
    Set bdtienda = OpenDatabase(App.Path & "\canela.mdb")
    FechaTrabajo = HaceFecha(Now)

    ' ========== NUEVO: Inicializar módulo PrestaShop ==========
    If InicializarModuloPS() Then
        MsgBox "Conectado con PrestaShop ✓", vbInformation
    Else
        MsgBox "Modo OFFLINE: No se pudo conectar con PrestaShop" & vbCrLf & _
               "El sistema funcionará solo con datos locales", vbExclamation
    End If
    ' ==========================================================

    BuscarPrestados
    BuscarApartados
    BlAlarmaQuitar = True
End Sub
```

### Paso 3: Actualizar Búsqueda de Productos en TPV

**Archivo:** `frmventa.frm`

**Modificar función de búsqueda (alrededor de línea 1185-1199):**

```vb
Private Sub cmdarticulo_Click()
    On Error GoTo sehodio

    ModoBusca = "articulos"
    If CodigoBusca = "" Then CodigoBusca = InputBox("Escriba el código")

    If CodigoBusca <> "" Then
        ' ========== NUEVO: Buscar primero en PrestaShop ==========
        Dim productoPS As ProductoPS
        productoPS = BuscarProductoPorCodigo(CodigoBusca)

        If productoPS.Encontrado Then
            ' Producto encontrado en PrestaShop
            MsgBox "Producto: " & productoPS.Nombre & vbCrLf & _
                   "Precio: " & Format(productoPS.PrecioConIVA, "0.00") & "€" & vbCrLf & _
                   "Stock disponible: " & productoPS.Stock, vbInformation

            ' Aquí puedes agregar lógica para añadir al ticket
            ' (lo implementaremos en siguientes pasos)

            Exit Sub
        End If
        ' ==========================================================

        ' Si no se encuentra en PrestaShop, buscar localmente
        SqlArticulos = "Select idart, codigo, tipo, precioventa, " _
        & " color, talla, extra from articulos where vendido = false and apartado = false and" _
        & " idart = " & CodigoBusca & " order by codigo"

        ' ... resto del código existente ...
    End If

sehodio:
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description
    End If
End Sub
```

---

## ✅ TESTING Y VERIFICACIÓN

### Test 1: Verificar API Bridge desde Navegador

**URL de test:**
```
https://www.canelamoda.es/api_bridge/bridge.php?action=test
```

**Resultado esperado:** JSON con `"success": true`

---

### Test 2: Buscar Producto Real

**Ejemplo con un producto que EXISTE en tu PrestaShop:**

```
https://www.canelamoda.es/api_bridge/bridge.php?action=buscar_producto&codigo=ABC-12345678
```

**Reemplazar `ABC-12345678` con un código real de tu catálogo**

**Resultado esperado:**

```json
{
    "success": true,
    "data": {
        "id": 123,
        "reference": "ABC-12345678",
        "nombre": "Vestido Rojo Talla M",
        "precio_con_iva": 45.50,
        "stock": 5,
        ...
    },
    "tiempo_ms": 250
}
```

---

### Test 3: Verificar Tablas en Access

1. Abrir `canela.mdb`
2. Verificar tabla `ConfigAPI`:
   ```sql
   SELECT * FROM ConfigAPI;
   ```
   Deberías ver 6 registros de configuración

3. Ejecutar búsqueda en VB6
4. Verificar tabla `ProductosPS` (debe tener 1+ registro)
5. Verificar tabla `LogSincronizacion` (debe registrar la búsqueda)

---

### Test 4: Prueba desde VB6

1. **Abrir VB6 y ejecutar el proyecto (F5)**
2. **En el formulario de ventas, buscar un producto:**
   - Introducir código de producto
   - Debería aparecer mensaje con información de PrestaShop
3. **Verificar log en Access:**
   ```sql
   SELECT * FROM LogSincronizacion ORDER BY FechaHora DESC;
   ```

---

## 🔍 TROUBLESHOOTING

### Error: "Error de configuración: API Key no configurada"

**Causa:** No has editado `api_config.php` o no tiene 32 caracteres

**Solución:**
1. Verificar que renombraste `api_config.php.example` a `api_config.php`
2. Editar línea 18 con tu API Key de 32 caracteres
3. Volver a subir por FTP

---

### Error: "Error al conectar con PrestaShop" (HTTP 401)

**Causa:** API Key incorrecta o sin permisos

**Solución:**
1. Verificar que la API Key es correcta (copiar/pegar desde PrestaShop)
2. En PrestaShop Admin, verificar que el Webservice está **Activado**
3. Verificar permisos GET en `products` y `stock_availables`

---

### Error: "cURL Error: SSL certificate problem"

**Causa:** Certificado SSL no válido

**Solución temporal (SOLO PARA TESTING):**

Editar `bridge.php`, línea ~410:

```php
// Cambiar de:
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, true);

// A (temporal):
curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
```

**⚠️ IMPORTANTE:** Esto solo para testing, en producción usar certificados válidos

---

### Error: "Producto no encontrado" (pero existe en PrestaShop)

**Causa:** Formato de código diferente

**Solución:**
1. Verificar en PrestaShop el formato exacto del campo `reference`
2. Probar con código EAN13 si está configurado
3. Revisar log: `api_bridge/bridge_debug.log` (si DEBUG_MODE = true)

---

### Error: "No se pudo conectar con API Bridge" desde VB6

**Causa:** Firewall, URL incorrecta, o servidor caído

**Solución:**
1. **Verificar URL en Access:**
   ```sql
   SELECT Valor FROM ConfigAPI WHERE Clave='API_BRIDGE_URL';
   ```
   Debe ser: `https://www.canelamoda.es/api_bridge/bridge.php`

2. **Verificar desde navegador:**
   Abrir URL de test (ver Test 1)

3. **Verificar firewall:**
   Permitir conexiones salientes HTTPS (puerto 443)

---

### Ver logs detallados del API Bridge

**Archivo de log:** `api_bridge/bridge_debug.log`

**Acceder vía FTP y descargar para revisar**

Ejemplo de líneas de log:
```
[2025-12-19 14:30:00] [BUSQUEDA] [192.168.1.100] [ABC-12345678] [250ms] Producto encontrado: Vestido Rojo
[2025-12-19 14:30:05] [STOCK] [192.168.1.100] [123] [120ms] Stock obtenido: 5 unidades
```

---

## 📞 SOPORTE

**Si encuentras errores no documentados:**

1. Revisar archivo `bridge_debug.log` en el servidor
2. Revisar tabla `LogSincronizacion` en Access
3. Anotar mensaje de error exacto
4. Verificar versión de PHP: `<?php phpinfo(); ?>`

---

## ✅ CHECKLIST DE INSTALACIÓN

- [ ] Tablas creadas en Access (5 tablas)
- [ ] ConfigAPI configurada con URL del bridge
- [ ] API Key generada en PrestaShop
- [ ] Permisos GET configurados (products, stock_availables)
- [ ] `api_config.php` editado con API Key
- [ ] Archivos subidos por FTP a `api_bridge/`
- [ ] Permisos 600 en `api_config.php`
- [ ] Directorio `cache/` creado con permisos 777
- [ ] Test desde navegador OK (action=test)
- [ ] Test de búsqueda de producto OK
- [ ] ModuloPrestaShop.bas agregado al proyecto VB6
- [ ] Form_Load modificado con inicialización
- [ ] Búsqueda de producto modificada en frmventa.frm
- [ ] Test desde VB6 exitoso

---

**¡Instalación completada!** 🎉

La Fase 1 (Solo Lectura) está lista para usar.
