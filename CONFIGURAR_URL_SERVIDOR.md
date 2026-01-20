# Configuración de URL del Servidor para Códigos de Barras

## 📋 Resumen

El sistema de códigos de barras necesita descargar imágenes desde el servidor PHP. Debes configurar las URLs correctamente según tu entorno.

## 🔧 URLs a Configurar

### En `frmetiquetasPS.frm`

Busca la función `GenerarImagenesCodigosBarras()` y modifica estas dos líneas:

#### 1. URL del API (línea ~764)

```vb
' IMPORTANTE: Cambiar esta URL según tu configuración
urlAPI = "http://canelamoda.es/api_bridge/bridge.php?action=generar_codigos_barras"
```

#### 2. URL base para descargar imágenes (línea ~802)

```vb
' URL base del servidor para descargar imágenes
urlBaseServidor = "http://canelamoda.es/api_bridge/temp_barcodes/"
```

## 🌍 Configuraciones Según Entorno

### Servidor en Producción (canelamoda.es)

```vb
' URL del API
urlAPI = "http://canelamoda.es/api_bridge/bridge.php?action=generar_codigos_barras"

' URL base para imágenes
urlBaseServidor = "http://canelamoda.es/api_bridge/temp_barcodes/"
```

### Servidor Local (localhost)

```vb
' URL del API
urlAPI = "http://localhost/CanelaPoS/api_bridge/bridge.php?action=generar_codigos_barras"

' URL base para imágenes
urlBaseServidor = "http://localhost/CanelaPoS/api_bridge/temp_barcodes/"
```

### Servidor en Red Local (IP específica)

```vb
' URL del API
urlAPI = "http://192.168.1.100/pos/api_bridge/bridge.php?action=generar_codigos_barras"

' URL base para imágenes
urlBaseServidor = "http://192.168.1.100/pos/api_bridge/temp_barcodes/"
```

### Servidor con HTTPS

```vb
' URL del API
urlAPI = "https://canelamoda.es/api_bridge/bridge.php?action=generar_codigos_barras"

' URL base para imágenes
urlBaseServidor = "https://canelamoda.es/api_bridge/temp_barcodes/"
```

## 📁 Carpetas Temporales

El sistema usa DOS carpetas:

### 1. Carpeta en el Servidor (PHP)
- **Ubicación:** `api_bridge/temp_barcodes/`
- **Propósito:** Generar imágenes PNG de códigos de barras
- **Acceso:** HTTP público
- **Limpieza:** Automática (archivos > 1 hora)

### 2. Carpeta Local (VB6)
- **Ubicación:** `[App.Path]\temp_barcodes\`
- **Propósito:** Descargar imágenes del servidor para uso local
- **Acceso:** Sistema de archivos local
- **Limpieza:** Al cerrar formulario

## 🔄 Flujo de Trabajo

```
1. VB6 envía JSON al API → http://canelamoda.es/api_bridge/bridge.php
                              ↓
2. PHP genera imágenes PNG → api_bridge/temp_barcodes/barcode_xxx.png
                              ↓
3. VB6 descarga cada imagen → GET http://canelamoda.es/api_bridge/temp_barcodes/barcode_xxx.png
                              ↓
4. VB6 guarda localmente → C:\...\CanelaPoS\temp_barcodes\barcode_xxx.png
                              ↓
5. VB6 carga con LoadPicture() desde disco local
                              ↓
6. Al imprimir: PaintPicture usa la imagen cargada
                              ↓
7. Al cerrar: Elimina archivos locales
```

## ⚠️ Problemas Comunes

### Error: "No se pudieron cargar las imágenes"

**Causa:** URLs incorrectas o servidor no accesible

**Solución:**
1. Verificar que las URLs son correctas
2. Probar en navegador:
   - `http://canelamoda.es/api_bridge/bridge.php?action=test`
   - Debe devolver JSON con `"success": true`
3. Verificar que la carpeta `temp_barcodes` tiene permisos de lectura HTTP

### Error HTTP 404

**Causa:** Ruta incorrecta al archivo PHP

**Solución:**
- Verificar que `bridge.php` existe en `api_bridge/`
- Verificar que la ruta incluye `/api_bridge/`
- Probar la URL completa en navegador

### Error HTTP 500

**Causa:** Error en PHP

**Solución:**
- Revisar logs de Apache/PHP
- Verificar que `barcode_generator.php` existe
- Verificar permisos de escritura en `temp_barcodes/`

### Imágenes no se descargan

**Causa:** Firewall o proxy bloqueando conexión

**Solución:**
- Verificar firewall de Windows
- Verificar que VB6 puede hacer peticiones HTTP
- Probar desde navegador en la misma máquina

## 🧪 Prueba de Conectividad

### Desde Navegador

1. Abrir: `http://canelamoda.es/api_bridge/bridge.php?action=test`
2. Debe mostrar:
   ```json
   {
     "success": true,
     "data": {
       "mensaje": "Conexión exitosa con PrestaShop",
       ...
     }
   }
   ```

### Desde VB6

En el Immediate Window (Ctrl+G) ejecutar:

```vb
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "http://canelamoda.es/api_bridge/bridge.php?action=test", False
http.send
? http.Status
? Left(http.responseText, 200)
```

Debe mostrar:
```
200
{"success":true,"data":{"mensaje":"Conexión exitosa con PrestaShop",...
```

## 🔒 Consideraciones de Seguridad

### Carpeta Temporal Pública

La carpeta `api_bridge/temp_barcodes/` es **accesible públicamente** por HTTP.

**Riesgos:**
- Cualquiera puede ver/descargar las imágenes si conoce el nombre
- Los nombres incluyen timestamp y número aleatorio para dificultar adivinación

**Mitigaciones:**
- Limpieza automática de archivos > 1 hora
- Nombres únicos con timestamp + aleatorio
- Solo imágenes PNG (sin datos sensibles)

### Alternativa: Autenticación

Si necesitas mayor seguridad, puedes:
1. Agregar token de autenticación al API
2. Usar carpeta fuera de DocumentRoot
3. Servir imágenes solo con autenticación válida

## 📝 Resumen

**Para que funcione correctamente:**

1. ✅ Configurar `urlAPI` con la URL correcta del servidor
2. ✅ Configurar `urlBaseServidor` con la URL de la carpeta temp_barcodes
3. ✅ Verificar que el servidor PHP es accesible desde VB6
4. ✅ Verificar permisos de escritura en `api_bridge/temp_barcodes/`
5. ✅ Verificar que VB6 puede descargar archivos HTTP

**El sistema descargará las imágenes del servidor a una carpeta temporal local y las cargará desde allí.**
