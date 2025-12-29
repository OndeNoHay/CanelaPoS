# Diagnóstico del Error 400 - API Bridge

## Problema Actual

El log muestra:
```
[2025-12-29 18:59:42] [INFO] ERROR: Error HTTP: 400 - Bad Request
```

Error HTTP 400 significa que el servidor rechaza la petición porque está mal formada.

## Causas Posibles

1. **Parámetros incorrectos**: El API Bridge espera `reference` o `ean13` en lugar de `code`
2. **Método incorrecto**: Espera POST en lugar de GET
3. **Formato JSON**: Espera JSON en el body, no query string
4. **Headers faltantes**: Necesita headers específicos
5. **Action incorrecta**: El action debe ser diferente a `search`

## Pasos para Resolver

### Opción 1: Compartir bridge.php

Por favor, comparte el contenido del archivo `api_bridge/bridge.php` (puedes ocultar la API Key).

### Opción 2: Ejecutar Script de Diagnóstico

He creado `DiagnosticoAPIBridge.bas` que prueba diferentes formatos:

1. En VB6, abre la ventana Immediate (Ctrl+G)
2. Ejecuta: `DiagnosticoAPIBridge`
3. Revisa los resultados en la ventana Immediate
4. Envíame el output completo

### Opción 3: Prueba Manual con test_bridge.html

1. Abre en tu navegador: `https://www.canelamoda.es/api_bridge/test_bridge.html`
2. Introduce el código: `2804389083757`
3. Haz clic en buscar
4. Abre las DevTools del navegador (F12)
5. Ve a la pestaña "Network"
6. Busca la petición al bridge.php
7. Haz clic derecho → Copy → Copy as cURL
8. Envíame el comando cURL completo

### Opción 4: Revisar Código PHP del test_bridge.html

Si tienes acceso al código HTML de `test_bridge.html`, revisa cómo hace la llamada al API:

Busca en el código JavaScript algo como:
```javascript
fetch('bridge.php?action=...
// o
$.ajax({ url: 'bridge.php', ...
```

## Formatos Probados por el Script

El script prueba:

1. `GET bridge.php?action=search&code=XXX`
2. `GET bridge.php?action=search&reference=XXX`
3. `GET bridge.php?action=search&ean13=XXX`
4. `POST bridge.php` con JSON: `{"action":"search","code":"XXX"}`
5. `GET bridge.php?code=XXX` (sin action)
6. `GET bridge.php?action=getProduct&code=XXX`

Uno de estos debería funcionar.

## Temporalmente: Desactivar Integración

Mientras investigamos, puedes desactivar la búsqueda en PrestaShop:

**Archivo:** `config/prestashop.ini`

```ini
[General]
BuscarEnPrestaShop=0    # Cambiar a 0
```

Esto permitirá que el POS funcione normalmente buscando solo en BD local.

## Información Necesaria

Para ayudarte mejor, necesito:

1. Contenido de `bridge.php` (sin API Key)
2. O resultado del script de diagnóstico
3. O cURL de la petición que funciona desde test_bridge.html

¿Cuál prefieres proporcionar?
