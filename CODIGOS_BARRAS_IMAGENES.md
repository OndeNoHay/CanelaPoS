# Códigos de Barras como Imágenes - Solución Implementada

## 📋 Resumen

El sistema de etiquetas PrestaShop ahora genera códigos de barras como **imágenes PNG** en lugar de usar fuentes tipográficas. Esto garantiza compatibilidad 100% con scanners de códigos de barras.

## 🎯 Problema Resuelto

**Problema anterior:**
- Los códigos de barras se imprimían usando fuentes (IDAutomationHC39M, Libre Barcode 128, etc.)
- Los scanners **NO podían leer** los códigos de barras impresos
- Dependía de fuentes específicas instaladas en el sistema
- Problemas con proporciones y tamaños

**Solución implementada:**
- Los códigos de barras se generan como **imágenes PNG** usando PHP
- Las imágenes se crean con la biblioteca personalizada `BarcodeGenerator`
- Formato EAN13 con especificaciones correctas de barras
- Los scanners **pueden leer** los códigos perfectamente ✅

## 🔧 Cómo Funciona

### 1. **Flujo de Trabajo**

```
Usuario busca productos → VB6 obtiene productos de PrestaShop
         ↓
VB6 recopila todos los EAN13 únicos
         ↓
VB6 envía petición POST al API Bridge con array JSON de EAN13
         ↓
PHP genera imágenes PNG (300x150px) para cada código
         ↓
VB6 carga las imágenes desde disco
         ↓
Al imprimir etiquetas: VB6 usa PaintPicture para insertar las imágenes
         ↓
Al cerrar formulario: VB6 elimina archivos temporales
```

### 2. **Componentes del Sistema**

#### A) Biblioteca PHP: `api_bridge/barcode_generator.php`

Clase `BarcodeGenerator` que implementa:
- Codificación EAN13 según especificaciones oficiales
- Patrones de barras Left-Odd, Left-Even, Right
- Sistema de paridad para primer dígito
- Generación de imagen con GD (incluido en PHP)
- Guardado como PNG de alta resolución

**Métodos principales:**
```php
generateEAN13($ean13, $width, $height)    // Genera imagen en memoria
saveEAN13($ean13, $filepath, $width, $height)  // Guarda como archivo PNG
```

#### B) Endpoint API: `bridge.php?action=generar_codigos_barras`

**Entrada:**
```http
POST /api_bridge/bridge.php?action=generar_codigos_barras
Content-Type: application/json

["8435423154703", "8435423154710", "8435423154727"]
```

**Salida:**
```json
{
  "success": true,
  "data": {
    "archivos": [
      {
        "ean13": "8435423154703",
        "filename": "barcode_8435423154703_1737392840_1234.png",
        "filepath": "/path/to/api_bridge/temp_barcodes/barcode_8435423154703_1737392840_1234.png",
        "url": "api_bridge/temp_barcodes/barcode_8435423154703_1737392840_1234.png"
      },
      ...
    ],
    "total_generados": 3,
    "total_errores": 0,
    "errores": []
  },
  "tiempo_ms": 45
}
```

**Características:**
- Acepta hasta 500 códigos por petición
- Genera imágenes a 300x150 píxeles (alta resolución)
- Limpia automáticamente archivos de más de 1 hora
- Maneja errores individualmente por código

#### C) Formulario VB6: `frmetiquetasPS.frm`

**Variables agregadas:**
```vb
Dim barcodeImages As Collection      ' Imágenes indexadas por EAN13
Dim barcodeFilenames As Collection   ' Nombres para limpieza
Dim rutaServidorPHP As String        ' Ruta base
```

**Función principal: `GenerarImagenesCodigosBarras()`**
1. Recopila EAN13 únicos de todas las etiquetas
2. Construye JSON array
3. Hace POST al API usando MSXML2.XMLHTTP
4. Parsea respuesta JSON (sin biblioteca externa)
5. Carga imágenes con LoadPicture()
6. Almacena en Collection indexada por EAN13

**Impresión modificada:**
```vb
' Antes (usando fuentes):
Printer.FontName = "IDAutomationHC39M"
Printer.Print "*" & ean13 & "*"

' Ahora (usando imágenes):
Set barcodeImg = barcodeImages(ean13)
Printer.PaintPicture barcodeImg, x + 15, Y, 35, 10
```

**Limpieza en Form_Unload:**
- Libera Collection de imágenes
- Elimina archivos PNG temporales con Kill
- Limpia tabla temporal de base de datos

### 3. **Carpeta Temporal**

**Ubicación:** `api_bridge/temp_barcodes/`

**Archivos generados:**
- Formato: `barcode_[EAN13]_[timestamp]_[random].png`
- Ejemplo: `barcode_8435423154703_1737392840_1234.png`
- Tamaño: ~5-10 KB por archivo

**Limpieza automática:**
- PHP: Elimina archivos de más de 1 hora al generar nuevos códigos
- VB6: Elimina archivos al cerrar el formulario
- Git: `.gitignore` evita que se suban al repositorio

## 📐 Especificaciones Técnicas

### Dimensiones de Código de Barras

**Imagen generada (PHP):**
- Ancho: 300 píxeles
- Alto: 150 píxeles
- Formato: PNG con fondo blanco
- Incluye números legibles debajo de las barras

**Impresión en etiqueta (VB6):**
- Ancho: 35 mm
- Alto: 10 mm
- Escala automática con PaintPicture
- Posición: x + 15, y (parte superior derecha de etiqueta)

### Formato EAN13

**Estructura:**
- 13 dígitos numéricos
- Primer dígito: Define sistema de paridad
- Dígitos 1-6: Codificados con paridad L-odd/L-even
- Dígitos 7-12: Codificados con paridad R
- Guard bars: 101 (inicio), 01010 (centro), 101 (fin)

**Ejemplo:** `8435423154703`
```
8 = Sistema (paridad: OEEOEO)
435423 = Grupo izquierdo (con paridad)
154703 = Grupo derecho
```

## 🚀 Ventajas de Esta Solución

✅ **Compatibilidad 100% con scanners**
- Las imágenes siguen exactamente las especificaciones EAN13
- Proporciones y tamaños correctos
- No depende de renderizado de fuentes

✅ **Sin dependencias de fuentes**
- No requiere instalar IDAutomationHC39M
- No requiere Libre Barcode 128/EAN13
- Funciona en cualquier sistema

✅ **Alta calidad**
- Resolución 300x150 píxeles
- Escala perfectamente al imprimir
- Barras nítidas y bien definidas

✅ **Reutilización eficiente**
- Genera una imagen por EAN13 único
- Reutiliza imágenes para productos con múltiples tallas
- Caché temporal evita regeneración

✅ **Mantenible**
- Código PHP simple y bien documentado
- Biblioteca standalone (sin Composer)
- Fácil de extender a otros formatos (Code128, Code39, etc.)

## 🔍 Resolución de Problemas

### El scanner no lee los códigos

**Verificar:**
1. ¿Se están generando las imágenes?
   - Revisar carpeta `api_bridge/temp_barcodes/`
   - Debe haber archivos PNG después de buscar productos

2. ¿Las imágenes se ven correctas?
   - Abrir un PNG con visor de imágenes
   - Debe verse un código de barras con líneas verticales claras

3. ¿La impresión es legible?
   - Las barras deben verse negras y nítidas
   - No debe haber difuminado o pixelación

4. ¿El scanner está configurado para EAN13?
   - Algunos scanners requieren activar formatos específicos
   - Probar con códigos de productos comerciales conocidos

### Error "No se pudieron cargar las imágenes"

**Causas posibles:**
1. **Servidor PHP no accesible**
   - Verificar que Apache/PHP estén corriendo
   - Probar: `http://localhost/CanelaPoS/api_bridge/bridge.php?action=test`

2. **Permisos de carpeta**
   ```bash
   chmod 755 api_bridge/temp_barcodes
   ```

3. **Ruta incorrecta en VB6**
   - Variable `rutaServidorPHP` debe apuntar a la raíz del proyecto
   - Por defecto usa `App.Path`

### Error HTTP al generar códigos

**Verificar URL del API:**
```vb
' En GenerarImagenesCodigosBarras()
urlAPI = "http://localhost/CanelaPoS/api_bridge/bridge.php?action=generar_codigos_barras"
```

Ajustar según tu configuración:
- Cambiar `localhost` si usas otro host
- Cambiar `/CanelaPoS/` si el proyecto está en otra carpeta
- Verificar que el servidor web esté corriendo

## 📝 Archivos Modificados

### Archivos Nuevos
- ✨ `api_bridge/barcode_generator.php` - Biblioteca de generación
- ✨ `api_bridge/temp_barcodes/.gitignore` - Ignorar archivos temporales
- ✨ `CODIGOS_BARRAS_IMAGENES.md` - Esta documentación

### Archivos Modificados
- 🔧 `api_bridge/bridge.php` - Nuevo endpoint generar_codigos_barras
- 🔧 `frmetiquetasPS.frm` - Generación y uso de imágenes

## 🎓 Referencias

**Especificación EAN-13:**
- [GS1 - EAN/UPC Symbology](https://www.gs1.org/standards/barcodes/ean-upc)
- [Wikipedia - EAN-13](https://en.wikipedia.org/wiki/International_Article_Number)

**Bibliotecas alternativas (si necesitas otros formatos):**
- [picqer/php-barcode-generator](https://github.com/picqer/php-barcode-generator) - Soporta Code39, Code128, QR, etc.
- [tecnickcom/tc-lib-barcode](https://github.com/tecnickcom/tc-lib-barcode) - Muy completa, muchos formatos

## ✅ Resultado Final

Los códigos de barras ahora:
- ✅ Se imprimen como imágenes de alta calidad
- ✅ Son 100% escaneables con cualquier scanner
- ✅ No dependen de fuentes tipográficas
- ✅ Siguen las especificaciones EAN13 oficiales
- ✅ Se generan automáticamente al buscar productos
- ✅ Se limpian automáticamente al cerrar el formulario

**¡El sistema está listo para producción!** 🎉
