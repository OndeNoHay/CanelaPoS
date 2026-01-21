# Sistema de Etiquetas con Códigos QR para PrestaShop

## Descripción

Este sistema permite imprimir etiquetas de productos de PrestaShop con códigos QR en lugar de códigos de barras tradicionales. Los códigos QR son más compactos y permiten una mejor organización del diseño de la etiqueta.

## Archivos del Sistema

### Nuevos archivos creados:

1. **frmetiquetasQR.frm** - Formulario VB6 para etiquetas con QR
2. **frmetiquetasQR.frx** - Archivo binario del formulario
3. **qr_generator.html** - Generador de códigos QR en JavaScript (100% local, sin servidor)

### Archivos existentes (sin modificar):

- **frmetiquetasPS.frm** - Formulario original con códigos de barras (sigue funcionando)
- **ModuloPrestaShop.bas** - Módulo de integración con PrestaShop API
- **api_bridge/bridge.php** - Bridge PHP para API de PrestaShop

## Características

### Ventajas del QR vs Código de Barras

✓ **Más compacto**: Un QR de 15x15mm puede contener un EAN13 completo
✓ **Mejor composición**: El formato cuadrado permite mejor uso del espacio
✓ **Más robusto**: Mayor tolerancia a errores de impresión
✓ **Escaneo multidireccional**: Se puede leer desde cualquier ángulo
✓ **100% local**: No requiere servidor ni conexión a internet para generar QR

### Diseño de Etiqueta con QR

```
┌─────────────────────────────────────┐
│ [QR]  PVP: 99.99€                  │
│ [15mm] Nombre producto - Talla     │
│                                     │
└─────────────────────────────────────┘
```

**Distribución:**
- **QR code**: Esquina superior izquierda (tamaño dinámico 10-25mm)
- **Precio**: A la derecha del QR, parte superior (tamaño fijo 14pt, negrita)
- **Nombre**: A la derecha del QR, debajo del precio (8pt)
- **Talla**: Incluida con el nombre si aplica

## Cómo Usar

### 1. Abrir el Formulario

Desde el proyecto VB6, abrir `frmetiquetasQR.frm`

### 2. Configurar Dimensiones

- **Ancho etiqueta**: 52.5mm (por defecto)
- **Alto etiqueta**: 29.7mm (por defecto)
- **Márgenes interiores H/V**: 1mm (ajustables)
- **Margen superior**: 0mm (ajustable)

### 3. Buscar Productos

1. Introducir rango de IDs de productos PrestaShop (Desde/Hasta)
2. Clic en "Buscar en PrestaShop"
3. El sistema mostrará productos activos con stock en el grid

### 4. Imprimir Etiquetas

1. Revisar productos en el grid
2. (Opcional) Configurar posición inicial de impresión
3. Clic en "Imprime con QR"
4. El sistema generará códigos QR e imprimirá las etiquetas

## Tecnología

### JavaScript Local (qr_generator.html)

El archivo HTML contiene:
- **Biblioteca QRCode.js completa** (embebida, sin CDN)
- **Generador de QR** que convierte texto a código QR
- **Exportador a formato BMP** compatible con VB6
- **Interfaz para VB6** vía WebBrowser control

### Comunicación VB6 ↔ JavaScript

```vb
' VB6 llama a función JavaScript
base64Data = WebBrowser1.Document.parentWindow.GenerateQRCode(texto, tamaño)

' JavaScript genera QR y devuelve imagen en base64
return canvas.toDataURL('image/bmp')
```

### Proceso de Generación

1. **VB6** carga `qr_generator.html` en WebBrowser invisible (Form_Load)
2. **Para cada etiqueta**:
   - VB6 llama a `GenerateQRCode(ean13, tamañoPx)`
   - JavaScript genera QR en Canvas
   - JavaScript convierte a base64 (BMP)
   - VB6 decodifica base64 usando MSXML2.DOMDocument
   - VB6 guarda temporalmente como archivo BMP
   - VB6 carga con LoadPicture()
   - VB6 imprime con Printer.PaintPicture()
   - VB6 elimina archivo temporal

## Configuración de Impresora

- **Formato papel**: A4 (210mm x 297mm)
- **Orientación**: Portrait (vertical)
- **Márgenes externos**: 0mm (configurar en impresora)
- **Resolución recomendada**: 300 DPI o superior
- **ScaleMode**: 6 (milímetros)

## Comparación con Sistema de Códigos de Barras

| Característica | Código de Barras (frmetiquetasPS) | Código QR (frmetiquetasQR) |
|---------------|-----------------------------------|----------------------------|
| Formato | Horizontal (fuente IDAutomationHC39M) | Cuadrado (imagen generada) |
| Espacio ocupado | ~40mm ancho x 7mm alto | ~15mm x 15mm |
| Dependencias | Fuente instalada en sistema | HTML + JavaScript local |
| Generación | Directa (Printer.Print) | Vía JavaScript + LoadPicture |
| Escáner | Code 39 específico | Cualquier lector QR |
| Complejidad | Simple | Media (pero 100% local) |

## Requisitos del Sistema

### VB6 Runtime:
- Visual Basic 6.0 SP6
- DAO 3.6 (Microsoft Data Access Objects)
- MSXML2.DOMDocument (para decodificar base64)
- ADODB.Stream (para guardar archivos)

### Controles ActiveX:
- **dbgrid32.ocx** - Microsoft Data Bound Grid Control
- **mswebbrw.dll** (SHDocVw) - Microsoft Web Browser Control

### Archivos requeridos en carpeta de aplicación:
- **qr_generator.html** - DEBE estar en la raíz de la aplicación (App.Path)
- **bdtienda.mdb** - Base de datos Access (tabla temporal TempEtiquetasPS)

## Solución de Problemas

### Error: "No se encontró qr_generator.html"

**Causa**: El archivo HTML no está en la carpeta correcta.
**Solución**: Copiar `qr_generator.html` a la carpeta raíz de la aplicación (misma carpeta que el .exe)

### Los QR no se generan (aparece texto plano)

**Causas posibles**:
1. WebBrowser no terminó de cargar el HTML
2. JavaScript bloqueado por seguridad
3. Error en decodificación base64

**Soluciones**:
1. Aumentar tiempo de espera en GenerarQRCode()
2. Verificar configuración de seguridad de IE (Internet Explorer)
3. Verificar que MSXML2 está disponible

### QR codes se imprimen pero el escáner no los lee

**Causas posibles**:
1. Resolución de impresión muy baja
2. Tamaño de QR muy pequeño
3. Escáner no compatible con QR

**Soluciones**:
1. Aumentar resolución de impresora a 300 DPI o más
2. Aumentar tamaño de etiqueta o reducir márgenes
3. Probar con escáner compatible con códigos QR estándar

### Error: "Type mismatch" o "Invalid picture"

**Causa**: Problema al decodificar o cargar la imagen.
**Solución**: Verificar que ADODB.Stream y MSXML2 están registrados correctamente

## Mejoras Futuras Posibles

1. **Caché de QR codes**: Generar todos los QR al inicio para imprimir más rápido
2. **Configuración de nivel de corrección**: Permitir elegir nivel L/M/Q/H
3. **Información adicional en QR**: Incluir precio y talla además del EAN13
4. **Vista previa**: Mostrar cómo quedará la etiqueta antes de imprimir
5. **Exportar a PDF**: Guardar etiquetas como PDF en lugar de imprimir

## Soporte

Para problemas o dudas:
1. Verificar que todos los archivos requeridos están presentes
2. Comprobar que los controles ActiveX están registrados
3. Revisar errores en el log de VB6
4. Probar primero con el sistema de códigos de barras (frmetiquetasPS) para aislar el problema

## Autor y Versión

- **Sistema original**: frmetiquetasPS.frm (códigos de barras)
- **Sistema QR**: frmetiquetasQR.frm
- **Fecha de creación**: Enero 2026
- **Versión**: 1.0
- **Compatibilidad**: VB6, Windows XP+, PrestaShop 8.1+
