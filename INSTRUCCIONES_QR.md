# Instrucciones para Usar Códigos QR

## ⚠️ ARCHIVO REQUERIDO

Para que el sistema de códigos QR funcione, necesitas descargar **un archivo adicional**:

### Descargar qrcode.js

**Opción 1: Descarga directa desde GitHub**
1. Ve a: https://github.com/davidshimjs/qrcodejs
2. Haz clic en el archivo `qrcode.js` (o `qrcode.min.js`)
3. Haz clic en el botón **"Raw"**
4. Guarda la página (Ctrl+S) como `qrcode.js`
5. Copia el archivo a la **carpeta raíz de la aplicación** (misma carpeta que qr_generator.html)

**Opción 2: Desde CDN (para probar)**
Si solo quieres probar, puedes descargar directamente:
```
https://cdn.rawgit.com/davidshimjs/qrcodejs/gh-pages/qrcode.min.js
```
Guarda como `qrcode.js` en la carpeta de la aplicación.

## 📁 Estructura de Archivos Requerida

La carpeta de la aplicación debe contener:
```
/CanelaPoS/
├── frmetiquetasQR.frm          ← Formulario VB6
├── qr_generator.html           ← Generador QR (ya existe)
└── qrcode.js                   ← ¡DESCARGAR ESTE ARCHIVO!
```

**IMPORTANTE**: El archivo debe llamarse exactamente `qrcode.js` (también funciona `qrcode.min.js` si cambias la línea 33 del HTML).

## ✅ Verificar la Instalación

### Método 1: Abrir en navegador (RECOMENDADO)
1. Abre `qr_generator.html` en tu navegador (Chrome, Firefox, Edge)
2. Deberías ver:
   - ✅ **"Biblioteca QR cargada correctamente"** (fondo verde)
   - **"Biblioteca: qrcodejs (davidshimjs)"**
   - Un código QR de prueba visible debajo
3. En la consola (F12) debe aparecer:
   - `QR library loaded successfully`
   - `Test QR for VB6 generated successfully`

**Si ves error rojo**, el archivo `qrcode.js` no está en la ubicación correcta o no se descargó bien.

### Método 2: Desde VB6
1. Abre el proyecto en VB6
2. Ejecuta el formulario `frmetiquetasQR`
3. Espera 1-2 segundos
4. El botón debe cambiar a **"Imprime con QR"**
5. Si dice **"ERROR: QR no disponible"**, revisa que `qrcode.js` esté presente

## 🧪 Probar el Sistema

1. Abre el formulario `frmetiquetasQR.frm` en VB6
2. Introduce rango de IDs de productos (ej: 1-10)
3. Clic en "Buscar en PrestaShop"
4. Clic en "Imprime con QR"
5. Deberías ver códigos QR **cuadrados negros escaneables** en las etiquetas

## 🔍 Probar con tu Código de Prueba

Puedes crear un archivo HTML simple para probar que la biblioteca funciona:

```html
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
</head>
<body>
    <div id="qrcode"></div>
    <script src="qrcode.js"></script>
    <script>
    var qrcode = new QRCode(document.getElementById("qrcode"), {
        text: "2410788252771",
        width: 128,
        height: 128,
        colorDark: "#000000",
        colorLight: "#ffffff",
        correctLevel: QRCode.CorrectLevel.H
    });
    </script>
</body>
</html>
```

Si esto funciona, entonces la biblioteca está bien instalada.

## ❓ Solución de Problemas

### Error: "No se pudo cargar qrcode.js"
**Causa**: El archivo no está en la carpeta correcta o tiene nombre incorrecto
**Solución**:
- Verifica que `qrcode.js` está en la **misma carpeta** que `qr_generator.html`
- NO debe estar en una subcarpeta
- El nombre debe ser exactamente `qrcode.js` (minúsculas)
- Si descargaste `qrcode.min.js`, renómbralo a `qrcode.js` o cambia la línea 33 del HTML

### Error: "Biblioteca QR no cargada"
**Causa**: El WebBrowser no pudo cargar el archivo JavaScript
**Solución**:
- Verifica que el archivo no está bloqueado por Windows:
  - Clic derecho en `qrcode.js` → Propiedades
  - Si hay un botón "Desbloquear" en la parte inferior, haz clic en él
  - Aplica y cierra
- Asegúrate de que el archivo no está corrupto (descárgalo de nuevo)
- Verifica que es un archivo JavaScript válido (ábrelo en un editor de texto)

### Los QR se generan pero no se imprimen
**Causa**: El WebBrowser necesita más tiempo para inicializar
**Solución**:
- Espera unos segundos después de abrir el formulario
- El botón debe decir "Imprime con QR" (no "Cargando QR...")
- Si sigue fallando, cierra y vuelve a abrir el formulario

### Los QR se imprimen pero el escáner no los lee
**Causa**: Resolución de impresión muy baja o tamaño muy pequeño
**Solución**:
- Aumenta la resolución de la impresora a 300 DPI o más
- Aumenta el tamaño de las etiquetas
- Prueba con un lector QR de smartphone para verificar que son válidos
- Usa nivel de corrección más alto (edita línea 91 del HTML: `QRCode.CorrectLevel.H`)

## 📊 Comparación: Códigos de Barras vs QR

| Característica | Código de Barras | Código QR |
|---------------|------------------|-----------|
| **Formulario** | frmetiquetasPS.frm | frmetiquetasQR.frm |
| **Formato** | Horizontal (Code 39) | Cuadrado |
| **Espacio usado** | ~40mm x 7mm | ~15mm x 15mm |
| **Dependencias** | Fuente IDAutomationHC39M | qrcode.js |
| **Configuración** | Ninguna | Descargar 1 archivo |
| **Complejidad** | Simple | Media |
| **Escaneabilidad** | Code 39 scanner | Cualquier lector QR o smartphone |
| **Ventaja** | Más simple, ya funciona | Más compacto, multidireccional |

## 🎯 Recomendación

- **Si tienes un escáner Code 39 funcionando**: usa `frmetiquetasPS.frm` (más simple)
- **Si quieres códigos más compactos o usar smartphone**: usa `frmetiquetasQR.frm` (requiere qrcode.js)
- **Ambos sistemas coexisten**: puedes tener los dos instalados y usar el que prefieras

## 📝 Notas Técnicas

**Biblioteca QR usada**: qrcodejs by David Shim
- **Repositorio**: https://github.com/davidshimjs/qrcodejs
- **Licencia**: MIT License (libre uso comercial y personal)
- **Tamaño**: ~12 KB (minificado)
- **API**: `new QRCode(elemento, opciones)`
- **Corrección de errores**:
  - L (7%) - Mínimo
  - M (15%) - Medio
  - Q (25%) - Bueno
  - H (30%) - Máximo (recomendado para impresión)

**Ventajas de esta biblioteca**:
- ✅ API muy simple y fácil de usar
- ✅ Genera automáticamente canvas o imagen
- ✅ Compatible con IE9+ y todos los navegadores modernos
- ✅ No requiere dependencias adicionales
- ✅ Ampliamente usada y probada

**Cómo funciona con VB6**:
1. VB6 carga `qr_generator.html` en WebBrowser (invisible)
2. El HTML carga `qrcode.js`
3. VB6 llama a función JavaScript: `GenerateQRCode(ean13, tamaño)`
4. JavaScript crea QR con `new QRCode()` en un div temporal
5. JavaScript extrae el canvas generado
6. JavaScript convierte canvas a base64 (PNG)
7. VB6 recibe el base64 y lo decodifica con MSXML2
8. VB6 guarda temporalmente como .png
9. VB6 carga con LoadPicture()
10. VB6 imprime con Printer.PaintPicture()
11. VB6 elimina archivo temporal

## 🆘 Soporte

Si sigues teniendo problemas:

1. **Verifica instalación**:
   - Abre `qr_generator.html` en navegador
   - Debe mostrar mensaje verde y QR de prueba

2. **Revisa consola del navegador**:
   - Presiona F12
   - Mira si hay errores en rojo
   - Copia los mensajes para depurar

3. **Archivos correctos**:
   - `qr_generator.html` - debe tener `<script src="qrcode.js">`
   - `qrcode.js` - debe existir en la misma carpeta
   - Ambos en la carpeta raíz de la aplicación

4. **Alternativa temporal**:
   - Usa `frmetiquetasPS.frm` (códigos de barras)
   - Funciona sin archivos adicionales
   - Ya probado y funcionando

## 📚 Recursos

- Repositorio biblioteca: https://github.com/davidshimjs/qrcodejs
- Demo online: https://davidshimjs.github.io/qrcodejs/
- Especificación QR: ISO/IEC 18004
- Probar QR codes: Usa cualquier app de QR en tu smartphone
