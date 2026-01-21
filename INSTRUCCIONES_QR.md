# Instrucciones para Usar Códigos QR

## ⚠️ ARCHIVO REQUERIDO

Para que el sistema de códigos QR funcione, necesitas descargar **un archivo adicional**:

### Descargar qrcode.min.js

**Opción 1: Descarga directa**
1. Ve a: https://raw.githubusercontent.com/kazuhikoarase/qrcode-generator/master/js/qrcode.min.js
2. Guarda el archivo como `qrcode.min.js`
3. Copia el archivo a la **carpeta raíz de la aplicación** (misma carpeta que qr_generator.html)

**Opción 2: Desde el repositorio**
1. Ve a: https://github.com/kazuhikoarase/qrcode-generator
2. Navega a: `js/qrcode.min.js`
3. Haz clic en "Raw" y guarda el archivo
4. Copia el archivo a la **carpeta raíz de la aplicación**

## 📁 Estructura de Archivos Requerida

La carpeta de la aplicación debe contener:
```
/CanelaPoS/
├── frmetiquetasQR.frm          ← Formulario VB6
├── qr_generator.html           ← Generador QR (ya existe)
└── qrcode.min.js              ← ¡DESCARGAR ESTE ARCHIVO!
```

## ✅ Verificar la Instalación

### Método 1: Abrir en navegador
1. Abre `qr_generator.html` en tu navegador (Chrome, Firefox, Edge)
2. Deberías ver:
   - ✅ **"Biblioteca QR cargada correctamente"** (fondo verde)
   - Un código QR de prueba visible
3. Si ves un **error rojo**, el archivo qrcode.min.js no está en la ubicación correcta

### Método 2: Desde VB6
1. Abre el proyecto en VB6
2. Ejecuta el formulario `frmetiquetasQR`
3. Espera 1-2 segundos
4. El botón debe cambiar a **"Imprime con QR"**
5. Si dice **"ERROR: QR no disponible"**, revisa que qrcode.min.js esté presente

## 🧪 Probar el Sistema

1. Abre el formulario `frmetiquetasQR.frm`
2. Introduce rango de IDs de productos (ej: 1-10)
3. Clic en "Buscar en PrestaShop"
4. Clic en "Imprime con QR"
5. Deberías ver códigos QR **cuadrados negros** en las etiquetas

## ❓ Solución de Problemas

### Error: "No se pudo cargar qrcode.min.js"
**Causa**: El archivo no está en la carpeta correcta
**Solución**:
- Verifica que `qrcode.min.js` está en la misma carpeta que `qr_generator.html`
- NO debe estar en una subcarpeta
- El nombre debe ser exactamente `qrcode.min.js` (minúsculas)

### Error: "Biblioteca QR no cargada"
**Causa**: El WebBrowser no pudo cargar el archivo JavaScript
**Solución**:
- Verifica que el archivo no está bloqueado por Windows (clic derecho → Propiedades → Desbloquear)
- Asegúrate de que el archivo no está corrupto (descárgalo de nuevo)

### Los QR no se imprimen
**Causa**: El WebBrowser necesita más tiempo para inicializar
**Solución**:
- Espera unos segundos después de abrir el formulario
- El botón debe decir "Imprime con QR" (no "Cargando QR...")
- Si sigue fallando, cierra y vuelve a abrir el formulario

## 📊 Comparación: Códigos de Barras vs QR

| Característica | Código de Barras | Código QR |
|---------------|------------------|-----------|
| **Formulario** | frmetiquetasPS.frm | frmetiquetasQR.frm |
| **Formato** | Horizontal (Code 39) | Cuadrado |
| **Espacio usado** | ~40mm x 7mm | ~15mm x 15mm |
| **Dependencias** | Fuente IDAutomationHC39M | qrcode.min.js |
| **Configuración** | Ninguna | Descargar archivo |
| **Complejidad** | Simple | Media |
| **Escaneabilidad** | Code 39 scanner | Cualquier lector QR |

## 🎯 Recomendación

- Si tienes un escáner Code 39 funcionando: **usa frmetiquetasPS.frm** (más simple)
- Si quieres códigos más compactos: **usa frmetiquetasQR.frm** (requiere qrcode.min.js)

## 📝 Notas Técnicas

**Biblioteca QR usada**: qrcode-generator by Kazuhiko Arase
- **Licencia**: MIT License (libre uso comercial y personal)
- **Versión**: Latest from master branch
- **Repositorio**: https://github.com/kazuhikoarase/qrcode-generator
- **Tamaño archivo**: ~10 KB (minificado)
- **Corrección de errores**: Level L (7%) - suficiente para EAN13

**Ventajas de esta biblioteca**:
- ✅ Lightweight (pequeña y rápida)
- ✅ No requiere dependencias adicionales
- ✅ Compatible con IE8+ y todos los navegadores modernos
- ✅ Genera QR codes válidos según ISO/IEC 18004
- ✅ Bien mantenida y ampliamente usada

## 🆘 Soporte

Si sigues teniendo problemas:
1. Verifica que `qrcode.min.js` existe y está en la ubicación correcta
2. Abre `qr_generator.html` en el navegador y verifica que el QR de prueba se genera
3. Revisa la consola del navegador (F12) para ver mensajes de error
4. Como alternativa, usa `frmetiquetasPS.frm` que funciona sin archivos adicionales
