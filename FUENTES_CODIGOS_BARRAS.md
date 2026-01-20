# Fuentes de Códigos de Barras para Etiquetas

## 🎯 Problema

El formulario de etiquetas (`frmetiquetasPS.frm`) imprime códigos de barras EAN13. Para que estos códigos sean **escaneables** por lectores de códigos de barras, necesitas instalar una fuente específica para EAN13 en Windows.

---

## ✅ Fuentes EAN13 Recomendadas

### **Opción 1: Libre Barcode EAN13 (GRATIS - Recomendada)**

**Fuente:** `Libre Barcode EAN13 Text`
**Licencia:** Open Font License (OFL) - Gratis y libre
**Descarga:** https://fonts.google.com/specimen/Libre+Barcode+EAN13+Text

**Instalación:**
1. Descargar fuente desde Google Fonts
2. Descomprimir el archivo ZIP
3. Hacer doble clic en `LibreBarcodeEAN13Text-Regular.ttf`
4. Clic en "Instalar"
5. Reiniciar la aplicación VB6

**Configuración en el código:**
```vb
Printer.FontName = "Libre Barcode EAN13 Text"
```

---

### **Opción 2: IDAutomation EAN13 (COMERCIAL)**

**Fuente:** `IDAutomation EAN13`
**Licencia:** Comercial (de pago)
**Web:** https://www.idautomation.com/

**Ventajas:**
- Soporte profesional
- Documentación completa
- Múltiples variantes

**Configuración en el código:**
```vb
Printer.FontName = "IDAutomation EAN13"
```

---

### **Opción 3: Code128 como alternativa**

Si no puedes instalar fuentes EAN13, puedes usar **Code128** que soporta números y es más universal:

**Fuente:** `Libre Barcode 128 Text`
**Descarga:** https://fonts.google.com/specimen/Libre+Barcode+128+Text

**Configuración en el código:**
```vb
Printer.FontName = "Libre Barcode 128 Text"
```

⚠️ **NOTA:** Code128 no es EAN13 estándar, pero funciona con la mayoría de lectores.

---

## 🔧 Configuración Actual

El código actualmente usa:
```vb
Printer.FontName = "Libre Barcode EAN13 Text"
```

Si esta fuente NO está instalada, el código imprimirá los números en **Arial** (legibles pero no escaneables).

---

## 📝 Cambiar la Fuente en el Código

Para cambiar la fuente usada, edita el archivo `frmetiquetasPS.frm`, línea **~370**:

```vb
' Cambiar esta línea:
Printer.FontName = "Libre Barcode EAN13 Text"

' Por tu fuente preferida:
Printer.FontName = "IDAutomation EAN13"  ' O la que tengas instalada
```

---

## ✅ Verificar Fuentes Instaladas

Para ver qué fuentes de códigos de barras tienes instaladas en Windows:

1. Abrir **Panel de Control**
2. Ir a **Apariencia y personalización** → **Fuentes**
3. Buscar fuentes que contengan "Barcode", "EAN", "Code128", etc.

---

## 🧪 Probar Códigos de Barras

Después de instalar la fuente:

1. Imprimir etiquetas de prueba
2. Usar lector de códigos de barras
3. Verificar que lee correctamente el EAN13

**Ejemplo de EAN13 válido:** `5901234123457`

---

## 🎨 Formato EAN13

- **Longitud:** 13 dígitos exactos
- **Sin espacios ni guiones**
- **Sin asteriscos** (a diferencia de Code 39)
- **Checksum incluido** (último dígito)

**Ejemplos válidos:**
```
8437016850015
8411082502016
5901234123457
```

**Ejemplos inválidos:**
```
*8437016850015*    ❌ (no usar asteriscos con EAN13)
843701685001       ❌ (solo 12 dígitos)
8437-0168-50015    ❌ (no usar guiones)
```

---

## 📞 Soporte

Si después de instalar la fuente los códigos no son escaneables:

1. **Verificar** que la fuente está instalada correctamente
2. **Reiniciar** la aplicación VB6
3. **Verificar** que el código EAN13 es válido (13 dígitos)
4. **Probar** con diferentes tamaños de fuente (18-30 puntos)
5. **Verificar** que la impresora tiene suficiente resolución (mínimo 300 DPI)

---

## 🔗 Enlaces Útiles

- **Google Fonts (gratis):** https://fonts.google.com/?query=barcode
- **IDAutomation (comercial):** https://www.idautomation.com/
- **Validador EAN13:** https://www.gs1.org/services/check-digit-calculator

---

**Última actualización:** Enero 2026
