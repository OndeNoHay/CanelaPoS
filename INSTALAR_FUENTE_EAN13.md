# 📥 INSTALAR FUENTE DE CÓDIGOS DE BARRAS PARA ETIQUETAS ESCANEABLES

## ⚠️ IMPORTANTE - PROBLEMA DETECTADO CON EAN13

**ACTUALIZACIÓN:** La fuente "Libre Barcode EAN13 Text" tiene un problema:
- ❌ **Solo funciona si el EAN13 tiene checksum VÁLIDO**
- ❌ Si el checksum es incorrecto, muestra barras verticales genéricas
- ❌ Muchos productos pueden tener EAN13 sin checksum válido

### ✅ SOLUCIÓN RECOMENDADA: Usar Code 128

**Usa "Libre Barcode 128 Text" en lugar de EAN13:**
- ✅ NO requiere checksum específico
- ✅ Funciona con CUALQUIER número
- ✅ Escaneable por la mayoría de lectores
- ✅ Más flexible y confiable

---

## 🎯 Comparación de Fuentes

### Libre Barcode 128 Text (RECOMENDADA) ⭐
- ✅ Acepta cualquier número
- ✅ No valida checksum
- ✅ Funciona siempre
- ✅ Compatible con todos los lectores modernos
- ✅ **Esta es la fuente que debes instalar**

### Libre Barcode EAN13 Text (Problemática)
- ⚠️ Solo funciona con EAN13 con checksum VÁLIDO
- ⚠️ Si el checksum no es correcto → barras genéricas
- ⚠️ Puede fallar con productos de Prestashop
- ❌ **NO recomendada**

---

## 🎯 SOLUCIÓN: Instalar Libre Barcode 128 Text

---

## 📥 PASOS DE INSTALACIÓN

### **Paso 1: Descargar Libre Barcode 128 Text**

1. Abre tu navegador
2. Ve a: **https://fonts.google.com/specimen/Libre+Barcode+128+Text**
3. Haz clic en el botón **"Download family"** (esquina superior derecha)
4. Se descargará un archivo ZIP llamado `Libre_Barcode_128_Text.zip`

---

### **Paso 2: Instalar en Windows**

**Método 1: Doble clic (más fácil)**

1. Abre la carpeta de Descargas
2. Busca el archivo **`Libre_Barcode_128_Text.zip`**
3. Haz doble clic para abrir el ZIP
4. Dentro verás un archivo: **`LibreBarcode128Text-Regular.ttf`**
5. Haz **doble clic** en el archivo .ttf
6. Se abrirá una ventana de vista previa
7. Haz clic en el botón **"Instalar"** (arriba a la izquierda)
8. Espera unos segundos hasta que diga "Fuente instalada"
9. ✅ **¡Listo!**

**Método 2: Copiar a carpeta de fuentes**

1. Extrae el archivo .ttf del ZIP
2. Abre **Panel de Control** → **Apariencia y personalización** → **Fuentes**
3. Arrastra el archivo .ttf a la ventana de Fuentes
4. Windows lo instalará automáticamente
5. ✅ **¡Listo!**

---

### **Paso 3: Reiniciar la Aplicación**

**MUY IMPORTANTE:**

1. **Cierra completamente** el programa VB6 (si está abierto)
2. **Cierra completamente** la aplicación CanelaPoS (si está ejecutándose)
3. Vuelve a abrir la aplicación
4. Ahora los códigos de barras deberían funcionar

---

### **Paso 4: Probar**

1. Abre el formulario de etiquetas
2. Busca productos (ejemplo: IDs 1 al 5)
3. Imprime una etiqueta de prueba
4. **Escanea el código de barras con tu lector**
5. ✅ Debería leer el EAN13 correctamente

---

## 🔍 Verificar que la Fuente Está Instalada

Para confirmar que la fuente se instaló correctamente:

1. Abre **Panel de Control**
2. Ve a **Apariencia y personalización** → **Fuentes**
3. Busca en la lista: **Libre Barcode 128 Text**
4. Si aparece = ✅ Está instalada
5. Si NO aparece = ❌ Repite la instalación

### Probar la fuente en Word

1. Abre Microsoft Word
2. Escribe cualquier número (ej: `1234567890123`)
3. Selecciona el texto
4. Cambia la fuente a **"Libre Barcode 128 Text"**
5. Deberías ver un código de barras con barras de diferentes anchos
6. ✅ Si se ve correcto, la fuente funciona

---

## 🖼️ Comparación Visual

### **SIN la fuente (incorrecto):**
```
|||||||||||||||||||||||
1234567890123
```
- Líneas todas iguales
- No escaneable
- Usa Arial (texto normal)

### **CON Libre Barcode 128 Text (correcto):**
```
| || ||| || | ||| | || ||
1234567890123
```
- Líneas con diferentes anchos
- ✅ Escaneable
- ✅ Funciona con cualquier número
- Código de barras real

### **Problema con EAN13 (barras genéricas):**
```
|||||||||||||||||||||||
1234567890789
```
- Si el checksum EAN13 no es válido
- La fuente EAN13 muestra barras genéricas
- ❌ No escaneable
- **Por eso recomendamos Code 128**

---

## ❓ Preguntas Frecuentes

### **P: ¿La fuente es gratis?**
R: Sí, **Libre Barcode 128 Text** es completamente gratuita y de código abierto (Open Font License).

### **P: ¿Funciona con cualquier lector de códigos de barras?**
R: Sí, Code 128 es compatible con prácticamente todos los lectores de códigos de barras modernos.

### **P: ¿Por qué Code 128 en lugar de EAN13?**
R: Code 128 NO requiere checksum específico. EAN13 solo funciona si el último dígito es el checksum válido, y muchos productos pueden tener EAN13 sin el checksum correcto.

### **P: ¿Necesito instalarla en cada PC?**
R: Sí, cada computadora que vaya a imprimir etiquetas necesita tener la fuente instalada.

### **P: ¿Y si no puedo instalar fuentes (permisos de administrador)?**
R: Necesitas permisos de administrador para instalar fuentes en Windows. Contacta con tu administrador de sistemas.

### **P: ¿Hay alternativas comerciales?**
R: Sí, hay opciones de pago:
- **IDAutomation Code 128** (comercial)
- **ConnectCode Barcode Software** (comercial)
- Pero **Libre Barcode 128 Text** es gratis y funciona perfectamente

### **P: ¿Puedo usar la fuente EAN13 que ya instalé?**
R: Solo si tus productos de Prestashop tienen EAN13 con checksum válido. Para evitar problemas, mejor usa Code 128.

### **P: Los códigos siguen sin funcionar después de instalar la fuente**
R: Verifica:
1. ¿Cerraste y reabriste la aplicación?
2. ¿La fuente aparece en Panel de Control → Fuentes?
3. ¿El código EAN13 tiene exactamente 13 dígitos?
4. ¿Tu lector de códigos está configurado para leer EAN13?

---

## 🔗 Enlaces Directos

### Fuente Recomendada (Code 128)
- **Descargar:** https://fonts.google.com/specimen/Libre+Barcode+128+Text
- **Vista previa online:** https://fonts.google.com/specimen/Libre+Barcode+128+Text

### Fuente Alternativa (EAN13 - requiere checksum válido)
- **Descargar:** https://fonts.google.com/specimen/Libre+Barcode+EAN13+Text
- ⚠️ Solo usar si tus EAN13 tienen checksum válido

### Más información
- **Google Fonts (más fuentes):** https://fonts.google.com/?query=barcode
- **Documentación completa:** Ver archivo `FUENTES_CODIGOS_BARRAS.md`

---

## 📞 Soporte

Si después de seguir todos los pasos los códigos siguen sin funcionar:

1. Verifica que la fuente está instalada (Panel de Control → Fuentes)
2. Reinicia el PC (a veces Windows necesita reinicio completo)
3. Prueba imprimir desde WordPad con la fuente "Libre Barcode EAN13 Text"
4. Si funciona en WordPad pero no en la app, hay un problema con el código VB6

---

**Última actualización:** Enero 2026
**Versión:** 1.0
