# 📥 INSTALAR FUENTE EAN13 PARA CÓDIGOS DE BARRAS ESCANEABLES

## ⚠️ IMPORTANTE

Si los códigos de barras en las etiquetas se ven como **líneas verticales iguales** y el scanner **NO puede leerlos**, es porque **NO tienes instalada la fuente EAN13**.

Sin la fuente correcta:
- ❌ El código se imprime en Arial (texto normal)
- ❌ El scanner no puede leer el código
- ❌ Las líneas se ven todas iguales

Con la fuente correcta:
- ✅ Código de barras escaneable
- ✅ El scanner lee el EAN13 perfectamente
- ✅ Líneas con diferentes anchos (código válido)

---

## 🎯 SOLUCIÓN: Instalar Fuente EAN13 (GRATIS)

### **Paso 1: Descargar la Fuente**

Opción más fácil y gratuita: **Libre Barcode EAN13 Text**

1. Abre tu navegador
2. Ve a: **https://fonts.google.com/specimen/Libre+Barcode+EAN13+Text**
3. Haz clic en el botón **"Download family"** (esquina superior derecha)
4. Se descargará un archivo ZIP

---

### **Paso 2: Instalar en Windows**

**Método 1: Doble clic (más fácil)**

1. Abre la carpeta de Descargas
2. Busca el archivo **`Libre_Barcode_EAN13_Text.zip`**
3. Haz doble clic para abrir el ZIP
4. Dentro verás un archivo: **`LibreBarcodeEAN13Text-Regular.ttf`**
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
3. Busca en la lista: **Libre Barcode EAN13 Text**
4. Si aparece = ✅ Está instalada
5. Si NO aparece = ❌ Repite la instalación

---

## 🖼️ Comparación Visual

### **SIN la fuente (incorrecto):**
```
|||||||||||||||||||||||
2808408419187
```
- Líneas todas iguales
- No escaneable
- Usa Arial (texto normal)

### **CON la fuente (correcto):**
```
| || ||| || | ||| | || ||
2808408419187
```
- Líneas con diferentes anchos
- ✅ Escaneable
- Código de barras real

---

## ❓ Preguntas Frecuentes

### **P: ¿La fuente es gratis?**
R: Sí, **Libre Barcode EAN13 Text** es completamente gratuita y de código abierto (Open Font License).

### **P: ¿Funciona con cualquier lector de códigos de barras?**
R: Sí, funciona con cualquier lector que soporte EAN13 (que es el estándar).

### **P: ¿Necesito instalarla en cada PC?**
R: Sí, cada computadora que vaya a imprimir etiquetas necesita tener la fuente instalada.

### **P: ¿Y si no puedo instalar fuentes (permisos de administrador)?**
R: Necesitas permisos de administrador para instalar fuentes en Windows. Contacta con tu administrador de sistemas.

### **P: ¿Hay alternativas?**
R: Sí, puedes usar otras fuentes EAN13:
- **IDAutomation EAN13** (comercial, de pago)
- **Code EAN13** (comercial, de pago)
- Pero **Libre Barcode EAN13 Text** es gratis y funciona perfectamente

### **P: Los códigos siguen sin funcionar después de instalar la fuente**
R: Verifica:
1. ¿Cerraste y reabriste la aplicación?
2. ¿La fuente aparece en Panel de Control → Fuentes?
3. ¿El código EAN13 tiene exactamente 13 dígitos?
4. ¿Tu lector de códigos está configurado para leer EAN13?

---

## 🔗 Enlaces Directos

- **Descargar fuente:** https://fonts.google.com/specimen/Libre+Barcode+EAN13+Text
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
