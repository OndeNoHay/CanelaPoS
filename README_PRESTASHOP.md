# CanelaPoS - Integración PrestaShop

## 🎯 Resumen del Proyecto

Integración completa entre el sistema POS legacy (Visual Basic 6 + Microsoft Access) y PrestaShop mediante API Bridge, permitiendo búsqueda de productos y sincronización automática de stock.

**Fecha de implementación:** 29 de Diciembre de 2025
**Desarrollado por:** Claude Code
**Branch:** `claude/setup-api-bridge-gj7BX`

---

## ✨ Características Implementadas

### 1. Búsqueda de Productos en PrestaShop
- Búsqueda automática al escanear código/EAN en el POS
- Detección de productos con combinaciones (tallas, colores)
- Creación temporal de artículos en BD local
- Mapeo automático de precios (con/sin IVA)
- Fallback a BD local si no se encuentra en PrestaShop

### 2. Sincronización de Stock
- Actualización automática de stock después de cada venta
- Soporte para productos simples y con combinaciones
- Manejo inteligente de errores (no bloquea ventas)
- Logging completo de todas las operaciones

### 3. Sistema de Configuración
- Archivo INI para configuración flexible
- Activación/desactivación de la integración sin cambiar código
- Timeouts configurables
- Modo debug para troubleshooting

### 4. Sistema de Logging
- Logs rotativos diarios
- Niveles: INFO, WARNING, ERROR, DEBUG
- Retención automática de 30 días
- Logs de búsquedas, ventas y sincronización

---

## 📁 Estructura de Archivos

```
CanelaPoS/
├── ModuloPrestaShop.bas       # API Bridge communication
├── ModuloLog.bas               # Logging system
├── ModuloConfig.bas            # Configuration management
├── ModuloIntegracion.bas       # Integration orchestration
├── frmventa.frm                # [MODIFICADO] Sales form
├── Module1.bas                 # [EXISTENTE] Global variables
├── config/
│   └── prestashop.ini          # [AUTO-CREADO] Configuration
├── logs/
│   └── prestashop_YYYYMMDD.log # [AUTO-CREADO] Daily logs
├── GUIA_INTEGRACION_PRESTASHOP.md    # Guía técnica detallada
└── README_PRESTASHOP.md              # Este archivo
```

---

## 🚀 Instalación y Configuración

### Requisitos Previos
- Visual Basic 6.0
- Microsoft Access Database Engine
- Conexión a Internet
- API Bridge configurado en `https://www.canelamoda.es/api_bridge/`

### Pasos de Instalación

1. **Compilar el proyecto VB6**
   - Abrir proyecto en VB6
   - Compilar ejecutable o ejecutar en modo debug

2. **Configuración automática**
   - Al ejecutar por primera vez, se crea `config/prestashop.ini`
   - Al ejecutar por primera vez, se crea la carpeta `logs/`

3. **Verificar configuración** (opcional)
   - Editar `config/prestashop.ini` si necesitas cambiar parámetros
   - Por defecto, la integración está ACTIVADA

### Configuración del API Bridge

El API Bridge debe estar configurado en el servidor con:
- API Key válida de PrestaShop (ya configurada en el servidor)
- Endpoints para búsqueda, consulta y actualización de stock

---

## 🔧 Uso del Sistema

### Flujo Normal de Venta

1. **Escanear código de producto**
   - El sistema busca primero en PrestaShop
   - Si encuentra, crea artículo temporal y muestra datos
   - Si no encuentra, busca en BD local (comportamiento normal)

2. **Completar venta**
   - Agregar cliente si es necesario
   - Seleccionar forma de pago
   - Hacer clic en "Cobrar" o "Venta"

3. **Sincronización automática**
   - El sistema actualiza stock en PrestaShop
   - Elimina artículos temporales de BD local
   - Registra operación en log

### Cancelar Venta

Si se cancela una venta:
- Los artículos temporales de PrestaShop se eliminan
- No se actualiza stock
- Se registra cancelación en log

---

## ⚙️ Configuración Avanzada

### Archivo: config/prestashop.ini

```ini
[General]
IntegracionHabilitada=1          # 1=Activo, 0=Desactivado
BuscarEnPrestaShop=1             # 1=Buscar en PS, 0=Solo local
ActualizarStockAutomatico=1      # 1=Sincronizar, 0=No sincronizar
MostrarMensajesError=0           # 1=Mostrar, 0=Solo log
TimeoutSegundos=30               # Timeout API
LogHabilitado=1                  # 1=Activar logs, 0=Desactivar
ModoDebug=0                      # 1=Debug detallado, 0=Normal

[API]
URLBridge=https://www.canelamoda.es/api_bridge/
```

### Desactivar Integración Temporalmente

Si necesitas desactivar la integración sin modificar código:

1. Abrir `config/prestashop.ini`
2. Cambiar `IntegracionHabilitada=0`
3. Guardar archivo
4. Reiniciar aplicación

El sistema funcionará 100% en modo local.

---

## 📊 Monitorización

### Ver Logs

Los logs se guardan en:
```
logs/prestashop_YYYYMMDD.log
```

Ejemplo de contenido:
```
[2025-12-29 14:23:15] [INFO] Sistema de integración PrestaShop iniciado
[2025-12-29 14:23:45] [INFO] BÚSQUEDA - Código: 12345 | Encontrado: SÍ
[2025-12-29 14:24:10] [INFO] Artículo creado desde PrestaShop - ID Local: -7890001
[2025-12-29 14:25:30] [INFO] SYNC STOCK - Producto PS: 789 | Stock: 5→4 | Éxito: SÍ
```

### Estadísticas

Para ver estadísticas de uso:
- Revisar logs diarios
- Buscar líneas con "BÚSQUEDA" para productos consultados
- Buscar líneas con "SYNC STOCK" para sincronizaciones
- Buscar líneas con "ERROR" para problemas

---

## 🐛 Resolución de Problemas

### Problema: No encuentra productos en PrestaShop

**Causas posibles:**
- Integración desactivada en configuración
- Sin conexión a Internet
- API Bridge no responde
- Código no existe en PrestaShop

**Solución:**
1. Verificar `IntegracionHabilitada=1` en INI
2. Verificar `BuscarEnPrestaShop=1` en INI
3. Revisar log para ver errores específicos
4. Verificar que el producto exista en PrestaShop admin

### Problema: Stock no se actualiza

**Causas posibles:**
- Actualización automática desactivada
- Error de permisos en API
- Timeout en la conexión

**Solución:**
1. Verificar `ActualizarStockAutomatico=1` en INI
2. Revisar log - buscar "SYNC STOCK"
3. Verificar permisos de API Key en PrestaShop
4. Aumentar `TimeoutSegundos` si hay timeouts

### Problema: Errores de conexión frecuentes

**Solución:**
1. Aumentar timeout a 60 segundos
2. Verificar estabilidad de conexión a Internet
3. Verificar que servidor PrestaShop responde
4. Activar `ModoDebug=1` para más información

### Problema: Aplicación lenta

**Solución:**
1. Reducir timeout a 15-20 segundos
2. Verificar velocidad de respuesta del API Bridge
3. Considerar cachear productos frecuentes

---

## 🔒 Seguridad

### API Key
- La API Key está almacenada en el servidor (`api_bridge.php`)
- NO se envía ni almacena en el cliente VB6
- Cambiar API Key solo en el servidor, no en VB6

### Logs
- Los logs pueden contener información sensible
- NO compartir logs públicamente
- Revisar logs regularmente y eliminar antiguos manualmente si necesario

### Base de Datos
- Hacer backup regular de la BD Access
- Los artículos temporales (ID negativo) no deben editarse
- Los artículos temporales se limpian automáticamente

---

## 📈 Rendimiento

### Optimizaciones Implementadas

- **Artículos temporales:** Se crean con ID negativos para evitar conflictos
- **Timeout configurables:** Evita bloqueos largos
- **Fail-safe:** Errores de API no bloquean ventas locales
- **Logs rotativos:** Eliminación automática de logs antiguos

### Métricas Esperadas

- **Tiempo de búsqueda:** < 2 segundos (depende de conexión)
- **Tiempo de sincronización:** < 1 segundo por producto
- **Tamaño de logs:** ~100KB por día (aprox)

---

## 🧪 Testing

### Casos de Prueba Recomendados

1. **Búsqueda exitosa**
   - Escanear código existente en PrestaShop
   - Verificar datos correctos (nombre, precio, stock)
   - Completar venta
   - Verificar actualización de stock en PrestaShop

2. **Búsqueda fallida**
   - Escanear código NO existente en PrestaShop
   - Verificar que busca en BD local
   - Completar venta normalmente

3. **Sin conexión**
   - Desconectar Internet
   - Escanear cualquier código
   - Verificar que funciona en modo local
   - Verificar error registrado en log

4. **Cancelación de venta**
   - Escanear producto de PrestaShop
   - Hacer clic en "Borrar Datos"
   - Verificar que artículo temporal se elimina

5. **Producto con combinaciones**
   - Escanear producto con tallas
   - Verificar que muestra combinación correcta
   - Completar venta
   - Verificar actualización de stock de combinación específica

---

## 📞 Soporte Técnico

### Información para Debugging

Cuando reportes un problema, incluye:
1. Contenido del archivo `config/prestashop.ini`
2. Últimas 50 líneas del log del día (archivo .log en `logs/`)
3. Descripción del problema paso a paso
4. Código del producto que causó el problema

### Archivos de Configuración de Ejemplo

**Modo Producción:**
```ini
IntegracionHabilitada=1
BuscarEnPrestaShop=1
ActualizarStockAutomatico=1
MostrarMensajesError=0
LogHabilitado=1
ModoDebug=0
```

**Modo Debug:**
```ini
IntegracionHabilitada=1
BuscarEnPrestaShop=1
ActualizarStockAutomatico=1
MostrarMensajesError=1
LogHabilitado=1
ModoDebug=1
```

**Modo Solo Local (Desactivado):**
```ini
IntegracionHabilitada=0
BuscarEnPrestaShop=0
ActualizarStockAutomatico=0
```

---

## 🔄 Actualizaciones Futuras

### Posibles Mejoras

- [ ] Cache de productos frecuentes
- [ ] Sincronización batch de múltiples productos
- [ ] Interfaz gráfica para configuración
- [ ] Estadísticas en tiempo real
- [ ] Integración con clientes de PrestaShop
- [ ] Sincronización bidireccional (PS → POS)

---

## 📝 Notas Técnicas

### Compatibilidad
- VB6 Runtime requerido
- MSXML2.ServerXMLHTTP.6.0 (incluido en Windows)
- DAO 3.6 (Microsoft Access Database Engine)

### Limitaciones Conocidas
- Parseo JSON simplificado (sin librería externa)
- Solo soporta actualización de stock (no precios)
- No sincroniza nuevos productos de PS a POS automáticamente
- Requiere conexión activa a Internet para búsqueda en PS

### Arquitectura
- **Patrón:** Fail-safe wrapper pattern
- **Conexión HTTP:** MSXML2.ServerXMLHTTP.6.0
- **Parseo JSON:** Custom (simplificado)
- **Storage:** Access MDB + Archivo INI

---

## 📜 Licencia y Créditos

**Proyecto:** CanelaPoS - Sistema POS
**Integración PrestaShop:** Desarrollado por Claude Code
**Cliente:** Canela Moda
**Fecha:** Diciembre 2025

---

## 📚 Documentación Adicional

Para información técnica detallada, consultar:
- **GUIA_INTEGRACION_PRESTASHOP.md** - Guía técnica completa
- **estructura_bd_20251219_181525.md** - Esquema de base de datos
- **Código fuente** - Módulos VB6 comentados

---

**¿Preguntas?** Consulta primero los logs y la guía de integración.

**¡Buena suerte con la integración!** 🚀
