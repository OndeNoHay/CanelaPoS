-- ========================================================================
-- SCRIPT SQL PARA CREAR TABLAS DE INTEGRACIÓN PRESTASHOP
-- ========================================================================
-- Base de datos: canela.mdb (Microsoft Access)
-- Fecha: 19/12/2025 - VERSIÓN CORREGIDA
-- Propósito: Crear tablas para caché y sincronización con PrestaShop 8.1
--
-- INSTRUCCIONES:
-- 1. Abrir canela.mdb en Microsoft Access
-- 2. Ir a: Crear > Diseño de consulta > Cerrar ventana de agregar tablas
-- 3. Ver > Vista SQL
-- 4. Copiar y pegar CADA BLOQUE por separado (entre líneas de ====)
-- 5. Ejecutar cada bloque con el icono ! (Ejecutar)
-- 6. Repetir para TODOS los bloques
--
-- IMPORTANTE: Access no acepta sintaxis SQL estándar en algunos casos
-- Esta versión está adaptada específicamente para Access
-- ========================================================================

-- ========================================================================
-- BLOQUE 1: Tabla ConfigAPI
-- ========================================================================

CREATE TABLE ConfigAPI (
    Clave TEXT(50),
    Valor MEMO,
    FechaModificacion DATETIME,
    CONSTRAINT PK_ConfigAPI PRIMARY KEY (Clave)
);

-- ========================================================================
-- BLOQUE 2: Insertar datos en ConfigAPI
-- ========================================================================
-- Ejecutar CADA INSERT por separado

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion)
VALUES ('API_BRIDGE_URL', 'https://www.canelamoda.es/api_bridge/bridge.php', Now());

-- ========================================================================
-- BLOQUE 3: Más datos ConfigAPI
-- ========================================================================

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion)
VALUES ('API_TIMEOUT', '30', Now());

-- ========================================================================
-- BLOQUE 4: Más datos ConfigAPI
-- ========================================================================

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion)
VALUES ('SYNC_ENABLED', 'True', Now());

-- ========================================================================
-- BLOQUE 5: Más datos ConfigAPI
-- ========================================================================

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion)
VALUES ('DEBUG_MODE', 'False', Now());

-- ========================================================================
-- BLOQUE 6: Más datos ConfigAPI
-- ========================================================================

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion)
VALUES ('CACHE_EXPIRATION_MINUTES', '60', Now());

-- ========================================================================
-- BLOQUE 7: Más datos ConfigAPI
-- ========================================================================

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion)
VALUES ('LAST_SYNC', '', Now());

-- ========================================================================
-- BLOQUE 8: Tabla ProductosPS
-- ========================================================================

CREATE TABLE ProductosPS (
    IDProductoPS LONG,
    Referencia TEXT(50),
    EAN13 TEXT(13),
    Nombre TEXT(255),
    Descripcion MEMO,
    PrecioSinIVA CURRENCY,
    PrecioConIVA CURRENCY,
    IVA INTEGER,
    StockPS LONG,
    StockLocal LONG,
    DiferenciaStock LONG,
    UltimaConsulta DATETIME,
    UltimaActualizacion DATETIME,
    EstadoSync TEXT(20),
    URLImagen TEXT(255),
    Activo YESNO,
    CONSTRAINT PK_ProductosPS PRIMARY KEY (IDProductoPS)
);

-- ========================================================================
-- BLOQUE 9: Índice único en Referencia
-- ========================================================================

CREATE UNIQUE INDEX idx_referencia ON ProductosPS(Referencia);

-- ========================================================================
-- BLOQUE 10: Índice en EstadoSync
-- ========================================================================

CREATE INDEX idx_estado ON ProductosPS(EstadoSync);

-- ========================================================================
-- BLOQUE 11: Tabla LogSincronizacion
-- ========================================================================

CREATE TABLE LogSincronizacion (
    ID AUTOINCREMENT,
    FechaHora DATETIME,
    TipoOperacion TEXT(50),
    IDProductoPS LONG,
    Referencia TEXT(50),
    Descripcion MEMO,
    RespuestaAPI MEMO,
    CodigoHTTP INTEGER,
    TiempoRespuesta INTEGER,
    UsuarioVB TEXT(50),
    CONSTRAINT PK_LogSincronizacion PRIMARY KEY (ID)
);

-- ========================================================================
-- BLOQUE 12: Índice en FechaHora
-- ========================================================================

CREATE INDEX idx_fecha ON LogSincronizacion(FechaHora);

-- ========================================================================
-- BLOQUE 13: Índice en TipoOperacion
-- ========================================================================

CREATE INDEX idx_tipo ON LogSincronizacion(TipoOperacion);

-- ========================================================================
-- BLOQUE 14: Tabla MapeoArticulosPS
-- ========================================================================

CREATE TABLE MapeoArticulosPS (
    IDArticuloLocal LONG,
    IDProductoPS LONG,
    Referencia TEXT(50),
    FechaMapeo DATETIME,
    MapeadoPor TEXT(50),
    CONSTRAINT PK_MapeoArticulosPS PRIMARY KEY (IDArticuloLocal)
);

-- ========================================================================
-- BLOQUE 15: Índice en IDProductoPS
-- ========================================================================

CREATE INDEX idx_idproductops ON MapeoArticulosPS(IDProductoPS);

-- ========================================================================
-- BLOQUE 16: Tabla ColaSyncStock (para Fase 2)
-- ========================================================================

CREATE TABLE ColaSyncStock (
    ID AUTOINCREMENT,
    IDVenta LONG,
    IDProductoPS LONG,
    Referencia TEXT(50),
    CantidadVendida INTEGER,
    FechaVenta DATETIME,
    Procesado YESNO,
    FechaProcesado DATETIME,
    Reintentos INTEGER,
    ErrorMensaje MEMO,
    CONSTRAINT PK_ColaSyncStock PRIMARY KEY (ID)
);

-- ========================================================================
-- BLOQUE 17: Índice en Procesado
-- ========================================================================

CREATE INDEX idx_procesado ON ColaSyncStock(Procesado);

-- ========================================================================
-- BLOQUE 18: Índice en FechaVenta
-- ========================================================================

CREATE INDEX idx_fecha_venta ON ColaSyncStock(FechaVenta);

-- ========================================================================
-- VERIFICACIÓN: Consultar tablas creadas
-- ========================================================================
-- Ejecutar esta consulta para verificar que todas las tablas existen:
--
-- SELECT MSysObjects.Name
-- FROM MSysObjects
-- WHERE MSysObjects.Type=1
--   AND MSysObjects.Name IN ('ConfigAPI','ProductosPS','LogSincronizacion','MapeoArticulosPS','ColaSyncStock')
-- ORDER BY MSysObjects.Name;
--
-- Deberías ver 5 resultados (las 5 tablas)

-- ========================================================================
-- CONSULTA DE PRUEBA
-- ========================================================================
-- Verifica que ConfigAPI tiene los 6 valores correctos:
--
-- SELECT * FROM ConfigAPI ORDER BY Clave;
--
-- Deberías ver 6 registros:
-- 1. API_BRIDGE_URL
-- 2. API_TIMEOUT
-- 3. CACHE_EXPIRATION_MINUTES
-- 4. DEBUG_MODE
-- 5. LAST_SYNC
-- 6. SYNC_ENABLED

-- ========================================================================
-- VALORES POR DEFECTO
-- ========================================================================
-- Nota: Access no permite DEFAULT con funciones en CREATE TABLE
-- Por eso los valores por defecto se deben establecer vía VB6 o manualmente
--
-- Si necesitas establecer valores por defecto después:
--
-- Para ProductosPS:
-- UPDATE ProductosPS SET PrecioSinIVA = 0 WHERE PrecioSinIVA IS NULL;
-- UPDATE ProductosPS SET PrecioConIVA = 0 WHERE PrecioConIVA IS NULL;
-- UPDATE ProductosPS SET IVA = 21 WHERE IVA IS NULL;
-- UPDATE ProductosPS SET StockPS = 0 WHERE StockPS IS NULL;
-- UPDATE ProductosPS SET EstadoSync = 'OK' WHERE EstadoSync IS NULL;
-- UPDATE ProductosPS SET Activo = True WHERE Activo IS NULL;

-- ========================================================================
-- FIN DEL SCRIPT
-- ========================================================================
-- Total de bloques a ejecutar: 18
-- Tablas creadas: 5 (ConfigAPI, ProductosPS, LogSincronizacion, MapeoArticulosPS, ColaSyncStock)
-- Índices creados: 7
-- Registros insertados: 6 (en ConfigAPI)
-- ========================================================================
