-- ========================================================================
-- SCRIPT SQL PARA CREAR TABLAS DE INTEGRACIÓN PRESTASHOP
-- ========================================================================
-- Base de datos: canela.mdb (Microsoft Access)
-- Fecha: 19/12/2025
-- Propósito: Crear tablas para caché y sincronización con PrestaShop 8.1
--
-- INSTRUCCIONES:
-- 1. Abrir canela.mdb en Microsoft Access
-- 2. Ir a: Crear > Diseño de consulta > Cerrar ventana de agregar tablas
-- 3. Ver > Vista SQL
-- 4. Copiar y pegar cada bloque CREATE TABLE por separado
-- 5. Ejecutar (icono !)
-- 6. Repetir para cada tabla
-- ========================================================================

-- ========================================================================
-- TABLA 1: ConfigAPI - Configuración de la API
-- ========================================================================
-- Almacena parámetros de configuración para evitar hardcodear valores
-- en el código VB6

CREATE TABLE ConfigAPI (
    Clave TEXT(50) CONSTRAINT PK_ConfigAPI PRIMARY KEY,
    Valor MEMO,
    FechaModificacion DATETIME DEFAULT Now()
);

-- Insertar datos iniciales
-- IMPORTANTE: Cambiar la URL cuando tengas el bridge instalado en tu servidor
INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion) VALUES
    ('API_BRIDGE_URL', 'https://www.canelamoda.es/api_bridge/bridge.php', Now());

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion) VALUES
    ('API_TIMEOUT', '30', Now());

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion) VALUES
    ('SYNC_ENABLED', 'True', Now());

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion) VALUES
    ('DEBUG_MODE', 'False', Now());

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion) VALUES
    ('CACHE_EXPIRATION_MINUTES', '60', Now());

INSERT INTO ConfigAPI (Clave, Valor, FechaModificacion) VALUES
    ('LAST_SYNC', '', Now());


-- ========================================================================
-- TABLA 2: ProductosPS - Caché de productos de PrestaShop
-- ========================================================================
-- Almacena información de productos consultados para funcionamiento offline
-- y mejor rendimiento

CREATE TABLE ProductosPS (
    IDProductoPS LONG CONSTRAINT PK_ProductosPS PRIMARY KEY,
    Referencia TEXT(50) NOT NULL,
    EAN13 TEXT(13),
    Nombre TEXT(255),
    Descripcion MEMO,
    PrecioSinIVA CURRENCY DEFAULT 0,
    PrecioConIVA CURRENCY DEFAULT 0,
    IVA INTEGER DEFAULT 21,
    StockPS LONG DEFAULT 0,
    StockLocal LONG DEFAULT 0,
    DiferenciaStock LONG DEFAULT 0,
    UltimaConsulta DATETIME,
    UltimaActualizacion DATETIME,
    EstadoSync TEXT(20) DEFAULT 'OK',
    URLImagen TEXT(255),
    Activo YESNO DEFAULT True
);

-- Crear índice único en Referencia
CREATE UNIQUE INDEX idx_referencia ON ProductosPS(Referencia);

-- Crear índice en EstadoSync para consultas rápidas
CREATE INDEX idx_estado ON ProductosPS(EstadoSync);


-- ========================================================================
-- TABLA 3: LogSincronizacion - Auditoría de operaciones API
-- ========================================================================
-- Registra todas las peticiones a PrestaShop para debugging y auditoría

CREATE TABLE LogSincronizacion (
    ID AUTOINCREMENT CONSTRAINT PK_LogSincronizacion PRIMARY KEY,
    FechaHora DATETIME DEFAULT Now(),
    TipoOperacion TEXT(50),
    IDProductoPS LONG,
    Referencia TEXT(50),
    Descripcion MEMO,
    RespuestaAPI MEMO,
    CodigoHTTP INTEGER,
    TiempoRespuesta INTEGER,
    UsuarioVB TEXT(50)
);

-- Índice en FechaHora para consultas por fecha
CREATE INDEX idx_fecha ON LogSincronizacion(FechaHora);

-- Índice en TipoOperacion para filtrar por tipo
CREATE INDEX idx_tipo ON LogSincronizacion(TipoOperacion);


-- ========================================================================
-- TABLA 4: MapeoArticulosPS - Mapeo entre artículos locales y PrestaShop
-- ========================================================================
-- Relaciona IDs locales (articulos.idart) con IDs de PrestaShop

CREATE TABLE MapeoArticulosPS (
    IDArticuloLocal LONG CONSTRAINT PK_MapeoArticulosPS PRIMARY KEY,
    IDProductoPS LONG,
    Referencia TEXT(50),
    FechaMapeo DATETIME DEFAULT Now(),
    MapeadoPor TEXT(50)
);

-- Índice en IDProductoPS para búsquedas inversas
CREATE INDEX idx_idproductops ON MapeoArticulosPS(IDProductoPS);


-- ========================================================================
-- TABLA 5: ColaSyncStock - Cola de sincronización offline (FASE 2)
-- ========================================================================
-- Esta tabla se usará en Fase 2 cuando implementemos escritura
-- Por ahora la creamos vacía para tener la estructura lista

CREATE TABLE ColaSyncStock (
    ID AUTOINCREMENT CONSTRAINT PK_ColaSyncStock PRIMARY KEY,
    IDVenta LONG,
    IDProductoPS LONG,
    Referencia TEXT(50),
    CantidadVendida INTEGER DEFAULT 1,
    FechaVenta DATETIME,
    Procesado YESNO DEFAULT False,
    FechaProcesado DATETIME,
    Reintentos INTEGER DEFAULT 0,
    ErrorMensaje MEMO
);

-- Índice en Procesado para consultas de cola pendiente
CREATE INDEX idx_procesado ON ColaSyncStock(Procesado);

-- Índice en FechaVenta para ordenar cola
CREATE INDEX idx_fecha_venta ON ColaSyncStock(FechaVenta);


-- ========================================================================
-- VERIFICACIÓN DE TABLAS CREADAS
-- ========================================================================
-- Ejecutar esta consulta para verificar que todas las tablas existen:
--
-- SELECT MSysObjects.Name
-- FROM MSysObjects
-- WHERE MSysObjects.Type=1
--   AND MSysObjects.Name IN ('ConfigAPI','ProductosPS','LogSincronizacion','MapeoArticulosPS','ColaSyncStock')
-- ORDER BY MSysObjects.Name;
--
-- Deberías ver 5 resultados

-- ========================================================================
-- CONSULTA DE PRUEBA
-- ========================================================================
-- Verifica que ConfigAPI tiene los valores correctos:
-- SELECT * FROM ConfigAPI ORDER BY Clave;

-- ========================================================================
-- FIN DEL SCRIPT
-- ========================================================================
