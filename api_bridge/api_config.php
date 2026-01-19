<?php
/**
 * ========================================================================
 * CONFIGURACIÓN DEL API BRIDGE - PRESTASHOP 8.1
 * ========================================================================
 *
 * IMPORTANTE:
 * 1. Renombrar este archivo a: api_config.php
 * 2. Completar los valores a continuación
 * 3. Subir a tu servidor por FTP
 * 4. ASEGURAR que este archivo NO sea accesible públicamente
 *    (usar .htaccess o permisos del servidor)
 *
 * ========================================================================
 */

// URL base de la API de PrestaShop (CON BARRA FINAL)
define('PRESTASHOP_API_URL', 'https://canelamoda.es/api/');

// API Key de PrestaShop (32 caracteres)
// Generar desde: PrestaShop Admin > Parámetros Avanzados > Webservice > Z4IT
define('PRESTASHOP_API_KEY', 'LUV2UKQLMS8S6RGP1FKTBSWS9SK3****');

// Idioma por defecto para nombres de productos
// 1 = Español (verificar en tu instalación de PrestaShop)
define('PRESTASHOP_LANGUAGE_ID', 1);

// ID del grupo de atributos para "Talla" (Size)
// Este ID se usa para filtrar las combinaciones y mostrar solo las tallas
// Consultar en: PrestaShop Admin > Catálogo > Atributos y Características
define('SIZE_ATTRIBUTE_GROUP_ID', 5);

// Timeout para peticiones a PrestaShop (segundos)
define('API_TIMEOUT', 30);

// Activar modo DEBUG (genera archivo de log detallado)
define('DEBUG_MODE', true);

// Archivo de log (ruta relativa o absoluta)
define('LOG_FILE', __DIR__ . '/bridge_debug.log');

// Tiempo de vida del caché en segundos (0 = sin caché)
define('CACHE_TTL', 3600); // 1 hora

// Directorio para caché (debe tener permisos de escritura)
define('CACHE_DIR', __DIR__ . '/cache/');

// Zona horaria
date_default_timezone_set('Europe/Madrid');

// ========================================================================
// SEGURIDAD: Bloquear acceso directo a este archivo
// ========================================================================
if (basename($_SERVER['PHP_SELF']) == basename(__FILE__)) {
    http_response_code(403);
    die('Acceso denegado');
}

// ========================================================================
// VERIFICACIÓN DE CONFIGURACIÓN
// ========================================================================
function verificarConfiguracion() {
    $errores = [];

    if (PRESTASHOP_API_KEY == 'TU_API_KEY_AQUI_32_CARACTERES') {
        $errores[] = 'API Key no configurada';
    }

    if (strlen(PRESTASHOP_API_KEY) != 32) {
        $errores[] = 'API Key debe tener 32 caracteres';
    }

    if (DEBUG_MODE && !is_writable(dirname(LOG_FILE))) {
        $errores[] = 'Directorio de log no tiene permisos de escritura';
    }

    if (CACHE_TTL > 0 && !is_dir(CACHE_DIR)) {
        if (!mkdir(CACHE_DIR, 0755, true)) {
            $errores[] = 'No se pudo crear directorio de caché';
        }
    }

    return $errores;
}
