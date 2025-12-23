<?php
/**
 * ========================================================================
 * CONFIGURACIÓN DEL API BRIDGE - PRESTASHOP 8.1
 * ========================================================================
 * Versión corregida - 19/12/2025
 * ========================================================================
 */

// URL base de la API de PrestaShop (CON BARRA FINAL)
define('PRESTASHOP_API_URL', 'https://www.canelamoda.es/api/');

// API Key de PrestaShop (32 caracteres)
// IMPORTANTE: Reemplaza esto con tu API Key real
define('PRESTASHOP_API_KEY', 'TU_API_KEY_DE_32_CARACTERES_AQUI');

// Idioma por defecto para nombres de productos
// 1 = Español (verificar en tu instalación de PrestaShop)
define('PRESTASHOP_LANGUAGE_ID', 1);

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
// VERIFICACIÓN DE CONFIGURACIÓN - VERSIÓN CORREGIDA
// ========================================================================
function verificarConfiguracion() {
    $errores = [];

    // Limpiar espacios en blanco de la API Key
    $apiKey = defined('PRESTASHOP_API_KEY') ? trim(PRESTASHOP_API_KEY) : '';

    // Verificar que no sea el valor por defecto
    if (empty($apiKey) || $apiKey == 'TU_API_KEY_DE_32_CARACTERES_AQUI' || $apiKey == 'TU_API_KEY_AQUI_32_CARACTERES') {
        $errores[] = 'API Key no configurada';
    }

    // Verificar longitud de 32 caracteres
    if (!empty($apiKey) && strlen($apiKey) != 32) {
        $errores[] = 'API Key debe tener 32 caracteres (tiene ' . strlen($apiKey) . ')';
    }

    // Verificar directorio de log si DEBUG_MODE está activado
    if (DEBUG_MODE && defined('LOG_FILE')) {
        $logDir = dirname(LOG_FILE);
        if (!is_writable($logDir)) {
            $errores[] = 'Directorio de log no tiene permisos de escritura: ' . $logDir;
        }
    }

    // Verificar directorio de caché
    if (CACHE_TTL > 0 && defined('CACHE_DIR')) {
        if (!is_dir(CACHE_DIR)) {
            // Intentar crear el directorio
            if (!@mkdir(CACHE_DIR, 0777, true)) {
                $errores[] = 'No se pudo crear directorio de caché: ' . CACHE_DIR;
            }
        } elseif (!is_writable(CACHE_DIR)) {
            $errores[] = 'Directorio de caché no tiene permisos de escritura: ' . CACHE_DIR;
        }
    }

    return $errores;
}
