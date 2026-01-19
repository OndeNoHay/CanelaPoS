<?php
/**
 * Script de prueba para diagnosticar actualización de stock
 *
 * USO:
 * 1. Subir este archivo a api_bridge/
 * 2. Acceder desde navegador: https://www.canelamoda.es/api_bridge/test_stock.php
 * 3. Verificar que no haya errores de sintaxis
 */

// Incluir el archivo principal
require_once 'config.php';

echo "<h1>Test de Actualización de Stock</h1>";
echo "<pre>";

// 1. Verificar que las funciones existen
echo "1. Verificando funciones...\n";
if (function_exists('actualizarStockEnPrestaShop')) {
    echo "   ✓ actualizarStockEnPrestaShop() existe\n";
} else {
    echo "   ✗ actualizarStockEnPrestaShop() NO EXISTE\n";
}

if (function_exists('callPrestaShop')) {
    echo "   ✓ callPrestaShop() existe\n";
} else {
    echo "   ✗ callPrestaShop() NO EXISTE\n";
}

if (function_exists('handleActualizarStock')) {
    echo "   ✓ handleActualizarStock() existe\n";
} else {
    echo "   ✗ handleActualizarStock() NO EXISTE\n";
}

// 2. Verificar configuración
echo "\n2. Verificando configuración...\n";
echo "   PRESTASHOP_API_URL: " . (defined('PRESTASHOP_API_URL') ? PRESTASHOP_API_URL : 'NO DEFINIDO') . "\n";
echo "   PRESTASHOP_API_KEY: " . (defined('PRESTASHOP_API_KEY') ? (substr(PRESTASHOP_API_KEY, 0, 10) . '...') : 'NO DEFINIDO') . "\n";
echo "   DEBUG_MODE: " . (defined('DEBUG_MODE') && DEBUG_MODE ? 'ACTIVADO' : 'DESACTIVADO') . "\n";
echo "   LOG_FILE: " . (defined('LOG_FILE') ? LOG_FILE : 'NO DEFINIDO') . "\n";

// 3. Verificar permisos de escritura del log
echo "\n3. Verificando permisos de log...\n";
if (defined('LOG_FILE')) {
    if (file_exists(LOG_FILE)) {
        echo "   ✓ Log existe: " . LOG_FILE . "\n";
        echo "   Permisos: " . substr(sprintf('%o', fileperms(LOG_FILE)), -4) . "\n";
        if (is_writable(LOG_FILE)) {
            echo "   ✓ Log es escribible\n";
        } else {
            echo "   ✗ Log NO es escribible\n";
        }
    } else {
        echo "   ⚠ Log no existe aún (se creará en primera escritura)\n";
        $logDir = dirname(LOG_FILE);
        if (is_writable($logDir)) {
            echo "   ✓ Directorio es escribible: $logDir\n";
        } else {
            echo "   ✗ Directorio NO es escribible: $logDir\n";
        }
    }
}

// 4. Probar conexión a PrestaShop
echo "\n4. Probando conexión a PrestaShop...\n";
if (function_exists('callPrestaShop')) {
    $resultado = callPrestaShop('products?limit=1');
    if ($resultado['code'] == 200) {
        echo "   ✓ Conexión exitosa a PrestaShop\n";
        echo "   HTTP Code: " . $resultado['code'] . "\n";
    } else {
        echo "   ✗ Error en conexión a PrestaShop\n";
        echo "   HTTP Code: " . $resultado['code'] . "\n";
        echo "   Error: " . substr($resultado['data'], 0, 200) . "\n";
    }
}

// 5. Verificar que bridge.php maneja actualizar_stock
echo "\n5. Verificando endpoint actualizar_stock...\n";
$bridgeContent = file_get_contents('bridge.php');
if (strpos($bridgeContent, "case 'actualizar_stock':") !== false) {
    echo "   ✓ Case 'actualizar_stock' encontrado en bridge.php\n";
} else {
    echo "   ✗ Case 'actualizar_stock' NO encontrado en bridge.php\n";
}

if (strpos($bridgeContent, 'function handleActualizarStock()') !== false) {
    echo "   ✓ Función handleActualizarStock() encontrada\n";
} else {
    echo "   ✗ Función handleActualizarStock() NO encontrada\n";
}

if (strpos($bridgeContent, 'function actualizarStockEnPrestaShop') !== false) {
    echo "   ✓ Función actualizarStockEnPrestaShop() encontrada\n";
} else {
    echo "   ✗ Función actualizarStockEnPrestaShop() NO encontrada\n";
}

// 6. Verificar Content-Type header
echo "\n6. Verificando headers para PUT...\n";
if (strpos($bridgeContent, "'Content-Type: text/xml'") !== false) {
    echo "   ✓ Header 'Content-Type: text/xml' encontrado\n";
} else {
    echo "   ✗ Header 'Content-Type: text/xml' NO encontrado (NECESARIO para PUT)\n";
}

echo "\n</pre>";
echo "<hr>";
echo "<p><strong>Siguiente paso:</strong> Si todo está OK, probar endpoint POST desde VB6.</p>";
echo "<p>Para activar DEBUG_MODE, editar config.php y establecer: <code>define('DEBUG_MODE', true);</code></p>";
?>
