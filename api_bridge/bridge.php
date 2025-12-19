<?php
/**
 * ========================================================================
 * API BRIDGE - PRESTASHOP 8.1 PARA POS VB6
 * ========================================================================
 *
 * Propósito: Intermediario entre VB6 y PrestaShop API
 * - Convierte peticiones HTTP simples (VB6) a XML (PrestaShop)
 * - Maneja autenticación con PrestaShop
 * - Implementa caché local para mejor rendimiento
 * - Proporciona respuestas JSON simples
 *
 * FASE 1: SOLO LECTURA
 * - Buscar producto por código (reference o EAN13)
 * - Obtener stock de producto
 * - Obtener información de producto
 *
 * Endpoints:
 * - GET  /bridge.php?action=buscar_producto&codigo=ABC-12345678
 * - GET  /bridge.php?action=obtener_stock&id=123
 * - GET  /bridge.php?action=info_producto&id=123
 * - GET  /bridge.php?action=test (verificar configuración)
 *
 * Autor: Claude Code
 * Fecha: 19/12/2025
 * ========================================================================
 */

// Cargar configuración
require_once __DIR__ . '/api_config.php';

// Configurar headers JSON
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *'); // Permitir peticiones desde cualquier origen

// Verificar configuración
$erroresConfig = verificarConfiguracion();
if (!empty($erroresConfig)) {
    responderError(500, 'Error de configuración', $erroresConfig);
}

// ========================================================================
// ENRUTADOR PRINCIPAL
// ========================================================================

$action = isset($_GET['action']) ? $_GET['action'] : '';

try {
    switch ($action) {
        case 'test':
            handleTest();
            break;

        case 'buscar_producto':
            handleBuscarProducto();
            break;

        case 'obtener_stock':
            handleObtenerStock();
            break;

        case 'info_producto':
            handleInfoProducto();
            break;

        default:
            responderError(400, 'Acción no válida', [
                'accion_recibida' => $action,
                'acciones_validas' => ['test', 'buscar_producto', 'obtener_stock', 'info_producto']
            ]);
    }
} catch (Exception $e) {
    registrarLog('ERROR', $action, $e->getMessage());
    responderError(500, 'Error interno del servidor', [
        'mensaje' => $e->getMessage(),
        'archivo' => $e->getFile(),
        'linea' => $e->getLine()
    ]);
}

// ========================================================================
// HANDLERS DE ENDPOINTS
// ========================================================================

/**
 * Test de conectividad y configuración
 */
function handleTest() {
    $inicio = microtime(true);

    // Verificar conexión a PrestaShop
    $resultado = callPrestaShop('products?limit=1');

    $tiempo = round((microtime(true) - $inicio) * 1000);

    if ($resultado['code'] == 200) {
        responderExito([
            'mensaje' => 'Conexión exitosa con PrestaShop',
            'prestashop_url' => PRESTASHOP_API_URL,
            'api_key_configurada' => strlen(PRESTASHOP_API_KEY) == 32,
            'tiempo_respuesta_ms' => $tiempo,
            'debug_mode' => DEBUG_MODE,
            'cache_enabled' => CACHE_TTL > 0,
            'php_version' => phpversion(),
            'curl_disponible' => function_exists('curl_init'),
            'timestamp' => date('Y-m-d H:i:s')
        ]);
    } else {
        responderError($resultado['code'], 'Error al conectar con PrestaShop', [
            'http_code' => $resultado['code'],
            'respuesta' => substr($resultado['data'], 0, 500),
            'tiempo_ms' => $tiempo
        ]);
    }
}

/**
 * Buscar producto por código (reference o EAN13)
 * GET /bridge.php?action=buscar_producto&codigo=ABC-12345678
 */
function handleBuscarProducto() {
    $inicio = microtime(true);

    // Validar parámetros
    if (!isset($_GET['codigo']) || empty($_GET['codigo'])) {
        responderError(400, 'Parámetro "codigo" requerido');
    }

    $codigo = trim($_GET['codigo']);
    registrarLog('BUSQUEDA', $codigo, "Buscando producto con código: $codigo");

    // Intentar buscar por reference primero
    $resultado = buscarProductoPorReference($codigo);

    if ($resultado) {
        $tiempo = round((microtime(true) - $inicio) * 1000);
        registrarLog('BUSQUEDA', $codigo, "Producto encontrado: {$resultado['nombre']}", $tiempo);
        responderExito($resultado, $tiempo);
    }

    // Si no se encontró por reference, buscar por EAN13
    $resultado = buscarProductoPorEAN13($codigo);

    if ($resultado) {
        $tiempo = round((microtime(true) - $inicio) * 1000);
        registrarLog('BUSQUEDA', $codigo, "Producto encontrado por EAN13: {$resultado['nombre']}", $tiempo);
        responderExito($resultado, $tiempo);
    }

    // Producto no encontrado
    $tiempo = round((microtime(true) - $inicio) * 1000);
    registrarLog('BUSQUEDA', $codigo, "Producto no encontrado", $tiempo);
    responderError(404, 'Producto no encontrado', [
        'codigo_buscado' => $codigo,
        'mensaje' => 'No se encontró ningún producto con ese código (reference o EAN13)'
    ]);
}

/**
 * Obtener stock de un producto
 * GET /bridge.php?action=obtener_stock&id=123
 */
function handleObtenerStock() {
    $inicio = microtime(true);

    // Validar parámetros
    if (!isset($_GET['id']) || !is_numeric($_GET['id'])) {
        responderError(400, 'Parámetro "id" numérico requerido');
    }

    $idProducto = intval($_GET['id']);
    registrarLog('STOCK', $idProducto, "Consultando stock del producto ID: $idProducto");

    // Consultar stock_availables
    $resultado = callPrestaShop("stock_availables?filter[id_product]=$idProducto&display=full");

    if ($resultado['code'] != 200) {
        responderError($resultado['code'], 'Error al consultar stock', [
            'id_producto' => $idProducto,
            'http_code' => $resultado['code']
        ]);
    }

    // Parsear XML
    try {
        $xml = simplexml_load_string($resultado['data']);

        if (!isset($xml->stock_availables->stock_available)) {
            responderError(404, 'Stock no encontrado', [
                'id_producto' => $idProducto
            ]);
        }

        $stock = $xml->stock_availables->stock_available;
        $cantidad = (int)$stock->quantity;
        $idStockAvailable = (int)$stock->id;

        $tiempo = round((microtime(true) - $inicio) * 1000);
        registrarLog('STOCK', $idProducto, "Stock obtenido: $cantidad unidades", $tiempo);

        responderExito([
            'id_producto' => $idProducto,
            'id_stock_available' => $idStockAvailable,
            'cantidad' => $cantidad,
            'disponible' => $cantidad > 0
        ], $tiempo);

    } catch (Exception $e) {
        responderError(500, 'Error al parsear respuesta de PrestaShop', [
            'error' => $e->getMessage()
        ]);
    }
}

/**
 * Obtener información completa de un producto
 * GET /bridge.php?action=info_producto&id=123
 */
function handleInfoProducto() {
    $inicio = microtime(true);

    // Validar parámetros
    if (!isset($_GET['id']) || !is_numeric($_GET['id'])) {
        responderError(400, 'Parámetro "id" numérico requerido');
    }

    $idProducto = intval($_GET['id']);
    registrarLog('INFO', $idProducto, "Consultando información del producto ID: $idProducto");

    $producto = obtenerProductoCompleto($idProducto);

    if ($producto) {
        $tiempo = round((microtime(true) - $inicio) * 1000);
        registrarLog('INFO', $idProducto, "Información obtenida: {$producto['nombre']}", $tiempo);
        responderExito($producto, $tiempo);
    } else {
        responderError(404, 'Producto no encontrado', [
            'id_producto' => $idProducto
        ]);
    }
}

// ========================================================================
// FUNCIONES DE BÚSQUEDA
// ========================================================================

/**
 * Buscar producto por campo "reference"
 */
function buscarProductoPorReference($reference) {
    // Escapar caracteres especiales para URL
    $referenceEscaped = urlencode($reference);

    // Consultar PrestaShop filtrando por reference
    $resultado = callPrestaShop("products?filter[reference]=$referenceEscaped&display=full");

    if ($resultado['code'] != 200) {
        return null;
    }

    try {
        $xml = simplexml_load_string($resultado['data']);

        if (!isset($xml->products->product)) {
            return null;
        }

        $product = $xml->products->product;
        return parsearProducto($product);

    } catch (Exception $e) {
        registrarLog('ERROR', 'buscarProductoPorReference', $e->getMessage());
        return null;
    }
}

/**
 * Buscar producto por campo "ean13"
 */
function buscarProductoPorEAN13($ean13) {
    // Escapar caracteres especiales para URL
    $ean13Escaped = urlencode($ean13);

    // Consultar PrestaShop filtrando por EAN13
    $resultado = callPrestaShop("products?filter[ean13]=$ean13Escaped&display=full");

    if ($resultado['code'] != 200) {
        return null;
    }

    try {
        $xml = simplexml_load_string($resultado['data']);

        if (!isset($xml->products->product)) {
            return null;
        }

        $product = $xml->products->product;
        return parsearProducto($product);

    } catch (Exception $e) {
        registrarLog('ERROR', 'buscarProductoPorEAN13', $e->getMessage());
        return null;
    }
}

/**
 * Obtener producto completo por ID
 */
function obtenerProductoCompleto($idProducto) {
    $resultado = callPrestaShop("products/$idProducto");

    if ($resultado['code'] != 200) {
        return null;
    }

    try {
        $xml = simplexml_load_string($resultado['data']);

        if (!isset($xml->product)) {
            return null;
        }

        return parsearProducto($xml->product);

    } catch (Exception $e) {
        registrarLog('ERROR', 'obtenerProductoCompleto', $e->getMessage());
        return null;
    }
}

/**
 * Parsear XML de producto a array JSON-friendly
 */
function parsearProducto($productXML) {
    // Extraer nombre en el idioma configurado
    $nombre = '';
    if (isset($productXML->name->language)) {
        foreach ($productXML->name->language as $lang) {
            if ((int)$lang['id'] == PRESTASHOP_LANGUAGE_ID) {
                $nombre = (string)$lang;
                break;
            }
        }
        // Si no se encontró el idioma, usar el primero disponible
        if (empty($nombre)) {
            $nombre = (string)$productXML->name->language[0];
        }
    }

    // Extraer descripción corta
    $descripcion = '';
    if (isset($productXML->description_short->language)) {
        foreach ($productXML->description_short->language as $lang) {
            if ((int)$lang['id'] == PRESTASHOP_LANGUAGE_ID) {
                $descripcion = strip_tags((string)$lang);
                break;
            }
        }
    }

    // Obtener stock
    $stock = 0;
    if (isset($productXML->associations->stock_availables->stock_available)) {
        $stockInfo = $productXML->associations->stock_availables->stock_available;
        $idStockAvailable = (int)$stockInfo->id;

        // Consultar stock_available para obtener cantidad exacta
        $resultadoStock = callPrestaShop("stock_availables/$idStockAvailable");
        if ($resultadoStock['code'] == 200) {
            $stockXML = simplexml_load_string($resultadoStock['data']);
            $stock = (int)$stockXML->stock_available->quantity;
        }
    }

    return [
        'id' => (int)$productXML->id,
        'reference' => (string)$productXML->reference,
        'ean13' => (string)$productXML->ean13,
        'nombre' => $nombre,
        'descripcion' => $descripcion,
        'precio_sin_iva' => (float)$productXML->price,
        'precio_con_iva' => (float)$productXML->price * 1.21, // IVA 21% (ajustar según configuración)
        'iva' => 21,
        'stock' => $stock,
        'activo' => (int)$productXML->active == 1,
        'url_imagen' => construirURLImagen((int)$productXML->id, $productXML),
        'fecha_consulta' => date('Y-m-d H:i:s')
    ];
}

/**
 * Construir URL de imagen del producto
 */
function construirURLImagen($idProducto, $productXML) {
    // Verificar si tiene imágenes
    if (!isset($productXML->associations->images->image)) {
        return '';
    }

    $idImagen = (int)$productXML->associations->images->image[0]->id;

    // Construir URL según estructura de PrestaShop
    // Formato: /img/p/1/2/3/123.jpg (donde 123 es el ID de imagen)
    $idStr = (string)$idImagen;
    $path = implode('/', str_split($idStr));

    return "https://www.canelamoda.es/ps/img/p/$path/$idImagen.jpg";
}

// ========================================================================
// FUNCIONES DE COMUNICACIÓN CON PRESTASHOP
// ========================================================================

/**
 * Realizar petición HTTP a PrestaShop API
 */
function callPrestaShop($resource, $method = 'GET', $data = null) {
    $url = PRESTASHOP_API_URL . $resource;

    // Inicializar cURL
    $ch = curl_init();

    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_TIMEOUT, API_TIMEOUT);
    curl_setopt($ch, CURLOPT_USERPWD, PRESTASHOP_API_KEY . ':');
    curl_setopt($ch, CURLOPT_HTTPHEADER, [
        'Output-Format: XML'
    ]);

    // SSL (producción debe verificar certificados)
    curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, true);
    curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, 2);

    // Método HTTP
    if ($method == 'POST' || $method == 'PUT') {
        curl_setopt($ch, CURLOPT_CUSTOMREQUEST, $method);
        curl_setopt($ch, CURLOPT_POSTFIELDS, $data);
    }

    // Ejecutar petición
    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    $error = curl_error($ch);

    curl_close($ch);

    if ($error) {
        registrarLog('ERROR', $resource, "cURL Error: $error");
        return ['code' => 0, 'data' => $error];
    }

    return ['code' => $httpCode, 'data' => $response];
}

// ========================================================================
// FUNCIONES DE RESPUESTA
// ========================================================================

/**
 * Responder con éxito
 */
function responderExito($data, $tiempoMs = null) {
    $respuesta = [
        'success' => true,
        'data' => $data
    ];

    if ($tiempoMs !== null) {
        $respuesta['tiempo_ms'] = $tiempoMs;
    }

    echo json_encode($respuesta, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
    exit;
}

/**
 * Responder con error
 */
function responderError($httpCode, $mensaje, $detalles = null) {
    http_response_code($httpCode);

    $respuesta = [
        'success' => false,
        'error' => [
            'codigo' => $httpCode,
            'mensaje' => $mensaje
        ]
    ];

    if ($detalles !== null) {
        $respuesta['error']['detalles'] = $detalles;
    }

    echo json_encode($respuesta, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
    exit;
}

// ========================================================================
// FUNCIONES DE LOGGING
// ========================================================================

/**
 * Registrar evento en log
 */
function registrarLog($tipo, $referencia, $mensaje, $tiempoMs = null) {
    if (!DEBUG_MODE) {
        return;
    }

    $timestamp = date('Y-m-d H:i:s');
    $ip = $_SERVER['REMOTE_ADDR'] ?? 'UNKNOWN';
    $tiempoStr = $tiempoMs !== null ? " [{$tiempoMs}ms]" : '';

    $linea = "[$timestamp] [$tipo] [$ip] [$referencia]$tiempoStr $mensaje\n";

    file_put_contents(LOG_FILE, $linea, FILE_APPEND);
}
