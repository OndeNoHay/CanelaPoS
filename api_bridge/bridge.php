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
require_once __DIR__ . '/barcode_generator.php';

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

        case 'actualizar_stock':
            handleActualizarStock();
            break;

        case 'buscar_productos_rango':
            handleBuscarProductosRango();
            break;

        case 'generar_codigos_barras':
            handleGenerarCodigosBarras();
            break;

        default:
            responderError(400, 'Acción no válida', [
                'accion_recibida' => $action,
                'acciones_validas' => ['test', 'buscar_producto', 'obtener_stock', 'info_producto', 'actualizar_stock', 'buscar_productos_rango', 'generar_codigos_barras']
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

/**
 * Handler: Actualizar stock de un producto en PrestaShop
 * Endpoint: POST /bridge.php?action=actualizar_stock
 * Parámetros POST: id_producto, cantidad, id_combinacion (opcional)
 */
function handleActualizarStock() {
    $inicio = microtime(true);

    try {
        // Validar método HTTP
        if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
            responderError(405, 'Método no permitido. Use POST', [
                'metodo_recibido' => $_SERVER['REQUEST_METHOD']
            ]);
        }

        // Validar parámetros
        if (!isset($_POST['id_producto']) || !is_numeric($_POST['id_producto'])) {
            responderError(400, 'Parámetro "id_producto" numérico requerido');
        }

        if (!isset($_POST['cantidad']) || !is_numeric($_POST['cantidad'])) {
            responderError(400, 'Parámetro "cantidad" numérico requerido');
        }

        $idProducto = intval($_POST['id_producto']);
        $cantidad = intval($_POST['cantidad']);
        $idCombinacion = isset($_POST['id_combinacion']) ? intval($_POST['id_combinacion']) : 0;

        // Determinar tipo de operación para log
        $operacion = $cantidad < 0 ? 'DECREMENTO' : 'INCREMENTO';
        $cantidadAbs = abs($cantidad);

        registrarLog(
            'STOCK_UPDATE',
            $idProducto,
            "Solicitud actualización stock: Producto=$idProducto, Combinacion=$idCombinacion, Operacion=$operacion, Cantidad=$cantidadAbs"
        );

        // Actualizar stock en PrestaShop
        $resultado = actualizarStockEnPrestaShop($idProducto, $cantidad, $idCombinacion);

        if ($resultado['success']) {
            $tiempo = round((microtime(true) - $inicio) * 1000);
            registrarLog(
                'STOCK_UPDATE',
                $idProducto,
                "Stock actualizado: {$resultado['stock_anterior']} -> {$resultado['stock_nuevo']}",
                $tiempo
            );

            responderExito([
                'id_producto' => $idProducto,
                'id_combinacion' => $idCombinacion,
                'stock_anterior' => $resultado['stock_anterior'],
                'stock_nuevo' => $resultado['stock_nuevo'],
                'operacion' => $operacion,
                'cantidad_modificada' => $cantidadAbs
            ], $tiempo);
        } else {
            registrarLog(
                'ERROR',
                $idProducto,
                "Error actualizando stock: {$resultado['error']}"
            );

            responderError(500, 'Error al actualizar stock', [
                'id_producto' => $idProducto,
                'id_combinacion' => $idCombinacion,
                'error' => $resultado['error']
            ]);
        }
    } catch (Exception $e) {
        registrarLog('ERROR', 'STOCK_UPDATE', "Excepción fatal: " . $e->getMessage());
        responderError(500, 'Error interno del servidor', [
            'excepcion' => $e->getMessage(),
            'archivo' => $e->getFile(),
            'linea' => $e->getLine()
        ]);
    }
}

/**
 * Buscar productos por rango de IDs
 * GET /bridge.php?action=buscar_productos_rango&id_inicio=1&id_fin=100
 * Retorna solo productos activos y con stock > 0
 */
function handleBuscarProductosRango() {
    $inicio = microtime(true);

    // Validar parámetros
    if (!isset($_GET['id_inicio']) || !is_numeric($_GET['id_inicio'])) {
        responderError(400, 'Parámetro "id_inicio" numérico requerido');
    }

    if (!isset($_GET['id_fin']) || !is_numeric($_GET['id_fin'])) {
        responderError(400, 'Parámetro "id_fin" numérico requerido');
    }

    $idInicio = intval($_GET['id_inicio']);
    $idFin = intval($_GET['id_fin']);

    // Validar rango
    if ($idFin < $idInicio) {
        responderError(400, 'El id_fin debe ser mayor o igual que id_inicio', [
            'id_inicio' => $idInicio,
            'id_fin' => $idFin
        ]);
    }

    // Limitar el rango a 500 productos máximo para evitar timeouts
    if (($idFin - $idInicio) > 500) {
        responderError(400, 'El rango máximo permitido es de 500 productos', [
            'id_inicio' => $idInicio,
            'id_fin' => $idFin,
            'rango_solicitado' => ($idFin - $idInicio + 1)
        ]);
    }

    registrarLog('BUSQUEDA_RANGO', "$idInicio-$idFin", "Buscando productos del ID $idInicio al $idFin");

    $productos = [];
    $encontrados = 0;
    $errores = 0;

    // Iterar sobre el rango de IDs
    for ($idProducto = $idInicio; $idProducto <= $idFin; $idProducto++) {
        try {
            $producto = obtenerProductoCompleto($idProducto);

            // Si el producto existe, está activo y tiene stock
            if ($producto && $producto['activo'] && $producto['stock'] > 0) {
                // Si tiene combinaciones, solo incluir las que tienen stock
                if ($producto['tiene_combinaciones'] && !empty($producto['combinaciones'])) {
                    // Filtrar combinaciones con stock > 0
                    $combosConStock = array_filter($producto['combinaciones'], function($combo) {
                        return $combo['stock'] > 0;
                    });

                    if (!empty($combosConStock)) {
                        $producto['combinaciones'] = array_values($combosConStock);
                        $productos[] = $producto;
                        $encontrados++;
                    }
                } else {
                    // Producto estándar con stock
                    $productos[] = $producto;
                    $encontrados++;
                }
            }
        } catch (Exception $e) {
            $errores++;
            registrarLog('ERROR', "Producto ID $idProducto", "Error: " . $e->getMessage());
        }
    }

    $tiempo = round((microtime(true) - $inicio) * 1000);
    registrarLog('BUSQUEDA_RANGO', "$idInicio-$idFin",
        "Búsqueda completada: $encontrados productos encontrados, $errores errores", $tiempo);

    responderExito([
        'productos' => $productos,
        'total_encontrados' => $encontrados,
        'rango_consultado' => [
            'id_inicio' => $idInicio,
            'id_fin' => $idFin,
            'total_ids' => ($idFin - $idInicio + 1)
        ],
        'errores' => $errores
    ], $tiempo);
}

/**
 * Generar códigos de barras como imágenes PNG
 * POST /bridge.php?action=generar_codigos_barras
 * Body: JSON array de códigos EAN13
 * Ejemplo: ["8435423154703", "8435423154710", "8435423154727"]
 *
 * Retorna: JSON con rutas de archivos generados
 */
function handleGenerarCodigosBarras() {
    $inicio = microtime(true);

    try {
        // Validar método HTTP
        if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
            responderError(405, 'Método no permitido. Use POST', [
                'metodo_recibido' => $_SERVER['REQUEST_METHOD']
            ]);
        }

        // Leer body JSON
        $json = file_get_contents('php://input');
        $codigos = json_decode($json, true);

        // Validar formato
        if (!is_array($codigos) || empty($codigos)) {
            responderError(400, 'Se requiere un array JSON de códigos EAN13', [
                'ejemplo' => ['8435423154703', '8435423154710']
            ]);
        }

        // Limitar cantidad para evitar timeouts (máximo 500 códigos)
        if (count($codigos) > 500) {
            responderError(400, 'Máximo 500 códigos por petición', [
                'recibidos' => count($codigos),
                'maximo' => 500
            ]);
        }

        // Crear carpeta temporal si no existe
        $tempDir = __DIR__ . '/temp_barcodes';
        if (!file_exists($tempDir)) {
            mkdir($tempDir, 0755, true);
        }

        // Limpiar archivos antiguos (más de 1 hora)
        $archivosAntiguos = glob("$tempDir/*.bmp");
        $ahora = time();
        foreach ($archivosAntiguos as $archivo) {
            if ($ahora - filemtime($archivo) > 3600) { // 1 hora
                @unlink($archivo);
            }
        }

        registrarLog('BARCODES', count($codigos), "Generando " . count($codigos) . " códigos de barras");

        // Generar códigos de barras
        $generator = new BarcodeGenerator();
        $archivosGenerados = [];
        $errores = [];

        foreach ($codigos as $index => $ean13) {
            try {
                // Limpiar el código (solo números)
                $ean13_clean = preg_replace('/[^0-9]/', '', $ean13);

                if (empty($ean13_clean)) {
                    $errores[] = [
                        'codigo' => $ean13,
                        'error' => 'Código vacío o sin dígitos'
                    ];
                    continue;
                }

                // Generar nombre único para archivo (BMP para compatibilidad VB6)
                $timestamp = time();
                $random = mt_rand(1000, 9999);
                $filename = "barcode_{$ean13_clean}_{$timestamp}_{$random}.bmp";
                $filepath = "$tempDir/$filename";

                // Generar imagen BMP (300x150 píxeles - alta resolución)
                $resultado = $generator->saveEAN13($ean13_clean, $filepath, 300, 150);

                if ($resultado) {
                    $archivosGenerados[] = [
                        'ean13' => $ean13_clean,
                        'filename' => $filename,
                        'filepath' => $filepath,
                        'url' => "api_bridge/temp_barcodes/$filename"
                    ];
                } else {
                    $errores[] = [
                        'codigo' => $ean13,
                        'error' => 'Error al guardar imagen'
                    ];
                }

            } catch (Exception $e) {
                $errores[] = [
                    'codigo' => $ean13,
                    'error' => $e->getMessage()
                ];
            }
        }

        $tiempo = round((microtime(true) - $inicio) * 1000);
        registrarLog('BARCODES', count($archivosGenerados),
            "Generados " . count($archivosGenerados) . " códigos, " . count($errores) . " errores",
            $tiempo);

        responderExito([
            'archivos' => $archivosGenerados,
            'total_generados' => count($archivosGenerados),
            'total_errores' => count($errores),
            'errores' => $errores,
            'directorio_temporal' => $tempDir
        ], $tiempo);

    } catch (Exception $e) {
        registrarLog('ERROR', 'generar_codigos_barras', "Excepción: " . $e->getMessage());
        responderError(500, 'Error al generar códigos de barras', [
            'error' => $e->getMessage()
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
 * Obtener combinaciones (tallas) de un producto
 */
function obtenerCombinaciones($idProducto, $productXML) {
    $combinaciones = [];

    // Extraer IDs de combinaciones del XML del producto
    $comboIds = [];
    if (isset($productXML->associations->combinations->combination)) {
        foreach ($productXML->associations->combinations->combination as $combo) {
            $comboIds[] = (int)$combo->id;
        }
    }

    if (empty($comboIds)) {
        return $combinaciones;
    }

    // Obtener detalles de todas las combinaciones
    $combosFilter = implode('|', $comboIds);
    $resultadoCombos = callPrestaShop("combinations?filter[id]=[$combosFilter]&display=full");

    if ($resultadoCombos['code'] != 200) {
        registrarLog('ERROR', 'obtenerCombinaciones', "Error al obtener combinaciones: HTTP {$resultadoCombos['code']}");
        return $combinaciones;
    }

    try {
        $combosXML = simplexml_load_string($resultadoCombos['data']);

        if (!isset($combosXML->combinations->combination)) {
            return $combinaciones;
        }

        // Obtener stock para todas las combinaciones
        $stockData = obtenerStockCombinaciones($idProducto);

        // Obtener nombres de las tallas (product_option_values)
        $tallaIds = [];
        foreach ($combosXML->combinations->combination as $combo) {
            if (isset($combo->associations->product_option_values->product_option_value)) {
                foreach ($combo->associations->product_option_values->product_option_value as $pov) {
                    $tallaIds[] = (int)$pov->id;
                }
            }
        }

        $tallasNombres = obtenerNombresTallas($tallaIds);

        // Parsear cada combinación
        foreach ($combosXML->combinations->combination as $combo) {
            $idCombinacion = (int)$combo->id;
            $idProductoAttr = (int)$combo->id;

            // Extraer talla de esta combinación
            $talla = '';
            $idTalla = 0;
            if (isset($combo->associations->product_option_values->product_option_value)) {
                foreach ($combo->associations->product_option_values->product_option_value as $pov) {
                    $idTalla = (int)$pov->id;
                    $talla = isset($tallasNombres[$idTalla]) ? $tallasNombres[$idTalla] : "Talla $idTalla";
                    break; // Solo tomamos la primera (que debería ser la talla)
                }
            }

            // Obtener stock de esta combinación
            $stockKey = "p{$idProducto}-a{$idProductoAttr}";
            $stock = isset($stockData[$stockKey]) ? $stockData[$stockKey] : 0;

            $combinaciones[] = [
                'id_combinacion' => $idCombinacion,
                'id_product_attribute' => $idProductoAttr,
                'talla' => $talla,
                'id_talla' => $idTalla,
                'stock' => $stock,
                'disponible' => $stock > 0
            ];
        }

    } catch (Exception $e) {
        registrarLog('ERROR', 'obtenerCombinaciones', "Error al parsear combinaciones: {$e->getMessage()}");
    }

    return $combinaciones;
}

/**
 * Obtener stock de todas las combinaciones de un producto
 */
function obtenerStockCombinaciones($idProducto) {
    $stockData = [];

    $resultado = callPrestaShop("stock_availables?filter[id_product]=$idProducto&display=full");

    if ($resultado['code'] != 200) {
        return $stockData;
    }

    try {
        $stockXML = simplexml_load_string($resultado['data']);

        if (isset($stockXML->stock_availables->stock_available)) {
            foreach ($stockXML->stock_availables->stock_available as $stock) {
                $idProd = (int)$stock->id_product;
                $idAttr = (int)$stock->id_product_attribute;
                $cantidad = (int)$stock->quantity;

                $key = "p{$idProd}-a{$idAttr}";
                $stockData[$key] = $cantidad;
            }
        }
    } catch (Exception $e) {
        registrarLog('ERROR', 'obtenerStockCombinaciones', "Error al parsear stock: {$e->getMessage()}");
    }

    return $stockData;
}

/**
 * Obtener nombres de tallas desde product_option_values
 */
function obtenerNombresTallas($tallaIds) {
    $nombres = [];

    if (empty($tallaIds)) {
        return $nombres;
    }

    // Filtrar solo IDs únicos
    $tallaIds = array_unique($tallaIds);
    $idsFilter = implode('|', $tallaIds);

    $resultado = callPrestaShop("product_option_values?filter[id]=[$idsFilter]&display=full");

    if ($resultado['code'] != 200) {
        return $nombres;
    }

    try {
        $valoresXML = simplexml_load_string($resultado['data']);

        if (isset($valoresXML->product_option_values->product_option_value)) {
            foreach ($valoresXML->product_option_values->product_option_value as $valor) {
                $id = (int)$valor->id;
                $idAttrGroup = (int)$valor->id_attribute_group;

                // Solo procesar si es del grupo de atributos de TALLA
                if (defined('SIZE_ATTRIBUTE_GROUP_ID') && $idAttrGroup == SIZE_ATTRIBUTE_GROUP_ID) {
                    // Extraer nombre en el idioma configurado
                    $nombre = '';
                    if (isset($valor->name->language)) {
                        foreach ($valor->name->language as $lang) {
                            if ((int)$lang['id'] == PRESTASHOP_LANGUAGE_ID) {
                                $nombre = (string)$lang;
                                break;
                            }
                        }
                        // Si no se encontró el idioma, usar el primero
                        if (empty($nombre)) {
                            $nombre = (string)$valor->name->language[0];
                        }
                    }

                    $nombres[$id] = $nombre;
                }
            }
        }
    } catch (Exception $e) {
        registrarLog('ERROR', 'obtenerNombresTallas', "Error al parsear tallas: {$e->getMessage()}");
    }

    return $nombres;
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

    // Detectar si tiene combinaciones (tallas)
    $tieneCombinaciones = isset($productXML->associations->combinations->combination);
    $combinaciones = [];
    $stock = 0;

    if ($tieneCombinaciones) {
        // Producto con combinaciones (tallas)
        $combinaciones = obtenerCombinaciones((int)$productXML->id, $productXML);
        // Stock total = suma de stock de todas las combinaciones
        foreach ($combinaciones as $combo) {
            $stock += $combo['stock'];
        }
    } else {
        // Producto estándar sin combinaciones
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
        'tiene_combinaciones' => $tieneCombinaciones,
        'combinaciones' => $combinaciones,
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

    // Headers
    $headers = ['Output-Format: XML'];
    if ($method == 'POST' || $method == 'PUT') {
        $headers[] = 'Content-Type: text/xml';
    }
    curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);

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
// FUNCIONES DE ACTUALIZACIÓN DE STOCK
// ========================================================================

/**
 * Actualizar stock de un producto en PrestaShop
 *
 * @param int $idProducto ID del producto en PrestaShop
 * @param int $cantidad Cantidad a sumar/restar (negativo para decrementar)
 * @param int $idCombinacion ID de combinación (0 para producto sin combinaciones)
 * @return array ['success' => bool, 'stock_anterior' => int, 'stock_nuevo' => int, 'error' => string]
 */
function actualizarStockEnPrestaShop($idProducto, $cantidad, $idCombinacion = 0) {
    try {
        // 1. Encontrar el id_stock_available correcto usando filtros
        $filtro = "filter[id_product]=$idProducto";
        if ($idCombinacion > 0) {
            $filtro .= "&filter[id_product_attribute]=$idCombinacion";
        } else {
            // Producto sin combinaciones: id_product_attribute = 0
            $filtro .= "&filter[id_product_attribute]=0";
        }

        registrarLog('DEBUG', $idProducto, "Buscando stock_available con filtro: $filtro");

        $resultado = callPrestaShop("stock_availables?$filtro&display=full");

        if ($resultado['code'] != 200) {
            return [
                'success' => false,
                'error' => "Error al obtener stock_available: HTTP {$resultado['code']}"
            ];
        }

        // 2. Parsear respuesta XML
        $xml = simplexml_load_string($resultado['data']);

        if (!isset($xml->stock_availables->stock_available)) {
            registrarLog('ERROR', $idProducto, "No se encontró stock_available en respuesta XML");
            return [
                'success' => false,
                'error' => "No se encontró stock_available para producto=$idProducto, combinacion=$idCombinacion"
            ];
        }

        $stockAvailable = $xml->stock_availables->stock_available;
        $idStockAvailable = (int)$stockAvailable->id;
        $stockActual = (int)$stockAvailable->quantity;

        registrarLog('DEBUG', $idProducto, "Stock actual: $stockActual (id_stock_available=$idStockAvailable)");

        // 3. Calcular nuevo stock
        $stockNuevo = $stockActual + $cantidad;

        // Validar que no sea negativo
        if ($stockNuevo < 0) {
            return [
                'success' => false,
                'error' => "Stock insuficiente. Actual: $stockActual, Solicitado: " . abs($cantidad)
            ];
        }

        // 4. Construir XML para actualización
        // IMPORTANTE: PrestaShop requiere TODOS los campos, no solo los que cambian
        // Por eso tomamos el XML completo y modificamos solo quantity
        $xmlUpdate = new SimpleXMLElement('<?xml version="1.0" encoding="UTF-8"?><prestashop></prestashop>');
        $stockElement = $xmlUpdate->addChild('stock_available');

        // Copiar TODOS los campos del registro original
        $stockElement->addChild('id', (string)$stockAvailable->id);
        $stockElement->addChild('id_product', (string)$stockAvailable->id_product);
        $stockElement->addChild('id_product_attribute', (string)$stockAvailable->id_product_attribute);
        $stockElement->addChild('id_shop', (string)$stockAvailable->id_shop);
        $stockElement->addChild('id_shop_group', (string)$stockAvailable->id_shop_group);
        $stockElement->addChild('quantity', $stockNuevo);  // Este es el único que cambia
        $stockElement->addChild('depends_on_stock', (string)$stockAvailable->depends_on_stock);
        $stockElement->addChild('out_of_stock', (string)$stockAvailable->out_of_stock);

        // Si hay location, incluirlo (puede ser vacío)
        if (isset($stockAvailable->location)) {
            $stockElement->addChild('location', (string)$stockAvailable->location);
        }

        $xmlData = $xmlUpdate->asXML();

        registrarLog('DEBUG', $idProducto, "Enviando PUT a stock_availables/$idStockAvailable con quantity=$stockNuevo");

        // 5. Enviar PUT a PrestaShop
        $resultadoPut = callPrestaShop("stock_availables/$idStockAvailable", 'PUT', $xmlData);

        if ($resultadoPut['code'] != 200) {
            registrarLog('ERROR', $idProducto, "PUT falló: HTTP {$resultadoPut['code']} - Respuesta: " . substr($resultadoPut['data'], 0, 200));
            return [
                'success' => false,
                'error' => "Error al actualizar stock: HTTP {$resultadoPut['code']}"
            ];
        }

        registrarLog('DEBUG', $idProducto, "PUT exitoso: HTTP 200");

        // 6. Retornar éxito con valores
        return [
            'success' => true,
            'stock_anterior' => $stockActual,
            'stock_nuevo' => $stockNuevo
        ];

    } catch (Exception $e) {
        return [
            'success' => false,
            'error' => $e->getMessage()
        ];
    }
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
