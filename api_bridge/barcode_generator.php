<?php
/**
 * ========================================================================
 * GENERADOR DE CÓDIGOS DE BARRAS - EAN13
 * ========================================================================
 *
 * Genera imágenes PNG de códigos de barras EAN13 usando GD
 * Compatible con scanners estándar
 *
 * Basado en especificaciones EAN-13
 * Autor: Claude Code
 * Fecha: 2026-01-20
 * ========================================================================
 */

class BarcodeGenerator {

    // Patrones EAN-13 para codificación
    private $left_odd = [
        '0001101', '0011001', '0010011', '0111101', '0100011',
        '0110001', '0101111', '0111011', '0110111', '0001011'
    ];

    private $left_even = [
        '0100111', '0110011', '0011011', '0100001', '0011101',
        '0111001', '0000101', '0010001', '0001001', '0010111'
    ];

    private $right = [
        '1110010', '1100110', '1101100', '1000010', '1011100',
        '1001110', '1010000', '1000100', '1001000', '1110100'
    ];

    private $parity = [
        'OOOOOO', 'OOEOEE', 'OOEEOE', 'OOEEEO', 'OEOOEE',
        'OEEOOE', 'OEEEOO', 'OEOEOE', 'OEOEEO', 'OEEOEO'
    ];

    /**
     * Generar imagen PNG de código de barras EAN13
     *
     * @param string $ean13 Código EAN13 (13 dígitos)
     * @param int $width Ancho de la imagen en píxeles
     * @param int $height Alto de la imagen en píxeles
     * @return resource Imagen GD
     */
    public function generateEAN13($ean13, $width = 300, $height = 150) {
        // Validar y limpiar EAN13
        $ean13 = preg_replace('/[^0-9]/', '', $ean13);

        // Rellenar con ceros si es necesario
        $ean13 = str_pad($ean13, 13, '0', STR_PAD_LEFT);

        // Si tiene más de 13 dígitos, tomar los últimos 13
        if (strlen($ean13) > 13) {
            $ean13 = substr($ean13, -13);
        }

        // Calcular checksum si el último dígito es 0 (código sin checksum válido)
        // Esto permite que códigos sin checksum válido también funcionen
        // $ean13 = $this->addChecksum(substr($ean13, 0, 12));

        // Generar patrón de barras
        $barcode = $this->encodeEAN13($ean13);

        // Crear imagen
        $image = $this->drawBarcode($barcode, $ean13, $width, $height);

        return $image;
    }

    /**
     * Calcular dígito de control EAN13
     */
    private function calculateChecksum($ean12) {
        $sum = 0;
        for ($i = 0; $i < 12; $i++) {
            $digit = (int)$ean12[$i];
            $sum += ($i % 2 == 0) ? $digit : $digit * 3;
        }
        $checksum = (10 - ($sum % 10)) % 10;
        return $checksum;
    }

    /**
     * Agregar dígito de control a código EAN12
     */
    private function addChecksum($ean12) {
        $checksum = $this->calculateChecksum($ean12);
        return $ean12 . $checksum;
    }

    /**
     * Codificar EAN13 a patrón binario de barras
     */
    private function encodeEAN13($ean13) {
        $code = '';

        // Primer dígito determina el patrón de paridad
        $first_digit = (int)$ean13[0];
        $parity_pattern = $this->parity[$first_digit];

        // Guard bar inicial (101)
        $code .= '101';

        // Codificar los 6 dígitos de la izquierda (posiciones 1-6)
        for ($i = 1; $i <= 6; $i++) {
            $digit = (int)$ean13[$i];
            $parity_char = $parity_pattern[$i - 1];

            if ($parity_char == 'O') {
                $code .= $this->left_odd[$digit];
            } else {
                $code .= $this->left_even[$digit];
            }
        }

        // Guard bar central (01010)
        $code .= '01010';

        // Codificar los 6 dígitos de la derecha (posiciones 7-12)
        for ($i = 7; $i <= 12; $i++) {
            $digit = (int)$ean13[$i];
            $code .= $this->right[$digit];
        }

        // Guard bar final (101)
        $code .= '101';

        return $code;
    }

    /**
     * Dibujar código de barras en imagen GD
     */
    private function drawBarcode($barcode, $text, $width, $height) {
        $barcode_length = strlen($barcode);

        // Crear imagen
        $image = imagecreatetruecolor($width, $height);

        // Colores
        $white = imagecolorallocate($image, 255, 255, 255);
        $black = imagecolorallocate($image, 0, 0, 0);

        // Fondo blanco
        imagefill($image, 0, 0, $white);

        // Calcular dimensiones
        $text_height = 20; // Altura reservada para el texto
        $barcode_height = $height - $text_height;
        $bar_width = $width / $barcode_length;

        // Dibujar barras
        $x = 0;
        for ($i = 0; $i < $barcode_length; $i++) {
            if ($barcode[$i] == '1') {
                imagefilledrectangle(
                    $image,
                    (int)($x),
                    0,
                    (int)($x + $bar_width),
                    $barcode_height,
                    $black
                );
            }
            $x += $bar_width;
        }

        // Dibujar texto (número)
        $font_size = 3; // Tamaño de fuente (1-5 para imagestring)
        $text_width = imagefontwidth($font_size) * strlen($text);
        $text_x = ($width - $text_width) / 2;
        $text_y = $barcode_height + 2;

        imagestring($image, $font_size, (int)$text_x, (int)$text_y, $text, $black);

        return $image;
    }

    /**
     * Guardar imagen de código de barras en archivo BMP (mejor compatibilidad VB6)
     */
    public function saveEAN13($ean13, $filepath, $width = 300, $height = 150) {
        $image = $this->generateEAN13($ean13, $width, $height);

        // Guardar como BMP en lugar de PNG para compatibilidad VB6
        $result = imagebmp($image, $filepath);
        imagedestroy($image);
        return $result;
    }

    /**
     * Generar imagen de código de barras y retornarla como base64
     */
    public function getEAN13Base64($ean13, $width = 300, $height = 150) {
        $image = $this->generateEAN13($ean13, $width, $height);

        ob_start();
        imagebmp($image);
        $imageData = ob_get_clean();
        imagedestroy($image);

        return base64_encode($imageData);
    }
}
