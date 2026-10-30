<?php
declare(strict_types=1);

set_time_limit(0);
ini_set('max_execution_time', '0');
ini_set('memory_limit', '1024M');

// Cargar Composer ANTES de declarar el filtro de PhpSpreadsheet.
$autoload = __DIR__ . '/vendor/autoload.php';

if (!is_file($autoload)) {
    http_response_code(500);
    exit('No existe vendor/autoload.php. Ejecuta composer install en la carpeta del proyecto.');
}

require_once $autoload;

require __DIR__ . '/db.php';
require __DIR__ . '/functions.php';

final class FiltroFilasImportacion implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter
{
    private int $desde = 1;
    private int $hasta = 1;
    private int $filaEncabezado = 1;

    public function configurar(int $desde, int $hasta, int $filaEncabezado): void
    {
        $this->desde = $desde;
        $this->hasta = $hasta;
        $this->filaEncabezado = $filaEncabezado;
    }

    public function readCell(string $columnAddress, int $row, string $worksheetName = ''): bool
    {
        return $row === $this->filaEncabezado
            || ($row >= $this->desde && $row <= $this->hasta);
    }
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Location: index.php');
    exit;
}

$fuentes = [
    'ulyses' => [
        'tabla' => 'ulyses',
        'nombre' => 'ULYSES',
        'documentos' => [
            'Identificación',
            'Identificacion',
            'Documento',
            'NUM_DOCUMENTO'
        ]
    ],
    'was' => [
        'tabla' => 'was',
        'nombre' => 'WAS',
        'documentos' => [
            'dniper',
            'Documento',
            'Identificación',
            'Identificacion'
        ]
    ],
    'unoau' => [
        'tabla' => 'unoau',
        'nombre' => 'UnoAUno',
        'documentos' => [
            'NUM_DOCUMENTO',
            'Documento',
            'Identificación',
            'Identificacion'
        ]
    ],
];

$fuente = strtolower(trim((string)($_POST['fuente'] ?? '')));

if (!isset($fuentes[$fuente])) {
    header('Location: index.php?error=' . urlencode('Fuente no válida.'));
    exit;
}

if (
    !isset($_FILES['archivo']) ||
    $_FILES['archivo']['error'] !== UPLOAD_ERR_OK
) {
    header('Location: index.php?error=' . urlencode('No se recibió correctamente el archivo.'));
    exit;
}

$ruta = $_FILES['archivo']['tmp_name'];
$nombreArchivo = basename((string)$_FILES['archivo']['name']);
$extension = strtolower(pathinfo($nombreArchivo, PATHINFO_EXTENSION));

if (!in_array($extension, ['xlsx', 'xls', 'csv'], true)) {
    header('Location: index.php?error=' . urlencode('Formato no permitido. Usa XLSX, XLS o CSV.'));
    exit;
}

$bloque = 500;
$filaEncabezado = max(1, (int)($_POST['fila_encabezados'] ?? 1));
$hojaSolicitada = trim((string)($_POST['hoja'] ?? ''));

try {
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($ruta);
    $reader->setReadDataOnly(true);

    $hojas = $reader->listWorksheetInfo($ruta);

    if (!$hojas) {
        throw new RuntimeException('No se encontraron hojas legibles en el archivo.');
    }

    $nombreHoja = $hojaSolicitada !== ''
        ? $hojaSolicitada
        : (string)$hojas[0]['worksheetName'];

    $infoHoja = null;

    foreach ($hojas as $info) {
        if ((string)$info['worksheetName'] === $nombreHoja) {
            $infoHoja = $info;
            break;
        }
    }

    if ($infoHoja === null) {
        throw new RuntimeException('No se encontró la hoja indicada: ' . $nombreHoja);
    }

    $ultimaFila = (int)$infoHoja['totalRows'];
    $ultimaColumna = (int)$infoHoja['totalColumns'];

    if ($ultimaFila < $filaEncabezado || $ultimaColumna < 1) {
        throw new RuntimeException('La hoja no contiene datos en la fila de encabezados indicada.');
    }

    $filtro = new FiltroFilasImportacion();
    $filtro->configurar($filaEncabezado, $filaEncabezado, $filaEncabezado);

    $reader->setReadFilter($filtro);

    if (method_exists($reader, 'setLoadSheetsOnly')) {
        $reader->setLoadSheetsOnly([$nombreHoja]);
    }

    // Leer únicamente los encabezados.
    $libro = $reader->load($ruta);
    $hoja = $libro->getSheetByName($nombreHoja);

    if (!$hoja) {
        $libro->disconnectWorksheets();
        throw new RuntimeException('No fue posible abrir la hoja seleccionada.');
    }

    $filaCabecera = $hoja->rangeToArrayByColumn(
        1,
        $ultimaColumna,
        $filaEncabezado,
        $filaEncabezado,
        null,
        true,
        false
    )[0] ?? [];

    $libro->disconnectWorksheets();
    unset($libro, $hoja);

    // Limpiar encabezados, limitar su longitud y evitar duplicados.
    $encabezados = [];
    $vistos = [];

    foreach ($filaCabecera as $i => $valor) {
        $nombre = (string)$valor;

        // Quitar caracteres de control.
        $nombre = preg_replace('/[\x00-\x1F\x7F]/u', '', $nombre) ?? $nombre;
        $nombre = trim($nombre);

        if ($nombre === '') {
            $nombre = 'columna_' . ($i + 1);
        }

        // MySQL permite nombres de columnas de hasta 64 caracteres.
        $nombre = trim(mb_substr($nombre, 0, 64, 'UTF-8'));

        if ($nombre === '') {
            $nombre = 'columna_' . ($i + 1);
        }

        $base = $nombre;
        $sufijo = 2;

        while (isset($vistos[mb_strtolower($nombre, 'UTF-8')])) {
            $final = '_' . $sufijo++;
            $nombre = trim(
                mb_substr(
                    $base,
                    0,
                    64 - mb_strlen($final, 'UTF-8'),
                    'UTF-8'
                )
            ) . $final;
        }

        $vistos[mb_strtolower($nombre, 'UTF-8')] = true;
        $encabezados[] = $nombre;
    }

    if (!$encabezados) {
        throw new RuntimeException('No se detectaron encabezados.');
    }

    // Identificar la columna que contiene el documento.
    $columnaDocumento = null;

    foreach ($fuentes[$fuente]['documentos'] as $candidato) {
        foreach ($encabezados as $encabezado) {
            if (
                mb_strtolower(trim($encabezado), 'UTF-8') ===
                mb_strtolower($candidato, 'UTF-8')
            ) {
                $columnaDocumento = $encabezado;
                break 2;
            }
        }
    }

    if ($columnaDocumento === null) {
        throw new RuntimeException(
            'No se encontró la columna de identificación. Encabezados detectados: ' .
            implode(', ', $encabezados)
        );
    }

    foreach ($encabezados as $encabezado) {
        if (in_array($encabezado, ['__id', '__documento_normalizado'], true)) {
            throw new RuntimeException(
                'El encabezado ' . $encabezado . ' está reservado por el sistema.'
            );
        }
    }

    $tabla = $fuentes[$fuente]['tabla'];

    /*
     * ATENCIÓN:
     * Esta operación reemplaza la tabla seleccionada.
     * Asegúrate de tener was_old respaldada antes de importar WAS.
     */
    $pdo->exec('DROP TABLE IF EXISTS ' . qi($tabla));

    $definiciones = [
        '`__id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY',
        '`__documento_normalizado` VARCHAR(100) NULL'
    ];

    foreach ($encabezados as $encabezado) {
        $definiciones[] = qi($encabezado) . ' LONGTEXT NULL';
    }

    $definiciones[] = 'INDEX `idx_documento_normalizado` (`__documento_normalizado`)';

    $pdo->exec(
        'CREATE TABLE ' . qi($tabla) . ' (' .
        implode(',', $definiciones) .
        ') ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci'
    );

    $columnasInsert = array_merge(['__documento_normalizado'], $encabezados);

    $sql = 'INSERT INTO ' . qi($tabla) . ' (' .
        implode(',', array_map('qi', $columnasInsert)) .
        ') VALUES (' . implode(',', array_fill(0, count($columnasInsert), '?')) . ')';

    $insertar = $pdo->prepare($sql);
    $indiceDocumento = array_search($columnaDocumento, $encabezados, true);
    $totalImportadas = 0;

    // Procesar el Excel por bloques para controlar el uso de memoria.
    for (
        $inicio = $filaEncabezado + 1;
        $inicio <= $ultimaFila;
        $inicio += $bloque
    ) {
        $fin = min($inicio + $bloque - 1, $ultimaFila);

        $filtro->configurar($inicio, $fin, $filaEncabezado);
        $reader->setReadFilter($filtro);

        $libro = $reader->load($ruta);
        $hoja = $libro->getSheetByName($nombreHoja);

        if (!$hoja) {
            $libro->disconnectWorksheets();
            throw new RuntimeException(
                'Error al leer el bloque de filas ' . $inicio . ' a ' . $fin . '.'
            );
        }

        $pdo->beginTransaction();

        try {
            for ($fila = $inicio; $fila <= $fin; $fila++) {
                $valoresFila = $hoja->rangeToArrayByColumn(
                    1,
                    $ultimaColumna,
                    $fila,
                    $fila,
                    null,
                    true,
                    false
                )[0] ?? [];

                $hayDatos = false;

                foreach ($valoresFila as $valor) {
                    if ($valor !== null && trim((string)$valor) !== '') {
                        $hayDatos = true;
                        break;
                    }
                }

                if (!$hayDatos) {
                    continue;
                }

                $valores = [];

                foreach ($encabezados as $i => $encabezado) {
                    $valor = $valoresFila[$i] ?? null;

                    if ($valor instanceof DateTimeInterface) {
                        $valor = $valor->format('Y-m-d H:i:s');
                    }

                    if (is_float($valor) && floor($valor) === $valor) {
                        $valor = (string)(int)$valor;
                    }

                    $valores[] = $valor === null
                        ? null
                        : trim((string)$valor);
                }

                $documento = normdoc($valores[$indiceDocumento] ?? '');

                $insertar->execute(
                    array_merge(
                        [$documento !== '' ? $documento : null],
                        $valores
                    )
                );

                $totalImportadas++;
            }

            $pdo->commit();

        } catch (Throwable $e) {
            if ($pdo->inTransaction()) {
                $pdo->rollBack();
            }

            throw $e;

        } finally {
            $libro->disconnectWorksheets();
            unset($libro, $hoja);

            if (function_exists('gc_collect_cycles')) {
                gc_collect_cycles();
            }
        }
    }

    // Registrar la importación realizada.
    $registrar = $pdo->prepare(
        'INSERT INTO importaciones (fuente, nombre_archivo, hoja, total_filas) VALUES (?, ?, ?, ?)'
    );

    $registrar->execute([
        $fuentes[$fuente]['nombre'],
        $nombreArchivo,
        $nombreHoja,
        $totalImportadas
    ]);

    header('Location: index.php?ok=' . urlencode(
        $fuentes[$fuente]['nombre'] .
        ': importadas ' . number_format($totalImportadas) .
        ' filas en la tabla ' . $tabla . ' mediante lectura por bloques.'
    ));

    exit;

} catch (Throwable $e) {
    header('Location: index.php?error=' . urlencode($e->getMessage()));
    exit;
}