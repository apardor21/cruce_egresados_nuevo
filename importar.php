<?php
declare(strict_types=1);

/**
 * Importador seguro ULYSES / WAS / UnoAUno.
 * - Lee Excel por bloques para reducir memoria.
 * - Usa rangeToArray(), compatible con versiones habituales de PhpSpreadsheet.
 * - Importa primero a una tabla temporal.
 * - Solo reemplaza la tabla actual cuando TODA la importación termina bien.
 * - Conserva la tabla anterior con un nombre de respaldo automático.
 */

set_time_limit(0);
ini_set('max_execution_time', '0');
ini_set('memory_limit', '1024M');

$autoload = __DIR__ . '/vendor/autoload.php';
if (!is_file($autoload)) {
    http_response_code(500);
    exit('No existe vendor/autoload.php. Verifica Composer en la carpeta del proyecto.');
}
require_once $autoload;

$dbFile = __DIR__ . '/db.php';
$functionsFile = __DIR__ . '/functions.php';
if (!is_file($dbFile)) {
    http_response_code(500);
    exit('No se encuentra db.php en la carpeta del proyecto.');
}
require_once $dbFile;
if (is_file($functionsFile)) {
    require_once $functionsFile;
}

if (!isset($pdo) || !($pdo instanceof PDO)) {
    http_response_code(500);
    exit('db.php debe crear la conexión PDO en la variable $pdo.');
}
$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

/** Cita un identificador SQL de forma segura. */
function qiImportador(string $identificador): string
{
    return '`' . str_replace('`', '``', $identificador) . '`';
}

/** Normaliza la identificación usando la función del proyecto si existe. */
function normalizarDocumentoImportador($valor): string
{
    $texto = trim((string)($valor ?? ''));
    if ($texto === '') {
        return '';
    }
    if (function_exists('normdoc')) {
        return trim((string)normdoc($texto));
    }
    // Alternativa básica: elimina espacios, puntos y guiones, sin perder ceros iniciales.
    $texto = preg_replace('/[\s.\-]+/u', '', $texto) ?? $texto;
    return function_exists('mb_strtoupper') ? mb_strtoupper($texto, 'UTF-8') : strtoupper($texto);
}

/** Redirige a index.php mostrando un mensaje. */
function volverImportador(string $tipo, string $mensaje): void
{
    header('Location: index.php?' . $tipo . '=' . urlencode($mensaje));
    exit;
}

if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') {
    header('Location: index.php');
    exit;
}

$fuentes = [
    'ulyses' => [
        'tabla' => 'ulyses',
        'nombre' => 'ULYSES',
        'documentos' => ['Identificación', 'Identificacion', 'Documento', 'NUM_DOCUMENTO'],
    ],
    'was' => [
        'tabla' => 'was',
        'nombre' => 'WAS',
        'documentos' => ['dniper', 'Documento', 'Identificación', 'Identificacion'],
    ],
    'unoau' => [
        'tabla' => 'unoau',
        'nombre' => 'UnoAUno',
        'documentos' => ['NUM_DOCUMENTO', 'Documento', 'Identificación', 'Identificacion'],
    ],
];

$fuente = strtolower(trim((string)($_POST['fuente'] ?? '')));
if (!isset($fuentes[$fuente])) {
    volverImportador('error', 'Fuente no válida.');
}
if (!isset($_FILES['archivo']) || !isset($_FILES['archivo']['error']) || $_FILES['archivo']['error'] !== UPLOAD_ERR_OK) {
    volverImportador('error', 'No se recibió correctamente el archivo. Código de carga: ' . (string)($_FILES['archivo']['error'] ?? 'desconocido'));
}

$ruta = (string)$_FILES['archivo']['tmp_name'];
$nombreArchivo = basename((string)$_FILES['archivo']['name']);
$extension = strtolower(pathinfo($nombreArchivo, PATHINFO_EXTENSION));
if (!in_array($extension, ['xlsx', 'xls', 'csv'], true)) {
    volverImportador('error', 'Formato no permitido. Usa XLSX, XLS o CSV.');
}
if (!is_uploaded_file($ruta) && !is_file($ruta)) {
    volverImportador('error', 'El archivo temporal de carga no existe o no es accesible.');
}

$bloque = 500;
$filaEncabezado = max(1, (int)($_POST['fila_encabezados'] ?? 1));
$hojaSolicitada = trim((string)($_POST['hoja'] ?? ''));
$tablaTemporal = null;
$tabla = $fuentes[$fuente]['tabla'];

try {
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReaderForFile($ruta);
    $reader->setReadDataOnly(true);
    $hojas = $reader->listWorksheetInfo($ruta);
    if (!$hojas) {
        throw new RuntimeException('No se encontraron hojas legibles en el archivo.');
    }

    $nombreHoja = $hojaSolicitada !== '' ? $hojaSolicitada : (string)$hojas[0]['worksheetName'];
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
    $ultimaColumnaLetra = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($ultimaColumna);

    // Filtro de lectura: solo carga la fila de encabezados y el bloque solicitado.
    $filtro = new class($filaEncabezado, $filaEncabezado) implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter {
        private int $desde;
        private int $hasta;
        private int $filaEncabezado;
        public function __construct(int $desde, int $hasta) {
            $this->desde = $desde;
            $this->hasta = $hasta;
            $this->filaEncabezado = $desde;
        }
        public function configurar(int $desde, int $hasta): void {
            $this->desde = $desde;
            $this->hasta = $hasta;
        }
        public function readCell(string $columnAddress, int $row, string $worksheetName = ''): bool {
            return $row === $this->filaEncabezado || ($row >= $this->desde && $row <= $this->hasta);
        }
    };

    $reader->setReadFilter($filtro);
    if (method_exists($reader, 'setLoadSheetsOnly')) {
        $reader->setLoadSheetsOnly([$nombreHoja]);
    }

    // Leer encabezados con rangeToArray (no rangeToArrayByColumn).
    $libro = $reader->load($ruta);
    $hoja = $libro->getSheetByName($nombreHoja);
    if (!$hoja) {
        $libro->disconnectWorksheets();
        throw new RuntimeException('No fue posible abrir la hoja seleccionada.');
    }
    $rangoCabecera = 'A' . $filaEncabezado . ':' . $ultimaColumnaLetra . $filaEncabezado;
    $filaCabecera = $hoja->rangeToArray($rangoCabecera, null, true, false, false)[0] ?? [];
    $libro->disconnectWorksheets();
    unset($libro, $hoja);

    // Limpiar encabezados y hacerlos únicos respetando el límite de 64 caracteres.
    $encabezados = [];
    $vistos = [];
    foreach ($filaCabecera as $i => $valor) {
        $nombre = (string)($valor ?? '');
        $nombre = preg_replace('/[\x00-\x1F\x7F]/u', '', $nombre) ?? $nombre;
        $nombre = trim($nombre);
        if ($nombre === '') {
            $nombre = 'columna_' . ($i + 1);
        }
        $nombre = trim(mb_substr($nombre, 0, 64, 'UTF-8'));
        if ($nombre === '') {
            $nombre = 'columna_' . ($i + 1);
        }
        // Evitar conflicto con columnas internas del importador.
        if (in_array(mb_strtolower($nombre, 'UTF-8'), ['__id', '__documento_normalizado'], true)) {
            $nombre = 'dato_' . mb_substr($nombre, 0, 59, 'UTF-8');
        }
        $base = $nombre;
        $sufijo = 2;
        while (isset($vistos[mb_strtolower($nombre, 'UTF-8')])) {
            $final = '_' . $sufijo++;
            $nombre = mb_substr($base, 0, 64 - mb_strlen($final, 'UTF-8'), 'UTF-8') . $final;
        }
        $vistos[mb_strtolower($nombre, 'UTF-8')] = true;
        $encabezados[] = $nombre;
    }
    if (!$encabezados) {
        throw new RuntimeException('No se detectaron encabezados.');
    }

    // Encontrar la columna de identificación antes de crear o reemplazar tablas.
    $columnaDocumento = null;
    foreach ($fuentes[$fuente]['documentos'] as $candidato) {
        foreach ($encabezados as $encabezado) {
            if (mb_strtolower(trim($encabezado), 'UTF-8') === mb_strtolower($candidato, 'UTF-8')) {
                $columnaDocumento = $encabezado;
                break 2;
            }
        }
    }
    if ($columnaDocumento === null) {
        throw new RuntimeException('No se encontró la columna de identificación. Encabezados detectados: ' . implode(', ', $encabezados));
    }
    $indiceDocumento = array_search($columnaDocumento, $encabezados, true);
    if ($indiceDocumento === false) {
        throw new RuntimeException('No fue posible identificar la posición de la columna de documento.');
    }

    // Crear tabla temporal con nombre único. La tabla actual queda intacta durante la carga.
    $sufijoUnico = date('Ymd_His') . '_' . bin2hex(random_bytes(3));
    $tablaTemporal = 'tmp_' . $tabla . '_' . $sufijoUnico;
    $definiciones = [
        '`__id` BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY',
        '`__documento_normalizado` VARCHAR(100) NULL',
    ];
    foreach ($encabezados as $encabezado) {
        $definiciones[] = qiImportador($encabezado) . ' LONGTEXT NULL';
    }
    $definiciones[] = 'INDEX `idx_documento_normalizado` (`__documento_normalizado`)';
    $pdo->exec('CREATE TABLE ' . qiImportador($tablaTemporal) . ' (' . implode(',', $definiciones) . ') ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci');

    $columnasInsert = array_merge(['__documento_normalizado'], $encabezados);
    $sql = 'INSERT INTO ' . qiImportador($tablaTemporal) . ' (' . implode(',', array_map('qiImportador', $columnasInsert)) . ') VALUES (' . implode(',', array_fill(0, count($columnasInsert), '?')) . ')';
    $insertar = $pdo->prepare($sql);
    $totalImportadas = 0;

    for ($inicio = $filaEncabezado + 1; $inicio <= $ultimaFila; $inicio += $bloque) {
        $fin = min($inicio + $bloque - 1, $ultimaFila);
        $filtro->configurar($inicio, $fin);
        $reader->setReadFilter($filtro);
        $libro = $reader->load($ruta);
        $hoja = $libro->getSheetByName($nombreHoja);
        if (!$hoja) {
            $libro->disconnectWorksheets();
            throw new RuntimeException('Error al leer el bloque de filas ' . $inicio . ' a ' . $fin . '.');
        }

        try {
            $pdo->beginTransaction();
            for ($fila = $inicio; $fila <= $fin; $fila++) {
                $rango = 'A' . $fila . ':' . $ultimaColumnaLetra . $fila;
                $valoresFila = $hoja->rangeToArray($rango, null, true, false, false)[0] ?? [];
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
                    $valores[] = $valor === null ? null : trim((string)$valor);
                }
                $documento = normalizarDocumentoImportador($valores[$indiceDocumento] ?? '');
                $insertar->execute(array_merge([$documento !== '' ? $documento : null], $valores));
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

    if ($totalImportadas < 1) {
        throw new RuntimeException('La importación no encontró filas con datos. La tabla anterior se conserva.');
    }

    // Sustitución final. La tabla previa se conserva como respaldo con nombre único.
    $existeTabla = $pdo->prepare('SELECT COUNT(*) FROM information_schema.TABLES WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ?');
    $existeTabla->execute([$tabla]);
    $tablaAnteriorExiste = (int)$existeTabla->fetchColumn() > 0;
    $respaldo = null;
    if ($tablaAnteriorExiste) {
        $respaldo = $tabla . '_backup_' . date('Ymd_His');
        // Evitar colisión excepcional de nombres.
        $n = 2;
        while (true) {
            $existeTabla->execute([$respaldo]);
            if ((int)$existeTabla->fetchColumn() === 0) {
                break;
            }
            $respaldo = $tabla . '_backup_' . date('Ymd_His') . '_' . $n++;
        }
        // Un RENAME TABLE con ambas operaciones mantiene la sustitución atómica en MariaDB/MySQL.
        $pdo->exec('RENAME TABLE ' . qiImportador($tabla) . ' TO ' . qiImportador($respaldo) . ', ' . qiImportador($tablaTemporal) . ' TO ' . qiImportador($tabla));
    } else {
        $pdo->exec('RENAME TABLE ' . qiImportador($tablaTemporal) . ' TO ' . qiImportador($tabla));
    }
    $tablaTemporal = null; // Ya fue renombrada; no borrar.

    // Registrar la importación como operación secundaria; un fallo de bitácora no revierte los datos.
    $avisoBitacora = '';
    try {
        $existeLog = $pdo->prepare('SELECT COUNT(*) FROM information_schema.TABLES WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = ?');
        $existeLog->execute(['importaciones']);
        if ((int)$existeLog->fetchColumn() > 0) {
            $registrar = $pdo->prepare('INSERT INTO importaciones (fuente, nombre_archivo, hoja, total_filas) VALUES (?, ?, ?, ?)');
            $registrar->execute([$fuentes[$fuente]['nombre'], $nombreArchivo, $nombreHoja, $totalImportadas]);
        }
    } catch (Throwable $logError) {
        $avisoBitacora = ' La tabla de datos se actualizó, pero no se pudo registrar la bitácora: ' . $logError->getMessage();
    }

    $mensaje = $fuentes[$fuente]['nombre'] . ': importadas ' . number_format($totalImportadas) . ' filas en la tabla ' . $tabla . '.';
    if ($respaldo !== null) {
        $mensaje .= ' La tabla anterior se conservó como respaldo: ' . $respaldo . '.';
    }
    $mensaje .= $avisoBitacora;
    volverImportador('ok', $mensaje);

} catch (Throwable $e) {
    // Si falla la carga, borrar únicamente la tabla temporal. La tabla original nunca se toca antes del éxito.
    if (is_string($tablaTemporal) && $tablaTemporal !== '') {
        try {
            $pdo->exec('DROP TABLE IF EXISTS ' . qiImportador($tablaTemporal));
        } catch (Throwable $ignorado) {
            // No ocultar el error original.
        }
    }
    volverImportador('error', 'No se completó la importación. La tabla original se conserva si existía. Detalle: ' . $e->getMessage());
}
