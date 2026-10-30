<?php require __DIR__ . '/db.php';

require __DIR__ . '/functions.php';
$tables = ['ulyses' => 'ULYSES', 'was' => 'WAS', 'unoau' => 'UnoAUno'];
$counts = [];
foreach ($tables as $t => $label) {
    try {
        $counts[$label] = (int)$pdo->query('SELECT COUNT(*) FROM ' . qi($t))->fetchColumn();
    } catch (Throwable $e) {
        $counts[$label] = null;
    }
}
try {
    $logs = $pdo->query('SELECT * FROM importaciones ORDER BY id DESC LIMIT 8')->fetchAll();
} catch (Throwable $e) {
    $logs = [];
}
?>
<!doctype html>
<html lang="es">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Cruce de egresados</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {
            background: #f4f7fb
        }

        .hero {
            background: linear-gradient(120deg, #132b4f, #176b87);
            color: white;
            border-radius: 18px;
            padding: 26px
        }

        .card {
            border: 0;
            border-radius: 15px;
            box-shadow: 0 5px 20px #0001
        }

        .btn-primary {
            background: #176b87;
            border-color: #176b87
        }
    </style>
</head>

<body>
    <main class="container py-4">
        <div class="hero mb-4">
            <h1 class="h3">Cruce de egresados</h1>
            <p class="mb-0">Tres fuentes separadas, columnas originales y cruce por identificación.</p>
        </div>
        <?php if (isset($_GET['ok'])): ?><div class="alert alert-success"><?= h($_GET['ok']) ?></div><?php endif; ?><?php if (isset($_GET['error'])): ?><div class="alert alert-danger"><?= h($_GET['error']) ?></div><?php endif; ?>
        <div class="row g-3 mb-4"><?php foreach ($counts as $label => $n): ?><div class="col-md-4">
                    <div class="card p-3">
                        <div class="text-muted">Registros <?= h($label) ?></div>
                        <div class="display-6"><?= $n === null ? '—' : number_format($n) ?></div>
                    </div>
                </div><?php endforeach; ?></div>
        <div class="row g-3"><?php foreach (['ulyses' => ['ULYSES', 'Identificación'], 'was' => ['WAS', 'dniper'], 'unoau' => ['UnoAUno', 'NUM_DOCUMENTO']] as $key => $meta): ?><div class="col-lg-4">
                    <div class="card p-4 h-100">
                        <h2 class="h5"><?= h($meta[0]) ?></h2>
                        <p class="text-muted small">La importación reemplaza solo la tabla de esta fuente.</p>
                        <form action="importar.php" method="post" enctype="multipart/form-data"><input type="hidden" name="fuente" value="<?= h($key) ?>"><label class="form-label">Archivo Excel/CSV</label><input class="form-control mb-3" type="file" name="archivo" accept=".xlsx,.xls,.csv" required><label class="form-label">Nombre de hoja (opcional)</label><input class="form-control mb-3" name="hoja" placeholder="Primera hoja por defecto"><label class="form-label">Fila de encabezados</label><input class="form-control mb-3" type="number" min="1" name="fila_encabezados" value="1" required><button class="btn btn-primary w-100">Importar <?= h($meta[0]) ?></button></form>
                    </div>
                </div><?php endforeach; ?></div>
        <div class="my-4 d-flex gap-2 flex-wrap"><a class="btn btn-success" href="cruzar.php">Cruzar identificación</a><a class="btn btn-outline-secondary" href="estructura.php">Ver tablas</a></div>
        <div class="card p-3">
            <h2 class="h5">Últimas importaciones</h2>
            <div class="table-responsive">
                <table class="table table-sm">
                    <thead>
                        <tr>
                            <th>Fuente</th>
                            <th>Archivo</th>
                            <th>Hoja</th>
                            <th>Filas</th>
                            <th>Fecha</th>
                        </tr>
                    </thead>
                    <tbody><?php foreach ($logs as $l): ?><tr>
                                <td><?= h($l['fuente']) ?></td>
                                <td><?= h($l['nombre_archivo']) ?></td>
                                <td><?= h($l['hoja']) ?></td>
                                <td><?= number_format($l['total_filas']) ?></td>
                                <td><?= h($l['fecha_carga']) ?></td>
                            </tr><?php endforeach; ?><?php if (!$logs): ?><tr>
                                <td colspan="5">Sin importaciones todavía.</td>
                            </tr><?php endif; ?></tbody>
                </table>
            </div>
        </div>
    </main>
</body>

</html>