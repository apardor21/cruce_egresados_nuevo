<?php require __DIR__ . '/db.php';
require __DIR__ . '/functions.php'; ?>
<!doctype html>
<html lang="es">

<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Tablas</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
</head>

<body class="bg-light">
    <main class="container py-4"><a class="btn btn-outline-secondary btn-sm mb-3" href="index.php">← Inicio</a>
        <h1 class="h3">Estructura de tablas</h1><?php foreach (['ulyses' => 'ULYSES', 'was' => 'WAS', 'unoau' => 'UnoAUno'] as $t => $label): ?><section class="card card-body mb-3">
                <h2 class="h5"><?= h($label) ?> — <?= h($t) ?></h2><?php try {
                                                                    $cols = $pdo->query('SHOW COLUMNS FROM ' . qi($t))->fetchAll(); ?><div class="table-responsive">
                        <table class="table table-sm">
                            <tr>
                                <th>Campo</th>
                                <th>Tipo</th>
                                <th>Clave</th>
                            </tr><?php foreach ($cols as $c): ?><tr>
                                    <td><?= h($c['Field']) ?></td>
                                    <td><?= h($c['Type']) ?></td>
                                    <td><?= h($c['Key']) ?></td>
                                </tr><?php endforeach; ?>
                        </table>
                    </div><?php } catch (Throwable $e) { ?><p>Tabla aún no creada.</p><?php } ?>
            </section><?php endforeach; ?>
    </main>
</body>

</html>