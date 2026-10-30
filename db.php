<?php

declare(strict_types=1);
$c = require __DIR__ . '/config.php';
try {
    $pdo = new PDO("mysql:host={$c['host']};dbname={$c['database']};charset={$c['charset']}", $c['username'], $c['password'], [PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION, PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC, PDO::ATTR_EMULATE_PREPARES => false]);
} catch (PDOException $e) {
    http_response_code(500);
    exit('Error de conexión. Revisa config.php y database.sql. ' . htmlspecialchars($e->getMessage()));
}
