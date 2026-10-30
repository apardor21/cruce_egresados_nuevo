# Cruce de egresados ULYSES + WAS + UnoAUno

Proyecto nuevo PHP + MySQL para importar tres Excel en tablas separadas (`ulyses`, `was`, `unoau`) conservando los encabezados como columnas. Se agregan dos columnas técnicas: `__id` y `__documento_normalizado`.

## Instalación
1. Copia la carpeta a `C:\xampp\htdocs\cruce_egresados_nuevo`.
2. En CMD ejecuta `cd C:\xampp\htdocs\cruce_egresados_nuevo` y luego `composer install`.
3. Inicia Apache y MySQL en XAMPP.
4. Importa `database.sql` en phpMyAdmin.
5. Si tu MySQL tiene otra clave, ajusta `config.php`.
6. Abre `http://localhost/cruce_egresados_nuevo/`.

## Uso
Carga cada Excel por separado. Por defecto toma la primera hoja y la primera fila como encabezado; puedes indicar otra hoja y fila. Una nueva importación reemplaza únicamente la tabla de esa fuente. Los valores se guardan como texto para no convertir identificaciones ni perder ceros iniciales. La columna de identificación se detecta con nombres habituales: ULYSES `Identificación`, WAS `dniper`, UnoAUno `NUM_DOCUMENTO`.

## Notas
- La base `cruce_egresados` debe existir; `database.sql` la crea junto con el registro de importaciones.
- El importador admite XLSX, XLS y CSV.
- No se sube ningún Excel de prueba; los datos permanecen en tu computador/servidor.
