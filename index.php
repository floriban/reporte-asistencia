<!DOCTYPE html>
<html lang="es">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte de Asistencia</title>

    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-EVSTQN3/azprG1Anm3QDgpJLIm9Nao0Yz1ztcQTwFspd3yD65VohhpuuCOmLASjC" crossorigin="anonymous">
</head>

<body>
    <div class="container mt-5">
        <h1 class="text-center">Reporte de Asistencia</h1>
        <p class="text-center">Suba un archivo en formato .xlsx para exportar el reporte de asistencia</p>

        <form action="src/asistencia.php" method="POST" class="form" enctype='multipart/form-data'>
            <?php
            if (isset($_GET['error'])) {
                if ($_GET['error'] == 1) {
                    echo '<div class="alert alert-danger" role="alert">Debe subir una archivo de extensión .xlsx</div>';
                }

                if ($_GET['error'] == 2) {
                    echo '<div class="alert alert-danger" role="alert">Sólo esta admitido archivos de extensión .xlsx</div>';
                }

                if ($_GET['error'] == 3) {
                    echo '<div class="alert alert-danger" role="alert">Debe ingresar el rango de fechas de la asistencia</div>';
                }
            }
            ?>

            <div class="row">
                <div class="col-md-3">
                    <div class="form-group">
                        <label for="fecha_inicio">Fecha Inicio</label>
                        <input type="date" class="form-control" id="fecha_inicio" name="fecha_inicio" required>
                    </div>
                </div>

                <div class="col-md-3">
                    <div class="form-group">
                        <label for="fecha_fin">Fecha Fin</label>
                        <input type="date" class="form-control" id="fecha_fin" name="fecha_fin" required>
                    </div>
                </div>

                <div class="col-md-6">
                    <div class="form-group">
                        <label for="file">Subir archivo</label>
                        <input type="file" class="form-control" id="file" name="file" accept=".xlsx" required>
                    </div>
                </div>
            </div>

            <div class="d-grid gap-2">
                <input type="submit" value="Subir y exportar" class="btn btn-secondary mt-2 btn-block">
            </div>
        </form>
    </div>
</body>

</html>