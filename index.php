<!DOCTYPE html>
<html lang="en">

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
            <div class="form-group">
                <label for="file">Subir archivo</label>
                <input type="file" class="form-control" id="file" name="file" accept=".xlsx" required>
            </div>

            <div class="d-grid gap-2">
                <input type="submit" value="Subir y exportar" class="btn btn-secondary mt-2 btn-block">
            </div>
        </form>
    </div>
</body>

</html>