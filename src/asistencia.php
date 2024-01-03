<?php

// require_once '../vendor/autoload.php';

// VERIFICAR SI SE SUBIO EL ARCHIVO
if ($_FILES['file']['name'] == '') {
    echo "NADA";
    exit();
}

$nombreArchivo = $_FILES['file']['name'];
$extension = explode('.', $nombreArchivo);

// VERIFICAR SI SE SUBE EL ARCHIVO CORRECTO
if (end($extension) != 'xlsx') {
    echo "NO ES LA EXTENSION";
    exit();
}

// SUBIR ARCHIVO
$archivo = $_FILES['file']['tmp_name'];
$carpeta = __DIR__ . "/../files";

$nuevoNombre = cambiarNombre();

echo $nuevoNombre;


// use PhpOffice\PhpSpreadsheet\IOFactory;
// use PhpOffice\PhpSpreadsheet\Spreadsheet;
// use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

// // CARGAR ARCHIVO EXTERNO
// $file = __DIR__ . "/../files/principal.xlsx";

// $cargaArchivo = IOFactory::load($file);
// $worksheet = $cargaArchivo->getActiveSheet();
// $highestRow = $worksheet->getHighestDataRow();

// // DATOS DEL ARCHIVO A DESCARGAR
// $spreadsheet = new Spreadsheet();

// $spreadsheet->getProperties()->setCreator('Deddy Ibañez')
//     ->setLastModifiedBy('Deddy Ibañez')
//     ->setTitle('Reporte Vencidos')
//     ->setSubject('AFOSECAT SAN MARTIN')
//     ->setDescription('Reporte Vencidos')
//     ->setKeywords('AFOSECAT reporte excel')
//     ->setCategory('reporte');

// $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(14);
// $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(14);
// $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(14);
// $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(14);
// $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(14);

// // ESTILOS DEL BORDE
// $border = [
//     'borders' => [
//         'allBorders' => [
//             'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
//         ],
//     ]
// ];

// // LOGO
// $drawing = new Drawing();
// $drawing->setName('Logo');
// $drawing->setDescription('Logo');
// $drawing->setPath(__DIR__ . '/assets/images/logo-excel.png');
// $drawing->setCoordinates('A1');
// $drawing->setHeight(50);
// $drawing->setOffsetX(10);
// $drawing->setOffsetY(50);
// $drawing->setWorksheet($spreadsheet->getActiveSheet());



// // for ($i = 2; $i <= $highestRow; $i++) {
// //     echo $worksheet->getCell('A' . $i)->getValue() . "<br>";
// // }

// $spreadsheet->getActiveSheet()->setTitle('Reporte de Vencidos');

// $nameDocument = 'Reporte_Asistencia.xlsx';

// header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
// header('Content-Disposition: attachment;filename="' . $nameDocument . '"');
// header('Cache-Control: max-age=0');
// header('Cache-Control: max-age=1');

// $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
// $writer->save('php://output');

// exit;


function cambiarNombre()
{
    return 'HOLA';
}
