<?php

require_once '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

$ruta = explode('/', $_SERVER['PHP_SELF']);
$server = 'http://' . $_SERVER['SERVER_NAME'] . '/' . $ruta[1];

// VERIFICAR SI SE SUBIO EL ARCHIVO
if ($_FILES['file']['name'] == '') {
    header('Location:' . $server . '?error=1');
    exit;
}

$nombreArchivo = $_FILES['file']['name'];
$extension = explode('.', $nombreArchivo);

// VERIFICAR SI SE SUBE EL ARCHIVO CORRECTO
if (end($extension) != 'xlsx') {
    header('Location:' . $server . '?error=2');
    exit;
}

if (empty($_POST['fecha_inicio']) && empty($_POST['fecha_fin'])) {
    header('Location:' . $server . '?error=3');
    exit;
}

// SUBIR ARCHIVO
$archivo = $_FILES['file']['tmp_name'];
$carpeta = __DIR__ . "/../files";

if (!file_exists($carpeta)) {
    mkdir($carpeta);
}

$subido = subirArchivo($archivo, $carpeta);

// CARGAR ARCHIVO EXTERNO
$file = __DIR__ . "/../files/" . $subido;

$cargaArchivo = IOFactory::load($file);

// DATOS DEL ARCHIVO A DESCARGAR
$spreadsheet = new Spreadsheet();

$spreadsheet->getProperties()->setCreator('Deddy Ibañez')
    ->setLastModifiedBy('Deddy Ibañez')
    ->setTitle('Reporte Vencidos')
    ->setSubject('AFOSECAT SAN MARTIN')
    ->setDescription('Reporte Vencidos')
    ->setKeywords('AFOSECAT reporte excel')
    ->setCategory('reporte');

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(14);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(14);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(14);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(14);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(14);

// ESTILOS DEL BORDE
$border = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ]
];

// LOGO
$drawing = new Drawing();
$drawing->setName('Logo');
$drawing->setDescription('Logo');
$drawing->setPath(__DIR__ . '/assets/images/logo-excel.png');
$drawing->setCoordinates('B1');
$drawing->setHeight(50);
$drawing->setOffsetX(60);
$drawing->setOffsetY(10);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

// DATOS DE FECHAS
$fechaInicio = formatoFecha($_POST['fecha_inicio']);
$fechaFin = formatoFecha($_POST['fecha_fin']);

// ENCABEZADO
$spreadsheet->getActiveSheet()->getStyle('A5')->getFont()->setBold(true);
$spreadsheet->getActiveSheet()->getStyle('A5')->getFont()->setSize(16);
$spreadsheet->getActiveSheet()->getStyle('A5')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$spreadsheet->getActiveSheet()->getStyle('A5')->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
$spreadsheet->getActiveSheet()->mergeCells('A5:E5');
$spreadsheet->setActiveSheetIndex(0)
    ->setCellValue('A5', 'REPORTE DE ASISTENCIA');

$spreadsheet->getActiveSheet()->getStyle('A6')->getFont()->setSize(11);
$spreadsheet->getActiveSheet()->getStyle('A6')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$spreadsheet->getActiveSheet()->mergeCells('A6:E6');
$spreadsheet->setActiveSheetIndex(0)
    ->setCellValue('A6', 'DEL ' . $fechaInicio . ' AL ' . $fechaFin);

// GENERAR CONTENIDO
$row = 8;
$arrayGeneral = armarArray($cargaArchivo);

foreach ($arrayGeneral as $item) {
    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FF000000');

    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)
        ->getFont()->getColor()->setARGB('FFFFFFFF');

    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getFont()->setSize(10);
    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->applyFromArray($border);
    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getFont()->setBold(true);
    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
    $spreadsheet->getActiveSheet()->mergeCells('A' . $row . ':E' . $row);
    $spreadsheet->setActiveSheetIndex(0)
        ->setCellValue('A' . $row, $item['empleado']);

    $row += 1;

    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FF808080');

    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)
        ->getFont()->getColor()->setARGB('FFFFFFFF');

    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getFont()->setSize(10);
    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->applyFromArray($border);
    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
    $spreadsheet->getActiveSheet()->mergeCells('B' . $row . ':C' . $row);
    $spreadsheet->getActiveSheet()->mergeCells('D' . $row . ':E' . $row);
    $spreadsheet->setActiveSheetIndex(0)
        ->setCellValue('A' . $row, 'FECHA')
        ->setCellValue('B' . $row, 'TURNO MAÑANA')
        ->setCellValue('D' . $row, 'TURNO TARDE');

    $row += 1;

    // PONER FECHAS Y HORAS
    foreach ($item['asistencia'] as $asistencia) {
        $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getFont()->setSize(10);
        $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->applyFromArray($border);
        $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
        $spreadsheet->getActiveSheet()->getStyle('A' . $row . ':E' . $row)->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

        $spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A' . $row, $asistencia['dia']);

        $letra = 66;

        foreach ($asistencia['horas'] as $hora) {
            $spreadsheet->setActiveSheetIndex(0)
                ->setCellValue(chr($letra) . $row, $hora);

            $letra += 1;
        }

        $row += 1;
    }

    $row += 2;
}

$spreadsheet->getActiveSheet()->setTitle('Reporte de Asistencia');

$nameDocument = 'Reporte_Asistencia.xlsx';

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="' . $nameDocument . '"');
header('Cache-Control: max-age=0');
header('Cache-Control: max-age=1');

$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');

exit;

///////////////////////////////////////////////////////////////////////////////////////
// FUNCIONES
///////////////////////////////////////////////////////////////////////////////////////

function subirArchivo($archivo, $carpeta)
{
    $nombre = date('YmdHis') . '.xlsx';
    $destino = $carpeta . '/' . $nombre;

    return move_uploaded_file($archivo, $destino) ? $nombre : false;
}

function armarArray($cargaArchivo)
{
    $arrayGeneral = [];
    $iGen = 0;
    $iAsi = 0;

    $itemEmpleado = '';
    $itemFecha = '';

    $worksheet = $cargaArchivo->getActiveSheet();
    $highestRow = $worksheet->getHighestDataRow();

    for ($i = 2; $i <= $highestRow; $i++) {
        $empleado = $worksheet->getCell('A' . $i)->getValue();
        $fechaHora = $worksheet->getCell('B' . $i)->getValue();

        $token = explode(' ', $fechaHora);
        $fecha = $token[0];
        $hora = $token[1];

        if ($itemEmpleado != $empleado) {
            if ($itemEmpleado != '') {
                $iGen += 1;
            }

            $arrayGeneral[$iGen] = array_merge([], ['empleado' => $empleado]);
            $arrayGeneral[$iGen] = array_merge($arrayGeneral[$iGen], ['asistencia' => []]);

            $itemEmpleado = $empleado;
        }

        if ($itemFecha != $fecha) {
            if ($itemFecha != '') {
                $iAsi += 1;
            }

            $arrayGeneral[$iGen]['asistencia'][$iAsi] = array_merge([], ['dia' => $fecha]);
            $arrayGeneral[$iGen]['asistencia'][$iAsi] = array_merge($arrayGeneral[$iGen]['asistencia'][$iAsi], ['horas' => []]);

            $itemFecha = $fecha;
        }


        array_push($arrayGeneral[$iGen]['asistencia'][$iAsi]['horas'], $hora);
    }

    return $arrayGeneral;
}

function formatoFecha($fecha)
{
    $nuevaFecha = new DateTime($fecha);
    return $nuevaFecha->format("d/m/Y");
}
