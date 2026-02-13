<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

$data = json_decode(file_get_contents("php://input"), true);

$folder = __DIR__ . "/data";
$file = $folder . "/sales.xlsx";

if (!file_exists($folder)) {
    mkdir($folder, 0777, true);
}

if (file_exists($file)) {
    $spreadsheet = IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();
    $row = $sheet->getHighestRow() + 1;
} else {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1', 'Date');
    $sheet->setCellValue('B1', 'Salesman');
    $sheet->setCellValue('C1', 'Sales Amount');
    $sheet->setCellValue('D1', 'Commission');

    $row = 2;
}

$sheet->setCellValue("A$row", $data['date']);
$sheet->setCellValue("B$row", $data['salesman']);
$sheet->setCellValue("C$row", $data['amount']);
$sheet->setCellValue("D$row", $data['commission']);

$writer = new Xlsx($spreadsheet);
$writer->save($file);

echo json_encode(["status" => "success"]);
