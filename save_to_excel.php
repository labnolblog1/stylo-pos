<?php
require 'vendor/autoload.php'; // only if PhpSpreadsheet is installed

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$data = json_decode(file_get_contents("php://input"), true);

$file = __DIR__ . "/data/sales.xlsx";

// Create folder if not exists
if (!file_exists(__DIR__ . "/data")) {
    mkdir(__DIR__ . "/data", 0777, true);
}

if (file_exists($file)) {
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();
    $row = $sheet->getHighestRow() + 1;
} else {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $sheet->fromArray(
        ["Date", "Salesman", "Sales Amount", "Commission"],
        NULL,
        "A1"
    );
    $row = 2;
}

$sheet->fromArray([
    $data['date'],
    $data['salesman'],
    $data['amount'],
    $data['commission']
], NULL, "A$row");

$writer = new Xlsx($spreadsheet);
$writer->save($file);

echo json_encode(["status" => "success"]);
