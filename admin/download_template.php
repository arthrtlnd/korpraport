<?php
session_start();
require_once '../app/koneksi.php';
check_admin();

// Load library
require_once '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

// Cek ekstensi ZIP
if (!extension_loaded('zip')) {
    die("Error: Ekstensi PHP 'zip' belum aktif. Tidak bisa generate file Excel.");
}

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// --- 1. SETUP HEADER ---
$headers = [
    'A1' => 'NRP',
    'B1' => 'Nama',
    'C1' => 'NIK',
    'D1' => 'Tempat Lahir',
    'E1' => 'Tanggal Lahir (YYYY-MM-DD)',
    'F1' => 'Kode Pangkat',
    'G1' => 'Kode Korp',
    'H1' => 'Kode Matra',
    'I1' => 'Kode Satker Baru',
    'J1' => 'Kode Satker Lama'
];

foreach ($headers as $cell => $value) {
    $sheet->setCellValue($cell, $value);
    $sheet->getStyle($cell)->getFont()->setBold(true);
    $sheet->getStyle($cell)->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FFCCCCCC'); // Warna abu-abu
}

// --- 2. DATA CONTOH (DUMMY) ---
// Baris 2: Contoh Data
$sheet->setCellValueExplicit('A2', '111111', DataType::TYPE_STRING); // NRP String
$sheet->setCellValue('B2', 'CONTOH PERSONEL AD');
$sheet->setCellValueExplicit('C2', '3175000000000001', DataType::TYPE_STRING); // NIK String
$sheet->setCellValue('D2', 'Jakarta');
$sheet->setCellValue('E2', '1990-01-01');
$sheet->setCellValueExplicit('F2', '73', DataType::TYPE_STRING); // Kapten
$sheet->setCellValueExplicit('G2', 'C1', DataType::TYPE_STRING); // Arm
$sheet->setCellValue('H2', '1'); // AD
$sheet->setCellValueExplicit('I2', 'D13', DataType::TYPE_STRING); // Pusinfolahta
$sheet->setCellValueExplicit('J2', 'B02', DataType::TYPE_STRING); // Itjen

// Baris 3: Contoh Data 2
$sheet->setCellValueExplicit('A3', '222222', DataType::TYPE_STRING);
$sheet->setCellValue('B3', 'CONTOH PERSONEL AL');
$sheet->setCellValueExplicit('C3', '3175000000000002', DataType::TYPE_STRING);
$sheet->setCellValue('D3', 'Surabaya');
$sheet->setCellValue('E3', '1992-05-20');
$sheet->setCellValueExplicit('F3', '81', DataType::TYPE_STRING); // Mayor
$sheet->setCellValueExplicit('G3', '62', DataType::TYPE_STRING); // K
$sheet->setCellValue('H3', '2'); // AL
$sheet->setCellValueExplicit('I2', 'B02', DataType::TYPE_STRING); 
$sheet->setCellValueExplicit('J2', 'A00', DataType::TYPE_STRING); 

// --- 3. FORMATTING ---
// Set kolom NRP dan NIK jadi Text agar angka 0 di depan tidak hilang
$sheet->getStyle('A:A')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT);
$sheet->getStyle('C:C')->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT);
$sheet->getStyle('E:E')->getNumberFormat()->setFormatCode('yyyy-mm-dd');

// Auto width
foreach(range('A','J') as $columnID) {
    $sheet->getColumnDimension($columnID)->setAutoSize(true);
}

// --- 4. OUTPUT DOWNLOAD ---
$filename = 'Template_Import_Personel.xlsx';

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="'.$filename.'"');
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;
?>