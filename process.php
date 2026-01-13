<?php

require __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use Google\Client;
use Google\Service\Sheets;
use PhpOffice\PhpSpreadsheet\Spreadsheet;


/* -----------------------------
   CONFIGURATION
----------------------------- */
$spreadsheetId = 'your_ID';
$sheetRange    = 'mb fb data!A:G';

/* -----------------------------
   1. CHECK FILE UPLOAD
----------------------------- */
if (!isset($_FILES['excel']) || $_FILES['excel']['error'] !== 0) {
    die('Excel file upload failed');
}

/* -----------------------------
   2. READ EXCEL FILE
----------------------------- */
$tmpFile = $_FILES['excel']['tmp_name'];

$spreadsheet = IOFactory::load($tmpFile);
$sheet       = $spreadsheet->getActiveSheet();
$rows        = $sheet->toArray();

$excelNumbers = [];

foreach ($rows as $row) {
    if (!empty($row[0])) {           // Column A
        $excelNumbers[] = trim($row[0]);
    }
}

/* -----------------------------
   3. GOOGLE SHEETS AUTH
----------------------------- */
$client = new Client();
$client->setApplicationName('Excel Checker');
$client->setScopes([Sheets::SPREADSHEETS_READONLY]);
$client->setAuthConfig(__DIR__ . '/credentials.json');
$client->setAccessType('offline');

$service = new Sheets($client);

/* -----------------------------
   4. READ GOOGLE SHEET
----------------------------- */
$response = $service->spreadsheets_values->get(
    $spreadsheetId,
    $sheetRange
);

$googleRows = $response->getValues();

/* -----------------------------
   5. FIND DUPLICATES
----------------------------- */
$duplicates = [];

foreach ($googleRows as $gRow) {
    $number = $gRow[1] ?? null;   // Column B
    $date   = $gRow[6] ?? null;   // Column G

    if ($number && in_array($number, $excelNumbers)) {

        if (!isset($duplicates[$number])) {
            $duplicates[$number] = [
                'count' => 0,
                'dates' => []
            ];
        }

        $duplicates[$number]['count']++;

        if ($date) {
            $duplicates[$number]['dates'][] = $date;
        }
    }
}

/* -----------------------------
   6. OUTPUT RESULT
----------------------------- */
// header('Content-Type: application/json');

// $result = [];

// foreach ($duplicates as $number => $info) {
//     $result[] = [
//         'number' => $number,
//         'duplicate_count' => $info['count'],
//         'dates_in_column_g' => $info['dates']
//     ];
// }

// echo json_encode($result, JSON_PRETTY_PRINT);


$exportSpreadsheet = new Spreadsheet();
$exportSheet = $exportSpreadsheet->getActiveSheet();

// Header row
$exportSheet->setCellValue('A1', 'Number');
$exportSheet->setCellValue('B1', 'Duplicate Count');
$exportSheet->setCellValue('C1', 'Dates (Column G)');

$rowNum = 2;

foreach ($duplicates as $number => $info) {
    $exportSheet->setCellValue("A$rowNum", $number);
    $exportSheet->setCellValue("B$rowNum", $info['count']);
    $exportSheet->setCellValue("C$rowNum", implode(", ", $info['dates']));
    $rowNum++;
}

/* -----------------------------
   7. SEND EXCEL TO BROWSER
----------------------------- */
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="yoursheetname.xlsx"');
header('Cache-Control: max-age=0');

$writer = IOFactory::createWriter($exportSpreadsheet, 'Xlsx');
$writer->save('php://output');
exit;

