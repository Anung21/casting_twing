<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

// Set Nama File
$fileName = "Casting_Twin_H.xlsx";

// Styling arrays
$titleStyle = [
    'font' => ['bold' => true, 'size' => 22],
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER]
];

$headerStyle = [
    'font' => ['bold' => true, 'size' => 12],
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER],
    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    'fill' => ['fillType' => Fill::FILL_SOLID, 'startColor' => ['rgb' => 'EFEFEF']]
];

$borderStyle = ['borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]]];
$textCenter = ['alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER, 'vertical' => Alignment::VERTICAL_CENTER]];

// Create spreadsheet
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Check Sheet');

try {
    // Set column widths
    $columnWidths = ['A' => 5, 'B' => 20, 'C' => 30, 'D' => 15, 'E' => 15, 'F' => 15, 'G' => 15, 'H' => 15, 'I' => 15, 'J' => 15, 'K' => 15, 'L' => 15, 'M' => 15, 'N' => 15, 'O' => 15, 'P' => 15, 'Q' => 15, 'R' => 15, 'S' => 15];
    foreach ($columnWidths as $col => $width) {
        $sheet->getColumnDimension($col)->setWidth($width);
    }

    // Set row heights
    $rowHeights = [1 => 15, 2 => 30, 3 => 30, 5 => 25];
    foreach ($rowHeights as $row => $height) {
        $sheet->getRowDimension($row)->setRowHeight($height);
    }

    // Set logo
    $logoPath = __DIR__ . '/gambar.jpg'; // Pastikan path benar
    if (file_exists($logoPath)) {
        // Merge area untuk gambar
        $sheet->mergeCells('A1:E4');

        // Hitung ukuran berdasarkan kolom yang di-merge
        $columnWidth = 5 + 20 + 30 + 25 + 25; // Total lebar kolom A-E
        $rowHeight = 15 + 30 + 30 + 25; // Total tinggi baris 1-4

        // Tambahkan gambar
        $drawing = new Drawing();
        $drawing->setName('Company Logo');
        $drawing->setDescription('Company Logo');
        $drawing->setPath($logoPath);
        $drawing->setCoordinates('A1');
        $drawing->setWidth($columnWidth * 7);
        $drawing->setHeight($rowHeight);
        $drawing->setOffsetX(0);
        $drawing->setOffsetY(0);
        $drawing->setWorksheet($sheet);
    }

    // Set title
    $sheet->mergeCells("F1:O5");
    $sheet->setCellValue('F1', 'CHECK SHEET VERIFIKASI JOB SET UP & PATROL CASTING');
    $sheet->getStyle("F1")->applyFromArray($titleStyle);

    // Set form header
    $sheet->mergeCells("A5:E5");
    $sheet->setCellValue('A5', 'FORM');
    $sheet->getStyle("A5")->applyFromArray(['font' => ['bold' => true, 'size' => 14], 'alignment' => $textCenter['alignment']]);

    // Set borders 
    $sheet->getStyle("A1:S31")->applyFromArray([
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['rgb' => '000000'],
            ],
        ],
    ]);

    // Set header data
    $headers = [
        ['P1:S1', 'F0/QAS/Q1/1907'],
        ['P2:Q2', 'Revisi'],
        ['P3:Q3', '0'],
        ['R2:S2', 'Hal'],
        ['R3:S3', '1 / 3'],
        ['P4:S5', 'Berlaku mulai: 04 November 2020']
    ];
    foreach ($headers as [$range, $value]) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(substr($range, 0, strpos($range, ':')), $value);
    }
    $sheet->getStyle("P2:S5")->applyFromArray($textCenter);

    // Set table header
    $tableData = [
        ['A6' => 'Nama part', 'D6' => 'TWIN HEAD', 'J6' => 'Custom', 'K6' => 'PT. DNP', 'P6' => 'Nama mesin'],
        ['A7' => 'No. part', 'D7' => 'P332204-710B', 'J7' => 'Model', 'K7' => '4D34G', 'P7' => 'No. mesin'],
        ['A8' => 'Code', 'D8' => 'AV – 1', 'J8' => 'No. die', 'P8' => 'No. jig']
    ];

    foreach ($tableData as $row) {
        foreach ($row as $cell => $value) {
            $sheet->setCellValue($cell, $value);
        }
    }

    // Apply styling
    $boldCells = ["A5", "D6", "D7", "D8", "K6", "K7", "P1", "P3", "R3"];
    foreach ($boldCells as $cell) {
        $sheet->getStyle($cell)->applyFromArray(['font' => ['bold' => true]]);
    }

    $centerCells = ["P1"];
    foreach ($centerCells as $cell) {
        $sheet->getStyle($cell)->applyFromArray($textCenter);
    }

    // Apply border for used range
    $sheet->getStyle("A1:S8")->applyFromArray($borderStyle);

    // Set default font size set ke dua
    $spreadsheet->getDefaultStyle()->getFont()->setSize(12);

    // Daftar header tabel dengan merge cells
    $Headers = [
        'A32:B33' => 'No.',
        'C32:D33' => 'Item',
        'E32:G33' => 'Standard',
        'H32:H33' => 'Setup',
        'I32:I33' => 'Patrol',
        'J32:J32' => 'Control',
        'J33:J33' => 'Method',
        'K32:S32' => 'Cavity Sample',
    ];

    // Styling untuk header tabel
    $headerStyle = [
        'font' => ['bold' => true],
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
        ],
        'borders' => [
            'allBorders' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN]
        ]
    ];

    // Memberikan border untuk setiap kolom di baris K33:S33
    for ($col = 'K'; $col <= 'S'; $col++) {
        $cell = $col . '33';
        $sheet->getStyle($cell)->applyFromArray([
            'borders' => [
                'allBorders' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN]
            ]
        ]);
    }

    // Loop untuk mengatur merge, nilai, dan style pada tiap header
    foreach ($tableHeaders as $range => $title) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(explode(':', $range)[0], $title);
        $sheet->getStyle($range)->applyFromArray($headerStyle);
    }

    // Merge sel A34:S34 untuk judul bagian
    $sheet->mergeCells('A34:S34');
    $sheet->setCellValue('A34', 'I. LUBANG & INSERT PIN AREA MOVE');

    // Styling untuk judul bagian
    $sectionTitleStyle = [
        'font' => [
            'bold' => true,
            'italic' => true
        ],
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
        ],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'startColor' => ['rgb' => 'D9F1F1'] // Warna latar biru muda seperti gambar
        ],
        'borders' => [
            'outline' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE]
        ]
    ];

    // Terapkan styling ke sel A34:S34
    $sheet->getStyle('A34:S34')->applyFromArray($sectionTitleStyle);

    // Daftar sel yang akan di-merge beserta nilainya
    $mergedCells = [
        ['range' => 'A35:B39', 'value' => '1'],
        ['range' => 'A40:B46', 'value' => '2'],
        ['range' => 'A47:B48', 'value' => '3'],
        ['range' => 'A49:B51', 'value' => '4'],
        ['range' => 'A52:B53', 'value' => '5'],
        ['range' => 'A54:B54', 'value' => '6'],
        ['range' => 'A55:B55', 'value' => '7'],
        ['range' => 'A56:B7', 'value' => '8'],
    ];

    // Looping untuk mengatur merge, nilai, style, dan border pada setiap sel
    foreach ($mergedCells as $cell) {
        $sheet->mergeCells($cell['range']);
        $sheet->setCellValue(explode(':', $cell['range'])[0], $cell['value']);
        $sheet->getStyle($cell['range'])->applyFromArray($textCenter);

        // Tambahkan border
        $sheet->getStyle($cell['range'])->applyFromArray([
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
                ]
            ]
        ]);
    }

} catch (\Exception $e) {
    die('Error: ' . $e->getMessage());
}

// Output the Excel file
$writer = new Xlsx($spreadsheet);
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"$fileName\"");
$writer->save('php://output');
exit();
