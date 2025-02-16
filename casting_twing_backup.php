<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Font;
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

// ======================== PAGE 1 ===============
// --------------------- Bagian 1 ------------- 
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
        ['A8' => 'Code', 'D8' => 'AV – 1', 'J8' => 'No. die', 'P8' => 'No. jig'],
    ];

    // Set marge keterangan-gambar B9
    $sheet->mergeCells("B9");
    $sheet->setCellValue('B9', 'Gambar : ');
    $sheet->getStyle("B9")->applyFromArray($titleStyle);

    //  style B9 
    $sheet->getStyle("B9")->getFont()->setSize(12);
    $sheet->getStyle("B9")->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
    $sheet->getStyle("B9")->getAlignment()->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);

    foreach ($tableData as $row) {
        foreach ($row as $cell => $value) {
            $sheet->setCellValue($cell, $value);
        }
    }

    $centerCells = ["P1"];
    foreach ($centerCells as $cell) {
        $sheet->getStyle($cell)->applyFromArray($textCenter);
    }

    // Apply border for used range
    $sheet->getStyle("A1:S8")->applyFromArray($borderStyle);

    // Set default font size set ke dua
    //  border di sekitar range A32:S62
    $sheet->getStyle('A32:S62')->applyFromArray([
        'borders' => [
            'outline' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
            ]
        ]
    ]);
    $spreadsheet->getDefaultStyle()->getFont()->setSize(12);

    // --------------------------- BAGIAN 2 -----------------------
    // Daftar header tabel dengan merge cells
    $tableHeaders = [
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

    // Daftar judul bagian dan rentang selnya
    $sections = [
        ['A34:S34', '   I. LUBANG & INSERT PIN AREA MOVE'],
        ['A58:G58', 'II. LUBANG & INSERT PIN AREA CORE 1'],
        ['I58:S58', '']
    ];

    // Styling untuk judul bagian
    $sectionTitleStyle = [
        'font' => ['bold' => true, 'italic' => true],
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
        ],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'startColor' => ['rgb' => 'D9F1F1'] // Warna latar biru muda
        ],
        'borders' => [
            'outline' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE]
        ]
    ];

    // Iterasi untuk mengatur judul bagian
    foreach ($sections as [$range, $title]) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(explode(':', $range)[0], $title);
        $sheet->getStyle($range)->applyFromArray($sectionTitleStyle);
    }


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
        ['range' => 'A56:B57', 'value' => '8'],
        ['range' => 'A59:B60', 'value' => '9'],
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

    // Definisi sel yang akan digabungkan dan diisi dengan teks
    $cells = [
        // Cells 1
        ['C35:D37', 'a.Ex. Ejector pin (8plcs)'],
        ['C38:D39', 'b.Height Ejector pin (3plcs)'],
        ['E35:G35', 'Tidak minus / amblas'],
        ['E36:F37', '( Reff. 0 (0 ~ +0.3))'],
        ['E38:F39', '1 [0 ~ +0.5]'],
        ['G36', 'Max'],
        ['G37', 'Min'],
        ['G38', 'Max'],
        ['G39', 'Min'],

        // Cells 1 bagian kanan 
        ['J35', 'Visual'],
        ['J36:J37', 'Caliper'],
        ['J38:J39', 'Caliper'],
        ['K35', 'OK / NG'],
        ['L35', 'OK / NG'],
        ['M35', 'OK / NG'],
        ['N35', 'OK / NG'],
        ['O35', 'OK / NG'],
        ['P35', 'OK / NG'],
        ['Q35', 'OK / NG'],
        ['R35', 'OK / NG'],
        ['S35', 'OK / NG'],

        // Cells 2
        ['C40:D41', 'a. Inner diameter(2plcs)'],
        ['C42:D43', 'b. Depth hole (2plcs)'],
        ['C44:D46', 'c.Ex. Ejector bushing(2plcs)'],
        ['E40:F41', 'Ø15 ±0.3'],
        ['E42:F43', '40.5 ±0.3'],
        ['E44:G44', 'Tidak minus / amblas'],
        ['E45:F46', '[ Reff. 0 (0 ~ +0.3)]'],
        ['G40', '1'],
        ['G41', '2'],
        ['G42', '1'],
        ['G43', '2'],
        ['G45', '1'],
        ['G46', '2'],

        // Cells 2 bagian kanan 
        ['I40', '-'],
        ['I41', '-'],
        ['I42', '-'],
        ['I43', '-'],
        ['J44', 'visual'],
        ['J40:J41', 'Caliper'],
        ['J42:J43', 'Caliper'],
        ['J45:J46', 'Caliper'],
        ['K44', 'OK / NG'],
        ['L44', 'OK / NG'],
        ['M44', 'OK / NG'],
        ['N44', 'OK / NG'],
        ['O44', 'OK / NG'],
        ['P44', 'OK / NG'],
        ['Q44', 'OK / NG'],
        ['R44', 'OK / NG'],
        ['S44', 'OK / NG'],

        // Cells 3
        ['C47:D47', 'a. Inner diameter ujung'],
        ['C48:D48', 'b. Depth hole'],
        ['E47:G47', 'Ø13 (0 ~ +0.3)'],
        ['E48:G48', '56 (0 ~ +0.5)'],

        // Cells 3 bagian kanan 
        ['J47', 'Caliper'],
        ['J48', 'Caliper'],

        // Cells 4
        ['C49:D49', 'a. Inner diameter'],
        ['C50:D51', 'b.Appearance hole'],
        ['E49:G49', 'Ø7 [-0.3 ~ 0]'],
        ['E50:G51', 'Tidak: undercut, overheat, bengkok, scrap tipis (max 0.3).'],

        // Cells 4 bagian kanan 
        ['I49', '-'],
        ['J49', 'Caliper'],
        ['J50:J51', 'Visual'],
        ['K50:K51', 'OK / NG'],
        ['L50:L51', 'OK / NG'],
        ['M50:M51', 'OK / NG'],
        ['N50:N51', 'OK / NG'],
        ['O50:O51', 'OK / NG'],
        ['P50:P51', 'OK / NG'],
        ['Q50:Q51', 'OK / NG'],
        ['R50:R51', 'OK / NG'],
        ['S50:S51', 'OK / NG'],

        // Cells 5
        ['C52:D52', 'a. Inner diameter'],
        ['C53:D53', 'b. Depth hole'],
        ['E52:G52', 'Ø7 [-0.3 ~ 0]'],
        ['E53:G53', '8 ±0.3'],

        // Cells 5 bagian kanan 
        ['I52', '-'],
        ['J52', 'Caliper'],
        ['J53', 'Caliper'],

        // Cells 6
        ['C54:D54', 'Lot Marking'],
        ['E54:G54', 'Ø7 [-0.3 ~ 0]'],

        // Cells 6 bagian kanan 
        ['J54', 'Visual'],
        ['K54:K54', 'OK / NG'],
        ['L54:L54', 'OK / NG'],
        ['M54:M54', 'OK / NG'],
        ['N54:N54', 'OK / NG'],
        ['O54:O54', 'OK / NG'],
        ['P54:P54', 'OK / NG'],
        ['Q54:Q54', 'OK / NG'],
        ['R54:R54', 'OK / NG'],
        ['S54:S54', 'OK / NG'],

        // Cells 7
        ['C55:D55', 'Appearance ex gate'],
        ['E55:G55', 'Tidak : gompal & keropos'],

        //  Cells 7 bagian kanan 
        ['J55', 'Visual'],
        ['K55:K55', 'OK / NG'],
        ['L55:L55', 'OK / NG'],
        ['M55:M55', 'OK / NG'],
        ['N55:N55', 'OK / NG'],
        ['O55:O55', 'OK / NG'],
        ['P55:P55', 'OK / NG'],
        ['Q55:Q55', 'OK / NG'],
        ['R55:R55', 'OK / NG'],
        ['S55:S55', 'OK / NG'],

        // Cells 8
        ['C56:D57', 'Appearance permukaan move'],
        ['E56:G57', 'Tidak : Retak, undercut, flowline, overheat & gompal.'],

        //  Cells 8 bagian kanan
        ['J56:J57', 'Visual'],
        ['K56:K57', 'OK / NG'],
        ['L56:L57', 'OK / NG'],
        ['M56:M57', 'OK / NG'],
        ['N56:N57', 'OK / NG'],
        ['O56:O57', 'OK / NG'],
        ['P56:P57', 'OK / NG'],
        ['Q56:Q57', 'OK / NG'],
        ['R56:R57', 'OK / NG'],
        ['S56:S57', 'OK / NG'],

        //  Cells 9
        ['C59:D59', 'b.Thickness'],
        ['C60:D60', 'c.Thickness'],
        ['E59:G59', '11.07 ±0.3'],
        ['E60:G60', '10.5 (0 ~ +0.3)'],

        //  Cells 9 bagian kanan
        ['J59', 'Caliper'],
        ['J60', 'Caliper'],
    ];

    // Loop tunggal untuk memproses semua sel
    foreach ($cells as [$range, $text]) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(explode(':', $range)[0], $text);

        $sheet->getStyle($range)->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, // Pusatkan teks horizontal
                'vertical' => Alignment::VERTICAL_CENTER // Pusatkan teks vertikal
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'FFFFFF'] // Warna putih
            ],
            'font' => ['bold' => false], // Hapus bold
            'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]]
        ]);
    }

    // Merge kolom H35:H60
    $sheet->mergeCells('H35:H60');
    $sheet->setCellValue('H35', 'All Point If Job Set Up');

    // Styling untuk teks vertikal dan center
    $sheet->getStyle('H35:H60')->applyFromArray([
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER, // Pusatkan secara horizontal
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER, // Pusatkan secara vertikal
            'textRotation' => 90 // Putar teks 90 derajat (vertikal)
        ],
        'font' => [
            'bold' => true
        ],
        'borders' => [
            'allBorders' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN]
        ]
    ]);

    // Daftar kolom yang akan di-merge dan teksnya
    $mergeCells = [
        'I35:I39' => '1x / Shift',
        'I44:I46' => '1x / Shift',
        'I47:I48' => '1x / Shift',
        'I50:I51' => '1x / Shift',
        'I53:I55' => '1x / Shift',
        'I56:I57' => '1x / Shift',
        'I59:I60' => '1x / Shift',

    ];

    // Terapkan merge, set nilai, dan styling
    foreach ($mergeCells as $range => $text) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(explode(':', $range)[0], $text);
        $sheet->getStyle($range)->applyFromArray([
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'textRotation' => 90
            ],
            'font' => ['bold' => false],
            'borders' => ['allBorders' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN]]
        ]);
    }

    // Daftar range yang akan diberikan border
    $ranges = ['K36:S39', 'K40:S41', 'K42:S43', 'K45:S46', 'K47:S48', 'K49:S49', 'K52:S53', 'K59:S60', 'K74:S76', 'K79:S79', 'K80:S80', 'K86:S89', 'K92:S94', 'K99:S100', 'K103:S110', 'K59:S60'];
    // Terapkan border dan alignment ke setiap range
    foreach ($ranges as $range) {
        $sheet->getStyle($range)->applyFromArray([
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN // Border tipis
                ]
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER, // Pusatkan teks secara horizontal
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER // Pusatkan teks secara vertikal
            ]
        ]);
    }

    // ==================================== PAGE 2 ====================
// --------------------- Bagian 1 ------------- 

    // Set column widths
    $columnWidths = [
        'A' => 5,
        'B' => 20,
        'C' => 30,
        'D' => 15,
        'E' => 15,
        'F' => 15,
        'G' => 15,
        'H' => 15,
        'I' => 15,
        'J' => 15,
        'K' => 15,
        'L' => 15,
        'M' => 15,
        'N' => 15,
        'O' => 15,
        'P' => 15,
        'Q' => 15,
        'R' => 15,
        'S' => 15
    ];
    foreach ($columnWidths as $col => $width) {
        $sheet->getColumnDimension($col)->setWidth($width);
    }

    // Set row heights
    $rowHeights = [1 => 15, 2 => 30, 3 => 30, 5 => 25];
    foreach ($rowHeights as $row => $height) {
        $sheet->getRowDimension($row)->setRowHeight($height);
    }

    // // Set logo
    $logoPath = __DIR__ . '/gambar.jpg'; // Pastikan path benar
    if (file_exists($logoPath)) {
        // Merge area untuk gambar
        $sheet->mergeCells('A63:E66');

        // Sesuaikan ukuran gambar agar ringkas
        $drawing = new Drawing();
        $drawing->setName('Company Logo');
        $drawing->setDescription('Company Logo');
        $drawing->setPath($logoPath);
        $drawing->setCoordinates('A63');
        $drawing->setWidth(480); // Atur lebar gambar secara proporsional
        $drawing->setHeight(70); // Sesuaikan tinggi gambar
        $drawing->setOffsetX(5);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($sheet);
    }

    // Set title
    $sheet->mergeCells("F63:O67");
    $sheet->setCellValue('F63', 'CHECK SHEET VERIFIKASI JOB SET UP & PATROL CASTING');
    $sheet->getStyle("F63")->applyFromArray($titleStyle);

    // Set form header
    $sheet->mergeCells("A67:C67");
    $sheet->setCellValue('A67', 'FORM');
    $sheet->getStyle("A67")->applyFromArray([
        'font' => ['bold' => true, 'size' => 14],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]
    ]);

    // Set header data
    $headers = [
        ['P63:S63', 'F0/QAS/Q1/1907'],
        ['P64:Q64', 'Revisi'],
        ['P65:Q65', '0'],
        ['R64:S64', 'Hal'],
        ['R65:S65', '2 / 3'],
        ['P66:S67', 'Berlaku mulai: 04 November 2020']
    ];

    foreach ($headers as [$range, $value]) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(substr($range, 0, strpos($range, ':')), $value);
    }
    $sheet->getStyle("P63:S67")->applyFromArray($textCenter);

    // Set table header
    $tableData = [
        ['A68' => 'Nama part', 'D68' => 'TWIN HEAD', 'J68' => 'Custom', 'K68' => 'PT. DNP', 'P68' => 'Nama mesin'],
        ['A69' => 'No. part', 'D69' => 'P332204-6910B', 'J69' => 'Model', 'K69' => '4D34G', 'P69' => 'No. mesin'],
        ['A70' => 'Code', 'D70' => 'AV – 1', 'J70' => 'No. die', 'P70' => 'No. jig'],
    ];

    foreach ($tableData as $row) {
        foreach ($row as $cell => $value) {
            $sheet->setCellValue($cell, $value);
        }
    }

    // Merge cells untuk keterangan
    $sheet->mergeCells('A71:S71');
    $sheet->setCellValue('A71', 'Keterangan: Untuk Job Set up, diambil sample 1 pcs / jig untuk pemeriksaan awal.');
    $sheet->getStyle("A71")->applyFromArray([
        'font' => ['size' => 10, 'underline' => Font::UNDERLINE_SINGLE],
        'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT]
    ]);

    // Apply styling ke teks tertentu
    $boldCells = ["A5", "D6", "D7", "D8", "K6", "K7", "P1", "P3", "R3", "A35", "A40", "A47", "A49", "A52", "A54", "A55", "D68", "D69", "D70", "K68", "K69", "P63", "P64", "P65", "R64", "R65"];
    foreach ($boldCells as $cell) {
        $sheet->getStyle($cell)->applyFromArray(['font' => ['bold' => true]]);
    }

    // Apply border sesuai kolom
    $sheet->getStyle("A63:S71")->applyFromArray([
        'borders' => [
            'allBorders' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
            ]
        ]
    ]);

    // Atur border untuk tabel utama
    $sheet->getStyle('A68:S70')->applyFromArray([
        'borders' => [
            'outline' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM
            ],
            'inside' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
            ]
        ]
    ]);

    // --------------------------- BAGIAN 2 -----------------------
    // Daftar header tabel dengan merge cells
    $tableHeaders = [
        'A72:B73' => 'No.',
        'C72:D73' => 'Item',
        'E72:G73' => 'Standard',
        'H72:H73' => 'Setup',
        'I72:I73' => 'Patrol',
        'J72:J72' => 'Control',
        'J73:J73' => 'Method',
        'K72:S72' => 'Cavity Sample',
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

    // Memberikan border untuk setiap kolom di baris K73:S73
    for ($col = 'K'; $col <= 'S'; $col++) {
        $cell = $col . '73';
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

    // Daftar judul bagian dan rentang selnya
    $sections = [
        ['A85:G85', 'III. LUBANG & INSERT PIN AREA CORE 2'],
        ['A92:G92', 'IV. LUBANG & INSERT PIN AREA CORE 3. LUBANG & INSERT PIN AREA CORE 1'],
        ['A97:G97', 'V. LUBANG & INSERT PIN AREA FIX'],
        ['A114:G114', 'VI. OTHERS'],
        ['I85:S85', ''],
        ['I92:S92', ''],
        ['I97:S97', ''],
        ['I114:S114', ''],

    ];

    // Styling untuk judul bagian
    $sectionTitleStyle = [
        'font' => ['bold' => true, 'italic' => true],
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER
        ],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'startColor' => ['rgb' => 'D9F1F1'] // Warna latar biru muda
        ],
        'borders' => [
            'outline' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE]
        ]
    ];

    $mergedCells = [
        ['range' => 'A74:B78', 'value' => '10'],
        ['range' => 'A79:B80', 'value' => '11'],
        ['range' => 'A81:B81', 'value' => '12'],
        ['range' => 'A82:B82', 'value' => '13'],
        ['range' => 'A83:B84', 'value' => '14'],
        ['range' => 'A86:B87', 'value' => '15'],
        ['range' => 'A88:B89', 'value' => '16'],
        ['range' => 'A90:B91', 'value' => '17'],
        ['range' => 'A93:B94', 'value' => '18'],
        ['range' => 'A95:B96', 'value' => '19'],
        ['range' => 'A98:B98', 'value' => '20'],
        ['range' => 'A99:B101', 'value' => '21'],
        ['range' => 'A102:B102', 'value' => '22'],
        ['range' => 'A103:B111', 'value' => '23'],
        ['range' => 'A112:B113', 'value' => '24'],
        ['range' => 'A115:B116', 'value' => '25'],
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

    // Iterasi untuk mengatur judul bagian
    foreach ($sections as [$range, $title]) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(explode(':', $range)[0], $title);
        $sheet->getStyle($range)->applyFromArray($sectionTitleStyle);
    }
    $cells = [
        // Cells 10
        ['C74:D76', 'a.Diameter ujung(3plcs)'],
        ['C77:D78', 'b.Appearance hole(3plcs)'],
        ['E74:F76', '11 ±0.3'],
        ['E77:G78', 'Tidak: undercut, overheat, bengkok, scrap tipis (max 0.3)'],
        ['G74', '1'],
        ['G75', '2'],
        ['G76', '3'],

        // Cells 10 bagian kanan 
        ['J74:J76', 'Caliper'],
        ['J77:J78', 'Visual'],
        ['K77:K78', 'OK / NG'],
        ['L77:L78', 'OK / NG'],
        ['M77:M78', 'OK / NG'],
        ['N77:N78', 'OK / NG'],
        ['O77:O78', 'OK / NG'],
        ['P77:P78', 'OK / NG'],
        ['Q77:Q78', 'OK / NG'],
        ['R77:R78', 'OK / NG'],
        ['S77:S78', 'OK / NG'],

        // Cells 11
        ['C79:D79', 'a. Inner diameter'],
        ['C80:D80', 'b. Depth hole'],
        ['E79:G79', 'Ø15.8 [-0.3 ~ 0]'],
        ['E80:G80', '18 ±0.3'],

        // Cells 11 bagian kanan 
        ['I79', '-'],
        ['J79', 'Caliper'],
        ['J80', 'Caliper'],

        // Cells 12
        ['C81:D81', 'Appearance ex overflow'],
        ['E81:G81', 'Tidak : gompal & keropos'],

        // Cells 12 bagian kanan 
        ['J81', 'Visual'],
        ['K81', 'OK / NG'],
        ['L81', 'OK / NG'],
        ['M81', 'OK / NG'],
        ['N81', 'OK / NG'],
        ['O81', 'OK / NG'],
        ['P81', 'OK / NG'],
        ['Q81', 'OK / NG'],
        ['R81', 'OK / NG'],
        ['S81', 'OK / NG'],

        // Cells 13
        ['C82:D82', 'Appearance partingline'],
        ['E82:G82', 'Tidak flash, tebal scrap max 0.3'],

        // Cells 13 bagian kanan 
        ['J82', 'Visual'],
        ['K82', 'OK / NG'],
        ['L82', 'OK / NG'],
        ['M82', 'OK / NG'],
        ['N82', 'OK / NG'],
        ['O82', 'OK / NG'],
        ['P82', 'OK / NG'],
        ['Q82', 'OK / NG'],
        ['R82', 'OK / NG'],
        ['S82', 'OK / NG'],

        // Cells 14
        ['C83:D84', 'Appearance permukaan Core 1'],
        ['E83:G84', 'Tidak : Retak, undercut, flowline, overheat & gompal. '],

        // Cells 14 bagian kanan 
        ['J83:J84', 'Visual'],
        ['K83:K84', 'OK / NG'],
        ['L83:L84', 'OK / NG'],
        ['M83:M84', 'OK / NG'],
        ['N83:N84', 'OK / NG'],
        ['O83:O84', 'OK / NG'],
        ['P83:P84', 'OK / NG'],
        ['Q83:Q84', 'OK / NG'],
        ['R83:R84', 'OK / NG'],
        ['S83:S84', 'OK / NG'],

        // Cells 15
        ['C86:D86', 'a. Inner diameter'],
        ['C87:D87', 'b. Depth hole'],
        ['E86:G86', 'Ø15.8 [-0.3 ~ 0]'],
        ['E87:G87', '18 ±0.3'],

        // Cells 15 bagian kanan 
        ['I86', '-'],
        ['J86', 'Caliper'],
        ['J87', 'Caliper'],

        // Cells 16
        ['C88:D88', 'a. Inner diameter'],
        ['C89:D89', 'b. Depth hole'],
        ['E88:G88', 'Ø15.8 [-0.3 ~ 0]'],
        ['E89:G89', '9.5 ±0.3'],

        // Cells 16 bagian kanan 
        ['I88', '-'],
        ['J88', 'Caliper'],
        ['J89', 'Caliper'],

        // Cells 17
        ['C90:D91', ' Appearance permukaan Core 2'],
        ['E90:G91', 'Tidak : Retak, undercut, flowline, overheat, gompal & step [max 0.3]'],

        //  Cells 17 bagian kanan 
        ['J90:J91', 'Visual'],
        ['K90:K91', 'OK / NG'],
        ['L90:L91', 'OK / NG'],
        ['M90:M91', 'OK / NG'],
        ['N90:N91', 'OK / NG'],
        ['O90:O91', 'OK / NG'],
        ['P90:P91', 'OK / NG'],
        ['Q90:Q91', 'OK / NG'],
        ['R90:R91', 'OK / NG'],
        ['S90:S91', 'OK / NG'],

        // Cells 18
        ['C93:D93', 'a. Inner diameter'],
        ['C94:D94', 'b. Depth hole'],
        ['E93:G93', 'Ø15.8 [-0.3 ~ 0]'],
        ['E94:G94', '9.5 ±0.3'],

        //  Cells 18 bagian kanan
        ['I93', '-'],
        ['J93', 'Caliper'],
        ['J94', 'Caliper'],

        // Cells 19
        ['C95:D96', ' Appearance permukaan Core 3'],
        ['E95:G96', 'Tidak : Retak, undercut, flowline, overheat, gompal & step [max 0.3]'],

        //  Cells 19 bagian kanan 
        ['J95:J96', 'Visual'],
        ['K95:K96', 'OK / NG'],
        ['L95:L96', 'OK / NG'],
        ['M95:M96', 'OK / NG'],
        ['N95:N96', 'OK / NG'],
        ['O95:O96', 'OK / NG'],
        ['P95:P96', 'OK / NG'],
        ['Q95:Q96', 'OK / NG'],
        ['R95:R96', 'OK / NG'],
        ['S95:S96', 'OK / NG'],

        // Cells 20
        ['C98:D98', 'b.Appearance datum D1'],
        ['E98:G98', 'Rata, tidak : gelombang, cacat.'],

        // Cells 20 bagian kanan 
        ['J98', 'Visual'],
        ['K98', 'OK / NG'],
        ['L98', 'OK / NG'],
        ['M98', 'OK / NG'],
        ['N98', 'OK / NG'],
        ['O98', 'OK / NG'],
        ['P98', 'OK / NG'],
        ['Q98', 'OK / NG'],
        ['R98', 'OK / NG'],
        ['S98', 'OK / NG'],

        // Cells 21
        ['C99:D99', 'a. Inner diameter (D2)'],
        ['C100:D100', 'b. Depth hole'],
        ['C101:D101', ' c. Appearance hole'],
        ['E99:G99', 'Ø6 [-0.3 ~ 0]'],
        ['E100:G100', '7 ±0.3'],
        ['E101:G101', 'Tidak : undercut, overheat'],

        //  Cells 21 bagian kanan
        ['I99', '-'],
        ['J99', 'Caliper'],
        ['J100', 'Caliper'],
        ['J101', 'Visual'],
        ['K101', 'OK / NG'],
        ['L101', 'OK / NG'],
        ['M101', 'OK / NG'],
        ['N101', 'OK / NG'],
        ['O101', 'OK / NG'],
        ['P101', 'OK / NG'],
        ['Q101', 'OK / NG'],
        ['R101', 'OK / NG'],
        ['S101', 'OK / NG'],

        // Cells 22
        ['C102:D102', 'b. Appearance surface'],
        ['E102:G102', 'Rata, tidak : undercut, cacat.'],

        // Cells 22 bagian kanan 
        ['J102', 'Visual'],
        ['K102', 'OK / NG'],
        ['L102', 'OK / NG'],
        ['M102', 'OK / NG'],
        ['N102', 'OK / NG'],
        ['O102', 'OK / NG'],
        ['P102', 'OK / NG'],
        ['Q102', 'OK / NG'],
        ['R102', 'OK / NG'],
        ['S102', 'OK / NG'],

        // Cells 23
        ['C103:D106', 'a. Inner diameter(4plcs)'],
        ['C107:D110', 'b. Depth hole (4plcs)'],
        ['C111:D111', ' c. Appearance hole'],
        ['E103:F103', 'Ø4.5 [-0.3 ~ 0]'],
        ['E104:F104', 'Ø4.5 [-0.3 ~ 0]'],
        ['E105:F105', 'Ø4.5 [-0.3 ~ 0]'],
        ['E106:F106', 'Ø4.5 [-0.3 ~ 0]'],
        ['G103', '1'],
        ['G104', '2'],
        ['G105', '3'],
        ['G106', '4'],
        ['E107:F107', '12.5 ±0.3'],
        ['E108:F108', '12.5 ±0.3'],
        ['E109:F109', '12.5 ±0.3'],
        ['E110:F110', '12.5 ±0.3'],
        ['G107', '1'],
        ['G108', '2'],
        ['G109', '3'],
        ['G110', '4'],
        ['E111:G111', 'Tidak : undercut, overheat, patah.'],

        // Cells 23 bagian kanan
        ['I103', '-'],
        ['I104', '-'],
        ['I105', '-'],
        ['I106', '-'],
        ['J103:J106', 'Caliper'],
        ['J107:J110', 'Caliper'],
        ['J111', 'Visual'],
        ['K111', 'OK / NG'],
        ['L111', 'OK / NG'],
        ['M111', 'OK / NG'],
        ['N111', 'OK / NG'],
        ['O111', 'OK / NG'],
        ['P111', 'OK / NG'],
        ['Q111', 'OK / NG'],
        ['R111', 'OK / NG'],
        ['S111', 'OK / NG'],

        // Cells 24
        ['C112:D113', ' Appearance permukaan fix'],
        ['E112:G113', 'Tidak : Retak, undercut, flowline, overheat & gompal.'],

        // Cells 24 bagian kanan
        ['J112:J113', 'Visual'],
        ['K112:K113', 'OK / NG'],
        ['L112:L113', 'OK / NG'],
        ['M112:M113', 'OK / NG'],
        ['N112:N113', 'OK / NG'],
        ['O112:O113', 'OK / NG'],
        ['P112:P113', 'OK / NG'],
        ['Q112:Q113', 'OK / NG'],
        ['R112:R113', 'OK / NG'],
        ['S112:S113', 'OK / NG'],

        // Cells 25
        ['C115:D116', ' Test Cutting'],
        ['E115:G116', 'Mach. OK, porosity sesuai standard'],

        // Cells 25 bagian kanan
        ['J115:J116', 'Machining Test'],
        ['K115:K116', 'OK / NG'],
        ['L115:L116', 'OK / NG'],
        ['M115:M116', 'OK / NG'],
        ['N115:N116', 'OK / NG'],
        ['O115:O116', 'OK / NG'],
        ['P115:P116', 'OK / NG'],
        ['Q115:Q116', 'OK / NG'],
        ['R115:R116', 'OK / NG'],
        ['S115:S116', 'OK / NG'],

    ];

    // Loop tunggal untuk memproses semua sel
    foreach ($cells as [$range, $text]) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(explode(':', $range)[0], $text);

        $sheet->getStyle($range)->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, // Pusatkan teks horizontal
                'vertical' => Alignment::VERTICAL_CENTER // Pusatkan teks vertikal
            ],
            'fill' => [
                'fillType' => Fill::FILL_SOLID,
                'startColor' => ['rgb' => 'FFFFFF'] // Warna putih
            ],
            'font' => ['bold' => false], // Hapus bold
            'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]]
        ]);
    }

    $sheet->mergeCells('H74:H116');
    $sheet->setCellValue('H74', 'All Point If Job Set Up');

    // Styling untuk teks vertikal dan center
    $sheet->getStyle('H74:H116')->applyFromArray([
        'alignment' => [
            'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER, // Pusatkan secara horizontal
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER, // Pusatkan secara vertikal
            'textRotation' => 90 // Putar teks 90 derajat (vertikal)
        ],
        'font' => [
            'bold' => true
        ],
        'borders' => [
            'allBorders' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN]
        ]
    ]);

    $mergeCells = [
        'I74:I78' => '1x / Shift',
        'I80:I81' => '1x / Shift',
        'I82:I84' => '1x / Shift',
        'I50:I51' => '1x / Shift',
        'I87' => '1x / Shift',
        'I89' => '1x / Shift',
        'I90:I91' => '1x / Shift',
        'I94' => '1x / Shift',
        'I95:I96' => '1x / Shift',
        'I98' => '1x / Shift',
        'I100:I102' => '1x / Shift',
        'I107:I111' => '1x / Shift',
        'I112:I113' => '1x / Shift',
        'I115:I116' => '1x / Shift',

    ];

    // Terapkan merge, set nilai, dan styling
    foreach ($mergeCells as $range => $text) {
        $sheet->mergeCells($range);
        $sheet->setCellValue(explode(':', $range)[0], $text);
        $sheet->getStyle($range)->applyFromArray([
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'textRotation' => 90
            ],
            'font' => ['bold' => false],
            'borders' => ['allBorders' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN]]
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