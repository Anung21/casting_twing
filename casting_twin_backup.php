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
$fileName = "TESTING TO EXPORT GENERATOR";


ini_set('memory_limit', '-1');

/* collect to array */
$dataParentArray = [
    'id' => isset($Qgenerate->id) ? $Qgenerate->id : '',
    'revisi' => isset($Qgenerate->revisi) ? $Qgenerate->revisi : '',
    'berlakuForm' => isset($Qgenerate->created_at) ? date('d-m-Y', strtotime($Qgenerate->created_at)) : '',
    'namaPart' => isset($dataPart->nama_part) ? $dataPart->nama_part : '',
    'customerMaterial' => isset($dataPart->customer_material) ? $dataPart->customer_material : '',
    'material' => isset($dataPart->material) ? $dataPart->material : '',
    'customer' => isset($dataPart->customer) ? $dataPart->customer : '',
    'proses' => isset($dataPart->proses) ? $dataPart->proses : '',
    'qty' => '',
    'status' => '',
    'jenisPengecekan' => isset($Qform) && $Qform->isNotEmpty() && isset($Qform->first()->qualityFormCheckType->name)
        ? strtoupper($Qform->first()->qualityFormCheckType->name)
        : '',
];

/* Buat Spreadsheet */
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Sheet1');

/* array for style */
$satuPageIsi = 50;
$jumlahPage = 5;
$boldCells = [];
$centerCells = [];
$alignTextTop = [];
$textLeft = [];
$textRight = [];
$fontSizes = [];
$ColoumsHeaderTable = [];

try {
    $row = 1;

    /* === Set header page 1 === */
    /* kasih bates ujung kanan dan kiri */
    $sheet->getColumnDimension('A')->setWidth(2);
    $sheet->getRowDimension(1)->setRowHeight(10);
    $sheet->getColumnDimension('U')->setWidth(2);

    /* set logo */
    $sheet->mergeCells("B2:E4");
    $drawing = new Drawing();
    $drawing->setName('Company Logo');
    $drawing->setDescription('Company Logo');
    $drawing->setCoordinates('B2');
    $drawing->setHeight(55);
    $drawing->setWorksheet($sheet);
    $sheet->setCellValue('B2', '');

    /* set judul */
    $sheet->mergeCells("B5:E5")->setCellValue('B5', "Form berlaku mulai $dataParentArray[berlakuForm].");
    $sheet->mergeCells("F2:P5")->setCellValue('F2', "QUALITY E-FORM REPORT ( $dataParentArray[jenisPengecekan] )");
    $sheet->mergeCells("Q2:R2")->setCellValue('Q2', 'REGISTRATION NO.');
    $sheet->mergeCells("Q3:R3")->setCellValue('Q3', 'FO/QAS/QL/091');
    $sheet->mergeCells("Q4:R5")->setCellValue('Q4', "REVISION : $dataParentArray[revisi]");
    $sheet->mergeCells("S2:T2")->setCellValue('S2', 'PAGE');
    $sheet->mergeCells("S4:T4")->setCellValue('S4', "1/$jumlahPage");

    $sheet->mergeCells("B6:E6")->setCellValue('B6', 'PART NAME');
    $sheet->mergeCells("B7:E7")->setCellValue('B7', 'PART NO.');
    $sheet->mergeCells("B8:E8")->setCellValue('B8', 'MODEL/DIES/CAV.');
    $sheet->mergeCells("B9:E9")->setCellValue('B9', 'QUANTITY');
    $sheet->mergeCells("B10:E10")->setCellValue('B10', 'STATUS');
    $sheet->mergeCells("B11:E11")->setCellValue('B11', 'CUSTOMER');
    $sheet->mergeCells("B12:E12")->setCellValue('B12', 'PROSES');
    $sheet->mergeCells("N7")->setCellValue('N7', 'DATE');
    $sheet->mergeCells("N8:N10")->setCellValue('N8', 'SIGN');
    $sheet->mergeCells("N11")->setCellValue('N11', 'NAME');
    $sheet->mergeCells("O6:P6")->setCellValue('O6', 'INSPECTED');
    $sheet->mergeCells("Q6:R6")->setCellValue('Q6', 'CHECKED');
    $sheet->mergeCells("S6:T6")->setCellValue('S6', 'APPROVAL');
    $sheet->mergeCells("N12:T12")->setCellValue('N12', 'UNIT : mm');

    /* isi untuk judul or header */
    $sheet->mergeCells("F6:M6")->setCellValue('F6', "$dataParentArray[namaPart]"); /* part name */
    $sheet->mergeCells("F7:M7")->setCellValue('F7', "$dataParentArray[customerMaterial]"); /* part no */
    $sheet->mergeCells("F8:M8")->setCellValue('F8', "$dataParentArray[material]"); /* model/dies/cavity */
    $sheet->mergeCells("F9:M9")->setCellValue('F9', "$dataParentArray[qty]"); /* Quantity */
    $sheet->mergeCells("F10:M10")->setCellValue('F10', "$dataParentArray[status]"); /* status */
    $sheet->mergeCells("F11:M11")->setCellValue('F11', "$dataParentArray[customer]"); /* customer */
    $sheet->mergeCells("F12:M12")->setCellValue('F12', "$dataParentArray[proses]"); /* proses */
    $sheet->mergeCells("O7:P7")->setCellValue('O7', ""); /* date inspected */
    $sheet->mergeCells("Q7:R7")->setCellValue('Q7', ""); /* date checked */
    $sheet->mergeCells("S7:T7")->setCellValue('S7', ""); /* date approval */
    $sheet->mergeCells("O8:P10")->setCellValue('O8', ""); /* sign inspected */
    $sheet->mergeCells("Q8:R10")->setCellValue('Q8', ""); /* sign checked */
    $sheet->mergeCells("S8:T10")->setCellValue('S8', ""); /* sign approval */
    $sheet->mergeCells("O11:P11")->setCellValue('O11', ""); /* name inspected */
    $sheet->mergeCells("Q11:R11")->setCellValue('Q11', ""); /* name checked */
    $sheet->mergeCells("S11:T11")->setCellValue('S11', ""); /* name approval */
    /* === Set header page 1 === */

    /* === set Header of table in content on page 1 === */
    $sheet->mergeCells("B13:B14")->setCellValue('B13', 'NO');
    $sheet->mergeCells("C13:G14")->setCellValue('C13', 'INSPECTION ITEM');
    $sheet->mergeCells("H13:H14")->setCellValue('H13', 'STD');
    $sheet->mergeCells("I13:J13")->setCellValue('I13', 'TOLERANSI');
    $sheet->setCellValue('I14', 'LOWER');
    $sheet->setCellValue('J14', 'UPPER');
    $sheet->mergeCells("K13:L14")->setCellValue('K13', 'INPECTION METHOD');
    $sheet->mergeCells("M13:Q13")->setCellValue('M13', 'SAMPLE');
    $sheet->setCellValue('M14', '1');
    $sheet->setCellValue('N14', '2');
    $sheet->setCellValue('O14', '3');
    $sheet->setCellValue('P14', '4');
    $sheet->setCellValue('Q14', '5');
    $sheet->mergeCells("R13:R14")->setCellValue('R13', 'X');
    $sheet->mergeCells("S13:S14")->setCellValue('S13', 'R');
    $sheet->mergeCells("T13:T14")->setCellValue('T13', 'JUDGE');
    /* === set Header of table in content on page 1 === */

    /* === Set Content on table on page 1 === */
    for ($i = 0; $i < $satuPageIsi; $i++) {
        $starts = 15 + $i;
        $number = $i + 1;
        $sheet->setCellValue('B' . $starts, ""); /* NO */
        $sheet->mergeCells("C$starts:G$starts")->setCellValue('C' . $starts, ''); /* inspection item */
        $sheet->setCellValue('H' . $starts, ''); /* STD */
        $sheet->setCellValue('I' . $starts, ''); /* lower */
        $sheet->setCellValue('J' . $starts, ''); /* upper */
        $sheet->mergeCells("K$starts:L$starts")->setCellValue('K' . $starts, '');
        $sheet->setCellValue('M' . $starts, ''); /* sample 1 */
        $sheet->setCellValue('N' . $starts, ''); /* sample 2 */
        $sheet->setCellValue('O' . $starts, ''); /* sample 3 */
        $sheet->setCellValue('P' . $starts, ''); /* sample 4 */
        $sheet->setCellValue('Q' . $starts, ''); /* sample 5 */
        $sheet->setCellValue('R' . $starts, ''); /* X */
        $sheet->setCellValue('S' . $starts, ''); /* R */
        $sheet->setCellValue('T' . $starts, ''); /* JUDGE */

        /* Center Alignment */
        $columnsToCenter = ['B', 'H', 'I', 'J', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T'];
        foreach ($columnsToCenter as $col) {
            $sheet->getStyle($col . $starts)->getAlignment()
                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
                ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
        }
    }
    /* === Set Content on table on page 1 === */

    /* === Set Footer page 1 === */
    $totalRows = 15 + $satuPageIsi;
    $footerStartRow = $totalRows;
    $sheet->mergeCells("B$footerStartRow:P" . ($footerStartRow + 3))->setCellValue('B' . $footerStartRow, 'Note :');
    $sheet->mergeCells("R$footerStartRow:T$footerStartRow")->setCellValue('R' . $footerStartRow, 'PART APPROVAL');
    $sheet->mergeCells("R" . ($footerStartRow + 1) . ":T" . ($footerStartRow + 1))->setCellValue('R' . ($footerStartRow + 1), 'COUNTERMEASURE APPROVAL');
    $sheet->mergeCells("R" . ($footerStartRow + 2) . ":T" . ($footerStartRow + 2))->setCellValue('R' . ($footerStartRow + 2), 'OTHERS');
    $sheet->mergeCells("Q" . ($footerStartRow + 3) . ":T" . ($footerStartRow + 3))->setCellValue("Q" . ($footerStartRow + 3), 'TOTAL JUDGEMENT');
    $approvalStartRow = $footerStartRow + 4;
    $sheet->mergeCells("B$approvalStartRow:C$approvalStartRow")->setCellValue("B$approvalStartRow", '');
    $sheet->mergeCells("B" . ($approvalStartRow + 1) . ":C" . ($approvalStartRow + 1))->setCellValue("B" . ($approvalStartRow + 1), 'MADE');
    $sheet->mergeCells("B" . ($approvalStartRow + 2) . ":C" . ($approvalStartRow + 2))->setCellValue("B" . ($approvalStartRow + 2), 'CHECKED');
    $sheet->mergeCells("B" . ($approvalStartRow + 3) . ":C" . ($approvalStartRow + 3))->setCellValue("B" . ($approvalStartRow + 3), 'APPROVED');
    $sheet->mergeCells("D$approvalStartRow:E$approvalStartRow")->setCellValue("D$approvalStartRow", 'DATE');
    $sheet->mergeCells("D" . ($approvalStartRow + 1) . ":E" . ($approvalStartRow + 1))->setCellValue("D" . ($approvalStartRow + 1), ''); /* date made */
    $sheet->mergeCells("D" . ($approvalStartRow + 2) . ":E" . ($approvalStartRow + 2))->setCellValue("D" . ($approvalStartRow + 2), ''); /* date checked */
    $sheet->mergeCells("D" . ($approvalStartRow + 3) . ":E" . ($approvalStartRow + 3))->setCellValue("D" . ($approvalStartRow + 3), ''); /* date approved */
    $sheet->mergeCells("F$approvalStartRow:H$approvalStartRow")->setCellValue("F$approvalStartRow", 'SIGN/NAME');
    $sheet->mergeCells("F" . ($approvalStartRow + 1) . ":H" . ($approvalStartRow + 1))->setCellValue("F" . ($approvalStartRow + 1), ''); /* sign/name made */
    $sheet->mergeCells("F" . ($approvalStartRow + 2) . ":H" . ($approvalStartRow + 2))->setCellValue("F" . ($approvalStartRow + 2), ''); /* sign/name checked */
    $sheet->mergeCells("F" . ($approvalStartRow + 3) . ":H" . ($approvalStartRow + 3))->setCellValue("F" . ($approvalStartRow + 3), ''); /* sign/name approved */
    $sheet->mergeCells("I$approvalStartRow:K$approvalStartRow")->setCellValue("I$approvalStartRow", 'REVISION');
    $sheet->mergeCells("I" . ($approvalStartRow + 1) . ":K" . ($approvalStartRow + 1))->setCellValue("I" . ($approvalStartRow + 1), ''); /* Revision made */
    $sheet->mergeCells("I" . ($approvalStartRow + 2) . ":K" . ($approvalStartRow + 2))->setCellValue("I" . ($approvalStartRow + 2), ''); /* Revision checked */
    $sheet->mergeCells("I" . ($approvalStartRow + 3) . ":K" . ($approvalStartRow + 3))->setCellValue("I" . ($approvalStartRow + 3), ''); /* Revision approved */
    $sheet->mergeCells("L$approvalStartRow:M$approvalStartRow")->setCellValue("L$approvalStartRow", 'DATE');
    $sheet->mergeCells("L" . ($approvalStartRow + 1) . ":M" . ($approvalStartRow + 1))->setCellValue("L" . ($approvalStartRow + 1), ''); /* date made */
    $sheet->mergeCells("L" . ($approvalStartRow + 2) . ":M" . ($approvalStartRow + 2))->setCellValue("L" . ($approvalStartRow + 2), ''); /* date checked */
    $sheet->mergeCells("L" . ($approvalStartRow + 3) . ":M" . ($approvalStartRow + 3))->setCellValue("L" . ($approvalStartRow + 3), ''); /* date approved */
    $sheet->mergeCells("N$approvalStartRow:P$approvalStartRow")->setCellValue("N$approvalStartRow", 'SIGN/NAME');
    $sheet->mergeCells("N" . ($approvalStartRow + 1) . ":P" . ($approvalStartRow + 1))->setCellValue("N" . ($approvalStartRow + 1), ''); /* sign/name made */
    $sheet->mergeCells("N" . ($approvalStartRow + 2) . ":P" . ($approvalStartRow + 2))->setCellValue("N" . ($approvalStartRow + 2), ''); /* sign/name checked */
    $sheet->mergeCells("N" . ($approvalStartRow + 3) . ":P" . ($approvalStartRow + 3))->setCellValue("N" . ($approvalStartRow + 3), ''); /* sign/name approved */
    $sheet->mergeCells("Q$approvalStartRow:R" . ($approvalStartRow + 3))->setCellValue("Q$approvalStartRow", 'OK');
    $sheet->mergeCells("S$approvalStartRow:T" . ($approvalStartRow + 3))->setCellValue("S$approvalStartRow", 'NG');
    /* === Set Footer page 1 === */

    /* push array to style */
    array_push($boldCells, "B13:T14", "F2", "Q2", "S2", "Q3", "S3", "Q4", "O6", "Q6", "S6", "N7", "N8", "N11", "N12", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12");
    array_push($boldCells, "B" . ($footerStartRow + 5), "B" . ($footerStartRow + 6), "B" . ($footerStartRow + 7), "D" . ($footerStartRow + 4), "F" . ($footerStartRow + 4), "I" . ($footerStartRow + 4), "L" . ($footerStartRow + 4), "N" . ($footerStartRow + 4), "R" . ($footerStartRow + 0), "R" . ($footerStartRow + 1), "R" . ($footerStartRow + 2), "Q" . ($footerStartRow + 3), "Q" . ($footerStartRow + 4), "S" . ($footerStartRow + 4));
    array_push($centerCells, "F2", "B13:T14", "Q2", "Q3", "Q4", "S3", "S2", "Q3", "Q4", "O6", "Q6", "S6", "N7", "N8", "N11");
    array_push($centerCells, "D" . ($footerStartRow + 4), "F" . ($footerStartRow + 4), "I" . ($footerStartRow + 4), "L" . ($footerStartRow + 4), "N" . ($footerStartRow + 4), "Q" . ($footerStartRow + 3), "Q" . ($footerStartRow + 4), "S" . ($footerStartRow + 4));
    array_push($alignTextTop, "B$footerStartRow");
    array_push($textLeft, "B$footerStartRow");
    array_push($textRight, "N12");
    $fontSizes = array_merge($fontSizes, ["F2" => 22, "S3" => 28, "Q" . ($footerStartRow + 4) => 48, "S" . ($footerStartRow + 4) => 48]);
    array_push($ColoumsHeaderTable, "B13:T14");


    /* ======================================================= */

    $rowPageTerakhir = ($approvalStartRow + 3); // Baris terakhir dari page 1 (72)
    $latestRow = $rowPageTerakhir;

    for ($page = 2; $page <= $jumlahPage; $page++) {

        /* === set new header to page 2 === */
        $pembatasAtas = $latestRow + 1;
        $sheet->getRowDimension($pembatasAtas)->setRowHeight(10);
        $sheet->mergeCells("B" . ($pembatasAtas) . ":T" . ($pembatasAtas));

        $startsHeaderPage = $pembatasAtas + 1;

        /* set logo */
        $sheet->mergeCells("B" . $startsHeaderPage . ":E" . ($startsHeaderPage + 2));
        $drawing = new Drawing();
        $drawing->setName('Company Logo');
        $drawing->setDescription('Company Logo');
        $drawing->setCoordinates('B' . $startsHeaderPage);
        $drawing->setHeight(55);
        $drawing->setWorksheet($sheet);
        $sheet->setCellValue('B' . $startsHeaderPage, '');

        /* set judul */
        $sheet->mergeCells("B" . ($startsHeaderPage + 3) . ":E" . ($startsHeaderPage + 3))->setCellValue('B' . ($startsHeaderPage + 3), "Form berlaku mulai $dataParentArray[berlakuForm].");
        $sheet->mergeCells("F" . ($startsHeaderPage + 0) . ":P" . ($startsHeaderPage + 3))->setCellValue('F' . ($startsHeaderPage + 0), "QUALITY E-FORM REPORT ( $dataParentArray[jenisPengecekan] )");
        $sheet->mergeCells("Q" . ($startsHeaderPage + 0) . ":R" . ($startsHeaderPage + 0))->setCellValue('Q' . ($startsHeaderPage + 0), "REGISTRATION NO.");
        $sheet->mergeCells("Q" . ($startsHeaderPage + 1) . ":R" . ($startsHeaderPage + 1))->setCellValue('Q' . ($startsHeaderPage + 1), "FO/QAS/QL/091");
        $sheet->mergeCells("Q" . ($startsHeaderPage + 2) . ":R" . ($startsHeaderPage + 3))->setCellValue('Q' . ($startsHeaderPage + 2), "REVISION : $dataParentArray[revisi]");
        $sheet->mergeCells("S" . ($startsHeaderPage + 0) . ":T" . ($startsHeaderPage + 0))->setCellValue('S' . ($startsHeaderPage + 0), "PAGE");
        $sheet->mergeCells("S" . ($startsHeaderPage + 1) . ":T" . ($startsHeaderPage + 3))->setCellValue('S' . ($startsHeaderPage + 1), "$page/$jumlahPage");
        $sheet->mergeCells("B" . ($startsHeaderPage + 4) . ":E" . ($startsHeaderPage + 4))->setCellValue('B' . ($startsHeaderPage + 4), 'PART NAME');
        $sheet->mergeCells("B" . ($startsHeaderPage + 5) . ":E" . ($startsHeaderPage + 5))->setCellValue('B' . ($startsHeaderPage + 5), 'PART NO.');
        $sheet->mergeCells("B" . ($startsHeaderPage + 6) . ":E" . ($startsHeaderPage + 6))->setCellValue('B' . ($startsHeaderPage + 6), 'MODEL/DIES/CAV.');
        $sheet->mergeCells("B" . ($startsHeaderPage + 7) . ":E" . ($startsHeaderPage + 7))->setCellValue('B' . ($startsHeaderPage + 7), 'QUANTITY');
        $sheet->mergeCells("B" . ($startsHeaderPage + 8) . ":E" . ($startsHeaderPage + 8))->setCellValue('B' . ($startsHeaderPage + 8), 'STATUS');
        $sheet->mergeCells("B" . ($startsHeaderPage + 9) . ":E" . ($startsHeaderPage + 9))->setCellValue('B' . ($startsHeaderPage + 9), 'CUSTOMER');
        $sheet->mergeCells("B" . ($startsHeaderPage + 10) . ":E" . ($startsHeaderPage + 10))->setCellValue('B' . ($startsHeaderPage + 10), 'PROSES');
        $sheet->mergeCells("N" . ($startsHeaderPage + 10) . ":T" . ($startsHeaderPage + 10))->setCellValue('N' . ($startsHeaderPage + 10), 'UNIT : mm');

        /* isi untuk judul or header */
        $sheet->mergeCells("F" . ($startsHeaderPage + 4) . ":T" . ($startsHeaderPage + 4))->setCellValue('F' . ($startsHeaderPage + 4), "$dataParentArray[namaPart]");
        $sheet->mergeCells("F" . ($startsHeaderPage + 5) . ":T" . ($startsHeaderPage + 5))->setCellValue('F' . ($startsHeaderPage + 5), "$dataParentArray[customerMaterial]");
        $sheet->mergeCells("F" . ($startsHeaderPage + 6) . ":T" . ($startsHeaderPage + 6))->setCellValue('F' . ($startsHeaderPage + 6), "$dataParentArray[material]");
        $sheet->mergeCells("F" . ($startsHeaderPage + 7) . ":T" . ($startsHeaderPage + 7))->setCellValue('F' . ($startsHeaderPage + 7), "$dataParentArray[qty]");
        $sheet->mergeCells("F" . ($startsHeaderPage + 8) . ":T" . ($startsHeaderPage + 8))->setCellValue('F' . ($startsHeaderPage + 8), "$dataParentArray[status]");
        $sheet->mergeCells("F" . ($startsHeaderPage + 9) . ":T" . ($startsHeaderPage + 9))->setCellValue('F' . ($startsHeaderPage + 9), "$dataParentArray[customer]");
        $sheet->mergeCells("F" . ($startsHeaderPage + 10) . ":M" . ($startsHeaderPage + 10))->setCellValue('F' . ($startsHeaderPage + 10), "$dataParentArray[proses]");
        /* === set new header to page 2 === */

        /* === set Header of table on page 2 === */
        $startHeaderTabel = $startsHeaderPage + 11;
        $sheet->mergeCells("B" . ($startHeaderTabel + 0) . ":B" . ($startHeaderTabel + 1))->setCellValue('B' . ($startHeaderTabel + 0), 'NO');
        $sheet->mergeCells("C" . ($startHeaderTabel + 0) . ":G" . ($startHeaderTabel + 1))->setCellValue('C' . ($startHeaderTabel + 0), 'INSPECTION ITEM');
        $sheet->mergeCells("H" . ($startHeaderTabel + 0) . ":H" . ($startHeaderTabel + 1))->setCellValue('H' . ($startHeaderTabel + 0), 'STD');
        $sheet->mergeCells("I" . ($startHeaderTabel + 0) . ":J" . ($startHeaderTabel + 0))->setCellValue('I' . ($startHeaderTabel + 0), 'TOLERANSI');
        $sheet->setCellValue('I' . ($startHeaderTabel + 1), 'LOWER');
        $sheet->setCellValue('J' . ($startHeaderTabel + 1), 'UPPER');
        $sheet->mergeCells("K" . ($startHeaderTabel + 0) . ":L" . ($startHeaderTabel + 1))->setCellValue('K' . ($startHeaderTabel + 0), 'INPECTION METHOD');
        $sheet->mergeCells("M" . ($startHeaderTabel + 0) . ":Q" . ($startHeaderTabel + 0))->setCellValue('M' . ($startHeaderTabel + 0), 'SAMPLE');
        $sheet->setCellValue('M' . ($startHeaderTabel + 1), '1');
        $sheet->setCellValue('N' . ($startHeaderTabel + 1), '2');
        $sheet->setCellValue('O' . ($startHeaderTabel + 1), '3');
        $sheet->setCellValue('P' . ($startHeaderTabel + 1), '4');
        $sheet->setCellValue('Q' . ($startHeaderTabel + 1), '5');
        $sheet->mergeCells("R" . ($startHeaderTabel + 0) . ":R" . ($startHeaderTabel + 1))->setCellValue('R' . ($startHeaderTabel + 0), 'X');
        $sheet->mergeCells("S" . ($startHeaderTabel + 0) . ":S" . ($startHeaderTabel + 1))->setCellValue('S' . ($startHeaderTabel + 0), 'R');
        $sheet->mergeCells("T" . ($startHeaderTabel + 0) . ":T" . ($startHeaderTabel + 1))->setCellValue('T' . ($startHeaderTabel + 0), 'JUDGE');
        /* === set Header of table on page 2 === */

        /* === Set Content on table on page 2 === */
        $startContentPage = $startHeaderTabel + 2; // Data tabel dimulai
        for ($i = 0; $i < $satuPageIsi; $i++) {
            $starts = $startContentPage + $i;
            $number = $i + 1;
            $sheet->setCellValue('B' . $starts, ""); /* NO */
            $sheet->mergeCells("C$starts:G$starts")->setCellValue('C' . $starts, ""); /* inspection item */
            $sheet->setCellValue('H' . $starts, ""); /* STD */
            $sheet->setCellValue('I' . $starts, ""); /* lower */
            $sheet->setCellValue('J' . $starts, ""); /* upper */
            $sheet->mergeCells("K$starts:L$starts")->setCellValue('K' . $starts, "");
            $sheet->setCellValue('M' . $starts, ""); /* sample 1 */
            $sheet->setCellValue('N' . $starts, ""); /* sample 2 */
            $sheet->setCellValue('O' . $starts, ""); /* sample 3 */
            $sheet->setCellValue('P' . $starts, ""); /* sample 4 */
            $sheet->setCellValue('Q' . $starts, ""); /* sample 5 */
            $sheet->setCellValue('R' . $starts, ""); /* X */
            $sheet->setCellValue('S' . $starts, ""); /* R */
            $sheet->setCellValue('T' . $starts, ""); /* JUDGE */

            /* Center Alignment */
            $columnsToCenter = ['B', 'H', 'I', 'J', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T'];
            foreach ($columnsToCenter as $col) {
                $sheet->getStyle($col . $starts)->getAlignment()
                    ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER)
                    ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
            }
        }
        /* === Set Content on table on page 2 === */

        /* === Set Footer page 2 === */
        $totalIsiTabel = $startContentPage + $satuPageIsi - 1;
        $footerStartRowPage = $totalIsiTabel + 1;
        $sheet->mergeCells("B$footerStartRowPage:P" . ($footerStartRowPage + 3))->setCellValue('B' . $footerStartRowPage, 'Note :');
        $sheet->mergeCells("R$footerStartRowPage:T$footerStartRowPage")->setCellValue('R' . $footerStartRowPage, 'PART APPROVAL');
        $sheet->mergeCells("R" . ($footerStartRowPage + 1) . ":T" . ($footerStartRowPage + 1))->setCellValue('R' . ($footerStartRowPage + 1), 'COUNTERMEASURE APPROVAL');
        $sheet->mergeCells("R" . ($footerStartRowPage + 2) . ":T" . ($footerStartRowPage + 2))->setCellValue('R' . ($footerStartRowPage + 2), 'OTHERS');
        $sheet->mergeCells("Q" . ($footerStartRowPage + 3) . ":T" . ($footerStartRowPage + 3))->setCellValue("Q" . ($footerStartRowPage + 3), 'TOTAL JUDGEMENT');

        $approvalStartRowPage = $footerStartRowPage + 4;

        $sheet->mergeCells("B$approvalStartRowPage:C$approvalStartRowPage")->setCellValue("B$approvalStartRowPage", '');
        $sheet->mergeCells("B" . ($approvalStartRowPage + 1) . ":C" . ($approvalStartRowPage + 1))->setCellValue("B" . ($approvalStartRowPage + 1), 'MADE');
        $sheet->mergeCells("B" . ($approvalStartRowPage + 2) . ":C" . ($approvalStartRowPage + 2))->setCellValue("B" . ($approvalStartRowPage + 2), 'CHECKED');
        $sheet->mergeCells("B" . ($approvalStartRowPage + 3) . ":C" . ($approvalStartRowPage + 3))->setCellValue("B" . ($approvalStartRowPage + 3), 'APPROVED');
        $sheet->mergeCells("D$approvalStartRowPage:E$approvalStartRowPage")->setCellValue("D$approvalStartRowPage", 'DATE');
        $sheet->mergeCells("D" . ($approvalStartRowPage + 1) . ":E" . ($approvalStartRowPage + 1))->setCellValue("D" . ($approvalStartRowPage + 1), ''); /* date made */
        $sheet->mergeCells("D" . ($approvalStartRowPage + 2) . ":E" . ($approvalStartRowPage + 2))->setCellValue("D" . ($approvalStartRowPage + 2), ''); /* date checked */
        $sheet->mergeCells("D" . ($approvalStartRowPage + 3) . ":E" . ($approvalStartRowPage + 3))->setCellValue("D" . ($approvalStartRowPage + 3), ''); /* date approved */
        $sheet->mergeCells("F$approvalStartRowPage:J$approvalStartRowPage")->setCellValue("F$approvalStartRowPage", 'SIGN/NAME');
        $sheet->mergeCells("F" . ($approvalStartRowPage + 1) . ":J" . ($approvalStartRowPage + 1))->setCellValue("F" . ($approvalStartRowPage + 1), ''); /* sign/name made */
        $sheet->mergeCells("F" . ($approvalStartRowPage + 2) . ":J" . ($approvalStartRowPage + 2))->setCellValue("F" . ($approvalStartRowPage + 2), ''); /* sign/name checked */
        $sheet->mergeCells("F" . ($approvalStartRowPage + 3) . ":J" . ($approvalStartRowPage + 3))->setCellValue("F" . ($approvalStartRowPage + 3), ''); /* sign/name approved */
        $sheet->mergeCells("K$approvalStartRowPage:M$approvalStartRowPage")->setCellValue("K$approvalStartRowPage", 'REVISION');
        $sheet->mergeCells("K" . ($approvalStartRowPage + 1) . ":M" . ($approvalStartRowPage + 1))->setCellValue("K" . ($approvalStartRowPage + 1), ''); /* Revision made */
        $sheet->mergeCells("K" . ($approvalStartRowPage + 2) . ":M" . ($approvalStartRowPage + 2))->setCellValue("K" . ($approvalStartRowPage + 2), ''); /* Revision checked */
        $sheet->mergeCells("K" . ($approvalStartRowPage + 3) . ":M" . ($approvalStartRowPage + 3))->setCellValue("K" . ($approvalStartRowPage + 3), ''); /* Revision approved */
        $sheet->mergeCells("N$approvalStartRowPage:O$approvalStartRowPage")->setCellValue("N$approvalStartRowPage", 'DATE');
        $sheet->mergeCells("N" . ($approvalStartRowPage + 1) . ":O" . ($approvalStartRowPage + 1))->setCellValue("N" . ($approvalStartRowPage + 1), ''); /* date made */
        $sheet->mergeCells("N" . ($approvalStartRowPage + 2) . ":O" . ($approvalStartRowPage + 2))->setCellValue("N" . ($approvalStartRowPage + 2), ''); /* date checked */
        $sheet->mergeCells("N" . ($approvalStartRowPage + 3) . ":O" . ($approvalStartRowPage + 3))->setCellValue("N" . ($approvalStartRowPage + 3), ''); /* date approved */
        $sheet->mergeCells("P$approvalStartRowPage:T$approvalStartRowPage")->setCellValue("P$approvalStartRowPage", 'SIGN/NAME');
        $sheet->mergeCells("P" . ($approvalStartRowPage + 1) . ":T" . ($approvalStartRowPage + 1))->setCellValue("P" . ($approvalStartRowPage + 1), ''); /* sign/name made */
        $sheet->mergeCells("P" . ($approvalStartRowPage + 2) . ":T" . ($approvalStartRowPage + 2))->setCellValue("P" . ($approvalStartRowPage + 2), ''); /* sign/name checked */
        $sheet->mergeCells("P" . ($approvalStartRowPage + 3) . ":T" . ($approvalStartRowPage + 3))->setCellValue("P" . ($approvalStartRowPage + 3), ''); /* sign/name approved */
        $latestRow = $approvalStartRowPage + 3;
        /* === Set Footer page 2 === */

        array_push($boldCells, "B" . ($startsHeaderPage + 11) . ":T" . ($startsHeaderPage + 12), "F" . ($startsHeaderPage + 0), "Q" . ($startsHeaderPage + 0), "S" . ($startsHeaderPage + 0), "Q" . ($startsHeaderPage + 1), "S" . ($startsHeaderPage + 1), "Q" . ($startsHeaderPage + 2), "O" . ($startsHeaderPage + 4), "Q" . ($startsHeaderPage + 4), "S" . ($startsHeaderPage + 4), "N" . ($startsHeaderPage + 5), "N" . ($startsHeaderPage + 6), "N" . ($startsHeaderPage + 9), "N" . ($startsHeaderPage + 10), "B" . ($startsHeaderPage + 3), "B" . ($startsHeaderPage + 4), "B" . ($startsHeaderPage + 5), "B" . ($startsHeaderPage + 6), "B" . ($startsHeaderPage + 7), "B" . ($startsHeaderPage + 8), "B" . ($startsHeaderPage + 9), "B" . ($startsHeaderPage + 10));
        array_push($boldCells, "B" . ($footerStartRowPage + 5), "B" . ($footerStartRowPage + 6), "B" . ($footerStartRowPage + 7), "D" . ($footerStartRowPage + 4), "F" . ($footerStartRowPage + 4), "K" . ($footerStartRowPage + 4), "N" . ($footerStartRowPage + 4), "P" . ($footerStartRowPage + 4), "R" . ($footerStartRowPage + 0), "R" . ($footerStartRowPage + 1), "R" . ($footerStartRowPage + 2), "Q" . ($footerStartRowPage + 3), "Q" . ($footerStartRowPage + 4), "S" . ($footerStartRowPage + 4));
        array_push($centerCells, "F" . ($startsHeaderPage + 0), "B" . ($startsHeaderPage + 11) . ":T" . ($startsHeaderPage + 12), "Q" . ($startsHeaderPage + 0), "Q" . ($startsHeaderPage + 1), "Q" . ($startsHeaderPage + 2), "S" . ($startsHeaderPage + 1), "S" . ($startsHeaderPage + 0), "Q" . ($startsHeaderPage + 1), "Q" . ($startsHeaderPage + 2), "O" . ($startsHeaderPage + 4), "Q" . ($startsHeaderPage + 4), "S" . ($startsHeaderPage + 4), "N" . ($startsHeaderPage + 5), "N" . ($startsHeaderPage + 6), "N" . ($startsHeaderPage + 9));
        array_push($centerCells, "D" . ($footerStartRowPage + 4), "F" . ($footerStartRowPage + 4), "K" . ($footerStartRowPage + 4), "N" . ($footerStartRowPage + 4), "P" . ($footerStartRowPage + 4), "Q" . ($footerStartRowPage + 3), "Q" . ($footerStartRowPage + 4), "S" . ($footerStartRowPage + 4));
        array_push($alignTextTop, "B$footerStartRowPage");
        array_push($textLeft, "B$footerStartRowPage");
        array_push($textRight, "N" . ($startsHeaderPage + 10));
        $fontSizes = array_merge($fontSizes, ["F" . ($startsHeaderPage + 0) => 22, "S" . ($startsHeaderPage + 1) => 28]);
        array_push($ColoumsHeaderTable, "B" . ($startHeaderTabel + 0) . ":T" . ($startHeaderTabel + 1));


        // echo "page ke $page / $jumlahPage <br>";
        // echo "Page $page: Pembatas atas di $pembatasAtas | Header di $startsHeaderPage | Table Header di $startHeaderTabel | Data mulai di $startContentPage | Data berhenti di $totalIsiTabel | Footer mulai di $footerStartRowPage | Approval di $approvalStartRowPage | Akhir Page di $latestRow <br>";
    }

    /* ========== styling excel ========== */
    /* memberikan border pada semua kotak yang dipilih */
    $sheet->getStyle("B2:T" . ($latestRow))->applyFromArray([
        'borders' => [
            'allBorders' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            ],
        ],
    ]);

    /* Tulisan menjadi tebal AS fw-bold */
    foreach ($boldCells as $cell) {
        $sheet->getStyle("$cell")->applyFromArray([
            'font' => ['bold' => true],
        ]);
    }

    /* text menjadi center AS text-center */
    foreach ($centerCells as $cell) {
        $sheet->getStyle($cell)->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ]);
    }

    /* text menjadi atas AS align-items-top  */
    foreach ($alignTextTop as $cell) {
        $sheet->getStyle($cell)->applyFromArray([
            'alignment' => [
                'vertical' => Alignment::VERTICAL_TOP,
            ],
        ]);
    }

    /* text menjadi rata kiri AS text-start  */
    foreach ($textLeft as $cell) {
        $sheet->getStyle($cell)->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
            ],
        ]);
    }

    /* text menjadi rata Kanan AS text-end  */
    foreach ($textRight as $cell) {
        $sheet->getStyle($cell)->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_RIGHT,
            ],
        ]);
    }

    /* style untuk fubah font size */
    foreach ($fontSizes as $cell => $size) {
        $sheet->getStyle($cell)->applyFromArray(['font' => ['size' => $size]]);
    }

    /* style untuk wrap text and color in table header */
    foreach ($ColoumsHeaderTable as $cell) {
        $sheet->getStyle($cell)->applyFromArray(['alignment' => ['wrapText' => true]]);
        $sheet->getStyle($cell)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()->setRGB('CCFFFF');
    }
} catch (\Exception $e) {
    dd($e);
}

/* Save Template */
$writer = new Xlsx($spreadsheet);
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $fileName . '.xlsx' . '"');
$writer->save('php://output');