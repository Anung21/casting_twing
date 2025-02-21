<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
function dd(...$x)
{
    array_map(function ($x) {
        var_dump($x);
    }, func_get_args());
    die;
}


// Set Nama File
$fileName = "CastingEFORM";

ini_set('memory_limit', '-1');

/* collect to array */
$dataParentArray = [
    'idsurat' => 'n/a',
    'revisi' => '0',
    'hal' => '1/3',
    'berlakuForm' => '04 November 2020',
    'namaPart' => ' TWIN HEAD',
    'noPart' => '  P332204–710B',
    'code' => ' AV-1',
    'customer' => ' PT. DNP',
    'model' => ' 4D34G',
    'noDie' => '',
    'namaMesin' => '',
    'noMesin' => '',
    'noJig' => '',
    'ket' => 'Untuk Job Set up, diambil sample 1 pcs / jig untuk pemeriksaan awal',

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
$textitalic = [];

try {
    $row = 1;

    /* === Set header page 1 === */
    /* kasih bates ujung kanan dan kiri */
    $sheet->getColumnDimension('A')->setWidth(2);
    $sheet->getRowDimension(1)->setRowHeight(10);
    $sheet->getColumnDimension('U')->setWidth(2);

    /* set logo */
    $sheet->mergeCells("B2:F5");
    $drawing = new Drawing();
    $drawing->setName('Company Logo');
    $drawing->setDescription('Company Logo');
    $drawing->setCoordinates('B2');
    $drawing->setHeight(55);
    $drawing->setWorksheet($sheet);
    $sheet->setCellValue('B2', '');

    /* set judul */
    $sheet->mergeCells("Q5:T6")->setCellValue('Q5', "Berlaku mulai :  $dataParentArray[berlakuForm].");
    $sheet->mergeCells("G2:P6")->setCellValue('G2', "CHECK SHEET VERIFIKASI  JOB SET UP & PATROL CASTING ");
    $sheet->mergeCells("B6:F6")->setCellValue('B6', 'FORM');
    $sheet->mergeCells("S3:T3")->setCellValue('S3', "Hal");
    $sheet->mergeCells("S4:T4")->setCellValue('S4', "$dataParentArray[hal]");
    $sheet->mergeCells("Q3:R3")->setCellValue('Q3', "Revisi");
    $sheet->mergeCells("Q4:R4")->setCellValue('Q4', "$dataParentArray[revisi]");
    $sheet->mergeCells("Q2:T2")->setCellValue('Q2', 'FO/QAS/Q1/1907');
    $sheet->mergeCells("B7:D7")->setCellValue('B7', 'Nama part');
    $sheet->mergeCells("B8:D8")->setCellValue('B8', 'No. part');
    $sheet->mergeCells("B9:D9")->setCellValue('B9', 'Code');

    $sheet->mergeCells("K7:L7")->setCellValue('K7', 'Customer');
    $sheet->mergeCells("K8:L8")->setCellValue('K8', 'Model');
    $sheet->mergeCells("K9:L9")->setCellValue('K9', 'No.die');
    $sheet->mergeCells("Q7:R7")->setCellValue('Q7', 'Nama mesin');
    $sheet->mergeCells("Q8:R8")->setCellValue('Q8', 'No.mesin');
    $sheet->mergeCells("Q9:R9")->setCellValue('Q9', 'No. jig');

    /* isi untuk judul or header */
    $sheet->mergeCells("E7:J7")->setCellValue('E7', "$dataParentArray[namaPart]");
    $sheet->mergeCells("E8:J8")->setCellValue('E8', "$dataParentArray[noPart]");
    $sheet->mergeCells("E9:J9")->setCellValue('E9', "$dataParentArray[code]");
    $sheet->mergeCells("M7:P7")->setCellValue('M7', "$dataParentArray[customer]");
    $sheet->mergeCells("M8:P8")->setCellValue('M8', "$dataParentArray[model]");
    $sheet->mergeCells("M9:P9")->setCellValue('M9', "$dataParentArray[noDie]");
    $sheet->mergeCells("S7:T7")->setCellValue('S7', "$dataParentArray[namaMesin]");
    $sheet->mergeCells("S8:T8")->setCellValue('S8', "$dataParentArray[noMesin]");
    $sheet->mergeCells("S9:T9")->setCellValue('S9', "$dataParentArray[noJig]");


    /* === Set header page 1 === */

    /* === set Header of table in content on page 1 === */
    $sheet->mergeCells("B13:B14")->setCellValue('B13', 'NO');
    $sheet->mergeCells("C13:E14")->setCellValue('C13', 'ITEM');
    $sheet->mergeCells("F13:H14")->setCellValue('F13', 'STD');
    $sheet->mergeCells("I13:J13")->setCellValue('I13', 'TOLERANSI');
    $sheet->setCellValue('I14', 'LOWER');
    $sheet->setCellValue('J14', 'UPPER');
    $sheet->mergeCells("K13:K14")->setCellValue('K13', 'CONTROL/METHOD');
    $sheet->mergeCells("L13:T13")->setCellValue('L13', 'SAMPLE');
    $sheet->setCellValue('L14', '1');
    $sheet->setCellValue('M14', '2');
    $sheet->setCellValue('N14', '3');
    $sheet->setCellValue('O14', '4');
    $sheet->setCellValue('P14', '5');
    $sheet->setCellValue('Q14', '6');
    $sheet->setCellValue('R14', '7');
    $sheet->setCellValue('S14', '8');
    $sheet->setCellValue('T14', '9');

    // $sheet->mergeCells("R15:R16")->setCellValue('R15', 'X');
    // $sheet->mergeCells("S15:S16")->setCellValue('S15', 'R');
    // $sheet->mergeCells("T15:T16")->setCellValue('T15', 'JUDGE');
    // $sheet->mergeCells("B13:B14")->setCellValue('B13', 'No');
    // $sheet->mergeCells("C13:E14")->setCellValue('C13', 'Item');
    // $sheet->mergeCells("F13:J14")->setCellValue('F13', 'Standart');
    // $sheet->mergeCells("K13:K14")->setCellValue('K13', 'Control/Method');
    // $sheet->mergeCells("L13:T13")->setCellValue('L13', 'Cavity Sample');
    // $sheet->mergeCells("L13:T13")->setCellValue('L13', 'Cavity Sample');
    // $sheet->mergeCells("B15:T16")->setCellValue('B15', 'I. LUBANG & INSERT PIN AREA MOVE');

    // $sheet->setCellValue('L14', '1');
    // $sheet->setCellValue('M14', '2');
    // $sheet->setCellValue('N14', '3');
    // $sheet->setCellValue('O14', '4');
    // $sheet->setCellValue('P14', '5');
    // $sheet->setCellValue('Q14', '6');
    // $sheet->setCellValue('R14', '7');
    // $sheet->setCellValue('S14', '8');
    // $sheet->setCellValue('T14', '9');


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


    /* === set Header of table in content on page 1 === */

    /* === Set Content on table on page 1 === */
    for ($i = 0; $i < $satuPageIsi; $i++) {
        $starts = 15 + $i;
        $number = $i + 1;

        $sheet->setCellValue('B' . $starts, $number); /* NO */
        $sheet->mergeCells("C$starts:E$starts")->setCellValue('C' . $starts, ''); /*item */
        $sheet->mergeCells("F$starts:H$starts")->setCellValue('F' . $starts, ''); /*item */
        $sheet->setCellValue('I' . $starts, ''); /* lower */
        $sheet->setCellValue('J' . $starts, ''); /* upper */
        $sheet->setCellValue('K' . $starts, '');
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

    /* push array to style */
    array_push($boldCells, "B13:T14", "F2", "Q2", "S2", "Q3", "S3", "S4", "Q4", "O6", "Q6", "S6", "N7", "N8", "N11", "N12", "B5", "B6", "B7", "B8", "B9", "B10", "B11", "B12", "K7", "K8", "K9", "Q7", "Q8", "Q9", "G2", "G74", "K78", "K79", "K80", "Q78", "Q79", "Q80");
    array_push($boldCells, "B" . ($footerStartRow + 5), "B" . ($footerStartRow + 6), "B" . ($footerStartRow + 7), "D" . ($footerStartRow + 4), "F" . ($footerStartRow + 4), "I" . ($footerStartRow + 4), "L" . ($footerStartRow + 4), "N" . ($footerStartRow + 4), "R" . ($footerStartRow + 0), "R" . ($footerStartRow + 1), "R" . ($footerStartRow + 2), "Q" . ($footerStartRow + 3), "Q" . ($footerStartRow + 4), "S" . ($footerStartRow + 4));
    array_push($centerCells, "F2", "B13:T14", "G2:P6", "Q5:T6", "B6:F6", "S4:T4", "Q2", "Q3", "Q4", "S3", "S2", "Q3", "Q4", "O6", "Q6", "S6", "N7", "N8", "N11");
    array_push($centerCells, "D" . ($footerStartRow + 4), "F" . ($footerStartRow + 4), "I" . ($footerStartRow + 4), "L" . ($footerStartRow + 4), "N" . ($footerStartRow + 4), "Q" . ($footerStartRow + 3), "Q" . ($footerStartRow + 4), "S" . ($footerStartRow + 4));
    array_push($alignTextTop, "B$footerStartRow");
    array_push($textLeft, "B$footerStartRow");
    $fontSizes = array_merge($fontSizes, ["G2" => 22, "S3" => 28, "B15" => 12, "Q",]);
    array_push($ColoumsHeaderTable, "B13:T14");


    // $rowPageTerakhir = ($approvalStartRow + 3); // Baris terakhir dari page 1 (72)
    // $latestRow = $rowPageTerakhir; // NILAINYA 72

    $rowPageTerakhir = ($approvalStartRow + 3); // Baris terakhir dari page 1 (72)
    $latestRow = $rowPageTerakhir;


    for ($page = 2; $page <= $jumlahPage; $page++) {

        /* === set new header to page 2 === */
        $pembatasAtas = $latestRow + 1; // NILAI 73
        $sheet->getRowDimension($pembatasAtas)->setRowHeight(10);
        $sheet->mergeCells("B" . ($pembatasAtas) . ":T" . ($pembatasAtas));

        $startsHeaderPage = $pembatasAtas + 1; // 74

        /* set logo */
        $sheet->mergeCells("B" . $startsHeaderPage . ":F" . ($startsHeaderPage + 3));
        $drawing = new Drawing();
        $drawing->setName('Company Logo');
        $drawing->setDescription('Company Logo');
        $drawing->setCoordinates('B' . $startsHeaderPage);
        $drawing->setHeight(55);
        $drawing->setWorksheet($sheet);
        $sheet->setCellValue('B' . $startsHeaderPage, '');

        /* set judul */
        $sheet->mergeCells("Q" . ($startsHeaderPage + 3) . ":T" . ($startsHeaderPage + 4))->setCellValue('Q' . ($startsHeaderPage + 3), "Berlaku mulai : $dataParentArray[berlakuForm].");
        $sheet->mergeCells("G" . ($startsHeaderPage + 0) . ":P" . ($startsHeaderPage + 4))->setCellValue('G' . ($startsHeaderPage + 0), "CHECK SHEET VERIFIKASI JOB SET UP & PATROL CASTING");
        $sheet->mergeCells("B" . ($startsHeaderPage + 4) . ":F" . ($startsHeaderPage + 4))->setCellValue('B' . ($startsHeaderPage + 4), 'FORM');
        $sheet->mergeCells("S" . ($startsHeaderPage + 1) . ":T" . ($startsHeaderPage + 1))->setCellValue('S' . ($startsHeaderPage + 1), "Hal");
        $sheet->mergeCells("S" . ($startsHeaderPage + 2) . ":T" . ($startsHeaderPage + 2))->setCellValue('S' . ($startsHeaderPage + 2), "$dataParentArray[hal]");
        $sheet->mergeCells("Q" . ($startsHeaderPage + 1) . ":R" . ($startsHeaderPage + 1))->setCellValue('Q' . ($startsHeaderPage + 1), "Revisi");
        $sheet->mergeCells("Q" . ($startsHeaderPage + 2) . ":R" . ($startsHeaderPage + 2))->setCellValue('Q' . ($startsHeaderPage + 2), "$dataParentArray[revisi]");
        $sheet->mergeCells("Q" . ($startsHeaderPage + 0) . ":T" . ($startsHeaderPage + 0))->setCellValue('Q' . ($startsHeaderPage + 0), 'FO/QAS/Q1/1907');
        $sheet->mergeCells("B" . ($startsHeaderPage + 5) . ":E" . ($startsHeaderPage + 5))->setCellValue('B' . ($startsHeaderPage + 5), 'Nama part');
        $sheet->mergeCells("B" . ($startsHeaderPage + 6) . ":E" . ($startsHeaderPage + 6))->setCellValue('B' . ($startsHeaderPage + 6), 'No. part');
        $sheet->mergeCells("B" . ($startsHeaderPage + 7) . ":E" . ($startsHeaderPage + 7))->setCellValue('B' . ($startsHeaderPage + 7), 'Code');
        $sheet->mergeCells("Q" . ($startsHeaderPage + 5) . ":R" . ($startsHeaderPage + 5))->setCellValue('Q' . ($startsHeaderPage + 5), 'Nama mesin');
        $sheet->mergeCells("Q" . ($startsHeaderPage + 6) . ":R" . ($startsHeaderPage + 6))->setCellValue('Q' . ($startsHeaderPage + 6), 'No.mesin');
        $sheet->mergeCells("Q" . ($startsHeaderPage + 7) . ":R" . ($startsHeaderPage + 7))->setCellValue('Q' . ($startsHeaderPage + 7), 'No.jig');
        $sheet->mergeCells("K" . ($startsHeaderPage + 5) . ":L" . ($startsHeaderPage + 5))->setCellValue('K' . ($startsHeaderPage + 5), 'Customer');
        $sheet->mergeCells("K" . ($startsHeaderPage + 6) . ":L" . ($startsHeaderPage + 6))->setCellValue('K' . ($startsHeaderPage + 6), 'Model');
        $sheet->mergeCells("K" . ($startsHeaderPage + 7) . ":L" . ($startsHeaderPage + 7))->setCellValue('K' . ($startsHeaderPage + 7), 'No.die');
        $sheet->mergeCells("B" . ($startsHeaderPage + 8) . ":T" . ($startsHeaderPage + 8))->setCellValue('B' . ($startsHeaderPage + 8), "Keterangan : $dataParentArray[ket]");

        /* isi untuk judul or header */
        $sheet->mergeCells("F" . ($startsHeaderPage + 5) . ":J" . ($startsHeaderPage + 5))->setCellValue('F' . ($startsHeaderPage + 5), $dataParentArray["namaPart"]);
        $sheet->mergeCells("F" . ($startsHeaderPage + 6) . ":J" . ($startsHeaderPage + 6))->setCellValue('F' . ($startsHeaderPage + 6), $dataParentArray["noPart"]);
        $sheet->mergeCells("F" . ($startsHeaderPage + 7) . ":J" . ($startsHeaderPage + 7))->setCellValue('F' . ($startsHeaderPage + 7), $dataParentArray["code"]);
        $sheet->mergeCells("M" . ($startsHeaderPage + 5) . ":P" . ($startsHeaderPage + 5))->setCellValue('M' . ($startsHeaderPage + 5), $dataParentArray["customer"]);
        $sheet->mergeCells("M" . ($startsHeaderPage + 6) . ":P" . ($startsHeaderPage + 6))->setCellValue('M' . ($startsHeaderPage + 6), $dataParentArray['model']);
        $sheet->mergeCells("M" . ($startsHeaderPage + 7) . ":P" . ($startsHeaderPage + 7))->setCellValue('M' . ($startsHeaderPage + 7), $dataParentArray["noDie"]);
        $sheet->mergeCells("S" . ($startsHeaderPage + 5) . ":T" . ($startsHeaderPage + 5))->setCellValue('S' . ($startsHeaderPage + 5), $dataParentArray["namaMesin"]);
        $sheet->mergeCells("S" . ($startsHeaderPage + 6) . ":T" . ($startsHeaderPage + 6))->setCellValue('S' . ($startsHeaderPage + 6), $dataParentArray["noMesin"]);
        $sheet->mergeCells("S" . ($startsHeaderPage + 7) . ":T" . ($startsHeaderPage + 7))->setCellValue('S' . ($startsHeaderPage + 7), $dataParentArray["noJig"]);

        /* === set new header to page 2 === */

        /* === set Header of table on page 2 === */
        $startHeaderTabel = $startsHeaderPage + 11;
        $sheet->mergeCells("B" . ($startHeaderTabel + 0) . ":B" . ($startHeaderTabel + 1))->setCellValue('B' . ($startHeaderTabel + 0), 'NO');
        $sheet->mergeCells("C" . ($startHeaderTabel + 0) . ":E" . ($startHeaderTabel + 1))->setCellValue('C' . ($startHeaderTabel + 0), 'ITEM');
        $sheet->mergeCells("F" . ($startHeaderTabel + 0) . ":H" . ($startHeaderTabel + 1))->setCellValue('F' . ($startHeaderTabel + 0), 'STD');
        $sheet->mergeCells("I" . ($startHeaderTabel + 0) . ":J" . ($startHeaderTabel + 0))->setCellValue('I' . ($startHeaderTabel + 0), 'TOLERANSI');
        $sheet->setCellValue('I' . ($startHeaderTabel + 1), 'LOWER');
        $sheet->setCellValue('J' . ($startHeaderTabel + 1), 'UPPER');
        $sheet->mergeCells("K" . ($startHeaderTabel + 0) . ":K" . ($startHeaderTabel + 1))->setCellValue('K' . ($startHeaderTabel + 0), 'CONTROL/METHOD');
        $sheet->mergeCells("L" . ($startHeaderTabel + 0) . ":T" . ($startHeaderTabel + 0))->setCellValue('L' . ($startHeaderTabel + 0), 'SAMPLE');
        $sheet->setCellValue('L' . ($startHeaderTabel + 1), '1');
        $sheet->setCellValue('M' . ($startHeaderTabel + 1), '2');
        $sheet->setCellValue('N' . ($startHeaderTabel + 1), '3');
        $sheet->setCellValue('O' . ($startHeaderTabel + 1), '4');
        $sheet->setCellValue('P' . ($startHeaderTabel + 1), '5');
        $sheet->setCellValue('Q' . ($startHeaderTabel + 1), '6');
        $sheet->setCellValue('R' . ($startHeaderTabel + 1), '7');
        $sheet->setCellValue('S' . ($startHeaderTabel + 1), '8');
        $sheet->setCellValue('T' . ($startHeaderTabel + 1), '9');



        // $sheet->mergeCells("B13:B14")->setCellValue('B13', 'NO');
        // $sheet->mergeCells("C13:E14")->setCellValue('C13', 'ITEM');
        // $sheet->mergeCells("F13:H14")->setCellValue('F13', 'STD');
        // $sheet->mergeCells("I13:J13")->setCellValue('I13', 'TOLERANSI');
        // $sheet->setCellValue('I14', 'LOWER');
        // $sheet->setCellValue('J14', 'UPPER');
        // $sheet->mergeCells("K13:K14")->setCellValue('K13', 'CONTROL/METHOD');
        // $sheet->mergeCells("L13:T13")->setCellValue('L13', 'SAMPLE');
        // $sheet->setCellValue('L14', '1');
        // $sheet->setCellValue('M14', '2');
        // $sheet->setCellValue('N14', '3');
        // $sheet->setCellValue('O14', '4');
        // $sheet->setCellValue('P14', '5');
        // $sheet->setCellValue('Q14', '6');
        // $sheet->setCellValue('R14', '7');
        // $sheet->setCellValue('S14', '8');
        // $sheet->setCellValue('T14', '9');
        /* === set Header of table on page 2 === */

        /* === Set Content on table on page 2 === */
        $startContentPage = $startHeaderTabel + 2; // Data tabel dimulai
        for ($i = 0; $i < $satuPageIsi; $i++) {
            $starts = $startContentPage + $i;
            $number = $i + 1;

            $sheet->setCellValue('B' . $starts, $number); /* NO */
            $sheet->mergeCells("C$starts:E$starts")->setCellValue('C' . $starts, ""); /* item */
            $sheet->setCellValue('H' . $starts, ""); /* STD */
            $sheet->setCellValue('I' . $starts, ""); /* lower */
            $sheet->setCellValue('J' . $starts, ""); /* upper */
            $sheet->mergeCells("K$starts:K$starts")->setCellValue('K' . $starts, "");
            $sheet->mergeCells("F$starts:H$starts")->setCellValue('F' . $starts, "");
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

        array_push(
            $boldCells,
            "B" . ($startsHeaderPage + 11),
            "T" . ($startsHeaderPage + 12),
            "F" . ($startsHeaderPage + 0),
            "Q" . ($startsHeaderPage + 0),
            "S" . ($startsHeaderPage + 0),
            "Q" . ($startsHeaderPage + 1),
            "S" . ($startsHeaderPage + 2),
            "Q" . ($startsHeaderPage + 2),
            "O" . ($startsHeaderPage + 4),
            "Q" . ($startsHeaderPage + 4),
            "S" . ($startsHeaderPage + 4),
            "N" . ($startsHeaderPage + 5),
            "N" . ($startsHeaderPage + 6),
            "N" . ($startsHeaderPage + 9),
            "N" . ($startsHeaderPage + 10),
            "B" . ($startsHeaderPage + 3),
            "B" . ($startsHeaderPage + 4),
            "B" . ($startsHeaderPage + 5),
            "B" . ($startsHeaderPage + 6),
            "B" . ($startsHeaderPage + 7),
            "B" . ($startsHeaderPage + 8),
            "B" . ($startsHeaderPage + 9),
            "B" . ($startsHeaderPage + 10),
            "K" . ($startsHeaderPage + 7),
            "Q" . ($startsHeaderPage + 7),
            "C" . ($startsHeaderPage + 11),
            "F" . ($startsHeaderPage + 11),
            "L" . ($startsHeaderPage + 11),
            "M" . ($startsHeaderPage + 11),
            "L" . ($startsHeaderPage + 11),
            "I" . ($startsHeaderPage + 11),
            "K" . ($startsHeaderPage + 11),
            "I" . ($startsHeaderPage + 12),
            "J" . ($startsHeaderPage + 12),
            "M" . ($startsHeaderPage + 11),
            "L" . ($startsHeaderPage + 12),
            "M" . ($startsHeaderPage + 12),
            "N" . ($startsHeaderPage + 12),
            "O" . ($startsHeaderPage + 12),
            "P" . ($startsHeaderPage + 12),
            "Q" . ($startsHeaderPage + 12),
            "R" . ($startsHeaderPage + 12),
            "S" . ($startsHeaderPage + 12),
            "T" . ($startsHeaderPage + 12),

        );

        array_push($boldCells, "B" . ($footerStartRowPage + 4), "B" . ($footerStartRowPage + 5), "B" . ($footerStartRowPage + 6), "B" . ($footerStartRowPage + 7), "D" . ($footerStartRowPage + 4), "F" . ($footerStartRowPage + 4), "K" . ($footerStartRowPage + 4), "N" . ($footerStartRowPage + 4), "P" . ($footerStartRowPage + 4), "R" . ($footerStartRowPage + 0), "R" . ($footerStartRowPage + 1), "R" . ($footerStartRowPage + 2), "Q" . ($footerStartRowPage + 3), "Q" . ($footerStartRowPage + 4), "S" . ($footerStartRowPage + 4));
        array_push($centerCells, "G" . ($startsHeaderPage + 0), "B" . ($startsHeaderPage + 4), "Q" . ($startsHeaderPage + 3), "B" . ($startsHeaderPage + 3), "Q" . ($startsHeaderPage + 2), "S" . ($startsHeaderPage + 2), "F" . ($startsHeaderPage + 0), "B" . ($startsHeaderPage + 11) . ":T" . ($startsHeaderPage + 12), "Q" . ($startsHeaderPage + 0), "Q" . ($startsHeaderPage + 1), "Q" . ($startsHeaderPage + 2), "S" . ($startsHeaderPage + 1), "S" . ($startsHeaderPage + 0), "Q" . ($startsHeaderPage + 1), "Q" . ($startsHeaderPage + 2), "O" . ($startsHeaderPage + 4), "S" . ($startsHeaderPage + 4), "N" . ($startsHeaderPage + 5), "N" . ($startsHeaderPage + 6), "N" . ($startsHeaderPage + 9), "G" . ($startsHeaderPage + 1));
        array_push($centerCells, "D" . ($footerStartRowPage + 4), "F" . ($footerStartRowPage + 4), "K" . ($footerStartRowPage + 4), "N" . ($footerStartRowPage + 4), "P" . ($footerStartRowPage + 4), "Q" . ($footerStartRowPage + 3), "S" . ($footerStartRowPage + 4));
        array_push($alignTextTop, "B$footerStartRowPage");
        array_push($textLeft, "B$footerStartRowPage");
        array_push($textRight, "N" . ($startsHeaderPage + 10));
        $fontSizes = array_merge($fontSizes, ["F" . ($startsHeaderPage + 0) => 22, "G" . ($startsHeaderPage + 0) => 22], );
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

    foreach ($textitalic as $cell) {
        $sheet->getStyle($cell)->applyFromArray([
            'font' => [
                'italic' => true,
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
header('Content-Disposition: attachment; filename="' . $fileName . '.xls' . '"');
$writer->save('php://output');