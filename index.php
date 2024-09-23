<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_POST['submit']) && isset($_FILES['file'])) {
    $valueToAdd = (float)$_POST['value'];

    $fileTmpPath = $_FILES['file']['tmp_name'];

    try {
        $spreadsheet = IOFactory::load($fileTmpPath);
        $sheet = $spreadsheet->getActiveSheet();
        $columnsToEdit = ['AO', 'AQ', 'AS'];
        $row = 2;
        while (true) {
            $cellValue = $sheet->getCell('AO' . $row)->getValue();

            if ($cellValue === null) {
                break;
            }

            foreach ($columnsToEdit as $column) {
                $cell = $sheet->getCell($column . $row);
                $originalValue = (float)$cell->getValue();
                $newValue = $originalValue + $valueToAdd;
                $sheet->setCellValue($column . $row, $newValue);
            }
            $row++;
        }

        $newFilename = 'modified_file.xlsx';
        $writer = new Xlsx($spreadsheet);

        $outputFile = tempnam(sys_get_temp_dir(), $newFilename);
        $writer->save($outputFile);

        header('Content-Description: File Transfer');
        header('Content-Disposition: attachment; filename="' . $newFilename . '"');
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Length: ' . filesize($outputFile));
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        readfile($outputFile);

        unlink($outputFile);
        exit();
    } catch (Exception $e) {
        echo 'Ошибка при обработке файла: ', $e->getMessage();
    }
} else {
    echo "Пожалуйста, загрузите файл.";
}
