<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Verifica se o arquivo foi enviado
    if (!isset($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
        die('Erro ao enviar o arquivo.');
    }

    // Verifica se o arquivo é um PDF
    $file = $_FILES['file'];
    $fileType = mime_content_type($file['tmp_name']);
    if ($fileType !== 'application/pdf') {
        die('O arquivo enviado não é um PDF.');
    }

    // Salva o arquivo na pasta "uploads"
    $uploadDir = 'uploads/';
    if (!is_dir($uploadDir)) {
        mkdir($uploadDir, 0777, true);
    }
    $filePath = $uploadDir . basename($file['name']);
    move_uploaded_file($file['tmp_name'], $filePath);

    // Extrai dados do PDF (simulação)
    $transactions = extractTransactionsFromPDF($filePath);

    if (empty($transactions)) {
        die('Nenhuma transação encontrada no PDF.');
    }

    // Cria o arquivo Excel
    $excelPath = str_replace('.pdf', '.xlsx', $filePath);
    createExcel($transactions, $excelPath);

    // Força o download do arquivo Excel
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment; filename="' . basename($excelPath) . '"');
    readfile($excelPath);

    // Remove os arquivos temporários
    unlink($filePath);
    unlink($excelPath);
    exit;
}

function extractTransactionsFromPDF($pdfPath) {
    // Simulação de extração de dados do PDF
    // Substitua por uma biblioteca como TCPDF ou FPDI para processar PDFs reais
    return [
        ['Data' => '01-01-2023', 'Descrição' => 'Compra A', 'Valor' => 100.50],
        ['Data' => '02-01-2023', 'Descrição' => 'Compra B', 'Valor' => 200.75],
    ];
}

function createExcel($transactions, $excelPath) {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Define os cabeçalhos
    $sheet->setCellValue('A1', 'Data');
    $sheet->setCellValue('B1', 'Descrição');
    $sheet->setCellValue('C1', 'Valor');

    // Adiciona os dados
    $row = 2;
    foreach ($transactions as $transaction) {
        $sheet->setCellValue("A{$row}", $transaction['Data']);
        $sheet->setCellValue("B{$row}", $transaction['Descrição']);
        $sheet->setCellValue("C{$row}", $transaction['Valor']);
        $row++;
    }

    // Formata a coluna de valores
    foreach ($sheet->getColumnIterator('C') as $column) {
        foreach ($column->getCellIterator() as $cell) {
            $cell->getStyle()->getNumberFormat()->setFormatCode('#,##0.00');
        }
    }

    // Salva o arquivo Excel
    $writer = new Xlsx($spreadsheet);
    $writer->save($excelPath);
}