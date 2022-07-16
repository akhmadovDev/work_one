<?php

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

require_once 'config.php';
require_once 'vendor/autoload.php';

/**
 * @var $host // hostname
 * @var $dbName // database name
 * @var $user // database user
 * @var $password // database password
 * @var  $dsn
 * Database connect
 */
$dsn = "mysql:host=$host;dbname=$dbName;charset=UTF8";
$options = [
    PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION
];

try {
    $pdo = new PDO($dsn, $user, $password, $options);
} catch (Exception $e) {
    echo $e->getMessage();
}

$Reader = new Xlsx();

$spreadSheet = $Reader->load('1.xlsx');
$excelSheet = $spreadSheet->getActiveSheet();

/** Read and import data */
foreach ($excelSheet->getRowIterator() as $row) {
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(false);
    $data = [];

    foreach ($cellIterator as $cell) {
        if (empty($cell->getValue())) {
            $data[] = 'empty';
        }
        $data[] = $cell->getValue();
    }

    try {
        $sql = /** @lang sql */
            "INSERT INTO `xls` (`name`, `email`) VALUES (:name, :email)";
        $stmt = $pdo->prepare($sql);
        $stmt->execute([
            ':name' => $data[0],
            ':email' => $data[1]
        ]);
    } catch (Exception $e) {
        echo $e->getMessage();
    }
}



