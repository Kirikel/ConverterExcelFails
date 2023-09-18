<?php

ini_set('default_charset', 'utf-8');
ini_set('display_errors', 0);
ini_set('display_startup_errors', 0);
error_reporting(E_ALL);

require('Classes/PHPExcel.php');


?>
<!doctype>
<html>
<head>
</head>

<style>
body {
    background-color: #f5f5f5;
    font-family: Arial, sans-serif;
}

.container {
    background-color: #333;
    color: #fff;
    padding: 20px;
    margin: 20px auto;
    max-width: 400px;
    border-radius: 5px;
    box-shadow: 0 0 10px rgba(0, 0, 0, 0.2);
}

.header {
    font-size: 24px;
    margin-bottom: 20px;
}

.form {
    display: flex;
    flex-direction: column;
}

.file-label {
    font-size: 16px;
    margin-bottom: 10px;
}

.file-input {
    padding: 10px;
    border: 1px solid #ddd;
    border-radius: 5px;
    font-size: 14px;
    margin-bottom: 10px;
}

.submit-button {
    background-color: #007bff;
    color: #fff;
    padding: 10px;
    border: none;
    border-radius: 5px;
    font-size: 16px;
    cursor: pointer;
    transition: background-color 0.2s;
}

.submit-button:hover {
    background-color: #0056b3;
}
</style>


<body>


<?php
if ($_SERVER["REQUEST_METHOD"] == "POST") {
$file_tmp = $_FILES['excel_file']['tmp_name'];
$tmpfname = $file_tmp;
$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
$excelObj = $excelReader->load($tmpfname);
$worksheet = $excelObj->getSheet(0);
$lastRow = $worksheet->getHighestRow();
$dataArray = array();
for ($row = 8; $row <= $lastRow; $row++) {
    $service = $worksheet->getCell('E' . $row)->getValue();
    $price = $worksheet->getCell('F' . $row)->getValue();
    $source = $worksheet->getCell('L' . $row)->getValue();
    $tariff = $worksheet->getCell('G' . $row)->getValue();
    $isWeekend = strpos($tariff, 'Выходные') !== false;
    $issource = $source === 'Да';
    $key = $service . ($isWeekend ? ' (Выходные)' : ' (Будни)') . ($issource ? ' (Источник)' : '');
	if (array_key_exists($key, $dataArray)) {
		if ($price != $dataArray[$key]['price']) {
			$key .= " (Цена: $price)";	
		}
    }
    if (array_key_exists($key, $dataArray)) {
        $dataArray[$key]['quantity']++;
        $dataArray[$key]['price'] = $price;
        $dataArray[$key]['total'] += $price;
    } else {
        $dataArray[$key] = [
            'service' => $service,
            'tariff' => $isWeekend ? 'Выходные' : 'Будни',
            'quantity' => 1,
            'price' => $price,
            'total' => $price,
            'source' => $issource ? 'Маяк' : 'Маламут'
        ];
    }
}

$objPHPExcel = new PHPExcel();
$activeSheet = $objPHPExcel->setActiveSheetIndex(0);
$activeSheet->setCellValue('A1', 'ID')
			->setCellValue('B1', 'Услуга')
            ->setCellValue('C1', 'Тариф')
			->setCellValue('D1', 'кол-во')
			->setCellValue('E1', 'Цена')
			->setCellValue('F1', 'Сумма')
            ->setCellValue('G1', 'Источник');

$countItems = 2;
foreach ($dataArray as $data) {
$activeSheet->setCellValue('A' . $countItems, $countItems)
			->setCellValue('B' . $countItems, $data['service'])
            ->setCellValue('C' . $countItems, $data['tariff'])
			->setCellValue('D' . $countItems, $data['quantity'])
			->setCellValue('E' . $countItems, $data['price'])
			->setCellValue('F' . $countItems, $data['total'])
            ->setCellValue('G' . $countItems, $data['source']);
	$countItems = $countItems + 1;
}
$filename = 'example.xlsx';
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($filename);
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment;filename=\"$filename\"");
header('Cache-Control: max-age=0');
$objWriter->save('php://output');
}else{?>
	    <div class="container">
        <h2 class="header">Загрузка и импорт файла Excel</h2>
        <form class="form" method="post" enctype="multipart/form-data">
            <label for="excel_file" class="file-label">Выберите файл Excel:</label>
            <input type="file" id="excel_file" name="excel_file" accept=".xlsx" class="file-input">
            <input type="submit" value="Отправить в обработку" class="submit-button">
        </form>
    </div>
	
<?php } ?>
</body>
</html>