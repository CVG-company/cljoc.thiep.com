<?php
// require '../../vendor/autoload.php';

// use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

// function readExcel($filePath)
// {
//     $reader = new Xlsx();
//     $spreadsheet = $reader->load($filePath);
//     $data = [];

//     foreach ($spreadsheet->getActiveSheet()->getRowIterator() as $row) {
//         $rowData = [];
//         foreach ($row->getCellIterator() as $cell) {
//             $rowData[] = $cell->getValue();
//         }
//         $data[] = $rowData;
//     }

//     return $data;
// }

// $excelData = readExcel('../../cljoc-event.xlsx');
// $listUser = [];
// $dateTime = new DateTime();
// if (isset($excelData) && is_array($excelData) && count($excelData)) {
//     $title = $excelData[0];
//     foreach ($excelData as $key => $value) {
//         if ($key == 0) continue;
//         $data = [];
//         foreach ($title as $keyTitle => $valueTitle) {
//             $data[$valueTitle] = $value[$keyTitle];
//         }
//         $listUser[] = $data;
//     }
// }

// Cấu hình Database
$serverName = "APPLAB01";
$connectionOptions = array(
    "Database" => "Registration_Form",
    "Uid" => "sa",
    "PWD" => "sa",
    "CharacterSet" => "UTF-8"
);

$conn = sqlsrv_connect($serverName, $connectionOptions);
if ($conn) {
    $sql = "SELECT * FROM registration";
    $isSql = true;
    $selectStmt = sqlsrv_query($conn, $sql);
    $row = sqlsrv_fetch_array($selectStmt, SQLSRV_FETCH_ASSOC);

    if ($selectStmt) {
        $listUser = array();

        while ($row = sqlsrv_fetch_array($selectStmt, SQLSRV_FETCH_ASSOC)) {
            $listUser[] = $row;
        }
    }

    sqlsrv_close($conn);
} else {
    // Connection failed
    echo json_encode(['success' => false, 'message' => 'Error connecting to the database!']);
}
?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <title>DataTables Example</title>
    <link rel="stylesheet" type="text/css" href="/resources/css/bootstrap.min.css">
    <link rel="stylesheet" type="text/css" href="/resources/css/dataTables.bootstrap.css">
    <link rel="stylesheet" type="text/css" href="/resources/css/dataTables.responsive.css">
    <link rel="stylesheet" type="text/css" href="/resources/style.css">
    <script src="/resources/js/jquery.min.js"></script>
    <script src="/resources/js/jquery.dataTables.min.js"></script>
    <script src="/resources/js/dataTables.bootstrap.js"></script>
    <script src="/resources/js/dataTables.responsive.js"></script>
    <style>
        body {
            font-size: 140%;
        }

        h2 {
            text-align: center;
            padding: 20px 0;
        }

        table caption {
            padding: .5em 0;
        }

        table.dataTable th,
        table.dataTable td {
            white-space: nowrap;
        }

        .p {
            text-align: center;
            padding-top: 140px;
            font-size: 14px;
        }

        .container {
            width: 80%;
        }
    </style>
</head>

<body>
    <h2>Responsive Table with DataTables</h2>

    <div class="container">
        <div class="row">
            <div class="col-xs-12">
                <table class="table table-bordered table-hover dt-responsive">
                    <thead>
                        <tr>
                            <th>Staff name</th>
                            <th>Title</th>
                            <th>Department</th>
                            <th>Telephone</th>
                            <th>Dependants name - Year of birth</th>
                            <th>Time Register</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php
                        if (isset($listUser) && is_array($listUser) && count($listUser)) {
                            foreach ($listUser as $row) :
                        ?>
                                <tr>
                                    <td><?php echo $row['staff_name']; ?></td>
                                    <td><?php echo $row['title']; ?></td>
                                    <td><?php echo $row['department']; ?></td>
                                    <td><?php echo $row['phone']; ?></td>
                                    <td><?php echo $row['dependent']; ?></td>
                                    <td><?php echo isset($isSql) && $isSql ? $row['created_on']->format('Y-m-d H:i:s') : date('d/m/Y H:i:s', strtotime($row['created_on'])); ?></td>
                                </tr>
                        <?php endforeach;
                        } ?>
                    </tbody>
                </table>

            </div>
        </div>
    </div>
    <script>
        $('table').DataTable({
            "pageLength": 50
        });
    </script>
</body>

</html>