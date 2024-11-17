<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Check if it's an AJAX request
    if (isset($_SERVER['HTTP_X_REQUESTED_WITH']) && strtolower($_SERVER['HTTP_X_REQUESTED_WITH']) === 'xmlhttprequest') {
        // It's an AJAX request

        // Lấy dữ liệu từ form
        $name = isset($_POST['name']) ? $_POST['name'] : '';
        $title = isset($_POST['title']) ? $_POST['title'] : '';
        $department = isset($_POST['department']) ? $_POST['department'] : '';
        $phone = isset($_POST['phone']) ? $_POST['phone'] : '';
        $dependants = isset($_POST['dependants']) ? $_POST['dependants'] : '';
        $answer = isset($_POST['answer']) ? $_POST['answer'] : '';
        $ip = $_SERVER['REMOTE_ADDR'];

        // Kiểm tra và tạo file Excel nếu chưa tồn tại
        $excelFile = 'cljoc-event.xlsx';
        if (!file_exists($excelFile)) {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->setCellValue('A1', 'Staff Name');
            $sheet->setCellValue('B1', 'Title');
            $sheet->setCellValue('C1', 'Department');
            $sheet->setCellValue('D1', 'Telephone');
            $sheet->setCellValue('E1', 'Dependants - DOB');
            $sheet->setCellValue('F1', 'IP');
            $sheet->setCellValue('G1', 'Created On');
        } else {
            // Nếu file tồn tại, mở nó để thêm dữ liệu mới
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($excelFile);
            $sheet = $spreadsheet->getActiveSheet();
        }

        // Thêm dữ liệu mới vào file Excel
        $lastRow = $sheet->getHighestRow() + 1;
        $sheet->setCellValue('A' . $lastRow, $name);
        $sheet->setCellValue('B' . $lastRow, $title);
        $sheet->setCellValue('C' . $lastRow, $department);
        $sheet->setCellValue('D' . $lastRow, $phone);
        $sheet->setCellValue('E' . $lastRow, $dependants);
        $sheet->setCellValue('F' . $lastRow, $ip);
        $sheet->setCellValue('G' . $lastRow, gmdate('Y-m-d H:i:s', time() + 7 * 3600));

        // // Lưu file Excel
        $writer = new Xlsx($spreadsheet);
        $writer->save($excelFile);

        $serverName = "APPLAB01";
        $connectionOptions = array(
            "Database" => "Registration_Form",
            "Uid" => "sa",
            "PWD" => "sa",
            "CharacterSet" => "UTF-8"
        );

        $conn = sqlsrv_connect($serverName, $connectionOptions);

        if ($conn) {
            $sql = "INSERT INTO registration (staff_name, title, department, phone, dependent, answer, ip, created_on)
                    VALUES (?, ?, ?, ?, ?, ?, GETDATE())";
            $params = array($name, $title, $department, $phone, $dependants, $answer, $ip);
            $stmt = sqlsrv_query($conn, $sql, $params);

            if ($stmt) {
                // Insertion successful
                echo json_encode(['success' => true, 'message' => 'You have successfully confirmed!']);
            } else {
                // Insertion failed
                echo json_encode(['success' => false, 'message' => 'Error inserting data into the database!']);
            }

            sqlsrv_close($conn);
        } else {
            // Connection failed
            echo json_encode(['success' => false, 'message' => 'Error connecting to the database!']);
        }
        exit;
    } else {
        echo json_encode(['success' => false, 'message' => 'An error occurred, please try again!']);
        exit;
    }
}
?>

<!DOCTYPE html>
<html lang="vi">

<head>
    <meta charset="utf-8">
    <title>Cửu Long Party </title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes" />
    <link rel="stylesheet" href="resources/library/css/library.css">
    <link rel="stylesheet" href="resources/uikit/css/uikit.modify.css">
    <link rel="stylesheet" href="resources/uikit/css/uikit.modify.css">
    <link rel="stylesheet" type="text/css" href="node_modules/toastify-js/src/toastify.css">
    <link rel="stylesheet" href="resources/sweet.css">
    <link rel="stylesheet" href="resources/style.css">
    <script src="resources/library/js/jquery.js"></script>
</head>

<body>
    <div class="homepage">
        <div class="uk-container uk-container-center">
            <div class="uk-grid uk-grid-collapse">
                <div class="uk-width-small-1-1 uk-width-medium-1-2" style="text-align: center;">
                    <div class="text bling-overlay">
                        <img src="/resources/img/TEXTTTT.png" alt="">
                    </div>
                </div>
                <div class="uk-width-small-1-1 uk-width-medium-1-2 relative">
                    <div class="invitation"><img src="/resources/img/invitation.png" alt=""></div>
                    <div class="timeline">
                        <div class="uk-grid uk-grid-large">
                            <div class="uk-width-2-5">
                                <div class="time-left uk-text-right">
                                    <div class="time-container">
                                        <div class="time-start">17:30 - 18:15</div>
                                        <div class="time-img uk-flex uk-flex-right">
                                            <div class="time-img-item img-scaledown"><img src="/resources/img/1.png" alt=""></div>
                                            <div class="time-img-item img-scaledown"><img src="/resources/img/2.png" alt=""></div>
                                            <div class="time-img-item img-scaledown"><img src="/resources/img/3.png" alt=""></div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="uk-width-3-5 time-right-wrap">
                                <div class="time-right uk-text-left">
                                    <div class="time-container">
                                        <h4 class="title"><span>Welcome</span></h4>
                                        <ul class="uk-clearfix timelist">
                                            <li>Traditional games for kids</li>
                                            <li>Photo booth activities</li>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="timeline">
                        <div class="uk-grid uk-grid-large">
                            <div class="uk-width-2-5">
                                <div class="time-left uk-text-right">
                                    <div class="time-container">
                                        <div class="time-start">18:15 - 20:45</div>
                                        <div class="time-img">
                                            <div class="uk-flex uk-flex-right" style="padding-right: 35px;">
                                                <div class="time-img-item img-scaledown mb30"><img src="/resources/img/4.png" alt=""></div>
                                                <div class="time-img-item img-scaledown mb30"><img src="/resources/img/5.png" alt=""></div>
                                            </div>
                                            <div class="uk-flex uk-flex-right">
                                                <div class="time-img-item img-scaledown"><img src="/resources/img/8.png" alt=""></div>
                                                <div class="time-img-item-2 img-scaledown"><img src="/resources/img/thia.png" alt=""></div>
                                                <div class="time-img-item img-scaledown"><img src="/resources/img/6.png" alt=""></div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="uk-width-3-5 time-right-wrap">
                                <div class="time-right uk-text-left">
                                    <div class="time-container">
                                        <h4 class="title"><span>Appreciation & Party Program</span></h4>
                                        <ul class="uk-clearfix timelist">
                                            <li>Openning performance.</li>
                                            <li>GM Speech on 2024 activities & 2025 work programs.</li>
                                            <li>Company highlights (ten initiative events of 2024).</li>
                                            <li>Toasting moment - Party Announcement.</li>
                                            <li>Party celebration.</li>
                                            <li>Lucky Draw.</li>
                                            <li>Magic performances and games, entertaiment.</li>
                                            <li>Other Activities.</li>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="timeline">
                        <div class="uk-grid uk-grid-large">
                            <div class="uk-width-2-5">
                                <div class="time-left uk-text-right">
                                    <div class="time-container">
                                        <div class="time-start">20:45 to Closing</div>
                                        <div class="time-img uk-flex uk-flex-right">

                                            <div class="time-img-item img-scaledown"><img src="/resources/img/9.png" alt=""></div>
                                            <div class="time-img-item img-scaledown"><img src="/resources/img/10.png" alt=""></div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="uk-width-3-5 time-right-wrap">
                                <div class="time-right uk-text-left">
                                    <div class="time-container">
                                        <h4 class="title"><span>After Party Session </span></h4>
                                        <ul class="uk-clearfix timelist">
                                            <li>Karaoke, Song, Music Performance</li>
                                        </ul>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="contact" style="margin-left: -150px;">
                        <div class="uk-grid uk-grid-small">
                            <div class="uk-width-small-1-1 uk-width-medium-1-2">
                                <div class="content ">
                                    <div class="content-img"><img src="/resources/img/Join Us_.png" alt=""></div>
                                    <div class="description uk-text-center">Cuu Long Year End Party will be more complete <br> with your presence. Please confirm your <br> participation so we can welcome you and your <br> family in the most thorough way!</div>
                                    <div class="uk-text-right">Best regards</div>
                                </div>
                            </div>
                            <div class="uk-width-small-1-1 uk-width-medium-1-2">
                                <form action="" method="post">
                                    <div class="uk-grid uk-grid-medium">
                                        <div class="uk-width-3-5">
                                            <div class="uk-grid uk-grid-medium">
                                                <div class="uk-width-1-2 mb20">
                                                    <div class="form-field">
                                                        <label for="">Staff name</label>
                                                        <input type="text" required id="name" name="name">
                                                    </div>
                                                </div>
                                                <div class="uk-width-1-2 mb20">
                                                    <div class="form-field">
                                                        <label for="">Company Email</label>
                                                        <input type="text" required id="Email" name="Email">
                                                    </div>
                                                </div>
                                                <div class="uk-width-1-2 ">
                                                    <div class="form-field">
                                                        <label for="">Department</label>
                                                        <div>
                                                            <select required id="department" name="department">
                                                                <option value="Administration">Administration</option>
                                                                <option value="Operation">Operation</option>
                                                                <option value="Finance&Accounting">Finance&Accounting</option>
                                                                <option value="Procurement">Procurement</option>
                                                                <option value="Sub-surface">Sub-surface</option>
                                                                <option value="Development">Development</option>
                                                                <option value="HSE">HSE</option>
                                                                <option value="OPR-STV">OPR-STV</option>
                                                                <option value="OPR-STD">OPR-STD</option>
                                                                <option value="OPR-STT">OPR-STT</option>
                                                                <option value="OPR-VungTauBase">OPR-VungTauBase</option>
                                                            </select>
                                                        </div>
                                                    </div>
                                                </div>
                                                <div class="uk-width-1-2 ">
                                                    <div class="form-field">
                                                        <label for="">Telephone</label>
                                                        <input type="text" required id="phone" name="phone">
                                                    </div>

                                                </div>

                                            </div>
                                            <button class="btn-submit" id="submitBtn">Feedback</button>
                                        </div>
                                        <div class="uk-width-2-5">
                                            <div class="form-field">
                                                <label for="">Dependants - DOB</label>
                                                <div class="relative" id="input-container">
                                                    <input type="text" required id="phone" name="dependants">
                                                    <!-- <textarea id="limitedTextarea" rows="8" name="dependants" style="resize: none; overflow: hidden;"></textarea> -->
                                                    <button class="butn add-btn" type="button">
                                                        <svg width="15px" height="15px" viewBox="0 0 15 15" version="1.1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
                                                            <g id="Home-2" stroke="none" stroke-width="1" fill="none" fill-rule="evenodd">
                                                                <g id="1-Home" transform="translate(-583.000000, -280.000000)" fill="#fff" stroke="#fff">
                                                                    <polygon id="+" points="591.264706 286.735294 597 286.735294 597 288.264706 591.264706 288.264706 591.264706 294 589.735294 294 589.735294 288.264706 584 288.264706 584 286.735294 589.735294 286.735294 589.735294 281 591.264706 281"></polygon>
                                                                </g>
                                                            </g>
                                                        </svg>
                                                    </button>
                                                </div>
                                            </div>
                                        </div>
                                    </div>

                                    <!-- <div class="form-field uk-flex mb20 form-radio-va">
                                        <input type="radio" name="answer" id="opt1" value="yes" required>
                                        <label for="opt1" class="label1" style="margin-right: 15px">
                                            <span>Yes</span>
                                        </label>
                                        <input type="radio" name="answer" id="opt2" value="no" required>
                                        <label for="opt2" class="label2">
                                            <span>No</span>
                                        </label>
                                    </div> -->
                                </form>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div style="height: 0; width: 0; position: absolute; visibility: hidden;">
        <svg xmlns="http://www.w3.org/2000/svg">
            <symbol id="icon-play" viewBox="0 0 24 24">
                <path d="M8 5v14l11-7z" />
            </symbol>
            <symbol id="icon-pause" viewBox="0 0 24 24">
                <path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z" />
            </symbol>
            <symbol id="icon-close" viewBox="0 0 24 24">
                <path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z" />
            </symbol>
            <symbol id="icon-settings" viewBox="0 0 24 24">
                <path
                    d="M19.43 12.98c.04-.32.07-.64.07-.98s-.03-.66-.07-.98l2.11-1.65c.19-.15.24-.42.12-.64l-2-3.46c-.12-.22-.39-.3-.61-.22l-2.49 1c-.52-.4-1.08-.73-1.69-.98l-.38-2.65C14.46 2.18 14.25 2 14 2h-4c-.25 0-.46.18-.49.42l-.38 2.65c-.61.25-1.17.59-1.69.98l-2.49-1c-.23-.09-.49 0-.61.22l-2 3.46c-.13.22-.07.49.12.64l2.11 1.65c-.04.32-.07.65-.07.98s.03.66.07.98l-2.11 1.65c-.19.15-.24.42-.12.64l2 3.46c.12.22.39.3.61.22l2.49-1c.52.4 1.08.73 1.69.98l.38 2.65c.03.24.24.42.49.42h4c.25 0 .46-.18.49-.42l.38-2.65c.61-.25 1.17-.59 1.69-.98l2.49 1c.23.09.49 0 .61-.22l2-3.46c.12-.22.07-.49-.12-.64l-2.11-1.65zM12 15.5c-1.93 0-3.5-1.57-3.5-3.5s1.57-3.5 3.5-3.5 3.5 1.57 3.5 3.5-1.57 3.5-3.5 3.5z" />
            </symbol>
            <symbol id="icon-sound-on" viewBox="0 0 24 24">
                <path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3c0-1.77-1.02-3.29-2.5-4.03v8.05c1.48-.73 2.5-2.25 2.5-4.02zM14 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77z" />
            </symbol>
            <symbol id="icon-sound-off" viewBox="0 0 24 24">
                <path
                    d="M16.5 12c0-1.77-1.02-3.29-2.5-4.03v2.21l2.45 2.45c.03-.2.05-.41.05-.63zm2.5 0c0 .94-.2 1.82-.54 2.64l1.51 1.51C20.63 14.91 21 13.5 21 12c0-4.28-2.99-7.86-7-8.77v2.06c2.89.86 5 3.54 5 6.71zM4.27 3L3 4.27 7.73 9H3v6h4l5 5v-6.73l4.25 4.25c-.67.52-1.42.93-2.25 1.18v2.06c1.38-.31 2.63-.95 3.69-1.81L19.73 21 21 19.73l-9-9L4.27 3zM12 4L9.91 6.09 12 8.18V4z" />
            </symbol>
        </svg>
    </div>

    <div class="fix-total">
        <div class="relative">
            <img src="resources/img/total.png" alt="" />
            <div class="total-fixed">
                Number of Registered People:
                <br />
                <span>02-Jan-2024: <strong>670</strong></span>
            </div>
        </div>
    </div>

    <script src="node_modules/toastify-js/src/toastify.js"></script>
    <script src="resources/sweet.js"></script>

    <script>
        $(document).ready(function() {
            const maxInputs = 3;
            $(document).on('click', '.add-btn', function() {
                const currentInputs = $('#input-container .input-group').length;
                if (currentInputs < maxInputs) {
                    const newInputGroup = `
                    <div class="input-group mt10">
                        <input type="text" required name="dependants">
                        <button class="butn remove-btn" type="button" aria-label="Xóa điểm dừng">
                            <svg width="15px" height="15px" viewBox="0 0 15 15" version="1.1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink">
                                <g id="Home-2" stroke="none" stroke-width="1" fill="none" fill-rule="evenodd">
                                    <g id="1-Home" transform="translate(-583.000000, -280.000000)" fill="#fff" stroke="#fff">
                                        <polygon id="+" points="591.264706 286.735294 597 286.735294 597 288.264706 591.264706 288.264706 591.264706 294 589.735294 294 589.735294 288.264706 584 288.264706 584 286.735294 589.735294 286.735294 589.735294 281 591.264706 281"></polygon>
                                    </g>
                                </g>
                            </svg>
                        </button>
                    </div>
                    `;
                    $('#input-container').append(newInputGroup);

                }

            });

            // Sự kiện click cho nút xóa (-)
            $(document).on('click', '.remove-btn', function() {
                // Xóa input-group hiện tại
                $(this).closest('.input-group').remove();
            });
        });

        $(document).ready(function() {
            const maxLines = 8;

            $('#limitedTextarea').on('input', function() {
                const lines = $(this).val().split('\n');

                if (lines.length > maxLines) {
                    $(this).val(lines.slice(0, maxLines).join('\n'));
                }
            });
        });

        document.getElementById('submitBtn').addEventListener('click', function(event) {
            event.preventDefault();
            if (validateForm()) {
                sendData();
            }
        });

        function validateForm() {
            var name = document.getElementById('name').value.trim();
            var Email = document.getElementById('Email').value.trim();
            var department = document.getElementById('department').value.trim();
            var phone = document.getElementById('phone').value.trim();
            // var answerYes = document.getElementById('opt1').checked;
            // var answerNo = document.getElementById('opt2').checked;
            var emailRegex = /^[a-zA-Z0-9._%+-]+@cljoc\.com\.vn$/;
            if (name === '') {
                alert('Please enter Staff name.');
                return false;
            }

            if (!emailRegex.test(Email)) {
                alert('Please enter a valid Email with domain cljoc.com.vn.');
                return false;
            }

            if (department === '') {
                alert('Please enter Department.');
                return false;
            }

            if (phone === '') {
                alert('Please enter Telephone.');
                return false;
            }

            // if (!answerYes && !answerNo) {
            //     alert('Please select an answer for Yes/No.');
            //     return false;
            // }

            return true;
        }

        function sendData() {
            // Prepare form data
            var formData = new FormData(document.querySelector('form'));

            // Make AJAX request
            var xhr = new XMLHttpRequest();
            xhr.open('POST', 'index.php', true);
            xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');

            // Set up event handler for successful response
            xhr.onload = function() {
                if (xhr.status >= 200 && xhr.status < 300) {
                    handleData(xhr.responseText);
                } else {
                    console.error('Error submitting form:', xhr.statusText);
                }
            };

            // Set up event handler for error
            xhr.onerror = function() {
                console.error('Network error while submitting form.');
            };

            // Send form data
            xhr.send(formData);
        }

        function handleData(response) {
            // Parse the response as needed
            var data = JSON.parse(response);
            console.log(data)

            if (data.success) {
                // Handle success
                $('form')[0].reset();
                swal("Success", data.message, "success");
            } else {
                swal("Error", data.message, "error");

            }
        }
    </script>
</body>

</html>