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
            $sheet->setCellValue('F1', 'Answer');
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
        $sheet->setCellValue('F' . $lastRow, $answer);

        // Lưu file Excel
        $writer = new Xlsx($spreadsheet);
        $writer->save($excelFile);

        // Return a JSON response for AJAX
        header('Content-Type: application/json');
        echo json_encode(['success' => true, 'message' => 'You have successfully confirmed!']);
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
        <link rel="stylesheet" href="resources/style.css">
        <script src="/resources/library/js/jquery.js"></script>
    </head>
    <body>
        <div class="homepage">
            <div class="top-bg">
                <img src="/resources/img/b1.png" alt="">
                <img src="/resources/img/b2.png" alt="">
            </div>
            <div class="uk-container uk-container-center">
                <div class="logo uk-text-center img-scaledown"><img src="/resources/img/logo.png" alt=""></div>
                <h1 class="heading-1 uk-text-center"><span>Year and party</span></h1>
                <h2 class="time uk-text-center"><span>13 Jan 2024</span></h2>
                <div class="clock uk-text-center img-scaledown"><img src="/resources/img/oclock.png" alt=""></div>
                <h2 class="title uk-text-center"><span>INVITATION</span></h2>
                <div class="description uk-text-center">Come celebrate with us as we ring in the Year End Party</div>
                <div class="block uk-text-center">
                    <div class="uk-grid uk-grid-collapse">
                        <div class=" uk-width-1-3">
                            <div class="block-item">
                                <h3 class="title"><span>Time</span></h3>
                                <div class="number">16:30</div>
                            </div>
                        </div>
                        <div class=" uk-width-1-3">
                            <div class="block-item">
                                <h3 class="title"><span>Date</span></h3>
                                <div class="number">13</div>
                            </div>
                        </div>
                        <div class=" uk-width-1-3">
                            <div class="block-item">
                                <h3 class="title"><span>Month</span></h3>
                                <div class="number">01.2024</div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="address">
                    <h3 class="small-title uk-text-center"><span>Adress</span></h3>
                    <h2 class="title uk-text-center"><span>GEM CENTER</span></h2>
                    <div class="address-detail uk-text-center">08 Nguyen Binh Khiem, Da Kao, District 1, HCMC</div>
                </div>
                <div class="line  img-scaledown"><img src="/resources/img/line.png" alt=""></div>
                <div class="timeline">
                    <h2 class="title uk-text-center"><span>Time line</span></h2>
                    <div class="uk-grid uk-grid-large">
                        <div class="uk-width-1-2">
                            <div class="time-left uk-text-right">
                                <div class="time-container">
                                    <div class="time-start">16:30 - 17:30</div>
                                    <div class="time-img uk-flex uk-flex-right">
                                      <div class="time-img-item img-scaledown"><img src="/resources/img/1.png" alt=""></div>
                                      <div class="time-img-item img-scaledown"><img src="/resources/img/2.png" alt=""></div>
                                      <div class="time-img-item img-scaledown"><img src="/resources/img/3.png" alt=""></div>
                                    </div>
                                </div>
                                <div class="time-container">
                                    <div class="time-start">17:30 - 20:00</div>
                                    <div class="time-img ">
                                        <div class="uk-flex uk-flex-right">
                                            <div class="time-img-item img-scaledown mb30"><img src="/resources/img/4.png" alt=""></div>
                                            <div class="time-img-item img-scaledown mb30"><img src="/resources/img/5.png" alt=""></div>
                                        </div>
                                        <div class="uk-flex uk-flex-right">
                                            <div class="time-img-item img-scaledown mb30"><img src="/resources/img/6.png" alt=""></div>
                                            <div class="time-img-item img-scaledown mb30"><img src="/resources/img/7.png" alt=""></div>
                                        </div>
                                    </div>
                                </div>
                                <div class="time-container">
                                    <div class="time-start">20:00 to Closing</div>
                                    <div class="time-img uk-flex uk-flex-right">
                                      <div class="time-img-item img-scaledown"><img src="/resources/img/8.png" alt=""></div>
                                      <div class="time-img-item img-scaledown"><img src="/resources/img/9.png" alt=""></div>
                                      <div class="time-img-item img-scaledown"><img src="/resources/img/10.png" alt=""></div>
                                    </div>
                                </div>

                            </div>
                        </div>
                        <div class="uk-width-1-2">
                            <div class="time-right uk-text-left">
                                <div class="time-container">
                                    <h4 class="title"><span>Welcome</span></h4>
                                    <ul class="uk-clearfix timelist">
                                        <li>Traditional games for kids</li>
                                        <li>Photo booth activities</li>
                                        <li>Company highlights review</li>
                                    </ul>
                                </div>
                                <div class="time-container">
                                    <h4 class="title"><span>Appreciation & Party Program</span></h4>
                                    <ul class="uk-clearfix timelist time-list-2" style="padding-top: 24px;padding-bottom:24px;">
                                        <li>Openning performance.</li>
                                        <li>2023 activities summary & orientation of 2024 work programs.</li>
                                        <li>Employee appreciation for individuals.</li>
                                        <!-- <li>Art performance contest with theme: "Cuu Long đạp gió, rẽ sóng".</li> -->
                                        <li>Toasting moment - Party Announcement.</li>
                                        <li>Celebration party.</li>
                                        <li>Magic performances and games.</li>
                                        <!-- <li>Speech from appreciated employee.</li> -->
                                        <li>Game shows/ Magic and circus shows for family.</li>
                                        <!-- <li>Art performance contest result.</li> -->
                                        <li>Lucky Draw.</li>
                                    </ul>
                                </div>
                                <div class="time-container">
                                    <h4 class="title"><span>After Party Session </span></h4>
                                    <ul class="uk-clearfix timelist">
                                        <li>Karaoke, Song performance, Art performance ...</li>
                                    </ul>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="contact">
                    <div class="uk-grid uk-grid-large">
                        <div class="uk-width-small-1-1 uk-width-medium-1-2">
                            <div class="content ">
                                <h2 class="title uk-text-center"><span>Join Us?</span></h2>
                                <div class="description uk-text-center">Cuu Long Year End Party will be more complete <br> with your presence. Please confirm your  <br> participation so we can welcome you and your  <br> family in the most thorough way! <br> <br> Please register before 03-Jan-23.</div>
                                <div class="uk-text-right">Best regards.</div>
                            </div>
                        </div>
                        <div class="uk-width-small-1-1 uk-width-medium-1-2">
                            <form action="" method="post" style="padding-left: 50px" id="formRequest">
                                <div class="uk-grid uk-grid-medium">
                                    <div class="uk-width-1-2 mb20">
                                        <div class="form-field">
                                            <label for="">Staff name</label>
                                            <input type="text" required id="name" name="name">
                                        </div>
                                    </div>
                                    <div class="uk-width-1-2 mb20">
                                        <div class="form-field">
                                            <label for="">Title</label>
                                            <div>
                                                <input type="text" required id="title" name="title">
                                            </div>
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
                                                </select>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="uk-width-1-2 ">
                                        <div class="form-field">
                                            <label for="">Telephone</label>
                                            <input type="text" required id="phone"  name="phone">
                                        </div>
                                    </div>
                                </div>
                                <div class="form-field mt20 mb30">
                                    <label for="">
                                        Dependants name - Year of birth
                                        <!-- <div class="description-dob" id="descriptionDob">
                                            Ex:
                                            <br> 
                                            cao nhat minh - 1989
                                            <br>
                                            cao nhat minh - 1990
                                        </div> -->
                                    </label>
                                    <textarea cols="30" rows="10" name="dependants" id="dependants" placeholder="Example: Cao Van A - 1989"></textarea>
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
                                <button class="btn-submit" id="submitBtn">Register</button>
                            </form>
                        </div>
                    </div>
                </div>
            </div>
            <div class="bot-bg">
                <img src="/resources/img/b3.png" alt="">
                <img src="/resources/img/b4.png" alt="">
                <img src="/resources/img/b5.png" alt="">
            </div>
            <audio autoplay style="display:none" loop>
                <source src="" type="audio/mpeg">
            </audio>
        </div>
        <script src="node_modules/toastify-js/src/toastify.js"></script>
        <script>
            document.getElementById('dependants').addEventListener('input', function() {
                var dependantsTextarea = document.getElementById('dependants');
                var descriptionDob = document.getElementById('descriptionDob');

                if (dependantsTextarea && descriptionDob) {
                    dependantsTextarea.addEventListener('input', function() {
                        if (dependantsTextarea.value.trim() !== '') {
                            descriptionDob.style.display = 'none';
                        } else {
                            descriptionDob.style.display = 'block';
                        }
                    });
                } else {
                    console.log('Could not find elements with the specified IDs');
                }
            });


            document.getElementById('submitBtn').addEventListener('click', function (event) {
                event.preventDefault();
                if (validateForm()) {
                    sendData();
                }
            });

            function validateForm() {
                var name = document.getElementById('name').value.trim();
                var title = document.getElementById('title').value.trim();
                var department = document.getElementById('department').value.trim();
                var phone = document.getElementById('phone').value.trim();
                // var answerYes = document.getElementById('opt1').checked;
                // var answerNo = document.getElementById('opt2').checked;

                 if (name === '') {
                    alert('Please enter Staff name.');
                    return false;
                }

                if (title === '') {
                    alert('Please enter Title.');
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
                xhr.onload = function () {
                    if (xhr.status >= 200 && xhr.status < 300) {
                        handleData(xhr.responseText);
                    } else {
                        console.error('Error submitting form:', xhr.statusText);
                    }
                };

                // Set up event handler for error
                xhr.onerror = function () {
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
                    var form = document.getElementById('formRequest')
                    // Handle success
                    for (var i = 0; i < form.elements.length; i++) {
                        var element = form.elements[i]
                        if(element.tagName === "INPUT" || element.tagName === "SELECT" || element.tagName === "TEXTAREA"){
                            element.value = ""
                        }
                    }
                    Toastify({
                        text: data.message,
                        duration: 3000,
                        newWindow: true,
                        close: true,
                        gravity: "top",
                        position: "right",
                        backgroundColor: "linear-gradient(to right, #00b09b, #96c93d)",
                    }).showToast();
                } else {
                    Toastify({
                        text: data.message,
                        duration: 3000,
                        newWindow: true,
                        close: true,
                        gravity: "top",
                        position: "right",
                        backgroundColor: "linear-gradient(to right, #e74c3c, #c0392b)",
                    }).showToast();
                }
            }
        </script>
    </body>
</html>
