<?php 

    error_reporting(E_ALL); 
     
    $err=""; 
     
    //Extract vbaProject.bin from a modern office file. 
    function getFromZip($fName){ 
        global $err; 
        $temp=""; 
        $zip=new ZipArchive(); 
         
        //Try to open the file as zip archive 
        if($res=$zip->open($fName) && $zip->numFiles>0){ 
            //try to figure out, where to extract vbaProject.bin from (excel,word,powerpoint) 
            if(($temp=$zip->getFromName("xl/vbaProject.bin"))==FALSE){ 
                if(($temp=$zip->getFromName("word/vbaProject.bin"))==FALSE){ 
                    if(($temp=$zip->getFromName("ppt/vbaProject.bin"))==FALSE){ 
                        $err="Không có mã VBA trong file được tải lên. Chương trình này chỉ loại bỏ bảo vệ VBA chứ không phải bảo vệ tài liệu (mật khẩu để mở hoặc thay đổi nội dung)";
                    } 
                } 
            } 
            $zip->close(); 
        } 
        else{ 
            $err="Không thể mở file của bạn. Nó có vẻ là từ Office 2007 rất cũ hoặc mới hơn nhưng file bị lỗi.."; 
        } 
        //return the vbaProject.bin content if it has been extracted. 
        return $temp===FALSE?"":$temp; 
    } 

    //Add vbaProject.bin back to a modern Office file 
    function addToZip($contents,$fName) 
    { 
        global $err; 
        $temp=""; 
        $zip=new ZipArchive; 
        //Open file as zip archive 
        if($res=$zip->open($fName)) 
        { 
            //Try to find where the original vbaProject.bin was located and replace it with decrypted blob. 
            if($zip->getFromName("xl/vbaProject.bin")==FALSE) 
            { 
                if($temp=$zip->getFromName("word/vbaProject.bin")==FALSE){ 
                    //Powerpoint 
                    $zip->deleteName("ppt/vbaProject.bin"); 
                    $zip->addFromString("ppt/vbaProject.bin",$contents); 
                } 
                else{ 
                    //Word 
                    $zip->deleteName("word/vbaProject.bin"); 
                    $zip->addFromString("word/vbaProject.bin",$contents); 
                } 
            } 
            else{ 
                //Excel 
                $zip->deleteName("xl/vbaProject.bin"); 
                $zip->addFromString("xl/vbaProject.bin",$contents); 
            } 
            $zip->close(); 
        } 
        else{ 
            $err="Không thể mở tệp để thay đổi cài đặt VBA. Thông báo cho người sở hữu."; 
        } 
    } 
     
    //provides file source code upon request 
    if(isset($_GET["source"])){ 
        echo '<html><head><meta http-equiv="X-UA-Compatible" content="IE=edge" /><meta name="viewport" content="width=device-width, initial-scale=1" /><title>VBA Unlocker Source</title></head><body>'; 
        highlight_file(__FILE__); 
        echo "</body></html>"; 
        exit(0); 
    } 
     
    //Check if file uploaded 
    if(isset($_FILES['excel'])){ 
        //Ensure a file name exists 
        if(isset($_FILES['excel']['tmp_name']) && $_FILES['excel']['tmp_name']!=""){ 
            //Try to read the file 
            if($fp=fopen($_FILES['excel']['tmp_name'],"rb")){ 
                //Read everything and close file handle 
                $contents=fread($fp,filesize($_FILES['excel']['tmp_name'])); 
                fclose($fp); 

                //If it starts with "PK" it is a modern file (O2007 and newer) 
                if(substr($contents,0,2)=="PK"){ 
                    //Create temporary zip file name 
                    $z=tempnam(dirname(__FILE__)."/TMP/","zip"); 
                    //Move temporary file to place where it's not readonly to us. 
                    move_uploaded_file($_FILES['excel']['tmp_name'], $z); 
                    //Get VBA blob from zip 
                    $contents=getFromZip($z); 
                    if($contents!="" && $err==""){ 
                        if(strpos($contents,"DPB=")===FALSE){ 
                            $err="Chúng tôi đã tìm thấy Mã VBA nhưng nó không được đặt mật khẩu."; 
                        } 
                        else{ 
                            $contents=str_replace("DPB=","DPx=",$contents); 
                            addToZip($contents,$z); 
                            if($err=="") 
                            { 
                                if($fp=fopen($z,"rb")){ 
                                    header("Content-Type: application/octet-stream"); 
                                    header("Content-Disposition: attachment; filename=\"" . $_FILES['excel']['name'] . "\""); 
                                    echo fread($fp,filesize($z)); 
                                    fclose($fp); 
                                    //Delete the uploaded file on success 
                                    unlink($z); 
                                    exit(0); 
                                } 
                                else{ 
                                    $err="Không thể gửi lại file office. Lỗi khi mở tệp tạm thời."; 
                                } 
                            } 
                        } 
                    } 
                    else{ 
                        $err="File này không mã hóa dự án VBA, hoặc file yêu cầu mật khẩu để mở file.."; 
                    } 
                    //Delete the uploaded file on error 
                    unlink($z); 
                } 
                else{ 
                    //Delete uploaded file because it's in $contents now 
                    unlink($_FILES['excel']['tmp_name']); 
                     
                    //assume classic file (O2003 and older) 
                    if(strpos($contents,"DPB=")===FALSE){ 
                        $err="There is no VBA code or it is not protected."; 
                    } 
                    else{ 
                        //This removes the protection 
                        $contents=str_replace("DPB=","DPx=",$contents); 
                        //Send back file 
                        header("Content-Disposition: attachment; filename=\"" . $_FILES['excel']['name'] . "\""); 
                        header("Content-Type: application/octet-stream"); 
                        echo $contents; 
                        exit(0); 
                    } 
                } 
            } 
            else{ 
                $err="Chúng tôi không thể mở file. Bộ nhớ của chúng tôi đã đầy hoặc nó đã bị chương trình Chống vi-rút của chúng tôi xóa."; 
            } 
        } 
        else{ 
            $err="Không nhận được file nào. Vui lòng chọn một file để giải mã"; 
        } 
    } 
?>
<!DOCTYPE html>
<html lang="en" class="h-100">
<head>
	<meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <link rel="icon" href="../passexcel/favicon1.ico"> 
    <meta name="description" content="Một công cụ đơn giản có thể giúp xóa mật khẩu khỏi Trang tính Excel.">
	<title>Xóa Mật Khẩu Trang Tính Excel</title>
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.0/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-KyZXEAg3QhqLMpG8r+8fhAXLRk2vvoC2f3B09zVXn8CA5QIVfZOJ3BCsw2P0p/We" crossorigin="anonymous">
	<link href="./css/styles.css" rel="stylesheet">
	<!-- Prevents some weird errors caused by caching -->
	<meta http-equiv="cache-control" content="max-age=0" />
	<meta http-equiv="cache-control" content="no-cache" />
	<meta http-equiv="expires" content="0" />
	<meta http-equiv="expires" content="Tue, 01 Jan 1980 1:00:00 GMT" />
	<meta http-equiv="pragma" content="no-cache" />
</head>
<body class="d-flex h-100 text-center text-white bg-dark">
	<div class="cover-container d-flex w-150 h-150 p-3 mx-auto flex-column">
		<header class="mb-auto">
		</header>
		<main class="px-4">
			<div id="warning">
				<h2>Lời Nói Đầu</h2>
				<p class="lead">
					Công cụ này được thiết kế để sử dụng với các tệp mà bạn có quyền xóa mật khẩu.<br>
					Tôi chỉ cung cấp một công cụ chung chung, bạn là người quyết định bạn làm gì với nó và chịu trách nhiệm khi sử dụng nó!<br><br>
					<a href="#" class="btn btn-lg btn-secondary fw-bold border-white bg-white" onclick="acceptTerms();">Đồng Ý Và Tiếp Tục Sử Dụng</a>
				</p>
			</div>
			<div id="file-select" hidden>
				<h2>1. Xóa Mật Khẩu Trang Tính Excel</h2>
				<p class="lead">
					Nhấp vào nút bên dưới để chọn file excel cần mở khóa trang tính (Định dạng:<b>.xslx</b> hoặc <b>.xlsm</b>).<br><br>
					<a href="#" class="btn btn-lg btn-primary fw-bold border-white" onclick="selectFile();">Chọn File Excel</a>
					<!-- Element hidden to only use it via JS -->
					<input id="input-file" type="file" style="display: none"/>
				</p><br>
				<h4><p style="text-align:left;"><span style="font-family: Arial">Hướng Dẫn Xóa Mật Khẩu Trang Tính Excel:</span></p></h4>
					<div class="how-to-paragraph">
						<ol>
							<li><p style="text-align:left;"><span style="font-family: Arial">Đầu tiên, hãy chọn file Excel bị quên mật khẩu cần mở khóa.</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Sau đó xác nhận lại đã chọn đúng file Excel cần mở khóa trang tính chưa?</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Sau khi xác nhận chương trình sẽ xử lý và xóa các mật khẩu trang tính trong file Excel đó.</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Cuối cùng bấm vào "Xóa Mật Khẩu" để tải xuống file Excel đã xóa tất cả mật khẩu trang tính.</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Bấm "Làm Lại" để tiếp tục xóa mật khẩu trang tính file Excel khác nếu bạn cần.</span></p></li>
						</ol>
					</div><br><br>
				<br><h2>2. Xóa Mật Khẩu Dự Án VBA</h2>
				<p class="lead">
					Nhấp vào nút bên dưới để chọn file dự án VBA cần mở khóa (Office:<b>Word</b>,<b>Excel</b>,<b>Powerpoint</b>).<br>
				<?php if($err){ echo "<div class='alert alert-danger'>$err</div>";} ?> 
				<form method="post" action="index.php" enctype="multipart/form-data" class="form-inline"> 
					<label class="control-label">Office File (doc,docm,xls,xlsm,ppt,pptm): 
					<input type="file" name="excel" class="form-control" required /></label><br/><br/>
					<input type="submit" class="btn btn-primary" value="Xóa Mật Khẩu VBA Và Tải Xuống" /> 
				</form>
				</p><br>
				

			
				
				<h4><p style="text-align:left;"><span style="font-family: Arial">Hướng Dẫn Xóa Mật Khẩu Trang Tính Excel:</span></p></h4>
					<div class="how-to-paragraph">
						<ol>
							<li><p style="text-align:left;"><span style="font-family: Arial">Tải lên tài liệu Office có dự án VBA của bạn.</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Sau đó bấm "Xóa Mật Khẩu VBA Và Tải Xuống" để tải xuống file VBA đã mở khóa.</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Mở tài liệu đã tải xuống và nhấn <b>ALT + F11</b>. Xác nhận tất cả các thông báo lỗi có thể xuất hiện.</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial"><b>Không mở dự án VBA</b> , hãy chuyển đến " Tools => VBA Project Properties"..</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Trên tab "Protection", xóa hộp kiểm và mật khẩu.</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Lưu file lại.</span></p></li>
							<li><p style="text-align:left;"><span style="font-family: Arial">Mật khẩu VBA đã bị xóa và bạn có thể xem hoặc thay đổi mã như bình thường.</span></p></li>
						</ol>
					</div>
				
			</div>
			<div id="file-confirm" hidden>
				<h2>Xác Nhận Cần Thiết</h2>
				<p class="lead">
					<span id="file-confirm-text"></span><br><br>
					<a href="#" class="btn btn-lg btn-danger fw-bold border-white" onclick="restart();">Không, Tôi Muốn Chọn File Khác!</a>&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="#" class="btn btn-lg btn-success fw-bold border-white" onclick="startProcessing();">Có, Hãy Tiến Hành Xử Lý File Này</a>
				</p>
			</div>
			<div id="file-process-waiting" hidden>
				<h2>Xin vui lòng chờ trong giây lát...</h2>
				<p class="lead">
					<span id="file-process-waiting-text">
						Tệp đang được phân tích và nếu tìm thấy mật khẩu, bạn sẽ được nhắc xóa chúng.<br><br>
						Nếu quá trình này mất hơn 30 giây hoặc một phút, có thể do lỗi không mong muốn, có nghĩa là công cụ này không thể xử lý tệp của bạn, bạn có thể thử lại.
					</span><br><br>
					<a href="#" class="btn btn-lg btn-danger fw-bold border-white" onclick="restart();">Hủy Tất Cả</a>
				</p>
			</div>
			<div id="file-process-finished" hidden>
				<h2>File Đã Được Phân Tích</h2>
				<p class="lead">
					<span id="file-process-finished-text">File đã được phân tích xong và sẵn sàng xóa mật khẩu.</span><br><br>
					<span id="file-process-finished-subtext">Bằng cách nhấp vào nút <b><i> "Xóa Mật Khẩu" </i> </b>, bạn sẽ tải xuống File mới mà không có mật khẩu trang tính.</span><br><br>
					<a href="#" class="btn btn-lg btn-danger fw-bold border-white" onclick="restart();">Hủy Tất Cả</a>&nbsp;&nbsp;&nbsp;&nbsp;
					<a href="#" class="btn btn-lg btn-success fw-bold border-white" onclick="downloadProcessedFile();">Xóa Mật Khẩu</a>
				</p>
				<!-- Used to download the final file with a specific name -->
				<a id="zip-downloader-tag" href="" hidden></a>
			</div>
			<div id="error" hidden>
				<h2>Có Một Lỗi Xảy Ra...</h2>
				<p class="lead">
					<span id="error-text"></span><br><br>
					<a href="#" class="btn btn-lg btn-warning fw-bold border-white" onclick="restart();">Làm Lại</a>
				</p>
			</div>
			<div id="end" hidden>
				<h2>Hoàn Thành</h2>
				<p class="lead">
					Chúc các bạn một ngày làm việc hiệu quả!<br><br>
					<a href="#" class="btn btn-lg btn-info fw-bold border-white" onclick="restart();">Làm Lại</a>
				</p>
			</div>
		</main>
		<footer class="mt-auto text-white-50">
			<br><br><p>QUAY LẠI <a href="../index.html" class="text-red">TRANG CHỦ</a><br></p>
		</footer>
	</div>
	<!-- Library used to manipulate ZIP files -->
	<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.7.1/jszip.min.js" integrity="sha512-xQBQYt9UcgblF6aCMrwU1NkVA7HCXaSN2oq0so80KO+y68M+n64FOcqgav4igHe6D5ObBLIf68DWv+gfBowczg==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
	<script>
		const excelFileRegex = /^.*\.xls[xm]$/gi;
		const excelWorksheetRegex = /^xl\/worksheets\/.*.xml$/gi;
		
		var outputZip;
		var outputZipFilename = "default-filename.error.zip";
		var filesTotalCount = 0;
		var filesProcessedCount = 0;
		var passwordsRemoved = 0;
		
		function acceptTerms() {
			document.getElementById("warning").hidden = true;
			document.getElementById("file-select").hidden = false;
		}
		
		function selectFile() {
			document.getElementById("input-file").click();
			fileSelectionHandler();
		}
		
		/* Waits for a file to be selected */
		function fileSelectionHandler() {
			if(document.getElementById("input-file").files.length == 0) {
				setTimeout(fileSelectionHandler, 500);
				return;
			} else if(document.getElementById("input-file").files.length > 1) {
				handleError("You selected more than one file !");
			} else {
				//console.log(document.getElementById("input-file").files[0].name);
				if(document.getElementById("input-file").files[0].name.match(excelFileRegex)) {
					document.getElementById("file-select").hidden = true;
					document.getElementById("file-confirm-text").textContent = "Bạn đã chọn một tệp có tên \""+document.getElementById("input-file").files[0].name+"\", bạn có chắc chắn đây là file cần mở khóa trang tính?";
					document.getElementById("file-confirm").hidden = false;
				} else {
					handleError("File bạn vừa chọn có lẽ không phải là file Excel !");
				}
			}
		}
		
		function handleError(errorMessage) {
			// Hiding everything
			document.getElementById("warning").hidden = true;
			document.getElementById("file-select").hidden = true;
			document.getElementById("file-confirm").hidden = true;
			document.getElementById("file-process-waiting").hidden = true;
			document.getElementById("file-process-finished").hidden = true;
			document.getElementById("end").hidden = true;
			
			// Preparing and showing error message
			document.getElementById("error-text").textContent = errorMessage;
			document.getElementById("error").hidden = false;
		}
		
		function restart() {
			location.reload();
		}
		
		function startProcessing() {
			document.getElementById("file-confirm").hidden = true;
			document.getElementById("file-process-waiting").hidden = false;
			// Done this way so as to not block the rendering of the "Please wait text".
			setTimeout(processFile, 100);
		}
		
		function processFile() {
			outputZipFilename = document.getElementById('input-file').files[0].name;
			outputZipExtension = "."+outputZipFilename.split(".").pop();
			outputZipFilename = outputZipFilename.substring(0, outputZipFilename.length - outputZipExtension.length);
			outputZipFilename = outputZipFilename + "_no-password" + outputZipExtension;
			
			JSZip.loadAsync(document.getElementById('input-file').files[0]).then(function(zip) {
				outputZip = new JSZip();
				filesTotalCount = 0;
				filesProcessedCount = 0;
				passwordsRemoved = 0;
				
				for(const[fileKey, fileValue] of Object.entries(zip.files)) {
					filesTotalCount++;
					
					if(fileKey.match(excelWorksheetRegex)) {
						//console.debug("Checking: "+fileKey);
						fileValue.async("string").then(function(fileText) {
							var startIndex = fileText.indexOf('<sheetProtection ');
							
							if(startIndex === -1) {
								// No password found.
								outputZip.file(fileKey, fileText);
								//console.debug("Analysed: "+fileKey);
							} else {
								// Removing the password.
								var endIndex = fileText.indexOf('/>', startIndex) + 2;
								fileText = fileText.replace(fileText.substr(startIndex, endIndex-startIndex), "");
								outputZip.file(fileKey, fileText);
								//console.debug("Processed: "+fileKey);
								passwordsRemoved++;
							}
							
							filesProcessedCount++;
						});
					} else {
						// Other files.
						//console.debug("Ignoring: "+fileKey);
						fileValue.async("string").then(function(fileText) {
							outputZip.file(fileKey, fileText);
							//console.debug("Copied: "+fileKey);
							filesProcessedCount++;
						});
					}
				}
				
				//console.debug("Waiting for all the files to be processed !");
				setTimeout(waitFilesBeingProcessed, 50);
			}, function (e) {
				handleError("Không giải nén được nội dung của file trong trình duyệt! ("+e.message+")");
			});
		}
		
		function waitFilesBeingProcessed() {
			//console.debug("Processed "+filesProcessedCount+" file(s) out of "+filesTotalCount);
			
			if(filesTotalCount != filesProcessedCount) {
				setTimeout(waitFilesBeingProcessed, 50);
			} else {
				//console.debug("Done, now switching the page !");
				if(passwordsRemoved > 0) {
					document.getElementById("file-process-waiting").hidden = true;
					document.getElementById("file-process-finished-text").textContent = "File đã được phân tích, "+ passwordsRemoved +" mật khẩu "+ (passwordsRemoved>1 ?"":"") +" được tìm thấy và sẵn sàng xóa!";
					document.getElementById("file-process-finished").hidden = false;
				} else {
					handleError("Không tìm thấy mật khẩu nào trong tài liệu Excel !");
				}
			}
		}
		
		function downloadProcessedFile() {
			outputZip.generateAsync({type:"base64"}).then(function(base64) {
				link = document.getElementById("zip-downloader-tag")
    			link.download = outputZipFilename;
    			link.href = 'data:application/zip;base64,' + base64;
    			link.click();
				document.getElementById("file-process-finished").hidden = true;
				document.getElementById("end").hidden = false;
			}, function(err) {
				console.error(err);
				handleError("Đã xảy ra lỗi khi tạo file của bạn, vui lòng kiểm tra bảng điều khiển để biết thêm thông tin !");
			});
		}
	</script>
</body>
</html>