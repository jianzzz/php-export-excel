<?php
@set_time_limit(0);
if (!defined('DIRECTORY_SEPARATOR')) define('DIRECTORY_SEPARATOR', '/');
if (!defined('BASE_PATH')) define('BASE_PATH',str_replace('\\',DIRECTORY_SEPARATOR,realpath(dirname(__FILE__).DIRECTORY_SEPARATOR)).DIRECTORY_SEPARATOR);
if (!defined('EXCELDIRPATH')) define('EXCELDIRPATH',BASE_PATH . 'db2excelDownload' . DIRECTORY_SEPARATOR);

class data2export
{
    private $excelWriter;
    private $sheetRowLimit = 65535;   //excel行数限制
    private $sheetRowCount = 0; //标记当前excel的sheet已经添加了多少行
    private $activeSheetIndex = 0;
    private $perQueryCount = 100;//每次查询的条数
    private $sheetBaseName = "Sheet";
    private $currentSheetName;
    private $prefix;//excel前缀
    private $suffix;//excel后缀
    function __construct()
    {
        include_once 'phpxlsxwriter'.DIRECTORY_SEPARATOR.'xlsxwriter.class.php';
        $this->excelWriter = new XLSXWriter();
        $this->prefix = date('Y-m-d.H.i.s').".";
        $this->suffix = '.xlsx';
        if (!file_exists(EXCELDIRPATH)) mkdir(EXCELDIRPATH);
    }

    function init(){
        $adminRight = array();
        $adminRight["测试-extraContent为空"]="Test1";
        $adminRight["测试-extraContent不为空"]="Test2";
        $adminRight["测试-Sheet行数-50行"]="Test3";
        include 'template'.DIRECTORY_SEPARATOR.'excel_akeyexport.tpl.php';
    }
    /**
     * 一键导出表单
     */
    public function aKeyExport()
    {
        $excelLine = $_POST['excelLine'];
        if(isset($excelLine))
            $this->sheetRowLimit = (int)($excelLine);
        else
            $this->sheetRowLimit = 65535;
        $adminRight = $_POST['adminRight'];
        if(!isset($adminRight)){
            var_dump("请先选择需要导出的表");
            include 'index.php';
            return;
        }
        foreach($adminRight as $ar){
            $this->$ar();
        }
        header("Location: excelListShow.php", true, 302);
    }

    /**
     * 导出Excel的列表
     */
    public function excelListShow(){
        $info = $infos = array();
        $dir = realpath(EXCELDIRPATH);
        $dir = $dir.DIRECTORY_SEPARATOR;
        $oDir = @opendir(EXCELDIRPATH);
        while($fileName = readdir($oDir)) {
            if (!(strcmp(basename($fileName), ".") == 0 || strcmp(basename($fileName), "..") == 0)) {
                $fullPath = $dir . $fileName;
                if (is_file($fullPath)) {
                    //basename会忽略文件名最开始的中文字段，本文件使用时间作为文件前缀
                    $info['filename'] = basename($fullPath);
                    //实验证明，excel表名需要转换为UTF-8字符才能成功显示
                    $info['filename'] = iconv('GB2312', 'UTF-8', $info['filename']);
                    $info['filesize'] = $this->sizecount(filesize($fullPath));
                    $info['maketime'] = date('Y-m-d H:i:s', filemtime($fullPath));
                    $infos[] = $info;
                }
            }
        }
        include 'template'.DIRECTORY_SEPARATOR.'excel_download.tpl.php';
    }

    /**
     * Excel下载
     */
    public function download(){
        $filename = trim($_GET['filename']);
        $this->excel_down(EXCELDIRPATH.$filename);
    }

    /**
     * 文件下载
     * @param $filepath 文件路径
     * @param $filename 文件名称
     */
    private function excel_down($filepath, $filename = '') {
        if(!$filename) $filename = basename($filepath);
        if($this->is_ie()) $filename = rawurlencode($filename);
        $filetype = $this->fileext($filename);
        $filesize = sprintf("%u", filesize($filepath));
        if(ob_get_length() !== false) @ob_end_clean();
        header('Pragma: public');
        header('Last-Modified: '.gmdate('D, d M Y H:i:s') . ' GMT');
        header('Cache-Control: no-store, no-cache, must-revalidate');
        header('Cache-Control: pre-check=0, post-check=0, max-age=0');
        header('Content-Transfer-Encoding: binary');
        header('Content-Encoding: none');
        header('Content-type: '.$filetype);
        header('Content-Disposition: attachment; filename="'.$filename.'"');
        header('Content-length: '.$filesize);
        ob_clean(); //!!!!!!!!!important, or it maybe cause messy code
        flush();
        readfile($filepath);
        exit;
    }
    /**
     * Excel删除
     */
    public function delete(){
        $filenames = $_POST['filenames'];
                var_dump($filenames);
        $dir = realpath(EXCELDIRPATH);
        $dir = $dir.DIRECTORY_SEPARATOR;
        if($filenames) {
            if(is_array($filenames)) {
                foreach($filenames as $filename) {
                    //实验证明，删除需要转换为GB2312字符才能成功删除
                    $fullPath = $dir.$filename;
                    @unlink(iconv('UTF-8','GB2312',$fullPath));
                }
            } else {
                //实验证明，删除需要转换为GB2312字符才能成功删除
                $fullPath = $dir .$filenames;
                @unlink(iconv('UTF-8','GB2312',$fullPath));
            }
        } else {
            var_dump('请先选择删除的excel');
        }
        header("Location: excelListShow.php", true, 302);
    }

    /**************************************************************************************************/

    public function Test1()
    {
        $this->clearSheetRecord();
        $filePath = $this->produceFilePath('测试-extraContent为空');
        $this->do_test1($filePath);
    }

    private function do_test1($filePath)
    {
        $header = array();
        $header["id"] = "ID";
        $header["username"] = "用户名";
        $header["email"] = "邮箱";
        $header["phone"] = "手机";
        //输入新的表头
        $this->setSheetName($this->sheetBaseName, $this->activeSheetIndex);
        $this->exportExcelHeader($header, $this->currentSheetName);

        $content = array(
            array("id"=>'01',"username"=>'user1',"email"=>'12345678@163.com',"phone"=>'12345678',"school"=>'华南理工大学'),
            array("id"=>'02',"username"=>'user2',"email"=>'12345678@163.com',"phone"=>'12345678',"school"=>'华南理工大学'),
            array("id"=>'03',"username"=>'user3',"email"=>'12345678@163.com',"phone"=>'12345678',"school"=>'华南理工大学'),
        );
        $extraContent = array();
        //如果header的数据可以从$extraContent里面取，则不从$content中取，此处extraContent为空，
        //另，$content中含有school的字段，但不会被导出，因为$header中没有定义它
        $this->exportExcelContent($content, $header, $extraContent);
        unset($extraContent);
        $this->writeToFile($filePath);
        unset($header);
    }

    public function Test2()
    {
        $this->clearSheetRecord();
        $filePath = $this->produceFilePath('测试-extraContent不为空');
        $this->do_test2($filePath);
    }

    private function do_test2($filePath)
    {
        $header = array();
        $header["id"] = "ID";
        $header["username"] = "用户名";
        $header["email"] = "邮箱";
        $header["phone"] = "手机";
        //输入新的表头
        $this->setSheetName($this->sheetBaseName, $this->activeSheetIndex);
        $this->exportExcelHeader($header, $this->currentSheetName);

        $content = array(
            array("id"=>'01',"username"=>'user1',"email"=>'12345678@163.com',"phone"=>'12345678',"school"=>'华南理工大学'),
            array("id"=>'02',"username"=>'user2',"email"=>'12345678@163.com',"phone"=>'12345678',"school"=>'华南理工大学'),
            array("id"=>'03',"username"=>'user3',"email"=>'12345678@163.com',"phone"=>'12345678',"school"=>'华南理工大学'),
        );
        $extraContent = array();
        $extraContent = array(
            array("id"=>'0001',"username"=>'user1111'),
            array("id"=>'0002',"username"=>'user2222'),
            array("id"=>'0003',"username"=>'user3333'),
        );
        //如果header的数据可以从$extraContent里面取，则不从$content中取，此处extraContent改变了id和username字段
        //另，$content中含有school的字段，但不会被导出，因为$header中没有定义它
        $this->exportExcelContent($content, $header, $extraContent);
        unset($extraContent);
        $this->writeToFile($filePath);
        unset($header);
    }

    public function Test3()
    {
        $this->clearSheetRecord();
        $filePath = $this->produceFilePath('测试-Sheet行数-50行');
        $this->do_test3($filePath);
    }

    private function do_test3($filePath)
    {
        $header = array();
        $header["id"] = "ID";
        $header["username"] = "用户名";
        $header["email"] = "邮箱";
        $header["phone"] = "手机";
        //输入新的表头
        $this->setSheetName($this->sheetBaseName, $this->activeSheetIndex);
        $this->exportExcelHeader($header, $this->currentSheetName);

        $content = array();
        for($i=0;$i<100;$i++){
            $content[] = array("id"=>$i,"username"=>'user'.$i,"email"=>'12345678@163.com',"phone"=>'12345678',"school"=>'华南理工大学');
        }
        $extraContent = array();
        for($i=0;$i<50;$i++){
            $extraContent[] = array("id"=>'0001',"username"=>'user1111');
        }
        //如果header的数据可以从$extraContent里面取，则不从$content中取，此处extraContent改变了前50行数据的id和username字段
        //另，$content中含有school的字段，但不会被导出，因为$header中没有定义它
        $this->sheetRowLimit = 50;//测试每个sheet存50行数据
        $this->exportExcelContent($content, $header, $extraContent);
        unset($extraContent);
        $this->writeToFile($filePath);
        unset($header);
    }


    /**************************************************************************************************/
    private function clearSheetRecord()
    {
        $this->sheetRowCount = 0;
        $this->activeSheetIndex = 0;
        unset($this->excelWriter);
        $this->excelWriter = new XLSXWriter();
    }

    //设置sheet名字
    private function setSheetName($baseName, $sheetIndex)
    {
        $this->currentSheetName = $baseName . $sheetIndex;
    }

    //输出行头
    private function exportExcelHeader($header, $sheetName)
    {
        $numFields = count($header);
        $headerKey = array_keys($header);
        $excelHeader = array();
        for ($j = 0; $j < $numFields; $j++) {
            $excelHeader[] = $header[$headerKey[$j]];
        }
        $this->excelWriter->writeSheetRow($sheetName, $excelHeader);
    }

    //输出内容
    private function exportExcelContent($infos, $header, $extraContent)
    {
        $arrCount = count($infos);
        $numFields = count($header);
        $headerKey = array_keys($header);
        $excelContent = array();
        $ifCheckExtraContent = isset($extraContent) && (count($extraContent)>0);

        for ($i = 0; $i < $arrCount; $i++) {
            $tempArr = array();
            for ($k = 0; $k < $numFields; $k++) {
                if ($ifCheckExtraContent && isset($extraContent[$i]) && array_key_exists($headerKey[$k], $extraContent[$i]))
                    $tempArr[] = $extraContent[$i][$headerKey[$k]];
                else $tempArr[] = $infos[$i][$headerKey[$k]];
            }
            $excelContent[] = $tempArr;
            unset($tempArr);
        }

        foreach ($excelContent as $i => $row) {
            if ($this->sheetRowCount >= $this->sheetRowLimit) {
                //超出限制，结束sheet输入
                $this->sheetRowCount = 0;
                $this->excelWriter->finalizeSheet($this->currentSheetName);
                $this->activeSheetIndex++;
                //输入新的表头
                $this->setSheetName($this->sheetBaseName, $this->activeSheetIndex);
                $this->exportExcelHeader($header, $this->currentSheetName);
            }
            $this->sheetRowCount++;
            $this->excelWriter->writeSheetRow($this->currentSheetName, $row);
        }
        unset($excelContent);
    }

    //保存到文件
    private function writeToFile($filePath)
    {
        ob_clean();//清除缓冲区
        $this->excelWriter->writeToFile($filePath);
    }

    //UTF-8转GB2312
    private function convertEnToCz($name){
        return iconv('UTF-8','GB2312',$name);
    }

    //构造文件完整路径
    private function produceFilePath($name){
        $this->prefix = date('Y-m-d.H.i.s').".";
        $fileName =  $this->prefix .$name .  $this->suffix;
        //实验证明，excel表名需要转换为GB2312字符才能成功保存文件名
        $fileName = $this->convertEnToCz($fileName);
        $filePath = EXCELDIRPATH . $fileName;
        return $filePath;
    }

    /**
    * 转换字节数为其他单位
    *
    *
    * @param    string  $filesize   字节大小
    * @return   string  返回大小
    */
    private function sizecount($filesize) {
        if ($filesize >= 1073741824) {
            $filesize = round($filesize / 1073741824 * 100) / 100 .' GB';
        } elseif ($filesize >= 1048576) {
            $filesize = round($filesize / 1048576 * 100) / 100 .' MB';
        } elseif($filesize >= 1024) {
            $filesize = round($filesize / 1024 * 100) / 100 . ' KB';
        } else {
            $filesize = $filesize.' Bytes';
        }
        return $filesize;
    }

    /**
    * IE浏览器判断
    */
    private function is_ie() {
        $useragent = strtolower($_SERVER['HTTP_USER_AGENT']);
        if((strpos($useragent, 'opera') !== false) || (strpos($useragent, 'konqueror') !== false)) return false;
        if(strpos($useragent, 'msie ') !== false) return true;
        return false;
    }
    /**
     * 取得文件扩展
     *
     * @param $filename 文件名
     * @return 扩展名
     */
    function fileext($filename) {
        return strtolower(trim(substr(strrchr($filename, '.'), 1, 10)));
    }

}

?>
