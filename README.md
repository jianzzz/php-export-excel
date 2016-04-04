
#php-export-excel
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This program simplifies the process of exporting excel, while providing access and download operation.  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;这个项目简化了excel导出的过程，同时提供了查看列表和下载操作。

##English description  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Suppose you have the following requirements:  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Your system involves a lot of form, you need to have a one-touch button to export all the data into Excel tables.
One-click-export mainly related to three processes: ***data query***, ***data re-organization*** and ***Excel exports***.
Suppose a two-dimensional array which stores the key - value relationship is the result of multi-table-joint queries (or single-table queries).  

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Therefore, we need to solve these problem:  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1.We need to ***map the column-key*** to its Chinese meaning when exporting Excel, because designers often use English or phoneticize to represent data.  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2.We need to ***remove the surplus fields*** that we don't need to export.  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3.We may ***reorganize the field contents*** with other information, and then export it.  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;In addition, perhaps you want to directly ***set the number of rows that each Sheet table stores***.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;In this project, an one-dimensional array ***$header*** stores the fields you want to export, its key => value stands for "DB Table Field => the Sheet table header." A two-dimensional array ***$content*** may represent multiple rows of the query data, $value = $content[$i][$key], $value will be exported if and only if $key is present in the array $header.*** $extraConten***t is a two-dimensional array, in principle, it has a consistent number with $content in the first dimension. For each row of the exported data, if $extraContent exist the corresponding key, not taken from the $content.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;In order to solve the "directly set the number of rows that each Sheet table stores" issue, we used https://github.com/mk-j/PHP_XLSXWriter source and extend it to export row by row instead the entire two-dimensional array for exported. Benefits of doing so is to easily determine whether the number of rows exceeds the limit of one's own Sheet table settings, and to create a new sheet form for storage when needed. This requires us to change function finalizeSheet() from private to public, which is defined in xlsxwriter.class.php file from https://github.com/mk-j/PHP_XLSXWriter.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Thus, data organization and export process are separated. You can query every 100 rows of db table data, then organize them, and transfer $header, $content, $extraContent to the export function exportExcelContent(). Function exportExcelContent() will export row by row, and calculates if the rows number is beyond the limit, if so, the new sheet is then stored. Each form exported as an Excel file.

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;In addition, we give a review page and a download page, view page link is: http://localhost/php-export-excel/formaction/index.php , download page link is：http://localhost/php-export-excel/formaction/excelListShow.php.

##Example
```php
    public function Test1()
    {
        $this->clearSheetRecord();
        $filePath = $this->produceFilePath('测试-extraContent为空');
        $this->do_test1($filePath);
    }  
```
```php
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
```
```php
    public function Test2()
    {
        $this->clearSheetRecord();
        $filePath = $this->produceFilePath('测试-extraContent不为空');
        $this->do_test2($filePath);
    } 
```
```php
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
```
```php
    public function Test3()
    {
        $this->clearSheetRecord();
        $filePath = $this->produceFilePath('测试-Sheet行数-50行');
        $this->do_test3($filePath);
    } 
```
```php
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
```

##以下是中文介绍：
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;假设您有以下的需求：  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;您的系统上涉及了很多表单，您需要有一个按钮可以一键导出所有数据到Excel表中。  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;一键导出主要涉及到三个过程：数据表查询、数据重新组织、导出Excel。  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;假设多表联合查询（或单表查询）的结果是二维array，array存储了键-值关系。  

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;因此，我们需要解决的问题是：  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1、需要将键映射为中文，因为设计者通常使用英文或拼音来表示数据表字段，导出为Excel时我们需要将其转为中文意思。  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2、将不需要导出的多余字段去除。  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3、根据其他信息重新组织字段内容，再进行导出。  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;另外，也许您希望能直接设置每个导出的Sheet表存储的行数。  

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本项目中，$header一维数组存储需要导出的字段，key=>value代表“数据表字段=>Sheet表首行的标头”。$content二维数组可代表查询出的多行数据表数据，$value = $content[$i][$key]，当且仅当$key存在于$header数组中时，$value才会被导出。$extraContent二维数组原则上与$content二维数组维数一致，对于每一行导出的数据，如果$extraContent中存在相应的key，则不从$content中取。

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;为了解决“直接设置每个导出的Sheet表存储的行数”的问题，我们使用了 https://github.com/mk-j/PHP_XLSXWriter 的源码，并将其扩展为按行导出，而不是将整个二维数组进行导出。这样做的好处是可以在按行导出过程中判断是否超出了己方设置的Sheet表行数限制，超出时自行启用新的sheet表进行存储。这需要将 https://github.com/mk-j/PHP_XLSXWriter 里xlsxwriter.class.php文件的finalizeSheet方法由私有改为共有。

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;由此，数据组织与导出过程被分离开来，您可以每次查询数据表100行数据后进行组织，并将$header、$content、$extraContent传递到导出函数exportExcelContent，exportExcelContent将按行导出，并计算是否超出行数限制，是则新建sheet表继续存储。每种表单导出为一个Excel文件。

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;另外提供了查看页面和下载页面，查看页面链接是：http://localhost/php-export-excel/formaction/index.php ,下载页面链接是：http://localhost/php-export-excel/formaction/excelListShow.php.
