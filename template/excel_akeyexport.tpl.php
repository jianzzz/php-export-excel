<?php include 'header.tpl.php';?>
<div class="pad_10">

<div class="table-list">
<form method="post" id="myform" name="myform" action="aKeyExport.php">
    <table width="100%" cellspacing="0">
        <thead>
            <tr>
            <th width="5%"  align="center">
            <input type="checkbox" value="" name="check_box" onchange="selectall('adminRight[]');"></th>
            <th width="10%"  align="center">序号</th>
            <th align="left" width="80%">名称</th>
            </tr>
        </thead>
    <tbody>
    <tr>
        <td align="center">*</td>
        <td align="center">*</td>
        <td align="left">
            <input type="submit" onclick="exportExcel()" name="dosubmitExcel" value="一键导出Excel" class="button">
            Excel行数：<input type="text" id="excelLine" name="excelLine" value="65535">
            <span style="display: none" id="imgLoading" >请稍候，正在为您导出...&nbsp;&nbsp;
            <img src="../template/static/image/excel_loading.gif"/></span>
        </td>
    </tr>
    <?php
        $index = 1;
        if(is_array($adminRight)) {
            foreach ($adminRight as $key => $value) {
                ?>
                <tr>
                    <td align="center">
                        <input type="checkbox" name="adminRight[]" value="<?php echo $value ?>">
                    </td>
                    <td align="center"><?php echo $index++ ?></td>
                    <td align="left"><?php echo $key ?></td>
                </tr>
                <?php
            }
        }
    ?>
    </tbody>
    </table>
</form>
</div>
</div>

</body>
</html>
<script type="text/javascript">
    function exportExcel(){
        $('#imgLoading').show();
    }

    /**
     * 全选checkbox,注意：标识checkbox id固定为为check_box
     * @param string name 列表check名称,如 uid[]
     */
    function selectall(name) {
        if ($("input[name=check_box]")[0].checked==false) {
            $("input[name='"+name+"']").each(function() {
                $(this).prop("checked",false);
            });
        } else {
            $("input[name='"+name+"']").each(function() {
                $(this).prop("checked",true);
            });
        }
    }
</script>
