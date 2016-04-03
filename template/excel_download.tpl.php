<?php include 'header.tpl.php';?>
<div class="pad_10">

<div class="table-list">
<form method="post" id="myform" name="myform" >
    <table width="100%" cellspacing="0">
        <thead>
            <tr>
            <th width="5%"  align="left">
            <input type="checkbox" value="" name="check_box" onclick="selectall('filenames[]');"></th>
            <th width="5%" ><?php echo '序号'?></th>
            <th width="30%"><?php echo '文件名称'?></th>
            <th width="10%"><?php echo '文件大小'?></th>
            <th width="15%"><?php echo '创建时间'?></th>
            <th width="15%"><?php echo '操作'?></th>
            </tr>
        </thead>
    <tbody>
 <?php
$index = 1;
if(is_array($infos)){
	foreach($infos as $info){
?>
	<tr>
	<td width="5%">
    <input type="checkbox" name="filenames[]" value="<?php echo $info['filename']?>">
	</td>
    <td  width="5%" align="center"><?php echo $index++?></td>
	<td  width="30%" align="left"><?php echo $info['filename']?></td>
	<td width="10%" align="center"><?php echo $info['filesize']?></td>
	<td width="15%" align="center"><?php echo $info['maketime']?></td>
	<td width="15%" align="center">
	<a href='downloadExcel.php?filename=<?php echo $info['filename']?>' target="_blank">下载Excel文件</a>
	</td>
	</tr>
<?php
	}
}
?>
    </tbody>
    </table>
<div class="btn">
<label for="check_box"><?php echo '全选'?>/<?php echo '取消'?></label>
<input type="submit" class="button" name="dosubmit" value="删除Excel文件" onclick="document.myform.action='deleteExcel.php';return confirm('确定删除这些Excel文件吗')"/>
</div>
</form>
</div>
</div>

</body>
</html>
<script type="text/javascript">

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
