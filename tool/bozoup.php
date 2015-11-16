<style type="text/css">
body,table,td{font-size:12px;color:#000090;}
body{margin:0;background: #f5f9fb;}
</style>
<?
$url=$_POST["Udir"];
$Burl=strtr("$url","/","\\");//表单传输的目录
$Surl=$_SERVER['DOCUMENT_ROOT']; //网站根目录
$dir = $Surl.$Burl;  //上传的目录
// $dir = getcwd() 获取当前目录  $Surl.$Burl指定目录
if($_POST["sub"]){//判断点击了提交按钮

 $nname = $_FILES["upfiles"]["name"];//获取上传的文件名称
 $tname = $_FILES["upfiles"]["tmp_name"];//获取上传文件的临时文件名  $_FILES["upfiles"]["tmp_name"]
 
 $fiearr=explode(".",$nname);       //将原文件名分成数组
 $key=count($fiearr)-1;                //计算出最后一个扩展名的主键
 $fie_extend=$fiearr[$key];             //列出上传文件的扩展名
 $fie_extend=strtolower($fie_extend);  //将扩展名统一为小写
 
 if($fie_extend=="zip"||$fie_extend=="rar"||$fie_extend=="7z"||$fie_extend=="iso"){  
 	 $path=$Surl.$Burl.$nname;//定义上传目录
 	 move_uploaded_file($tname,$path);//移动上传文件,在这之前其实文件已经上传成功!此处作一个命名处理而已!此处还是以原来的名称命名文件!
 $obj= new com("wscript.shell");//实例化COM组件
 $obj->run("winrar x $dir\\".$nname." ".$dir , 0 ,true);//执行RUN方法来执行winrar命令来解压文件!
 //unlink($nname);//此命令为删除文件,意思上传后删除原来上传的压缩文件,只留解压后的文件夹!

 }else{
 echo "对不起，上传格式必须是压缩包格式文件，请调整格式后重新上传，谢谢 ！";
 }
 


echo  "<script>alert('".$Burl.$nname."上传成功');</script>";
// 获取网站根目录
//echo $Surl.$Burl."<br>";
//echo "winrar x $dir\\$nname $dir";

 echo "<script>parent.location.reload();</script>";

}
?>