<style type="text/css">
body,table,td{font-size:12px;color:#000090;}
body{margin:0;background: #f5f9fb;}
</style>
<?
$url=$_POST["Udir"];
$Burl=strtr("$url","/","\\");//�������Ŀ¼
$Surl=$_SERVER['DOCUMENT_ROOT']; //��վ��Ŀ¼
$dir = $Surl.$Burl;  //�ϴ���Ŀ¼
// $dir = getcwd() ��ȡ��ǰĿ¼  $Surl.$Burlָ��Ŀ¼
if($_POST["sub"]){//�жϵ�����ύ��ť

 $nname = $_FILES["upfiles"]["name"];//��ȡ�ϴ����ļ�����
 $tname = $_FILES["upfiles"]["tmp_name"];//��ȡ�ϴ��ļ�����ʱ�ļ���  $_FILES["upfiles"]["tmp_name"]
 
 $fiearr=explode(".",$nname);       //��ԭ�ļ����ֳ�����
 $key=count($fiearr)-1;                //��������һ����չ��������
 $fie_extend=$fiearr[$key];             //�г��ϴ��ļ�����չ��
 $fie_extend=strtolower($fie_extend);  //����չ��ͳһΪСд
 
 if($fie_extend=="zip"||$fie_extend=="rar"||$fie_extend=="7z"||$fie_extend=="iso"){  
 	 $path=$Surl.$Burl.$nname;//�����ϴ�Ŀ¼
 	 move_uploaded_file($tname,$path);//�ƶ��ϴ��ļ�,����֮ǰ��ʵ�ļ��Ѿ��ϴ��ɹ�!�˴���һ�������������!�˴�������ԭ�������������ļ�!
 $obj= new com("wscript.shell");//ʵ����COM���
 $obj->run("winrar x $dir\\".$nname." ".$dir , 0 ,true);//ִ��RUN������ִ��winrar��������ѹ�ļ�!
 //unlink($nname);//������Ϊɾ���ļ�,��˼�ϴ���ɾ��ԭ���ϴ���ѹ���ļ�,ֻ����ѹ����ļ���!

 }else{
 echo "�Բ����ϴ���ʽ������ѹ������ʽ�ļ����������ʽ�������ϴ���лл ��";
 }
 


echo  "<script>alert('".$Burl.$nname."�ϴ��ɹ�');</script>";
// ��ȡ��վ��Ŀ¼
//echo $Surl.$Burl."<br>";
//echo "winrar x $dir\\$nname $dir";

 echo "<script>parent.location.reload();</script>";

}
?>