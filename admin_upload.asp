<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<html>
<head>
<title><%=sysname%>--�ļ��ϴ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body,table,td{font-size:12px;color:#000090;}
body{margin:0;background: #f5f9fb;}
table{width:100%;}
td{vertical-align:top;padding-top:30px;font-size:16px;}
form,p{margin:0;padding:0}
a{color:#000080;text-decoration:none;}
a:hover{color:#ff3333;text-decoration:underline}
.button{width: 180px;
height: 25px;
border: 1px solid #B5BBBF;
background-color: #fff;
margin: 3px 0;}
.button1{width: 400px;
height: 25px;
border: 1px solid #B5BBBF;
background-color: #fff;
margin: 3px 0;}
//-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000">

<form action="tool/bozoup.php" method="post" enctype="multipart/form-data" name="form1" id="form1">
<table width="100%" border="0" align="center"><tr><td align="center">
  <input type="file" name="upfiles" class="button1">
  <input type="submit" name="sub" value="�ύ����ѹRARѹ���ļ�"  class="button" >
  <input type="hidden" name="Udir" value="<%=Request.QueryString("Path")%>" >
  </td></tr>
  <tr><td align="left" style="padding-left:100px;line-height:30px;">1����ǰ�ϴ��ļ�����<span style="color:red">http://<%=Request.ServerVariables("SERVER_NAME")%><%=Request.QueryString("Path")%></span> Ŀ¼�� <br>2���ϴ�ѹ�����ļ����ó���2M��ѹ������ʽ֧��: rar zip 7z iso��<br>3��Ŀ¼��<span style="color:red">û��ͬ���ļ����ļ��У����߽�ѹ��ͬ���ļ����ļ���</span>������ᱨ��<br>4�������ļ����ֱ��ʹ���ұߡ��ϴ��ļ�����ť���в�����</td></tr></table>
</form>

</body>
</html>
