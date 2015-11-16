<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<html>
<head>
<title><%=sysname%>--文件上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body,table,td{font-size:12px;color:#000090;}
body{margin:0;background: #f5f9fb;}
table{width:100%;}
td{vertical-align:top;padding-left:12px;}
td td{padding-left:4px;vertical-align:middle;}
img{vertical-align:bottom}
.tx{margin:6px 6px 3px;border:solid 1px #000099;}
form,p{margin:0;padding:0}
a{color:#000080;text-decoration:none;}
a:hover{color:#ff3333;text-decoration:underline}
.button{width: 100px;
height: 25px;
border: 1px solid #B5BBBF;
background-color: #fff;
margin: 3px 0;}
//-->
</style>
</head>

<body>
<form name="form1" method="post" action="upprocess.asp?path=<%=Request.QueryString("Path")%>" enctype="multipart/form-data" >
  <input type="hidden" name="act" value="uploadfile">
  <table width="100%" border="0">
    <tr align="left" valign="middle" bgcolor="#f5f5f5"> 
      <td bgcolor="#f5f5f5" height="92"> 
        <script language="javascript">
	  function setid()
	  {
	  str='<br>';
	  if(!window.form1.upcount.value)
	   window.form1.upcount.value=1;
 	  for(i=1;i<=window.form1.upcount.value;i++)
	     str+='文件'+i+':<input type="file" name="file'+i+'" style="width:400" class="tx1"><br><br>';
	  window.upid.innerHTML=str+'<br>';
	  }
	  </script>
        <li> 需要上传的个数 
          <input type="text" name="upcount" class="tx" value="1" onkeydown="if(window.event.keyCode==13) {setid();return false}">
          <input type="button" name="Button" class="button" onclick="setid();" value="· 设定 ·">*最好10个以下
        </li>
        <br>
        <br>
        <li>上传到: 
          <input type="text" name="filepath" class="tx" style="width:350" value="<%=Request.QueryString("Path")%>" readonly>
        </li>
      </td>
    </tr>
    <tr align="center" valign="middle"> 
      <td align="left" id="upid" height="122"> 文件1: 
        <input type="file" name="file1" style="width:400" class="tx1" value="">
      </td>
    </tr>
    <tr align="center" valign="middle"> 
      <td height="24"> 
        <input type="submit" name="Submit" value="· 提交 ·" class="button">
        <input type="reset" name="Submit2" value="· 重置 ·" class="button">
      </td>
    </tr>
  </table>
</form>
<script language="javascript">
setid();
</script>
</body>
</html>
