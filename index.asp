<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="admin_md5.asp"-->
<%
if session("loginstatus")="islogined" and left(session("pathaccess"),1)="/" then
	response.redirect "admin_fsoexplorer.asp?ntime="&ntime
	response.end
end if

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
If Not IsObjInstalled("Scripting.FileSystemObject") Then 
	logininfo="�÷�������֧��FSO�����޷�ʹ�ñ�����"
elseif not IsObjInstalled("adodb.stream") Then
	logininfo="������ADO����汾̫�ͣ��㽫�޷�ʹ���ϴ����ܣ�"
end if

comeurl=trim(request.querystring("url"))
if request("submit")="�ύ" then
%>
<!--#include file="admin_fsoconn.asp"-->
<%
	dim loginname,loginpassword
	loginname=request.form("loginname")
	loginpassword=request.form("loginpassword")
	if trim(loginname)="" or instr(loginname,chr(39)) > 0 or instr(loginname,chr(34)) > 0 or instr(loginpassword,chr(39)) > 0 or instr(loginpassword,chr(34)) > 0 then
		response.redirect selfname&"?ntime="&ntime
		response.end
	end if
	comeurl=trim(request.form("comeurl"))
	sql="select * from userinfo where username="&chr(39)&loginname&chr(39)
	set rs=conn.execute(sql)
	if rs.bof and rs.eof then
		set rs=nothing
		conn.close
		set conn=nothing
		logininfo="���û�������"
	elseif rs("password")<>md5(loginpassword) then
		set rs=nothing
		conn.close
		set conn=nothing
		logininfo="���벻��ȷ"
	else
		session("loginname")=loginname
		session("grade")=rs("grade")
		session("pathaccess")=trim(rs("pathaccess"))
		session("loginstatus")="islogined"
		set rs=nothing
		conn.close
		set conn=nothing
		if comeurl="" then
			response.redirect "admin_fsoexplorer.asp?path="&session("pathaccess")&"&ntime="&ntime
		elseif instr(comeurl,"?")>0 then
			response.redirect(comeurl)
		else
			response.redirect(comeurl&"?ntime="&ntime)
		end if
		response.end
	end if
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title>��ӭ��¼<%=sysname%>�ļ�����ϵͳ</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="lib.js"></script>
<script src="cloud.js" type="text/javascript"></script>

<script language="javascript">
	$(function(){
    $('.loginbox').css({'position':'absolute','left':($(window).width()-692)/2});
	$(window).resize(function(){  
    $('.loginbox').css({'position':'absolute','left':($(window).width()-692)/2});
    })  
});  
</script> 

</head>

<body style="background-color:#1c77ac; background-image:url(images/light.png); background-repeat:no-repeat; background-position:center top; overflow:hidden;">



    <div id="mainBody">
      <div id="cloud1" class="cloud"></div>
      <div id="cloud2" class="cloud"></div>
    </div>  


<div class="logintop">    
    <span>��ӭ��¼<%=sysname%>�ļ��������ƽ̨</span>    
    <ul>
    <li><a href="http://bozoyan.com" target="_blank">����ţ����</a></li>
    <li><a href="http://www.net0575.cn" target="_blank">����΢��</a></li>
    <li><a href="http://www.yanbo.cc" target="_blank">����ʥ��</a></li>
    </ul>    
    </div>
    
    <div class="loginbody">
    
    <span class="systemlogo"></span> 
       			<form method=POST action="<%=selfname%>">
    <div class="loginbox">
    <%=logininfo%>
    <ul>
    <li><input name="loginname" type="text" class="loginuser"  onclick="JavaScript:this.value=''"/></li>
    <li><input name="loginpassword" type="password" class="loginpwd" onclick="JavaScript:this.value=''"/><input type="hidden" name="comeurl" value="<%=comeurl%>"></li>
    <li><input type=submit value=�ύ name="submit"  class="loginbtn"  /><label style="display:none;"><input name="" type="checkbox" value="" checked="checked" />��ס����</label></li>
    </ul>
   
    
    </div>
    </div>
    
    </div>
    
    
    
    <div class="loginbm">��Ȩ����  2015-2020  <a href="http://bozoyan.com"><%=sysname%></a> Design By Bozoyan  QQ��201782</div>
	
    
</body>

</html>
<!--#include file="admin_fsofoot.asp"-->