<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="admin_md5.asp"-->
<%
if session("loginstatus")<>"islogined" or session("grade")<>3 then
	response.redirect "index.asp?ntime="&ntime
	response.end
end if
%>
<!--#include file="admin_fsoconn.asp"-->
<%
dim username,password,password2
set rs=server.createobject("adodb.recordset")
username=trim(request("username"))
if request.form("submit")="ȷ��" then
	password=trim(request.form("password"))
	grade=trim(request.form("grade"))
	if grade<>"" and isnumeric(grade) then
		grade=cint(grade)
		if grade<1 or grade>3 then
			grade=1
		end if
	else
		grade=1
	end if
	pathaccess=trim(request.form("pathaccess"))
	if left(pathaccess,1)<>"/" then
		pathaccess="/"
	end if
	if right(pathaccess,1)<>"/" then
		pathaccess=pathaccess&"/"
	end if
	if username="" or instr(username,"'")>0 then
		response.write "<script language='javascript'>"&chr(13)
		response.write "<!--"&chr(13)
		response.write "alert('�û�������Ϊ�ջ�������ţ���������д��');"&chr(13)
		response.write "window.location.href='"&selfname&"?ntime="&ntime&"';"&chr(13)
		response.write "//-->"&chr(13)
		response.write "</script>"&chr(13)
		response.end
	end if
	if password="" or instr(password,"'")>0 then
		response.write "<script language='javascript'>"&chr(13)
		response.write "<!--"&chr(13)
		response.write "alert('���벻��Ϊ�ջ�������ţ�');"&chr(13)
		response.write "window.location.href='"&selfname&"?username="&username&"&ntime="&ntime&"';"&chr(13)
		response.write "//-->"&chr(13)
		response.write "</script>"&chr(13)
		response.end
	end if
	sql="select * from userinfo where username='"&username&"'"
	rs.open sql,conn,0,3
	if not (rs.bof and rs.eof) then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script language='javascript'>"&chr(13)
		response.write "<!--"&chr(13)
		response.write "alert('�Բ����û���"&username&"���Ѿ����ڣ���ѡ�������û�����');"&chr(13)
		response.write "window.location.href='"&selfname&"?username="&username&"';"&chr(13)
		response.write "//-->"&chr(13)
		response.write "</script>"&chr(13)
	else
		rs.addnew
		rs("username")=username
		rs("password")=md5(password)
		rs("grade")=grade
		rs("pathaccess")=pathaccess
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.redirect "admin_bozo.asp?ntime="&ntime
	end if
	response.end
end if
%>
<html>
<head>
<title><%=sysname%>--�����û�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
BODY{background:#ffffff url(images/bgbrick.gif);color:#770011;text-decoration: none;margin:10px}
A{COLOR:#3333ff;text-decoration:none}
A:hover{COLOR:#ff3333;text-decoration:underline}
form{padding:0;margin:0}
table{font-size:12px;width:250px;}
td{padding:0 0 0 12px;line-height:17px;}
.mybutton{font-size:12px; height:21px; width:60px; color: #E0E0E0; background-color: #006898; border: 2 solid #3399FF}
.myinput{font-size:12px;height:21px;width:130px;border:solid 1px #000080;}
.rdborder{border:solid #000080;border-width:1px 0 0 1px;padding:0}
.rdborder td{border:solid #000080;border-width:0 1px 1px 0;margin:0;}
.rowbg1 td{background:#eee3ee;}
.rowbg2 td{background:#eeeedd;}
.rowbg3 td{background:#f3f3ee;}
//-->
</style>
<script language="javascript"></script>
<script language="vbscript">
<!--
function formcheck()
	v_username=document.forms(0).username.value
	v_password=document.forms(0).password.value
	v_pathaccess=trim(document.forms(0).pathaccess.value)
	if v_username="" then
		msgbox("�û�������Ϊ��Ŷ")
		document.forms(0).username.focus()
		formcheck=false
	elseif trim(v_username)<>v_username or instr(v_username,"%")>0 or instr(v_username,"?")>0 or instr(v_username,"'")>0 or instr(v_username,chr(34))>0 then
		msgbox("�û����в��ܰ����ո񼰡�%������?������'������"&chr(34)&"���ȷ���")
		document.forms(0).username.focus()
		formcheck=false
	elseif v_password="" then
		msgbox("���벻��Ϊ�հ���")
		document.forms(0).password.focus()
		formcheck=false
	elseif v_pathaccess="" or left(v_pathaccess,1)<>"/" then
		msgbox("�û�Ŀ¼����Ϊ���ұ�����б��'/'��ͷ")
		document.forms(0).pathaccess.focus()
		formcheck=false
	elseif instr(v_pathaccess,":")>0 or instr(v_pathaccess,"\")>0 or instr(v_pathaccess,"*")>0 or instr(v_pathaccess,"?")>0 or instr(v_pathaccess,"%")>0 or instr(v_pathaccess,chr(34))>0 or instr(v_pathaccess,chr(39))>0 then
		msgbox("Ŀ¼�к��зǷ��ַ���")
		document.forms(0).pathaccess.focus()
		formcheck=false
	end if
end function
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table align="center" cellpadding=0 cellspacing=0 style="background:#00aead url(images/leadtop.gif);border:solid #000080;border-width:1px 1px 0;height:20px;color:#ffffff">
	<tr>
		<td align=center><%=sysname%>--������û�</td>
	</tr>
</table>
<table align="center" class="rdborder" cellspacing=0 height="140">
<form method="post" action="<%=selfname%>" onsubmit="return formcheck()">
<tr>
<td width="35%">�û�����</td>
<td><input name="username" value="<%=username%>" type="text" class="myinput"></td>
</tr>
<tr>
<td width="35%">���룺</td>
<td><input name="password" type="password" class="myinput"></td>
</tr>
<tr>
<td width="35%">�û�����</td>
<td><select name="grade" class="myinput" onchange="javascript:if(this.form.grade.options[this.form.grade.options.selectedIndex].value==3) this.form.pathaccess.value='/';">
<option value="1" selected>��ͨ�û�
<option value="3">�����û�
</select>
</td>
</tr>
<tr>
<td width="35%">Ŀ¼Ȩ�ޣ�</td>
<td><input name="pathaccess" type="text" class="myinput"></td>
</tr>
<tr>
<td colspan=2 align="center"><input type="submit" name="submit" value="ȷ��" class="mybutton" onMouseOver ="javascript:this.style.backgroundColor='#ff0000'" onMouseOut ="javascript:this.style.backgroundColor='#006898'">&nbsp;&nbsp;<input type="reset" value="����" class="mybutton" onMouseOver ="javascript:this.style.backgroundColor='#ff0000'" onMouseOut ="javascript:this.style.backgroundColor='#006898'">&nbsp;&nbsp;<input type="button" value="����" class="mybutton" onMouseOver ="javascript:this.style.backgroundColor='#ff0000'" onMouseOut ="javascript:this.style.backgroundColor='#006898'" onclick="window.location.href='admin_bozo.asp?ntime=<%=ntime%>'"></td>
</tr>
</form>
</table>
<table align="center" cellpadding=0 cellspacing=0 style="background:#00aead url(images/leadbottom.gif);border:solid #000080;border-width:0 1px 1px;height:20px;color:#ffffff">
	<tr>
		<td align=center><a href="javascript:void(0)" onclick="window.close()">�رձ�����</td>
	</tr>
</table>
</body>
</html>
