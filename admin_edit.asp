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
dim username,password
username=trim(request.querystring("username"))
if request.form("submit")="ȷ��" and username=request.form("username") then
	password=trim(request.form("password"))
	grade=trim(request.form("grade"))
	pathaccess=trim(request.form("pathaccess"))
	if grade<>"" and isnumeric(grade) then
		grade=fix(grade)
		if grade<1 or grade>3 then
			grade=3
		end if
	else
		grade=3
	end if
	if left(pathaccess,1)<>"/" then
		pathaccess="/"
	end if
	if right(pathaccess,1)<>"/" then
		pathaccess=pathaccess&"/"
	end if
	sql="select * from userinfo where username='"&replace(username,"'","''")&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,0,3
	if rs.bof and rs.eof then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.write "<script language='javascript'>"&chr(13)
		response.write "<!--"&chr(13)
		response.write "alert('�Բ����û���"&username&"�������ڣ������Ѿ���ɾ����');"&chr(13)
		response.write "window.location.href='admin_bozo.asp?ntime="&ntime&"';"&chr(13)
		response.write "//-->"&chr(13)
		response.write "</script>"&chr(13)
	else
		if password<>"" then
			rs("password")=md5(password)
		end if
		rs("grade")=grade
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		response.redirect "admin.asp?ntime="&ntime
	end if
	response.end
end if

sql="select * from userinfo where username='"&username&"'"
set rs=conn.execute(sql)
if rs.bof and rs.eof then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<script language='javascript'>"&chr(13)
	response.write "<!--"&chr(13)
	response.write "alert('�Բ��𣬷��������뷵���û�����ҳ���ԣ�');"&chr(13)
	response.write "window.location.href='admin_bozo.asp?ntime="&ntime&"';"&chr(13)
	response.write "//-->"&chr(13)
	response.write "</script>"&chr(13)
else
%>
<html>
<head>
<title><%=sysname%>--�û��༭</title>
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
	elseif instr(v_password,"'")>0 then
		msgbox("�����в��ܰ������Ű���")
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
		<td align=center><%=sysname%>--�û��༭</td>
	</tr>
</table>
<table align="center" class="rdborder" cellspacing=0 height="130">
<form method="post" action="<%=selfname%>?username=<%=username%>" onsubmit="return formcheck()">
<tr>
<td width="35%">�û�����</td>
<td><input name="username" value="<%=rs("username")%>" type="text" readonly class="myinput"></td>
</tr>
<tr>
<td width="35%">���룺</td>
<td><input name="password" type="password" value="" class="myinput"></td>
</tr>
<tr>
<td width="35%">�û�����</td>
<td><select name="grade" class="myinput" onchange="if(this.form.grade.options[this.form.grade.options.selectedIndex].value==3) this.form.pathaccess.value='/'">
<option value="1" <%if rs("grade")<>3 then response.write "selected"%>>һ���û�
<option value="3" <%if rs("grade")=3 then response.write "selected"%>>�����û�
</select>
</td>
</tr>
<tr>
<td width="35%">Ŀ¼Ȩ�ޣ�</td>
<td><input name="pathaccess" type="text" value="<%=rs("pathaccess")%>" class="myinput"></td>
</tr>
<%
end if
set rs=nothing
conn.close
set conn=nothing
%>
<tr>
<td colspan=2 align="center"><input type="submit" name="submit" value="ȷ��" class="mybutton" onMouseOver ="this.style.backgroundColor='#ff0000'" onMouseOut ="this.style.backgroundColor='#006898'">&nbsp;&nbsp;<input type="reset" value="����" class="mybutton" onMouseOver ="this.style.backgroundColor='#ff0000'" onMouseOut ="this.style.backgroundColor='#006898'">&nbsp;&nbsp;<input type="button" value="����" class="mybutton" onMouseOver ="javascript:this.style.backgroundColor='#ff0000'" onMouseOut ="javascript:this.style.backgroundColor='#006898'" onclick="window.location.href='admin_bozo.asp?ntime=<%=ntime%>'"></td>
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
