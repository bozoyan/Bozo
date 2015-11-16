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
username=trim(request.querystring("username"))
act=request.querystring("act")
isdel=(request.querystring("isdel"))
if act<>"deluser" or username="" then
%>
<html>
<head>
<title><%=sysname%>--用户管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
BODY{background:#ffffff url(images/bgbrick.gif);color:#770011;text-decoration: none;margin-top:30px}
A{COLOR:#3333ff;text-decoration:none}
A:hover{COLOR:#ff3333;text-decoration:underline}
form{padding:0;margin:0}
table{font-size:12px;width:250px;}
td{padding:0 0 0 12px;line-height:17px;}
.mybutton{font-size:12px; height:21px; width:60px; color: #E0E0E0; background-color: #006898; border: 2 solid #3399FF}
.rdborder{border:solid #000080;border-width:1px 0 0 1px;padding:0}
.rdborder td{border:solid #000080;border-width:0 1px 1px 0;margin:0;}
.rowbg1 td{background:#eee3ee;}
.rowbg2 td{background:#eeeedd;}
.rowbg3 td{background:#f3f3ee;}
//-->
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table align="center" cellpadding=0 cellspacing=0 style="background:#00aead url(images/leadtop.gif);border:solid #000080;border-width:1px 1px 0;height:20px;color:#ffffff">
	<tr>
		<td align=center><%=sysname%>--用户管理</td>
	</tr>
</table>
<table align="center" class="rdborder" cellspacing=0 height="60">
<form>
<tr>
<td width="35%">用户及权限：</td>
<td><select name="username" style="width:120px">
<%
	sql="select * from userinfo order by username asc"
	set rs=conn.execute(sql)
	do while not (rs.eof or err)
		response.write "<option value='"&rs("username")&"'>"&rs("username")&" ("
	if rs("grade")=3 then
		response.write "/"
	else
		response.write rs("pathaccess")
	end if
	response.write ")"&chr(13)
	rs.movenext
	loop
	set rs=nothing
	conn.close
	set conn=nothing
%>
</select>
</td>
</tr>
<tr>
<td colspan=2><input type="button" value="添加" class="mybutton" onMouseOver ="this.style.backgroundColor='#ff0000'" onMouseOut ="javascript:this.style.backgroundColor='#006898'" onclick="window.location.href='admin_add.asp?ntime=<%=ntime%>'">&nbsp;&nbsp;&nbsp;<input type="button" value="删除" class="mybutton" onMouseOver ="this.style.backgroundColor='#ff0000'" onMouseOut ="javascript:this.style.backgroundColor='#006898'" onclick="window.location.href='<%=selfname%>?act=deluser&username='+this.form.username.options[this.form.username.selectedIndex].value+'&ntime=<%=ntime%>'">&nbsp;&nbsp;&nbsp;<input type="button" value="编辑" class="mybutton" onMouseOver ="this.style.backgroundColor='#ff0000'" onMouseOut ="javascript:this.style.backgroundColor='#006898'" onclick="window.location.href='admin_edit.asp?username='+this.form.username.options[this.form.username.selectedIndex].value+'&ntime=<%=ntime%>'"></td>
</tr>
</form>
</table>
<table align="center" cellpadding=0 cellspacing=0 style="background:#00aead url(images/leadbottom.gif);border:solid #000080;border-width:0 1px 1px;height:20px;color:#ffffff">
	<tr>
		<td align=center><a href="javascript:void(0)" onclick="window.close()">关闭本窗口</td>
	</tr>
</table>
</body>
</html>
<%
elseif session("loginname")=username then
	conn.close
	set conn=nothing
%>
<script language="javascript">
<!--
alert("对不起，不能删除自己！");
window.history.go(-1);
//-->
</script>
<%
elseif isdel<>"true" then
	conn.close
	set conn=nothing
%>
<script language="javascript">
<!--
var question = window.confirm("你真的要删除用户“<%=username%>”吗？")
if (question) 
	window.location.href="<%=selfname%>?act=deluser&username=<%=username%>&isdel=true&ntime=<%=ntime%>";
else
	window.location.href="<%=selfname%>?ntime=<%=ntime%>";
//-->
</script>
<%
else
	sql="delete from userinfo where username='"&username&"'"
	conn.execute(sql)
	conn.close
	set conn=nothing
	response.redirect selfname&"?ntime="&ntime
end if
%>
