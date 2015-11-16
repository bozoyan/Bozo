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
	logininfo="该服务器不支持FSO，故无法使用本程序！"
elseif not IsObjInstalled("adodb.stream") Then
	logininfo="服务器ADO组件版本太低，你将无法使用上传功能！"
end if

comeurl=trim(request.querystring("url"))
if request("submit")="提交" then
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
		logininfo="该用户不存在"
	elseif rs("password")<>md5(loginpassword) then
		set rs=nothing
		conn.close
		set conn=nothing
		logininfo="密码不正确"
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
<html>
<title><%=sysname%>--登陆入口</title>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<style type="text/css">
<!--
input,body,table{font-size:12px}
form{margin:0;padding:0}
.awhite,.awhite a{color:#ffffff;text-decoration:none}
.awhite a:hover{color:#ff3333;text-decoration:underline}
-->
</style>
<body topmargin=150 background=images/bgbrick.gif  oncontextmenu=window.event.returnValue=false 
onselectstart=window.event.returnValue=false 
oncopy="alert('Sorry，信风网盟禁止复制！版權所有');event.returnValue=false;">
<table border=1 align="center" cellpadding=0 cellspacing=0 style="border-collapse:collapse" bordercolor=#111111 width=280 height=170>
	<tr>
		<td align=center class="awhite" style="background:#00aead url(images/leadtop.gif);height:19px;"><%=sysname%>--用户登陆</td>
	</tr>
	<tr>
		<td align=center>
			<form method=POST action="<%=selfname%>">
				<table border=0 cellpadding=0 cellspacing=0 style=border-collapse:collapse bordercolor=#111111 width=100% height=128 bgcolor=#E0E0E0>
					<tr>
						<td colspan=2 height="18" style="text-align:center;color:#ff3333;fong-weight:600"><%=logininfo%></td>
					</tr>
					<tr>
						<td width=32% height=31 align=center>用户名</td>
						<td width=68% height=31><input type="text" name="loginname" style="border:solid 1px;padding:1px;width:132px"></td>
					</tr>
					<tr>
						<td width=32% height=31 align=center>密&nbsp; 码</td>
						<td width=68% height=31><input type="password" name="loginpassword" style="border:solid 1px;padding:1px;width:132px"><input type="hidden" name="comeurl" value="<%=comeurl%>"></td>
					</tr>
					<tr>
						<td width=100% height=26 colspan=2 align=center>　 <input type=submit value=提交 name="submit" style="font-size:12px;height:21px;width:60px;color:#E0E0E0;background-color: #006898; border: 2 solid #3399FF" onMouseOver ="this.style.backgroundColor='#ff0000'" onMouseOut ="this.style.backgroundColor='#006898'">&nbsp;<input type=reset value=重置 name=reset style="font-size:12px; height:21px; width:60px; color: #E0E0E0; background-color: #006898; border: 2 solid #3399FF" onMouseOver ="this.style.backgroundColor='#ff0000'" onMouseOut ="this.style.backgroundColor='#006898'"> <input type="button" onclick="window.location='index.asp'" style="font-size:12px; height:21px; width:60px; color: #E0E0E0; background-color: #006898; border: 2 solid #3399FF" onMouseOver ="this.style.backgroundColor='#ff0000'" onMouseOut ="this.style.backgroundColor='#006898'" value="那我走了"></td>
					</tr>
					<tr>
						<td colspan=2 height="12"></td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
	<tr>
		<td align="center" class="awhite" style="background:#00aead url(images/leadbottom.gif);height:19px;"><a href="http://www.51xf.org" target="_blank">本程序由贤少信风提供</a></td>
	</tr>
</table>
<!--#include file="admin_fsofoot.asp"-->

