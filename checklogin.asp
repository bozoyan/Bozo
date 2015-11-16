<%
path=trim(request.querystring("path"))
if path="" or instr(path,"	")>0 or instr(path,vbcrlf)>0 or instr(path,":")>0 or instr(path,"\")>0 then
	path="/"
elseif right(path,1)<>"/" then
	path=path&"/"
end if
if left(path,1)<>"/" then
	path="/"&path
end if
if session("loginstatus")<>"islogined" then
	comeurl=selfname&"?ntime="&ntime
	for each vars in request.querystring
		if lcase(vars)<>"ntime" and len(request.querystring(vars))>=1 then
			comeurl=comeurl&"&"&vars&"="&server.urlencode(request.querystring(vars))
		end if
	next
	response.redirect "index.asp?url="&server.urlencode(comeurl)
	response.end
elseif session("grade")<>3 and session("pathaccess")<>left(path,len(session("pathaccess"))) then
%>
<html>
<head>
<title><%=sysname%>--出错信息</title>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<script language="javascript">
<!--
function alert_msg()
{
alert("抱歉，你只有管理目录 “<%=session("pathaccess")%>” 及其子目录的权限！");
window.location.href='admin_fsoexplorer.asp?path=<%=server.urlencode(session("pathaccess"))%>&ntime=<%=ntime%>';
}
//-->
</script>
</head>
<body onload="alert_msg()">
</body>
</html>
<%
	response.end
end if
%>
