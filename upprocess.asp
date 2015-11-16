<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="incupload.asp"-->
<%
set upload=new upload_5xsoft
if upload.form("act")="uploadfile" then
	filepath=trim(upload.form("filepath"))
	if left(filepath,1)<>"/" or instr(filepath,"\")>0 or instr(filepath,"*")>0 or instr(filepath,"?")>0 or instr(filepath,"'")>0 or instr(filepath,chr(34))>0 then
		response.write "<center><div style='width:600px;margin-top:100px;padding:3px 0;text-align:center;font-size:12px;color:#990000;border:solid 1px;border-color:#000090 #000090 #cccccc;background:#cccccc'>绍兴蓝蜗牛在线文件管理器--上传结果报告</div><div style='width:600px;padding:30px;font-size:12px;color:#000090;border:solid 1px;border-color:#000090 #000090 #cccccc;'>"
		response.write "对不起，没有填写路径或路径中含有非法字符<br><br>"
		response.write "<input type='button' style='width:65px;height:20px;font-size:12px' value='返回' onclick='window.history.go(-1)'>　　　<input type='button' style='width:65px;height:20px;font-size:12px' value='关闭' onclick='window.close()'></div>"&vbcrlf
		response.write "<div style='width:600px;padding:3px 0;font-size:12px;color:#000090;border:solid 1px;border-color:#cccccc #000090 #000090;background:#cccccc;'>版权所有：<a href='http://www.bozoyan.com' target='_blank'>www.bozoyan.com</a>　　　程序设计：<a href=mailto:yanbo@qq.com>绍兴蓝蜗牛</a></div>"&vbcrlf&"</center>"
		response.end
	end if
	if right(filepath,1)<>"/" then
		filepath=filepath&"/"
	end if
	if session("grade")<>3 and session("pathaccess")<>left(filepath,len(session("pathaccess"))) then
		set upload=nothing
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
window.close();
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
	basepath=Server.mappath(filepath)
	set obj_fso=server.createobject("scripting.filesystemobject")
	if not obj_fso.folderexists(basepath) then
		set obj_fso=nothing
		set upload=nothing
		response.write "<center><div style='width:600px;margin-top:100px;padding:3px 0;text-align:center;font-size:12px;color:#990000;border:solid 1px;border-color:#000090 #000090 #cccccc;background:#cccccc'>绍兴蓝蜗牛在线文件管理器--上传结果报告</div><div style='width:600px;padding:30px;font-size:12px;color:#000090;border:solid 1px;border-color:#000090 #000090 #cccccc;'>"
		response.write "对不起，目录“"&filepath&"”不存在，请先创建该目录！<br>"&vbcrlf
		response.write "<br><input type='button' style='width:65px;height:20px;font-size:12px' value='返回' onclick='window.history.go(-1)'>　　　<input type='button' style='width:65px;height:20px;font-size:12px' value='关闭' onclick='window.close()'></div>"
		response.write "<div style='width:600px;padding:3px 0;font-size:12px;color:#000090;border:solid 1px;border-color:#cccccc #000090 #000090;background:#cccccc;'>版权所有：<a href='http://www.bozoyan.com' target='_blank'>www.bozoyan.com</a>　　　程序设计：<a href=mailto:yanbo@qq.com>绍兴蓝蜗牛</a></div>"&vbcrlf&"</center>"
		response.end
	end if
%>
<html>
<head>
<title><%=sysname%>--上传文件成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body,table,td{font-size:12px;color:#000090;}
body{margin:0;background:#ffffff url(images/bgbrick.gif);}
table{width:760px}
table table{width:100%;}
td{vertical-align:top;padding-left:12px;}
td td{padding-left:4px;vertical-align:middle;}
img{vertical-align:bottom}
form,p{margin:0;padding:0}
a{color:#000080;text-decoration:none;}
a:hover{color:#ff3333;text-decoration:underline}
.button{display:inline-block; *display:inline; *zoom:1; line-height:30px; padding:0 20px; background-color:#56B4DC; color:#fff; font-size:14px; border-radius:3px; cursor:pointer; font-weight:normal;}
//-->
</style>
</head>
<body onload="opener.window.location.reload()">
<center>
<div style="width:600px;margin-top:100px;padding:3px 0;text-align:center;font-size:12px;color:#990000;border:solid 1px;border-color:#000090 #000090 #cccccc;background:#cccccc">绍兴蓝蜗牛在线文件管理器--上传结果报告</div>
<div style="width:600px;padding:30px;font-size:12px;color:#000090;border:solid 1px;border-color:#000090 #000090 #cccccc;">
<%
	i=0
	for each formName in upload.objFile
		set file=upload.objFile(formName)
		if file.FileSize>0 then
			file.SaveAs basepath&"\"&file.FileName
			response.write file.FileName&" ("&formatnumber(file.FileSize/1024,2,-1)&" K )　　上传至　　"
			response.write filepath&file.FileName&"　　成功!<br>"
			i=i+1
 		end if
		set file=nothing
	next
	set upload=nothing
	response.write "<br><br><br><b>共"&i&"个文件上传成功！</b><br>"
	response.write "<br><input type='button' style='width:65px;height:20px;font-size:12px' value='返回' onclick='window.history.go(-1)'>　　　<input type='button' style='width:65px;height:20px;font-size:12px' value='关闭' onclick='window.close()'></div>"&vbcrlf
	response.write "<div style='width:600px;padding:3px 0;font-size:12px;color:#000090;border:solid 1px;border-color:#cccccc #000090 #000090;background:#cccccc;'>版权所有：<a href='http://bozoyan.com' target='_blank'>绍兴蓝蜗牛网络</a>　　　程序设计：<a href=http://bozo.51xf.org>绍兴蓝蜗牛</a></div>"&vbcrlf&"</center>"&vbcrlf
	response.write "</body>"&vbcrlf
	response.write "</html>"&vbcrlf
	response.write "<script>parent.location.reload();</script>"&vbcrlf
else
	set upload=nothing
	response.redirect "admin_fsoexplorer.asp?ntime="&ntime
end if
%>
