<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<%
act=request.querystring("act")
fname=request.querystring("fname")
newname=trim(request.querystring("newname"))
epath=trim(request.querystring("epath"))
epage=trim(request.querystring("epage"))
if epath<>"" then
	epath=server.urlencode(epath)
end if
if isnumeric(epage) then
	epage=fix(epage)
else
	epage=1
end if
if newname="" then
	response.write "<script language='javascript'>"&vbcrlf
	response.write "<!--"&vbcrlf
	response.write "alert('指定的文件（夹）名称不能为空');"&vbcrlf
	response.write "window.history.go(-1);"&vbcrlf
	response.write "//-->"&vbcrlf
	response.write "</script>"&vbcrlf
	response.end
end if
if left(path,1)<>"/" or instr(fname,"*")>0 or instr(fname,"/")>0 or instr(fname,"\")>0 or instr(newname,"*")>0 or instr(newname,"/")>0 or instr(newname,"\")>0 then
	response.write "<script language='javascript'>"&vbcrlf
	response.write "<!--"&vbcrlf
	response.write "alert('指定的路径或文件名称非法');"&vbcrlf
	response.write "window.history.go(-1);"&vbcrlf
	response.write "//-->"&vbcrlf
	response.write "</script>"&vbcrlf
	response.end
end if
bpath=server.mappath(path)
if act="renfolder" then
	folder_path=bpath&"\"&fname
	new_path=bpath&"\"&newname
	set obj_fso=server.createobject("scripting.filesystemobject")
	if obj_fso.folderexists(folder_path) and obj_fso.folderexists(new_path)=false then
		obj_fso.getfolder(folder_path).name=newname
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		'response.write "alert('重命名文件夹 ["&fname&"] 为 ["&newname&"] 成功！');"&vbcrlf
		response.write "window.location.href='admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"';"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	elseif obj_fso.folderexists(new_path) then
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "alert('重命名文件夹 ["&fname&"] 失败，文件夹 ["&newname&"] 已存在！');"&vbcrlf
		response.write "window.history.go(-1);"&vbcrlf
	response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	else
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "alert('重命名文件夹 ["&fname&"] 失败，文件夹 ["&fname&"] 不存在！');"&vbcrlf
		response.write "window.history.go(-1);"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	end if
	set obj_fso=nothing
elseif act="renfile" then
	file_path=bpath&"\"&fname
	new_path=bpath&"\"&newname
	set obj_fso=server.createobject("scripting.filesystemobject")
	if obj_fso.fileexists(file_path) and obj_fso.fileexists(new_path)=false then
		obj_fso.getfile(file_path).name=newname
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		'response.write "alert('文件 ["&fname&"] 更名为 ["&newname&"] 成功！');"&vbcrlf
		response.write "window.location.href='admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"';"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	elseif obj_fso.fileexists(file_path) then
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "alert('文件 ["&fname&"] 更名失败，文件 ["&newname&"] 已经存在！');"&vbcrlf
		response.write "window.history.go(-1);"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	else
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "alert('文件 ["&fname&"] 更名失败，该文件不存在！');"&vbcrlf
		response.write "window.history.go(-1);"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	end if
	set obj_fso=nothing
else
	response.redirect "admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime
end if
%>
