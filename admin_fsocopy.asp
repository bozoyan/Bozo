<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<%
act=request.form("act")
basepath=trim(request.form("basepath"))
path=trim(request.querystring("path"))
if path<>"" then
	path=server.urlencode(path)
end if
page=request.querystring("page")
if (act="copy" or act="cut") and left(basepath,1)="/" then
	session("folders")=""
	session("files")=""
	n_folder=0
	n_file=0
	for each folder_name in request.form("folders")
		session("folders")=session("folders")&chr(9)&folder_name
		n_folder=n_folder+1
	next
	if left(session("folders"),1)=chr(9) then
		session("folders")=right(session("folders"),len(session("folders"))-1)
	end if
	for each file_name in request.form("files")
		session("files")=session("files")&chr(9)&file_name
		n_file=n_file+1
	next
	if left(session("files"),1)=chr(9) then
		session("files")=right(session("files"),len(session("files"))-1)
	end if
	if session("folders")="" and session("files")="" then
		session("act")=""
		session("basepath")=""
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "alert('你没有选中任何文件和文件夹哦！');"&vbcrlf
		response.write "window.history.go(-1);"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	else
		session("act")=act
		session("basepath")=server.mappath(basepath)
		response.redirect "admin_fsoexplorer.asp?path="&path&"&page="&page&"&ntime="&ntime
	end if
else
	response.write "<script language='javascript'>"&vbcrlf
	response.write "<!--"&vbcrlf
	response.write "alert('操作失败，请返回重试！');"&vbcrlf
	response.write "window.history.go(-1);"&vbcrlf
	response.write "//-->"&vbcrlf
	response.write "</script>"&vbcrlf
end if
%>
