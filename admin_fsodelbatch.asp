<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<%
basepath=trim(request.form("basepath"))
if request.form("act")="delbatch" and left(basepath,1)="/" then
	path=request.querystring("path")
	page=request.querystring("page")
	basepath=server.mappath(basepath)
	set obj_fso=server.createobject("scripting.filesystemobject")
	set obj_folders=request.form("folders")
	set obj_files=request.form("files")
	if obj_folders.count>0 then
		for each foldername in obj_folders
			if obj_fso.folderexists(basepath&"\"&foldername) then
				obj_fso.deletefolder basepath&"\"&foldername,true
			end if
		next
	end if
	if obj_files.count>0 then
		for each filename in obj_files
			if obj_fso.fileexists(basepath&"\"&filename) then
				obj_fso.deletefile basepath&"\"&filename,true
			end if
		next
	end if
	set obj_fso=nothing
	response.redirect "admin_fsoexplorer.asp?path="&path&"&page="&page&"&ntime="&ntime
else
	response.redirect "admin_fsoexplorer.asp?ntime="&ntime
end if
%>
