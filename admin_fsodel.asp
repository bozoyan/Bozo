<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="admin_fsofunctions.asp"-->
<%
act=request.querystring("act")
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
if path="" or instr(path,":")>0 or instr(path,"\")>0 or instr(path,"?")>0 or instr(path,"*")>0 then
	response.write "<script language='javascript'>"&vbcrlf
	response.write "<!--"&vbcrlf
	response.write "alert('ָ�����ļ�����·���к��зǷ��ַ�');"&vbcrlf
	response.write "window.history.go(-1);"&vbcrlf
	response.write "//-->"&vbcrlf
	response.write "</script>"&vbcrlf
	response.end
end if
if len(path)>1 and right(path,1)="/" then
	path=left(path,len(path)-1)
end if
full_path=server.mappath(path)
set obj_fso=server.createobject("scripting.filesystemobject")
select case act
	case "delfile"
		if obj_fso.fileexists(full_path) then
			obj_fso.deletefile(full_path)
			response.write "<script language='javascript'>"&vbcrlf
			response.write "<!--"&vbcrlf
			response.write "window.location.href='admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"';"&vbcrlf
			response.write "//-->"&vbcrlf
			response.write "</script>"&vbcrlf
		else
			response.write "<script language='javascript'>"&vbcrlf
			response.write "<!--"&vbcrlf
			response.write "alert('��Ҫɾ�����ļ� "&getname(path,"/")&" û���ҵ���');"&vbcrlf
			response.write "window.history.go(-1);"&vbcrlf
			response.write "//-->"&vbcrlf
			response.write "</script>"&vbcrlf
		end if
	case "delfolder"
		if obj_fso.folderexists(full_path) then
			obj_fso.deletefolder(full_path)
			response.write "<script language='javascript'>"&vbcrlf
			response.write "<!--"&vbcrlf
			response.write "window.location.href='admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"';"&vbcrlf
			response.write "//-->"&vbcrlf
			response.write "</script>"&vbcrlf
		else
			response.write "<script language='javascript'>"&vbcrlf
			response.write "<!--"&vbcrlf
			response.write "alert('��Ҫɾ������Ŀ¼ "&getname(path,"/")&" û���ҵ���');"&vbcrlf
			response.write "window.history.go(-1);"&vbcrlf
			response.write "//-->"&vbcrlf
			response.write "</script>"&vbcrlf
		end if
	case "createfolder"
		if obj_fso.folderexists(full_path) then
			response.write "<script language='javascript'>"&vbcrlf
			response.write "<!--"&vbcrlf
			response.write "alert(""��Ҫ��������Ŀ¼ "&getname(path,"/")&" �Ѿ����ڣ�"");"&vbcrlf
			response.write "window.location.href='admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"';"&vbcrlf
			response.write "//-->"&vbcrlf
			response.write "</script>"&vbcrlf
		else
			obj_fso.createfolder(full_path)
			response.write "<script language='javascript'>"&vbcrlf
			response.write "<!--"&vbcrlf
			response.write "window.location.href='admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"';"&vbcrlf
			response.write "//-->"&vbcrlf
			response.write "</script>"&vbcrlf
		end if
	case "createfile"
		if obj_fso.fileexists(full_path) then
			response.write "<script language='javascript'>"&vbcrlf
			response.write "<!--"&vbcrlf
			response.write "alert(""��Ҫ�������ļ� "&getname(path,"/")&" �Ѿ����ڣ�"");"&vbcrlf
			response.write "window.location.href='admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"';"&vbcrlf
			response.write "//-->"&vbcrlf
			response.write "</script>"&vbcrlf
		else
			obj_fso.createtextfile(full_path)
			response.write "<script language='javascript'>"&vbcrlf
			response.write "<!--"&vbcrlf
			response.write "window.location.href='admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"';"&vbcrlf
			response.write "//-->"&vbcrlf
			response.write "</script>"&vbcrlf
		end if
	case else
		response.redirect "admin_fsoexplorer.asp?path="&epath&"&page="&epage&"&ntime="&ntime
end select
%>
