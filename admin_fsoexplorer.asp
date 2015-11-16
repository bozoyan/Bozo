<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<!--#include file="admin_fsofunctions.asp"-->
<%
page=trim(request.querystring("page"))
act=trim(request.querystring("act"))
if page<>"" and isnumeric(page) then
	page=fix(page)
else
	page=1
end if
if path="/" then
	goparent="↑回上级目录..."
else
	goparent="<a href='"&selfname&"?path="&server.urlencode(left(path,instrrev(path,"/",len(path)-1)))&"&ntime="&ntime&"'>↑回上级目录...</a>"
end if
parent_url=server.urlencode(parent_url)
pathurl=server.urlencode(path)
s_folderpath=server.mappath(path)
set obj_fso=server.createobject("scripting.filesystemobject")
if act="paste" and session("act")="copy" and session("basepath")<>s_folderpath and obj_fso.folderexists(s_folderpath)=true then
	n_errfile=0
	n_errfolder=0
	if session("folders")<>"" then
		a_folders=split(session("folders"),chr(9))
		m=ubound(a_folders)
		for i=0 to m
			s_sourcefolder=session("basepath")&"\"&a_folders(i)
			s_targetfolder=s_folderpath&"\"&a_folders(i)
			if obj_fso.folderexists(s_sourcefolder) then
				obj_fso.copyfolder s_sourcefolder,s_targetfolder,true
			else
				n_errfolder=n_errfolder+1
			end if
		next
	else
		m=0-1
	end if
	if session("files")<>"" then
		a_files=split(session("files"),chr(9))
		n=ubound(a_files)
		for i=0 to n
			s_sourcefile=session("basepath")&"\"&a_files(i)
			s_targetfile=s_folderpath&"\"&a_files(i)
			if obj_fso.fileexists(s_sourcefile) then
				obj_fso.copyfile s_sourcefile,s_targetfile,true
			else
				n_errfile=n_errfile+1
			end if
		next
	else
		n=0-1
	end if
elseif act="paste" and session("act")="cut" and session("basepath")<>s_folderpath and obj_fso.folderexists(s_folderpath)=true then
	n_errfile=0
	n_errfolder=0
	if session("folders")<>"" then
		a_folders=split(session("folders"),chr(9))
		m=ubound(a_folders)
		for i=0 to m
			s_sourcefolder=session("basepath")&"\"&a_folders(i)
			s_targetfolder=s_folderpath&"\"&a_folders(i)
			if obj_fso.folderexists(s_sourcefolder) then
				obj_fso.copyfolder s_sourcefolder,s_targetfolder,true
				obj_fso.deletefolder s_sourcefolder,true
			else
				n_errfolder=n_errfolder+1
			end if
		next
	end if
	if session("files")<>"" then
		a_files=split(session("files"),chr(9))
		n=ubound(a_files)
		for i=0 to n
			s_sourcefile=session("basepath")&"\"&a_files(i)
			s_targetfile=s_folderpath&"\"&a_files(i)
			if obj_fso.fileexists(s_sourcefile) then
				obj_fso.copyfile s_sourcefile,s_targetfile,true
				obj_fso.deletefile s_sourcefile,true
			else
				n_errfile=n_errfile+1
			end if
		next
	end if
	session("act")=""
	session("basepath")=""
	session("folders")=""
	session("files")=""
end if
%>
<!doctype html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gbk" />
<title><%=sysname%>文件管理中心</title>
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script src="lib.js"></script>
<script src="layer.min.js"></script>
<script>
layer.use('src/layer.ext.js'); //载入layer拓展模块
</script>

<!--#include file="js.asp"-->
<style type="text/css">
<!--
/*
body,table,td{font-size:12px;color:#000090;}
body{margin:0 0 1px;padding:0;background:#ffffff url(images/bgbrick.gif);}
table{width:760px}
table table{width:100%;background:#eeeeee}
td{vertical-align:top;}
table table td{padding-left:4px;vertical-align:middle;background:#ffffff}
img{vertical-align:bottom}
form{margin:0;padding:0}
a{color:#000080;text-decoration:none;}
a:hover{color:#ff3333;text-decoration:underline}
div{width:100%;background:#eeeeee}
*/
.button0{width:100px;height:25px;border:1px solid #B5BBBF; background-color: #fff;margin:3px 0;}
.button1{width:70px;height:25px;border:1px solid #B5BBBF; background-color: #fff;margin:3px 0;}
.button2{width:57px;height:25px;border:1px solid #B5BBBF; background-color: #fff;margin:3px 0;}
.imgbutton{width:32px;height:32px;cursor:hand;}
.imgbt{border:solid 1px;border-color:#ffffff #999999 #999999 #ffffff;cursor:hand;}
pre{font-family:'微软雅黑'}
.box{padding:20px; background-color:#fff; margin:50px 100px; border-radius:5px;}
.box a{padding-right:15px;}
#about_hide{display:none}
.layer_text{background-color:#fff; padding:20px;}
.layer_text p{margin-bottom: 10px; text-indent: 2em; line-height: 23px;}
.button{display:inline-block; *display:inline; *zoom:1; line-height:30px; padding:0 20px; background-color:#56B4DC; color:#fff; font-size:14px; border-radius:3px; cursor:pointer; font-weight:normal;}
.imgs img{width:300px;}
.bozo a{
width:60px;height:25px;border:1px solid #B5BBBF; background-color: #fff;margin:5px 5px;padding-right:15px;float:left;
text-align: center;
display: block;
;}
//-->
</style>
</head>
<body style="background:url(images/topbg.gif) repeat-x;">
<table width="100%"><tr><td colspan="2">
    <div class="topleft">
    <a href="admin_fsoexplorer.asp" target="_parent"><img src="images/logo.png" title="系统首页" /></a>
    </div>
        
    <ul class="nav">
    <li><a href="admin_fsoexplorer.asp" class="selected"><img src="images/icon01.png" title="工作台" /><h2>工作台</h2></a></li>
     <%if session("loginstatus")="islogined" and session("grade")=3 then%><li><a href="javascript:;" id="mdb"><img src="images/icon02.png" title="数据库管理" /><h2>数据库管理</h2></a></li>
    <li><a href="javascript:;" id="code"><img src="images/icon03.png" title="代码设计" /><h2>代码设计</h2></a></li>
    <li><a href="javascript:;" id="gj0"><img src="images/icon04.png" title="常用工具" /><h2>常用工具</h2></a></li>
    <li><a href="javascript:;" id="yun"><img src="images/icon06.png" title="云系统总管理" /><h2>云 管 理</h2></a></li>
    <%end if%>
    <li> <a href="javascript:;" id="piup"><img src="images/icon05.png" title="压缩上传文件并解压" /><h2>压缩上传</h2></a></li>
    </ul>
            
    <div class="topright">    
    <ul>
    <li><span><img src="images/help.png" title="帮助"  class="helpimg"/></span><a href="javascript:;" id="help">帮助</a></li>
    <%if session("loginstatus")="islogined" and session("grade")=3 then%><li>
<a href="javascript:;" id="admin0">用户管理</a></li><%end if%>
    </ul>
     
     
     
    <div class="user">
    <span>管理员</span>
    <a href="loginout.asp?ntime=<%=ntime%>"><i>退出</i>
    <b>X</b></a>
    </div>    
    </div>
</td></tr>
<tr>
<td  bgcolor="#f0f9fd" width="300" align="left" valign="top">
	<div class="lefttop">
	<img src="images/del.gif" class="imgbutton" onmouseover="this.className='imgbt'" onmouseout="this.className='imgbutton'" onclick="delbatch()" alt="删除选中文件"><img src="images/copy.gif" class="imgbutton" onmouseover="this.className='imgbt'" onmouseout="this.className='imgbutton'" onclick="document.delthis.action='admin_fsocopy.asp?path=<%=pathurl%>&page=<%=page%>';document.delthis.act.value='copy';document.delthis.submit();" alt="复制选中文件"><img src="images/cut.gif" class="imgbutton" onmouseover="this.className='imgbt'" onmouseout="this.className='imgbutton'" onclick="document.delthis.action='admin_fsocopy.asp?path=<%=pathurl%>&page=<%=page%>';document.delthis.act.value='cut';document.delthis.submit();" alt="剪切选中文件"><%if session("act")="copy" or session("act")="cut" then%><img src="images/paste.gif" class="imgbutton" onmouseover="this.className='imgbt'" onmouseout="this.className='imgbutton'" onclick="window.location.href='<%=selfname%>?act=paste&path=<%=pathurl%>&page=<%=page%>&ntime=<%=ntime%>';" alt="粘贴文件到当前目录"><%end if%><img src="images/refresh.gif" class="imgbutton" onmouseover="this.className='imgbt'" onmouseout="this.className='imgbutton'" onclick="window.location.reload()" alt="刷新">
	</div>
</td>
<td>
	<div class="place">
	<form method="get" action="<%=selfname%>" style="padding:0;margin:0">
    <span>文件位置：</span>
<input type="text" name="path" style="margin:6px 6px 3px;width:600px;background-color: #ffffff;
border: solid 1px #eff1f5;height:25px;line-height:25px;" value="<%=path%>">
<input type="submit" class="button1" value="跳 转" > <input type="button" value="刷新目录文件" onclick="history.go(0)" class="button0" >
    </form>
    </div>
</td>
</table>
<table>
<tr>
<td  bgcolor="#f0f9fd" width="300" align="left" valign="top">
  <%if obj_fso.folderexists(s_folderpath) then%>  
    <dl class="leftmenu">
        
    <dd>
    <div class="title" style="padding-left:20px;">
  <input type="button" value="全选目录" class="button2" onclick="selectall(document.delthis.folders,true)"> 
<input type="button" value="清除" class="button2" onclick="selectall(document.delthis.folders,false)"> 
<input type="button" value="新建" class="button2" onclick="create_it('createfolder')">
    </div>
    	<div class="menuson" style="padding-left:10px;min-height:440px;">
   <%
	set obj_folder=obj_fso.getfolder(s_folderpath)
	if obj_folder.files.count mod pagesize=0 then
		totalpage=obj_folder.files.count\pagesize
	else
		totalpage=obj_folder.files.count\pagesize+1
	end if
	if page<1 then
		page=1
	end if
	if page>totalpage then
		page=totalpage
	end if
	response.write "<table cellspacing=1 border=0 width=300>"&vbcrlf&"<form name='delthis' method='post' action='admin_fsodelbatch.asp?page="&page&"&path="&pathurl&"'><input type='hidden' name='act' value='delbatch'><input type='hidden' name='basepath' value='"&path&"'>"&vbcrlf&"<tr>"&vbcrlf&"<td colspan=2>"&vbcrlf
	response.write goparent&"<br><br>"&vbcrlf
	for each s_folder in obj_folder.subfolders
		response.write "<tr>"&vbcrlf&"<td width='70%'>"&vbcrlf
		response.write "<input type='checkbox' name='folders' value='"&s_folder.name&"'>&nbsp;<img src='images/folder.gif'> <a href='"&selfname&"?path="&pathurl&server.urlencode(s_folder.name)&"&ntime="&ntime&"'>"&s_folder.name&"</a>"&vbcrlf
		response.write "<td><a href=""javascript:delfile('delfolder','"&s_folder.name&"')"">删除</a> <a href=""javascript:renameit('"&s_folder.name&"','renfolder')"">更名</a></td>"&vbcrlf&"</tr>"&vbcrlf
	next
%>
		</table>
        </div>    
    </dd>
    
    </dl></td>
<td width="100%" align="left" valign="top">
	<div class="place" style="text-align:right;padding-top:5px;padding-right:30px;">
	<input type="button" value="全选文件" class="button0" onclick="selectall(document.delthis.files,true)">　
<input type="button" value="清除选择" class="button0" onclick="selectall(document.delthis.files,false)">　
<input type="button" value="新建文件" class="button0" onclick="create_it('createfile')">　
<input type="button" value="上传文件" class="button0" id="up">
    </div>

    <table class="filetable">
    
    <thead>
    	<tr>
        <th width="31%">名称</th>
        <th width="34%">操作编辑</th>
        <th width="11%">大小</th>
        <th>修改日期</th>
        </tr>    	
    </thead>
    
    <tbody>
    
    <%
	i=1
	startnum=(page-1)*pagesize
	for each s_file in obj_folder.files
		if i>startnum then
			response.write "<tr>"&vbcrlf&"<td width='31%'><input type='checkbox' name='files' value='"&s_file.name&"'>"&vbcrlf
			file_type=getname(s_file,".")
			select case file_type
				case "htm","html"
					response.write "<img src='images/html.gif'> "
				case "css"
					response.write "<img src='images/css.gif'> "
				case "asp","aspx","inc"
					response.write "<img src='images/asp.gif'> "
				case "txt"
					response.write "<img src='images/text.gif'> "
				case "ai"
					response.write "<img src='images/f04.png'> "
				case "js","xml","yml","json","php","bo"
					response.write "<img src='images/f03.png'> "
				case "jpg","gif","png"
					response.write "<img src='images/img.gif'> "
				case "mdb","sql"
					response.write "<img src='images/access.gif'> "
				case "doc"
					response.write "<img src='images/f06.png'> "
				case "xls"
					response.write "<img src='images/f05.png'> "
				case "mid","mp3"
					response.write "<img src='images/midi.gif'> "
				case "zip","rar","exe","7z","ios","tar"
					response.write "<img src='images/f02.png'> "
				case "chm"
					response.write "<img src='images/chm.gif'> "
				case else
					response.write "<img src='images/unknown.gif'> "
			end select%>
		<a href="<%=path&s_file.name%>" target="_blank"><%=s_file.name%></a></td>
			<%	
			select case file_type
				case "zip","rar","tar","exe","gif","jpg","jpe","jpeg","png","bmp","psd","mdb","doc","ppt","xls","mid","mp3","avi","sql","ai","chm","7z","ios","cs","url","config","swf"
					response.write "<td width='24%'><a href=""javascript:dis_edit();"">"
				case "txt","bat","c","htm","html","css","js","vbs","inc","stm","shtm","shtml","asp","asa","aspx","asax","ascx","asmx","aspa","php","php3","php4","jsp","cgi"	
				%>
<td width='34%'>
<script>
    $(document).ready(function() {
    $('#<%=Left(s_file.name,inStrRev(s_file.name,".")-1)&Mid(s_file.name, InstrRev(s_file.name, ".") + 1) 
%>').on('click', function(){
    layer.tab({
    area: ['1000px', '560px'],
    data: [
        {title: '编辑文件', content:'<iframe width=1000 height=520 frameborder=0 scrolling=auto src=admin_fsoedit.asp?path=<%=pathurl&s_file.name%>></iframe>'},<%if session("loginstatus")="islogined" and session("grade")=3 then%>
        {title: '更换编辑界面', content:'<iframe width=1000 height=520 frameborder=0 scrolling=auto src=admin_fsoedit1.asp?path=<%=pathurl&s_file.name%>></iframe>'},
        {title: '代码', content:'<iframe width=1000 height=520 frameborder=0 scrolling=auto src=code/></iframe>'}<%end if%> 
    ]
});
});
});
    </script>
    
    <div class="bozo"><a href="javascript:;" id="<%=Left(s_file.name,inStrRev(s_file.name,".")-1)&Mid(s_file.name, InstrRev(s_file.name, ".") + 1) %>">编辑</a>
	
<%	
				case else
					response.write "<td width='34%'><div class='bozo'><a href="&path&s_file.name&" target='_blank'>编辑</a>"
			end select
			%>
			<%	response.write "<div class='bozo'><a href=""javascript:delfile('delfile','"&s_file.name&"')"">删除</a> <a href=""javascript:renameit('"&s_file.name&"','renfile')"">更名</a> <a href=""javascript:downfile('"&s_file.name&"')"">下载</a></div></td>"&vbcrlf
			dim bozokb
			bozokb=FormatNumber(s_file.size/1024,2,-1)
			response.write "<td width='11%'>"&bozokb&" kb</td>"&vbcrlf
			response.write "<td>"&s_file.datelastmodified&"</td>"&vbcrlf&"</tr>"&vbcrlf
		end if
		if i>startnum+pagesize then
			exit for
		end if
		i=i+1
	next
	%>
	

	
	
		</td>
	</tr>
		<tr>
		<td colspan=4 align=center height=50><%
	if page>1 then
		response.write "<a href='"&selfname&"?path="&pathurl&"&page="&(page-1)&"&ntime="&ntime&"'>上一页</a>　"
	end if
	response.write "共"&obj_folder.files.count&"个文件　当前第 "
	response.write "<select name='jtp' style='line-height:12px;border:none;height:20px;padding:0' onchange="&chr(34)&"window.location.href='"&selfname&"?page='+(this.options.selectedIndex+1)+'&path="&pathurl&"&ntime="&ntime&"'"&chr(34)&">"&vbcrlf
	for i=1 to totalpage
		if i=page then
			response.write "<option selected>"&i&vbcrlf
		else
			response.write "<option>"&i&vbcrlf
		end if
	next
	response.write "</select>"&vbcrlf
	response.write " 页"&vbcrlf
	response.write "　共 "&totalpage&" 页"
	if page<totalpage then
		response.write "　<a href='"&selfname&"?path="&pathurl&"&page="&(page+1)&"&ntime="&ntime&"'>下一页</a> "
	end if
	response.write "</td>"&vbcrlf&"</tr>"&vbcrlf
	response.write "</form>"&vbcrlf&"</table>"&vbcrlf
	set obj_folder=nothing
else
	response.write "<div style='width:100%;padding:25px 0 15px;text-align:center;background:transparent;color:#ff3333;font-weight:600'>文件夹不存在或者你没有访问权限</div>"
end if
set obj_fso=nothing
%>
</td>
</tr>
    
    </tbody>
    
    
    
    
    </table>
    
    
</td>
</tr></table>
</body>
</html>
<!--#include file="admin_fsofoot.asp"-->
