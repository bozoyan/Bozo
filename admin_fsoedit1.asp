<!--#include file="admin_fsoconfig.asp"-->
<!--#include file="checklogin.asp"-->
<%
if request.form("act")="editfile" then
	fcontent=request.form("fcontent")
	if path="" or instr(path,":")>0 or instr(path,"\")>0 then
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "alert('指定的文件名或路径中含有非法字符');"&vbcrlf
		response.write "window.close();"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
		response.end
	end if
	if len(path)>1 and right(path,1)="/" then
		path=left(path,len(path)-1)
	end if
	filepath=server.mappath(path)
	set obj_fso=server.createobject("scripting.filesystemobject")
	if obj_fso.fileexists(filepath) then
		set obj_file=obj_fso.opentextfile(filepath,2,false)
		obj_file.write fcontent
		obj_file.close
		set obj_file=nothing
     	response.write "<script language='javascript'>alert('编辑成功');</script>"&vbcrlf
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "parent.location.reload();"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	else
		response.write "<script language='javascript'>"&vbcrlf
		response.write "<!--"&vbcrlf
		response.write "alert('找不到指定的文件，可能该文件已经被删除');"&vbcrlf
		response.write "window.close();"&vbcrlf
		response.write "//-->"&vbcrlf
		response.write "</script>"&vbcrlf
	end if
	response.end

end if
%>
<!doctype html>
<head>
<title><%=sysname%>--文件编辑</title>
<meta http-equiv="Content-Type" content="text/html;charset=gbk">
<meta name="author" content="bozoyan;EMAIL:yanbo@qq.com">
<style type="text/css">
<!--
body,table,td{font-size:12px;color:#000090;}
body{margin:0;background: #f5f9fb;}
table{width:760px}
table table{width:100%;}
td{vertical-align:top;padding-left:12px;}
td td{padding-left:4px;vertical-align:middle;}
img{vertical-align:bottom}
form,p{margin:0;padding:0}
a{color:#000080;text-decoration:none;}
a:hover{color:#ff3333;text-decoration:underline}

.button{width:100px;height:25px;border:1px solid #B5BBBF; background-color: #fff;margin:3px 0;}
.botton a{
padding: 0 20px;
height: 25px;
line-height: 23px;
margin-right: 10px;
float: left;
padding-top:6px;
text-align: center;
display: block;
color:#000000;
border: 1px solid #aaa;
background: #F2F2F2;
cursor: pointer;}
      .CodeMirror {border: 1px solid black; font-size:13px}
//-->
</style>

<link rel="stylesheet" href="css/codemirror.css">
<link rel="stylesheet" href="css/3024-day.css">
<link rel="stylesheet" href="css/3024-night.css">
<link rel="stylesheet" href="css/ambiance.css">
<link rel="stylesheet" href="css/base16-dark.css">
<link rel="stylesheet" href="css/base16-light.css">
<link rel="stylesheet" href="css/blackboard.css">
<link rel="stylesheet" href="css/cobalt.css">
<link rel="stylesheet" href="css/eclipse.css">
<link rel="stylesheet" href="css/elegant.css">
<link rel="stylesheet" href="css/erlang-dark.css">
<link rel="stylesheet" href="css/lesser-dark.css">
<link rel="stylesheet" href="css/mbo.css">
<link rel="stylesheet" href="css/mdn-like.css">
<link rel="stylesheet" href="css/midnight.css">
<link rel="stylesheet" href="css/monokai.css">
<link rel="stylesheet" href="css/neat.css">
<link rel="stylesheet" href="css/neo.css">
<link rel="stylesheet" href="css/night.css">
<link rel="stylesheet" href="css/paraiso-dark.css">
<link rel="stylesheet" href="css/paraiso-light.css">
<link rel="stylesheet" href="css/pastel-on-dark.css">
<link rel="stylesheet" href="css/rubyblue.css">
<link rel="stylesheet" href="css/solarized.css">
<link rel="stylesheet" href="css/the-matrix.css">
<link rel="stylesheet" href="css/tomorrow-night-eighties.css">
<link rel="stylesheet" href="css/twilight.css">
<link rel="stylesheet" href="css/vibrant-ink.css">
<link rel="stylesheet" href="css/xq-dark.css">
<link rel="stylesheet" href="css/xq-light.css">
<script src="css/codemirror.js"></script>
<script src="src/javascript/javascript.js"></script>
<script src="src/selection/active-line.js"></script>
<script src="src/edit/matchbrackets.js"></script>
<head>
<body>
<script>
function add()
{
var ress=document.forms[0].java.value
window.location="view-source:"+ress;
}
</script>
<table align="center" cellspacing=0>
<tr>
<td style="padding-top:10px;">
<form method="get" action="<%=selfname%>">
<input type="text" name="path" style="margin:6px 6px 3px 0px;width:615px;height:25px;border:solid 1px #000099;float:left;" <%
path=trim(request.querystring("path"))
response.write "value='"&path&"' readonly> "
response.write "<div class='botton'><a href="&path&" target=_blank>查看文件</a></div>"
response.write "</form>"&vbcrlf&"</td>"&vbcrlf&"</tr>"&vbcrlf&"<tr>"&vbcrlf&"<td>"&vbcrlf
if path="" or instr(path,":")>0 or instr(path,"\")>0 then
	response.write "　　　　<font color=#ff3333>你要编辑的文件名或路径错误！</font>"
	response.end
end if
filepath=server.mappath(path)
set obj_fso=server.createobject("scripting.filesystemobject")
if obj_fso.fileexists(filepath) then
	set obj_file=obj_fso.opentextfile(filepath,1,false)
	if not obj_file.atendofstream then
		s_content=replace(obj_file.readall(),"&","&amp;")
		s_content=replace(s_content,"<","&lt;")
		s_content=replace(s_content,">","&gt;")
	end if
	obj_file.close
	set obj_file=nothing
	response.write "<form action='"&selfname&"?path="&server.urlencode(path)&"' method=post>"&vbcrlf
	response.write "<input type='hidden' name='act' value='editfile'>"&vbcrlf
	response.write "<input type='hidden' name='path' value='"&path&"'>"&vbcrlf
	%>
<article>
<textarea id="code" name="fcontent">
<%=s_content%>
</textarea>
<p style="padding-top:10px;">选择界面风格: <select onchange="selectTheme()" id=select>
    <option selected>cobalt</option>
    <option>pastel-on-dark</option>
    <option>default</option>
    <option>3024-day</option>
    <option>3024-night</option>
    <option>ambiance</option>
    <option>base16-dark</option>
    <option>base16-light</option>
    <option>blackboard</option>
    <option>eclipse</option>
    <option>elegant</option>
    <option>erlang-dark</option>
    <option>lesser-dark</option>
    <option>mbo</option>
    <option>mdn-like</option>
    <option>midnight</option>
    <option>monokai</option>
    <option>neat</option>
    <option>neo</option>
    <option>night</option>
    <option>paraiso-dark</option>
    <option>paraiso-light</option>
    <option>rubyblue</option>
    <option>solarized dark</option>
    <option>solarized light</option>
    <option>the-matrix</option>
    <option>tomorrow-night-eighties</option>
    <option>twilight</option>
    <option>vibrant-ink</option>
    <option>xq-dark</option>
    <option>xq-light</option>
</select>
</p>

<script>
  var editor = CodeMirror.fromTextArea(document.getElementById("code"), {
    lineNumbers: true,
    styleActiveLine: true,
    matchBrackets: true
  });
  var input = document.getElementById("select");
  function selectTheme() {
    var theme = input.options[input.selectedIndex].innerHTML;
    editor.setOption("theme", theme);
  }
  var choice = document.location.search &&
               decodeURIComponent(document.location.search.slice(1));
  if (choice) {
    input.value = choice;
    editor.setOption("theme", choice);
  }
</script>
  </article>



	<% response.write "<p style='text-align:right;padding-right:24px;'><input type='submit' class='button' value='保存编辑'>　<input type='reset' class='button' value='复原'></p>"&vbcrlf
	response.write "</form>"&vbcrlf
else
	response.write "　　　　<font color=#ff3333>你要编辑的文件不存在</font>"
end if
set obj_fso=nothing

%>
</td>
</tr>
</table>
</body>
</html>
<!--#include file="admin_fsofoot.asp"-->
