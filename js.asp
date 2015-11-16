<script language="javascript">
<!--
function viewfile(path)
{
window.open(path,'','');
}
function editfile(fname)
{
	fname=urlencoding(fname);
	window.open("admin_fsoedit.asp?path=<%=pathurl%>"+fname,'','');
}
function is_edit(fname)
{
	if(window.confirm('你确定该文件为ASCII类型文件，并编辑它吗？'))
		editfile(fname);
}
function dis_edit()
{
	alert('该类型文件为非ASCII文件，无法编辑！');
}
function renameit(fname,act)
{
	var newname=window.prompt("请输入文件(夹) ["+fname+"] 的新名称：","");
	if(newname!=null)
	{
		fname=urlencoding(fname);
		newname=urlencoding(newname);
		window.location.href="admin_fsorename.asp?act="+act+"&path=<%=pathurl%>&fname="+fname+"&newname="+newname+"&epath=<%=pathurl%>&epage=<%=page%>&ntime=<%=ntime%>";
	}
}

function create_it(act)
{
	var fname=window.prompt("请输入要创建的文件(夹)的名称","");
	if(fname!=null)
	{
		fname=urlencoding(fname);
		window.location.href="admin_fsodel.asp?act="+act+"&path=<%=pathurl%>"+fname+"&epath=<%=pathurl%>&epage=<%=page%>&ntime=<%=ntime%>";
	}
}


function downfile(fname)
{
	fname=urlencoding(fname);
	window.open("downfile.asp?path=<%=pathurl%>"+fname,'','');
}
function delbatch()
{
	if (window.confirm("你真的要删除该文件或目录吗？"))
	{
		document.delthis.submit();
	}
}
function delfile(act,fname)
{
	fname=urlencoding(fname);
	if (window.confirm("你真的要删除该文件或目录吗？")==true)
	{
	window.location.href="admin_fsodel.asp?act="+act+"&path=<%=pathurl%>"+fname+"&epath=<%=pathurl%>&epage=<%=page%>&ntime=<%=ntime%>";
	}
}
function selectall(objname,act)
{
	if(objname==null)
		alert("对不起，当前目录下没有文件（夹）可选择/清除");
	else if(objname.length==null)
		objname.checked=act;
	else
	{
		var n=objname.length;
		for (i=0;i<n;i++)
		{
			objname[i].checked=act;
		}
	}
}
//-->
</script>
<script language="vbscript">
<!--
function urlencoding(vstrin)
    dim i,strreturn,strSpecial
    strSpecial = "!""#$%&'()*+,/:;<=>?@[\]^`{|}~%"
    strreturn = ""
    for i = 1 to len(vstrin)
        thischr = mid(vstrin,i,1)
        if abs(asc(thischr)) < &hff then
        	if thischr=" " then
        		strreturn = strreturn & "+"
            elseif instr(strSpecial,thischr)>0 then
                strreturn = strreturn & "%" & hex(asc(thischr))
            else
                strreturn = strreturn & thischr
            end if
        else
            innercode = asc(thischr)
            if innercode < 0 then
                innercode = innercode + &h10000
            end if
            hight8 = (innercode  and &hff00)\ &hff
            low8 = innercode and &hff
            strreturn = strreturn & "%" & hex(hight8) &  "%" & hex(low8)
        end if
    next
    urlencoding = strreturn
end function
' function create_it(act)
'	fname=trim(inputbox("请输入要创建的文件(夹)的名称：","输入待创建的文件(夹)名"))
'	if fname<>"" then
'		fname=urlencoding(fname)
'		window.location.href="admin_fsodel.asp?act="&act&"&path=<%=pathurl%>"&fname&"&epath=<%=pathurl%>&page=<%=page%>&ntime=<%=ntime%>"
'	end if
'end function
//-->
</script>
<script>
    $(document).ready(function() {
    
    $('#mdb').on('click', function(){
    $.layer({
        type: 2,
        title: '数据库管理',
        maxmin: true,
        shadeClose: true, //开启点击遮罩关闭层
        area : ['1000px' , '550px'],
        offset : ['10px', ''],
        iframe: {src: 'tool/DBM.asp'}
    });
});
   
    $('#yun').on('click', function(){
    $.layer({
        type: 2,
        title: '云总管理',
        maxmin: true,
        shadeClose: true, //开启点击遮罩关闭层
        area : ['1000px' , '550px'],
        offset : ['10px', ''],
        iframe: {src: 'tool/yun.php'}
    });
});

 $('#admin0').on('click', function(){
    $.layer({
        type: 2,
        title: '用户管理',
        maxmin: true,
        shadeClose: true, //开启点击遮罩关闭层
        area : ['310px' , '210px'],
        offset : ['10px', ''],
        iframe: {src: 'admin_bozo.asp'}
    });
});

 $('#up').on('click', function(){
    $.layer({
        type: 2,
        title: '文件上传',
        maxmin: true,
        shadeClose: true, //开启点击遮罩关闭层
        area : ['800px' , '500px'],
        offset : ['50px', ''],
        iframe: {src: 'admin_fsoupload.asp?path=<%=pathurl%>'}
    });
});


 $('#piup').on('click', function(){
    $.layer({
        type: 2,
        title: '压缩文件上传解压',
        maxmin: true,
        shadeClose: true, //开启点击遮罩关闭层
        area : ['800px' , '300px'],
        offset : ['50px', ''],
        iframe: {src: 'admin_upload.asp?path=<%=pathurl%>'}
    });
});

 $('#code').on('click', function(){
    $.layer({
        type: 2,
        title: '代码模块设计',
        maxmin: true,
        shadeClose: true, //开启点击遮罩关闭层
        area : ['1000px' , '500px'],
        offset : ['50px', ''],
        iframe: {src: 'code/'}
    });
});


 $('#gj0').on('click', function(){
layer.tab({
    area: ['1250px', '500px'],
    data: [
        {title: '工具箱集合', content:'<iframe width=1240 height=460 frameborder=0 scrolling=auto src=tool/></iframe>'},
        {title: 'ASP探针', content:'<iframe width=1240 height=460 frameborder=0 scrolling=auto src=tool/U.asp></iframe>'},
        {title: 'PHP探针', content:'<iframe width=1240 height=460 frameborder=0 scrolling=auto src=tool/U.php></iframe>'},
        {title: '云管理', content:'<iframe width=1240 height=460 frameborder=0 scrolling=auto src=tool/yun.php></iframe>'}                   
    ]
});

});

$('#help').on('click', function(){
    layer.alert('客服：13033695711（严总监）<br>QQ:201782 &nbsp; yanbo@qq.com',9,'帮助中心');
});

});
    </script>