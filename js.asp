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
	if(window.confirm('��ȷ�����ļ�ΪASCII�����ļ������༭����'))
		editfile(fname);
}
function dis_edit()
{
	alert('�������ļ�Ϊ��ASCII�ļ����޷��༭��');
}
function renameit(fname,act)
{
	var newname=window.prompt("�������ļ�(��) ["+fname+"] �������ƣ�","");
	if(newname!=null)
	{
		fname=urlencoding(fname);
		newname=urlencoding(newname);
		window.location.href="admin_fsorename.asp?act="+act+"&path=<%=pathurl%>&fname="+fname+"&newname="+newname+"&epath=<%=pathurl%>&epage=<%=page%>&ntime=<%=ntime%>";
	}
}

function create_it(act)
{
	var fname=window.prompt("������Ҫ�������ļ�(��)������","");
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
	if (window.confirm("�����Ҫɾ�����ļ���Ŀ¼��"))
	{
		document.delthis.submit();
	}
}
function delfile(act,fname)
{
	fname=urlencoding(fname);
	if (window.confirm("�����Ҫɾ�����ļ���Ŀ¼��")==true)
	{
	window.location.href="admin_fsodel.asp?act="+act+"&path=<%=pathurl%>"+fname+"&epath=<%=pathurl%>&epage=<%=page%>&ntime=<%=ntime%>";
	}
}
function selectall(objname,act)
{
	if(objname==null)
		alert("�Բ��𣬵�ǰĿ¼��û���ļ����У���ѡ��/���");
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
'	fname=trim(inputbox("������Ҫ�������ļ�(��)�����ƣ�","������������ļ�(��)��"))
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
        title: '���ݿ����',
        maxmin: true,
        shadeClose: true, //����������ֹرղ�
        area : ['1000px' , '550px'],
        offset : ['10px', ''],
        iframe: {src: 'tool/DBM.asp'}
    });
});
   
    $('#yun').on('click', function(){
    $.layer({
        type: 2,
        title: '���ܹ���',
        maxmin: true,
        shadeClose: true, //����������ֹرղ�
        area : ['1000px' , '550px'],
        offset : ['10px', ''],
        iframe: {src: 'tool/yun.php'}
    });
});

 $('#admin0').on('click', function(){
    $.layer({
        type: 2,
        title: '�û�����',
        maxmin: true,
        shadeClose: true, //����������ֹرղ�
        area : ['310px' , '210px'],
        offset : ['10px', ''],
        iframe: {src: 'admin_bozo.asp'}
    });
});

 $('#up').on('click', function(){
    $.layer({
        type: 2,
        title: '�ļ��ϴ�',
        maxmin: true,
        shadeClose: true, //����������ֹرղ�
        area : ['800px' , '500px'],
        offset : ['50px', ''],
        iframe: {src: 'admin_fsoupload.asp?path=<%=pathurl%>'}
    });
});


 $('#piup').on('click', function(){
    $.layer({
        type: 2,
        title: 'ѹ���ļ��ϴ���ѹ',
        maxmin: true,
        shadeClose: true, //����������ֹرղ�
        area : ['800px' , '300px'],
        offset : ['50px', ''],
        iframe: {src: 'admin_upload.asp?path=<%=pathurl%>'}
    });
});

 $('#code').on('click', function(){
    $.layer({
        type: 2,
        title: '����ģ�����',
        maxmin: true,
        shadeClose: true, //����������ֹرղ�
        area : ['1000px' , '500px'],
        offset : ['50px', ''],
        iframe: {src: 'code/'}
    });
});


 $('#gj0').on('click', function(){
layer.tab({
    area: ['1250px', '500px'],
    data: [
        {title: '�����伯��', content:'<iframe width=1240 height=460 frameborder=0 scrolling=auto src=tool/></iframe>'},
        {title: 'ASP̽��', content:'<iframe width=1240 height=460 frameborder=0 scrolling=auto src=tool/U.asp></iframe>'},
        {title: 'PHP̽��', content:'<iframe width=1240 height=460 frameborder=0 scrolling=auto src=tool/U.php></iframe>'},
        {title: '�ƹ���', content:'<iframe width=1240 height=460 frameborder=0 scrolling=auto src=tool/yun.php></iframe>'}                   
    ]
});

});

$('#help').on('click', function(){
    layer.alert('�ͷ���13033695711�����ܼࣩ<br>QQ:201782 &nbsp; yanbo@qq.com',9,'��������');
});

});
    </script>