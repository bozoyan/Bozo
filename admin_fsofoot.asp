<%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
response.write "<!--����ִ��ʱ�䣺"&(endtime-starttime)*1000&"����-->"
%>
