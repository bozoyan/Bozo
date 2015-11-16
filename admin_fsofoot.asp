<%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
response.write "<!--程序执行时间："&(endtime-starttime)*1000&"毫秒-->"
%>
