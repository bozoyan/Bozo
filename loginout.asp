<%
session("loginname")=""
session("grade")=""
session("pathaccess")=""
session("loginstatus")=""
response.redirect "index.asp?ntime="&server.urlencode(now())
%>
