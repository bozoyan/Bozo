<%
dim conn,rs,sql
'set conn=server.createobject("adodb.connection")
'dbpath=server.mappath("admin_fsodb.asp")
'conn.open "Provider=Microsoft.jet.OLEDB.4.0;Data Source="&dbpath


set conn = server.createobject("adodb.connection") 
conn.open "admin" 
%>
