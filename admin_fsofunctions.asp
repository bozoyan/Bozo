<%
function getname(s_string,s_clipchar)
	n_strpos=instrrev(s_string,s_clipchar)
	getname=lcase(right(s_string,len(s_string)-n_strpos))
end function
%>
