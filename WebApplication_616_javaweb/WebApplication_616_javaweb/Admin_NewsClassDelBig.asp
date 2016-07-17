<%@language=vbscript codepage=936 %>
<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim BigClassName,sql
BigClassName=trim(Request("BigClassName"))
BigClassID=trim(Request("BigClassID"))
if BigClassName<>"" and BigClassID<>"" then
	sql="delete from NewsClass where ClassName='" & BigClassName & "'"
	conn.Execute sql
	sql="delete from News where ClassID=" & BigClassID
	conn.Execute sql
end if
call CloseConn()
response.redirect "Admin_NewsClassManage.asp"
%>


