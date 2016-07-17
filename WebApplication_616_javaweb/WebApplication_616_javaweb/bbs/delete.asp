<%
if session("admin")<>"ok" then response.Redirect("../admin/")
s = request.querystring("s")
p = request.querystring("p")
id = request.querystring("id")
parent = request.querystring("parent")
%>
<html><head><title>Java精品课程网站交流论坛-管理登陆</title>
<LINK href="images/css.css" rel=stylesheet type=text/css>
<script src="images/skins.js"></script></head>
<body bgcolor="#ffffff" background=images/bg.jpg>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="300" align=center>
<tr><td width="2%"><img border="0" src="images/T_left.gif"></td><td width="96%" background="images/Tt_bg.gif"></td><td width="2%"><img border="0" src="images/T_right.gif"></td></tr>
</table>
<TABLE border=0 cellPadding=5 cellSpacing=1 width=300 style="border-collapse: collapse" class=a2 align="center">
<form method="POST" action="delete.asp?id=<%=id%>&s=<%=s%>&p=<%=p%>&parent=<%=parent%>">
<tr class=a1><td align="center" colspan="2"><b>管理员控制</b></td></tr>
<tr class=a4><td>用户：</td><td><input type="text" name="user" size="19"></td></tr>
<tr class=a4><td>密码：</td><td><input type="password" name="pass" size="19"></td></tr>
<tr class=a3><td colspan="2" align="center"><input type="submit" value="确定删除" name="delete"></td></tr></form></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse"  width="300" align=center>
<tr><td width="1%"><img border="0" src="images/T_bottomleft.gif"></td><td width="97%" background="images/T_bottombg.gif"></td><td width="2%"><img border="0" src="images/T_bottomright.gif"></td></tr></table>
<%
if session("admin")="ok" and (p="all") then 
set db = Server.CreateObject("ADODB.Connection")
connect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data.asp")
db.Open connect
sql = "delete from postings where id = " & id & " or connected = " & id
db.Execute(sql)
db.Close
set db=Nothing%>
<script>
opener.location.href="index.asp?s=<%=s%>";
self.close();
</script>
<%end if
if session("admin")="ok" and (p="one") then 
set db = Server.CreateObject("ADODB.Connection")
connect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data.asp")
db.Open connect
sql = "delete from postings where id = " & id
db.Execute(sql)
sql = "update postings set replies = replies -1 where id = " & parent
db.Execute(sql)
db.Close
set db=Nothing%>
<script>
opener.location.reload();
self.close();
</script>
<%end if%>
</body></html>