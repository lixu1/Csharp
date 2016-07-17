<!--#include file="Replace.asp"-->
<html><head><title>Java精品课程网站交流论坛</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/css.css" rel=stylesheet type=text/css>
<script src="images/skins.js"></script></head>
<body bgcolor="#ffffff" background=images/bg.jpg>
<SCRIPT>valigntop()</SCRIPT> 
<table cellspacing="1" cellpadding="4" width="760" border="0" class="a2" bgcolor="#FFFFFF" align="center">
<tr><td colspan = "5"  class="a4" height="90"><center><img src=images/banner.jpg></center></td></tr>
<tr class="a1"><td width="54%"><a href="#write"><font color="#FF0000">【点这里发表帖子】</font></a>　【帖子标题】</td><td width="13%" align="center">【发贴人】</td><td width="8%" align="center">【浏览】</td>
<td width="8%" align="center">【回复】</td><td width="17%" align="center">【最后更新】</td></tr>      
<%s = request.querystring("s")
if len(s)<1 then s = 0
s=s*20
set db = Server.CreateObject("ADODB.Connection")
connect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data.asp")
'connect="Driver={Microsoft Access Driver (*.mdb)}; DBQ="& server.mappath("data.asp")
db.Open connect
sql = "select title, autor, mail, lastdate, views, replies, datum_updated, id from postings where connected = 0 order by updated desc" 
set rs = db.Execute(sql)
while not rs.eof
lastdate=rs("lastdate")
shicha=DateDiff("n",lastdate,date) 
count = count +1
if count>s and count<=(s+20) then
if count mod 2 = 0 then 
bgcolor = "a3"
else
bgcolor = "a4" 
end if%>
<tr class="<%=bgcolor%>">
<td>&nbsp;<img src=images/icon.gif> <a class="text" href="view.asp?id=<%=rs("id")%>&s=<%=s/20%>"><%=formatierung(rs("title"))%></a><%if shicha<=1 then %> <img src=images/new.gif><%end if%><%if rs("views")>=50 then%> <img src=images/icon4.gif><%end if%><%if rs("replies")>=10 then%> <img src=images/icon3.gif><%end if%></td>
<td align="center"><%if len(rs("mail"))>1 then%><a class="text" href="mailto:<%=formatierung(rs("mail"))%>"><%=formatierung(rs("autor"))%></a><%else%><%=formatierung(rs("autor"))%><%end if%></td>
<td align="center"><font color="#0000FF"><%=rs("views")%></font></td><td align="center">
<font color="#FF0000"><%=rs("replies")%></font></td><td align="center"><%=rs("datum_updated")%></td></tr>
<%end if
rs.movenext
wend
sql = "select count(id) as alleEintraege from postings where connected = 0" 
set rs = db.Execute(sql)
alleEintraege = int(rs("alleEintraege"))%>
<tr class="a4"><td width="100%" colspan = "5"><table width=100%><tr><td width=55%><font color="#FF0000">图片说明：</font>【今日新帖】 <img src=images/new.gif> 【点击超过50】 <img src=images/icon4.gif>  【回复超过10】 <img src=images/icon3.gif></td>
<td align="right">跳到 <select onChange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<% for t = 0 to fix((alleEintraege-1)/20)%> 
<option value="index.asp?s=<%=t%>" <%if t=(s/20) then%>selected<%end if%>><%=t+1%></option>  
<%next%></select> 页　当前第<font color=red> <%=s/20+1%> </font>页　<% if s>0 then %><a href=index.asp?s=<%=s/20-1%>>[上一页]</a>&nbsp;<%end if%><% if fix((alleEintraege-1)/20)-s/20 <> 0 then %><a href="index.asp?s=<%=s/20+1%>">[下一页]</a><%end if%></td></tr></table></td></tr></table>
<%rs.close
db.close
set rs = nothing
set db = nothing%>
<SCRIPT>valignbottom()</SCRIPT> 
<br>
<SCRIPT>valigntop()</SCRIPT> 
<table cellspacing="1" cellpadding="5" width="760" border="0" class="a2"  bgcolor="#FFFFFF" align="center">
<form method="POST" action="new.asp">
<tr class="a1"><td align="center" colspan="2"><a name=#write></a><b>新发帖/新留言 ( 带 
	<font color="#FF0000">*</font> 为必填选项 )</b></td></tr>      
<tr class="a4"><td width="120">　发贴者姓名：</td><td>&nbsp;<input type="text" name="autor" size="21"  maxlength="8" value="[游客]"> <font color="#FF0000">*</font> （<font color="#0000FF">无须注册即可发表</font>）</td></tr>      
<tr class="a3"><td>　发贴者邮箱：</td><td>&nbsp;<input type="text" name="mail" size="34" maxlength="25"></td></tr>
<tr class="a4"><td>　帖子标题：</td><td>&nbsp;<input type="text" name="title" size="34"  maxlength="25"> <font color="#FF0000">*</font></td></tr>
<tr class="a3"><td>　帖子内容：</td><td>&nbsp;<textarea rows="6" name="message" cols="72"></textarea> <font color="#FF0000">*</font></td></tr>
<tr class="a4"><td colspan="2" align="center"><input type="submit" value="确认发表" name="B1"> 　　<input type="reset" value="取消发送" name="B2"></td></tr>
</form></table>
<SCRIPT>valignbottom()</SCRIPT> 
<%call copyrights%>
</body></html>