<!--#include file="Replace.asp"-->
<%id = request.querystring("id")
s = request.querystring("s")
set db = Server.CreateObject("ADODB.Connection")
connect="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data.asp")
db.Open connect
sql = "update postings set views = views + 1 where id = " & id
db.Execute(sql)
sql = "select title, id, autor, mail, views, replies, datum, datum_updated, message from postings where id = " & id 
set rs = db.Execute(sql)%>
<html><head><title>��ӭ����΢����̳-TBBS-<%=rs("title")%></title>
<LINK href="images/css.css" rel=stylesheet type=text/css>
<script src="images/skins.js"></script></head>
<body bgcolor="#ffffff" background=images/bg.jpg>
<script>
function delete_eintrag(id, p, s, parent) {
	msg=window.open("delete.asp?id="+id+"&p="+p+"&s="+s+"&parent="+parent);
}
</script>
<%while not rs.eof
retitle = rs("title") 
parent = rs("id")%>
<SCRIPT>valigntop()</SCRIPT> 
<table cellspacing="1" cellpadding="4" width="760" border="0" class="a2" bgcolor="#FFFFFF" align="center">
<tr><td align="center" colspan = "4"  class="a4" height="60"><!--<font face="����" style="font-size: 22pt"><font color="#0060BF">˼���</font><font color="#FF0000"> ΢����̳</font> TBBS</font>--><img src=images/banner.jpg></td></tr>
<tr class="a1"><td colspan="4"><table width=100%><tr><td><a href="index.asp?s=<%=s%>"><font color="#FF0000"><b><< ������̳</b></font></a>��<b>���⣺<font color="#0000FF"><%=formatierung(rs("title"))%></font></b></td><td width=10% align=center><font color="#FF0000">[ ¥ �� ]</font></td></tr></table></td></tr>
<tr class="a3"><td width="35%">&nbsp;���ߣ�<%=formatierung(rs("autor"))%></td><td width="25%"><img src=images/email.gif> <%=formatierung(rs("mail"))%></td>
<td width="20%" align="center"><img src=images/time.gif> <%=rs("datum")%></td>
<td width="20%" align="center"><a href=#replay><img src=images/replay.gif border=0></a>��<a href="javascript:delete_eintrag('<%=id%>','all','<%=s%>','<%=parent%>');"><img src=images/del.gif border=0></a></td></tr>
<tr class="a4"><td colspan="4" height=40 valign=top class=365ye><img src=images/icon.gif> <%=formatierung(rs("message"))%></td></tr>
<%rs.movenext
wend
sql = "select title, autor, mail, views, replies, datum, datum_updated, id, message from postings where connected = "& id &" order by id" 
set rs = db.Execute(sql)
while not rs.eof
count = count +1%>
<tr class="a1"><td colspan="4"><table width=100%><tr><td>&nbsp;<b><%=formatierung(rs("title"))%></b></td><td width=10% align=center>�� <font color="#FF0000"> <%=count%></font> ¥</td></tr></table></td></tr>
<tr class="a3"><td>&nbsp;���ߣ�<%=formatierung(rs("autor"))%></td><td><img src=images/email.gif> <%=formatierung(rs("mail"))%></td>
<td align="center"><img src=images/time.gif> <%=rs("datum")%></td>
<td align="center"><a href=#replay><img src=images/replay.gif border=0></a>��<a href="javascript:delete_eintrag('<%=rs("id")%>','one','<%=s%>','<%=parent%>');"><img src=images/del.gif border=0></a></td></tr>
<tr class="a4"><td colspan="4" height=40 valign=top class=365ye><img src=images/icon2.gif> <%=formatierung(rs("message"))%></td></tr>
<%rs.movenext
wend
rs.close
db.close
set rs = nothing
set db = nothing%></table>
<SCRIPT>valignbottom()</SCRIPT> 
<br>
<SCRIPT>valigntop()</SCRIPT> 
<table cellspacing="1" cellpadding="5" width="760" border="0" class="a2" bgcolor="#FFFFFF" align="center">
<form method="POST" action="answer.asp?id=<%=id%>"><tr class="a1">
<td align="center" colspan="2"><b>�ظ���������</b><a name=#replay></a><b> ( �� 
<font color="#FF0000">*</font> Ϊ����ѡ�� )</b></td></tr>
<tr class="a4"><td width="120">��������������</td><td>&nbsp;<input type="text" name="autor" size="21"  maxlength="8"> <font color="#FF0000">*</font> ��<font color="#0000FF">����ע�ἴ�ɷ���</font>��</td></tr>      
<tr class="a3"><td>�����������䣺</td><td>&nbsp;<input type="text" name="mail" size="34" maxlength="25"></td></tr>
<tr class="a4"><td>�����ӱ��⣺</td><td>&nbsp;<input type="text" name="title" size="34"  maxlength="25" value="�ظ���<%=retitle%>"> <font color="#FF0000">*</font></td></tr>
<tr class="a3"><td>���������ݣ�</td><td>&nbsp;<textarea rows="6" name="message" cols="72"></textarea> <font color="#FF0000">*</font></td></tr>
<tr class="a4"><td colspan="2" align="center"><input type="submit" value="ȷ�Ϸ���" name="B1"> ����<input type="reset" value="ȡ������" name="B2"></td></tr>
</form></table>
<SCRIPT>valignbottom()</SCRIPT> 
<%call copyrights%>
</body></html>