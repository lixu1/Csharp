<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<title></title>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
</head>
<%
dim action,flag,content
dim Rs,sqlStr
flag=request.querystring("flag")
action=request.form("action")
if action="Add" then
	flag1=request.form("flag")
	content=request.form("content")
	set Rs=server.createobject("adodb.recordset")
	sqlStr="select * from config where flag=" & clng(flag1)
	Rs.open sqlStr,conn,1,3
	if rs.eof then rs.addnew
	rs("flag")=flag1
	Rs("content")=content
	Rs.update
	Rs.close
	call CloseConn()
	response.redirect "Admin_ConfigManage.asp?flag=" & flag1
end if

%>
<body>
<form method="POST" name="myform" onSubmit="return CheckForm();" action="Admin_ConfigManage.asp">
  <table border="0" align="center" cellpadding="0" cellspacing="0" class="border">
    <tr align="center">
      <td>
	<table width="100%" border="0" cellpadding="2" cellspacing="1">
		<%
		set Rs=server.createobject("adodb.recordset")
		sqlStr="select content from config where flag=" & clng(flag)
		Rs.open sqlStr,conn,1,1
		%>
          <tr class="tdbg"> 
            <td width="100" align="right" valign="middle"><p><strong>文章内容：</strong></p>
              <p align="left"><font color="#006600">                &middot;　换行请按Shift+Enter</font><br>
                <font color="#006600">&middot;　另起一段请按Enter</font><br>
              </p></td>
            <td colspan="2"><textarea name="Content" style="display:none"><%if not rs.eof then response.Write Rs("content")%></textarea> 
              <iframe ID="editor" src="WebEditor/ewebeditor.htm?id=content&style=coolblue" frameborder=1 scrolling=no width="600" height="405"></iframe> </td>
          </tr>
         <%
         Rs.close
         %>
        </table>
      </td>
    </tr>
  </table>
  <div align="center"> 
    <p>
      <input name="Action" type="hidden" id="Action" value="Add">
      <input name="flag" type="hidden" value="<%=flag%>">
      <input  name="Add" type="submit"  id="Add" value=" 修 改 "  style="cursor:hand;">
      &nbsp; 
      &nbsp; 
      <input name="Cancel" type="reset" id="Cancel" value=" 取 消 "  style="cursor:hand;">
    </p>
  </div>
</form>
</body>
</html>
<%
call CloseConn()
%>