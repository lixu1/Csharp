<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim Rs,sqlStr

if request("Action")="aa" then
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select * from news where newsid="&request("newsid")
	rs.open sql,conn,1,3
	rs("Adddate")=now()
	rs.Update
	rs.Close
	Response.Redirect "admin_NewsManage.asp?page=" & request("page") & "&ClassID=" & request("ClassID")
end if

%>
<html>
<head>
<title>文章管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled==false)
       e.checked = form.chkAll.checked;
    }
  }
function ConfirmDel()
{
	if(document.myform.Action.value=="Del")
	{
		document.myform.action="Admin_NewsDel.asp";
		if(confirm("确定要删除选中的新闻吗？"))
		    return true;
		else
			return false;
	}
}

</SCRIPT>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2"  align="center"><strong>新 闻 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30" ><strong>管理导航：</strong></td>
    <td><a href="Admin_NewsManage.asp">新闻管理首页</a>&nbsp;|&nbsp;<a href="Admin_NewsAdd.asp">添加新闻</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22"><%
    set Rs=server.createobject("adodb.recordset")
    sqlStr="select * from newsclass"
    Rs.open sqlStr,conn,1,1
    response.write "|&nbsp;&nbsp;"
    do while not Rs.eof 
    	response.write "&nbsp;&nbsp<a href='Admin_NewsManage.asp?ClassID="&Rs("ClassID")&"'>" & Rs("ClassName") & "</a>&nbsp;&nbsp;|"
    	Rs.movenext
    loop
    Rs.close
    %></td>
  </tr>
</table>
<br><br>
<table width='80%' border="0" cellpadding="0" cellspacing="0" align=center><tr>
    <form name="myform" method="Post" action="Admin_NewsDel.asp" onSubmit="return ConfirmDel();">
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22"> 
            <td height="22" width="39" align="center"><strong>选中</strong></td>
            <td width="32" align="center"  height="22"><strong>ID</strong></td>
            <td width="129" align="center"  height="22"><strong>类别</strong></td>
            <td align="center"  width=441><strong>新闻标题</strong></td>
            <td width="289" align="center" ><strong>操作</strong></td>
          </tr>
          <%
          dim ClassID,page,i
          ClassID=request("ClassID")
          if ClassID="" then
		      sqlStr="select * from news order by adddate desc, newsid desc"
          else
          	  sqlStr="select * from news where ClassID=" & clng(ClassID) & " order by adddate desc, newsid desc"
          end if
          set Rs=server.createobject("adodb.recordset")
	      'sqlStr="select * from news"
	      Rs.open sqlStr,conn,1,1
		  '此处为分页
		  Rs.pagesize=20
								      
		  page =clng(Request.QueryString("page"))
								      
		  if page<1 then page=1
		  if page>Rs.pagecount then page=Rs.pagecount
								      
		  if not Rs.eof then   
		  rs.AbsolutePage = page
							        
		  for i=1 to Rs.pagesize 
          %>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="39" align="center"><input name='NewsID' type='checkbox' onClick="unselectall()" id="ArticleID" value='<%=Rs("newsid")%>'></td>
            <td width="32" align="center"><%=Rs("newsid")%></td>
            <td width="129" align="center"><a href="Admin_NewsManage.asp?ClassID=<%=Rs("classid")%>"><%=ShowclassName(Rs("classid"))%></a></td>
            <td> <%=Rs("title")%></td>
            <td width="289" align="center"> <%
            	response.write "<a href='Admin_NewsModify.asp?newsID=" & Rs("newsid") &"'>修改</a>&nbsp;&nbsp;"
            	response.write "<a href='Admin_NewsDel.asp?newsID=" & Rs("newsid") & "&Action=Del' onclick='return ConfirmDel();'>删除</a>&nbsp;"
            %>
            <a href="Admin_NewsManage.asp?Action=aa&newsid=<%=Rs("newsid")%>&page=<%=page%>&ClassID=<%=ClassID%>">排序</a>
            </td>
          </tr>
          <%
			  Rs.movenext
			  if rs.EOF then exit for 
		  Next
          else
          	  response.write "<tr class=tdbg><td colspan=4>还没有新闻。</td></tr>"
          end if
          'Rs.close
		  %>
		  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
			<td align=center  colspan=5  height=30>
			<%
			if ClassID<>"" then
				if Rs.RecordCount<>0 then
					if rs.pagecount=1 then
						response.Write "<font color=#666666>第一页&nbsp;&nbsp;"
						response.Write "上一页&nbsp;&nbsp;"
						response.Write "下一页&nbsp;&nbsp;"
						response.Write "最后一页</font>&nbsp;"
					else
					  	if page=1 then
					 		response.Write "<font color=#666666>第一页&nbsp;&nbsp;"
							response.Write "上一页</font>&nbsp;&nbsp;"
					  		response.Write "<a href=?page="& (page+1) &"&ClassID="&ClassID&">下一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & rs.pagecount &"&ClassID="&ClassID&">最后一页</a>&nbsp;"
					  	elseif page=Rs.pagecount then
					 		response.Write "<a href=?page=1&ClassID="&ClassID&">第一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & (page-1) &"&ClassID="&ClassID&">上一页</a>&nbsp;&nbsp;"
					  		response.Write "<font color=#666666>下一页&nbsp;&nbsp;"
							response.Write "最后一页</font>&nbsp;"
					  	else
					  		response.Write "<a href=?page=1&ClassID="&ClassID&">第一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & (page-1) &"&ClassID="&ClassID&">上一页</a>&nbsp;&nbsp;"
					  		response.Write "<a href=?page="& (page+1) &"&ClassID="&ClassID&">下一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & rs.pagecount &"&ClassID="&ClassID&">最后一页</a>&nbsp;"
						end if
					end if
				end if
			else
				if Rs.RecordCount<>0 then
					if rs.pagecount=1 then
						response.Write "<font color=#666666>第一页&nbsp;&nbsp;"
						response.Write "上一页&nbsp;&nbsp;"
						response.Write "下一页&nbsp;&nbsp;"
						response.Write "最后一页</font>&nbsp;"
					else
					  	if page=1 then
					 		response.Write "<font color=#666666>第一页&nbsp;&nbsp;"
							response.Write "上一页</font>&nbsp;&nbsp;"
					  		response.Write "<a href=?page="& (page+1) &">下一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & rs.pagecount &">最后一页</a>&nbsp;"
					  	elseif page=Rs.pagecount then
					 		response.Write "<a href=?page=1>第一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & (page-1) &">上一页</a>&nbsp;&nbsp;"
					  		response.Write "<font color=#666666>下一页&nbsp;&nbsp;"
							response.Write "最后一页</font>&nbsp;"
					  	else
					  		response.Write "<a href=?page=1>第一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & (page-1) &">上一页</a>&nbsp;&nbsp;"
					  		response.Write "<a href=?page="& (page+1) &">下一页</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & rs.pagecount &">最后一页</a>&nbsp;"
						end if
					end if
				end if
			end if
			%>&nbsp;&nbsp;当前页<%=page%>/<%=rs.pagecount%>，&nbsp;共<%=rs.RecordCount%>条纪录 &nbsp;
			<%Rs.close%>
			</td>
		  </tr>
        </table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="250" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              选中本页显示的所有新闻 </td>
    <td><input name="submit" type='submit' value='删除选定的新闻' onClick="document.myform.Action.value='Del'">
              <input name="Action" type="hidden" id="Action" value="Del">
            </td>
  </tr>
</table>
</td>
</form></tr></table>
</body>
</html>
<%
call CloseConn()
%>