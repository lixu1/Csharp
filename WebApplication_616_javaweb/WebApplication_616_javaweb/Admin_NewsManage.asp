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
<title>���¹���</title>
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
		if(confirm("ȷ��Ҫɾ��ѡ�е�������"))
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
    <td height="22" colspan="2"  align="center"><strong>�� �� �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30" ><strong>��������</strong></td>
    <td><a href="Admin_NewsManage.asp">���Ź�����ҳ</a>&nbsp;|&nbsp;<a href="Admin_NewsAdd.asp">�������</a></td>
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
            <td height="22" width="39" align="center"><strong>ѡ��</strong></td>
            <td width="32" align="center"  height="22"><strong>ID</strong></td>
            <td width="129" align="center"  height="22"><strong>���</strong></td>
            <td align="center"  width=441><strong>���ű���</strong></td>
            <td width="289" align="center" ><strong>����</strong></td>
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
		  '�˴�Ϊ��ҳ
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
            	response.write "<a href='Admin_NewsModify.asp?newsID=" & Rs("newsid") &"'>�޸�</a>&nbsp;&nbsp;"
            	response.write "<a href='Admin_NewsDel.asp?newsID=" & Rs("newsid") & "&Action=Del' onclick='return ConfirmDel();'>ɾ��</a>&nbsp;"
            %>
            <a href="Admin_NewsManage.asp?Action=aa&newsid=<%=Rs("newsid")%>&page=<%=page%>&ClassID=<%=ClassID%>">����</a>
            </td>
          </tr>
          <%
			  Rs.movenext
			  if rs.EOF then exit for 
		  Next
          else
          	  response.write "<tr class=tdbg><td colspan=4>��û�����š�</td></tr>"
          end if
          'Rs.close
		  %>
		  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
			<td align=center  colspan=5  height=30>
			<%
			if ClassID<>"" then
				if Rs.RecordCount<>0 then
					if rs.pagecount=1 then
						response.Write "<font color=#666666>��һҳ&nbsp;&nbsp;"
						response.Write "��һҳ&nbsp;&nbsp;"
						response.Write "��һҳ&nbsp;&nbsp;"
						response.Write "���һҳ</font>&nbsp;"
					else
					  	if page=1 then
					 		response.Write "<font color=#666666>��һҳ&nbsp;&nbsp;"
							response.Write "��һҳ</font>&nbsp;&nbsp;"
					  		response.Write "<a href=?page="& (page+1) &"&ClassID="&ClassID&">��һҳ</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & rs.pagecount &"&ClassID="&ClassID&">���һҳ</a>&nbsp;"
					  	elseif page=Rs.pagecount then
					 		response.Write "<a href=?page=1&ClassID="&ClassID&">��һҳ</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & (page-1) &"&ClassID="&ClassID&">��һҳ</a>&nbsp;&nbsp;"
					  		response.Write "<font color=#666666>��һҳ&nbsp;&nbsp;"
							response.Write "���һҳ</font>&nbsp;"
					  	else
					  		response.Write "<a href=?page=1&ClassID="&ClassID&">��һҳ</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & (page-1) &"&ClassID="&ClassID&">��һҳ</a>&nbsp;&nbsp;"
					  		response.Write "<a href=?page="& (page+1) &"&ClassID="&ClassID&">��һҳ</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & rs.pagecount &"&ClassID="&ClassID&">���һҳ</a>&nbsp;"
						end if
					end if
				end if
			else
				if Rs.RecordCount<>0 then
					if rs.pagecount=1 then
						response.Write "<font color=#666666>��һҳ&nbsp;&nbsp;"
						response.Write "��һҳ&nbsp;&nbsp;"
						response.Write "��һҳ&nbsp;&nbsp;"
						response.Write "���һҳ</font>&nbsp;"
					else
					  	if page=1 then
					 		response.Write "<font color=#666666>��һҳ&nbsp;&nbsp;"
							response.Write "��һҳ</font>&nbsp;&nbsp;"
					  		response.Write "<a href=?page="& (page+1) &">��һҳ</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & rs.pagecount &">���һҳ</a>&nbsp;"
					  	elseif page=Rs.pagecount then
					 		response.Write "<a href=?page=1>��һҳ</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & (page-1) &">��һҳ</a>&nbsp;&nbsp;"
					  		response.Write "<font color=#666666>��һҳ&nbsp;&nbsp;"
							response.Write "���һҳ</font>&nbsp;"
					  	else
					  		response.Write "<a href=?page=1>��һҳ</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & (page-1) &">��һҳ</a>&nbsp;&nbsp;"
					  		response.Write "<a href=?page="& (page+1) &">��һҳ</a>&nbsp;&nbsp;"
							response.Write "<a href=?page=" & rs.pagecount &">���һҳ</a>&nbsp;"
						end if
					end if
				end if
			end if
			%>&nbsp;&nbsp;��ǰҳ<%=page%>/<%=rs.pagecount%>��&nbsp;��<%=rs.RecordCount%>����¼ &nbsp;
			<%Rs.close%>
			</td>
		  </tr>
        </table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="250" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ѡ�б�ҳ��ʾ���������� </td>
    <td><input name="submit" type='submit' value='ɾ��ѡ��������' onClick="document.myform.Action.value='Del'">
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