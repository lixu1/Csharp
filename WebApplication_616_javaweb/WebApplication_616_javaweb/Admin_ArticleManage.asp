<!--#include file="inc/conn.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim strFileName,FileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim i,j
dim keyword,strField
dim sql,Rs,sqlStr
dim BigID,SmallID,ClassID
FileName="Admin_ArticleManage.asp"

BigID=Trim(request.querystring("BigID"))
if BigID<>"" then
	BigID=clng(BigID)
else
	BigID=0
end if
strFileName=FileName & "?BigID=" & BigID

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
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
		document.myform.action="Admin_ArticleDel.asp";
		if(confirm("ȷ��Ҫɾ��ѡ�е������𣿱���������ѡ�е����³���ɾ�������ɻָ���"))
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
    <td height="22" colspan="2"  align="center"><strong>�� Ϣ �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="116" height="30" ><strong>��Ϣ��������</strong></td>
    <td width="876"><a href="Admin_ArticleManage.asp">��Ϣ������ҳ</a>&nbsp;|&nbsp;<a href="Admin_ArticleAdd.asp">�����Ϣ</a></td>
  </tr>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22"><%call Admin_ShowBigClass()%></td>
  </tr>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="22"><%
	call Admin_ShowPath2(BigID)%></td>
    <td width="200" height="22" align="right">
	
    </td>
  </tr>
</table>
<%
sqlStr="select * from Article where articleid>0"
if BigID>0 then
	sqlStr=sqlStr & " and BigID=" & BigID
end if
sqlStr=sqlStr & " order by adddate desc,articleid desc"
Set Rs= Server.CreateObject("ADODB.Recordset")
Rs.open sqlStr,conn,1,1
if Rs.eof and Rs.bof then
		response.write "<p align='center'><br>�����û���κ���Ϣ��<br></p>"
else
   	totalPut=Rs.recordcount
	if currentpage<1 then
   		currentpage=1
   	end if
   	if (currentpage-1)*MaxPerPage>totalput then
   		if (totalPut mod MaxPerPage)=0 then
     		currentpage= totalPut \ MaxPerPage
	  	else
	      	currentpage= totalPut \ MaxPerPage + 1
   		end if
   	end if
    if currentPage=1 then
       	showContent
       	showpage strFileName,totalput,MaxPerPage,true,true,"����Ϣ"
 	else
     	if (currentPage-1)*MaxPerPage<totalPut then
       	   	Rs.move  (currentPage-1)*MaxPerPage
       		dim bookmark
           	bookmark=Rs.bookmark
            showContent
            showpage strFileName,totalput,MaxPerPage,true,true,"����Ϣ"
       	else
	        currentPage=1
           	showContent
          	showpage strFileName,totalput,MaxPerPage,true,true,"����Ϣ"
	    end if
	end if
end if
Rs.close
set Rs=nothing  


sub showContent
   	dim ArticleNum
    ArticleNum=0
%>
<table width='90%' border="0" cellpadding="0" cellspacing="0" align="center"><tr>
    <form name="myform" method="Post" action="Admin_ArticleDel.asp" onSubmit="return ConfirmDel();">
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22"> 
            <td height="22" width="42" align="center"><strong>ѡ��</strong></td>
            <td width="32" align="center"  height="22"><strong>ID</strong></td>
            <td width="100" align="center"><strong>����</strong></td>
            <td width="413" align="center" ><strong>����</strong></td>
            <td width="108" align="center" ><strong>�����</strong></td>
            <td width="125" align="center" ><strong>����</strong></td>
            <td width="225" align="center" ><strong>����</strong></td>
          </tr>
          <%do while not Rs.eof%>
          <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
            <td width="42" align="center"><input name='ArticleID' type='checkbox' onClick="unselectall()" id="ArticleID" value='<%=cstr(Rs("articleID"))%>'></td>
            <td align="center"><%=Rs("articleid")%></td>
            <td align="center"><a href="Admin_ArticleManage.asp?BigID=<%=rs("bigid")%>"><%=ShowBigName(rs("bigid"))%></a></td>
            <td> <%
			if Rs("Defaultpic")<>"" then
				response.write "<font color=blue>[ͼ��]</font>"
			end if
			response.write  Rs("title") 
			%></td>
            <td width="108" align="center"><%= Rs("Hits") %></td>
            <td width="125" align="center"> <%
			if Rs("OnTop")=true then
				response.Write "<font color=blue>��</font> "
			else
				response.write "&nbsp;&nbsp;&nbsp;"
			end if
			if Rs("Elite")=true then
				response.write "<font color=green>��</a>"
			else
				response.write "&nbsp;&nbsp;"
			end if
			%> </td>
            <td width="225" align="center"> <%
            	response.write "<a href='Admin_ArticleModify.asp?ArticleID=" & Rs("articleid") &"'>�޸�</a>&nbsp;"
            	response.write "<a href='Admin_ArticleDel.asp?ArticleID=" & Rs("ArticleID") & "&Action=Del' onclick='return ConfirmDel();'>ɾ��</a>&nbsp;"
           		response.write "<a href='Admin_ArticleProperty.asp?ArticleID=" & Rs("ArticleID") & "&Action=aa&page="&request("page")&"&bigid="&bigid&"'>����</a>&nbsp;"
				if Rs("OnTop")=False then	
					'response.write "<a href='Admin_ArticleProperty.asp?ArticleID=" & Rs("ArticleID") & "&Action=SetOnTop'>�̶�</a>&nbsp;"
				else
					'response.write "<a href='Admin_ArticleProperty.asp?ArticleID=" & Rs("ArticleID") & "&Action=CancelOnTop'>���</a>&nbsp;"
				end if
            	if Rs("Elite")=False then	
					response.write "<a href='Admin_ArticleProperty.asp?ArticleID=" & Rs("ArticleID") & "&Action=SetElite'>��Ϊ�Ƽ�</a>"
				else
					response.write "<a href='Admin_ArticleProperty.asp?ArticleID=" & Rs("ArticleID") & "&Action=CancelElite'>ȡ���Ƽ�</a>"
				end if
            %></td>
          </tr>
          <%
		ArticleNum=ArticleNum+1
	   	if ArticleNum>=MaxPerPage then exit do
	   	Rs.movenext
	loop
%>
        </table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="250" height="30"><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
              ѡ�б�ҳ��ʾ��������Ϣ </td>
    <td><input name="submit" type='submit' value='ɾ��ѡ������Ϣ' onClick="document.myform.Action.value='Del'">
              <input name="Action" type="hidden" id="Action" value="Del">
               
            </td>
  </tr>
</table>
</td>
</form></tr></table>
<%
end sub

%>
</body>
</html>
<%
'��ʾ��Ʒ������Ŀ
sub Admin_ShowBigClass()
	set Rs=server.CreateObject("adodb.recordset")
	sqlStr="select * from BigClass order by adddate asc"
	Rs.open sqlStr,conn,1,1
	if Rs.eof and Rs.bof then
		response.Write "��û���κ���Ŀ�������������Ŀ��"
	else
		response.write "|&nbsp;"
		do while not Rs.eof
			if Rs("BigID")=BigID then
				response.Write("<a href='" & FileName & "?BigID=" & Rs("BigID") & "'><font color=red>" & Rs("BigName") & "</font></a> | ")
			else
				response.Write("<a href='" & FileName & "?BigID=" & Rs("BigID") & "'>" & Rs("BigName") & "</a> | ")
			end if
			Rs.movenext
		loop
	end if
	rs.close
	set rs=nothing
end sub

'��ʾ����λ��
sub Admin_ShowPath2(Big)
	dim strPathName
	strPathName=""
	strPathName=strPathName & "�����ڵ�λ�ã�&nbsp;&gt;&gt;&nbsp;"
	if Big>0 then
		set Rs=server.CreateObject("adodb.recordset")
		sqlStr="select BigName from BigClass where BigID=" & clng(Big)
		Rs.open sqlStr,conn,1,1
			strPathName=strPathName & Rs("BigName") & "&nbsp;&gt;&gt;&nbsp;"
		Rs.close
		set rs=nothing
	end if
	strPathName=strPathName & "&nbsp;������Ϣ"
	response.Write strPathName
end sub

call CloseConn()
%>