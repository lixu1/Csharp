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
dim BigID,SmallID,ClassID,Passed

FileName="Admin_ArticleCheck.asp"

BigID=Trim(request.querystring("BigID"))
Passed=trim(request("Passed"))
if Passed="" then
	Passed="False"
else
	Passed=Passed
end if
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
<title>��Ʒ����</title>
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
</SCRIPT>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2"  align="center"><strong>�� Ʒ �� ��</strong></td>
  </tr>
<form name="Passed" method="Post" action="<%=strFileName%>">
  <tr class="tdbg"> 
      <td width="70" height="30" ><strong>��Ʒѡ�</strong></td>
    <td >
  <input name="Passed" type="radio" value="False" onclick="submit();" <%if Passed="False" then response.write " checked"%>>
  δ��˵Ĳ�Ʒ&nbsp;&nbsp;&nbsp;&nbsp;
  <input name="Passed" type="radio" value="True" onclick="submit();" <%if Passed="True" then response.write " checked"%>>
  ����˵Ĳ�Ʒ</td>
  </tr>
</form>
</table>
<br>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="title"> 
    <td height="22"><%call Admin_ShowBigClass()%></td>
  </tr>
</table>
<br>
<%
sqlStr="select * from Article where Passed=" & Passed
if BigID>0 then
	sqlStr=sqlStr & " and BigID=" & BigID
end if
sqlStr=sqlStr & " order by articleid desc"
Set Rs= Server.CreateObject("ADODB.Recordset")
Rs.open sqlStr,conn,1,1
if Rs.eof and Rs.bof then
		if Passed="True" then
			response.write "<p align='center'><br>û���κ�����˵Ĳ�Ʒ��<br></p>"
		else
			response.write "<p align='center'><br>û���κδ���˵Ĳ�Ʒ��<br></p>"
		end if
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
       	showpage strFileName,totalput,MaxPerPage,true,true,"����Ʒ"
 	else
     	if (currentPage-1)*MaxPerPage<totalPut then
       	   	Rs.move  (currentPage-1)*MaxPerPage
       		dim bookmark
           	bookmark=Rs.bookmark
            showContent
            showpage strFileName,totalput,MaxPerPage,true,true,"����Ʒ"
       	else
	        currentPage=1
           	showContent
          	showpage strFileName,totalput,MaxPerPage,true,true,"����Ʒ"
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
    <form name="myform" method="Post" action="Admin_ArticleProperty.asp">
     <td><table class="border" border="0" cellspacing="1" width="100%" cellpadding="0">
          <tr class="title" height="22"> 
            <td height="22" width="30" align="center"><strong>ѡ��</strong></td>
            <td width="25" align="center"  height="22"><strong>ID</strong></td>
            <td width="327" align="center" ><strong>��Ʒ����</strong></td>
            <td width="53" align="center" ><strong>�����</strong></td>
            <td width="75" align="center" ><strong>��Ʒ����</strong></td>
            <td width="50" align="center" ><strong>�����</strong></td>
            <td width="234" align="center" ><strong>����</strong></td>
          </tr>
          <%do while not Rs.eof%>
          <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
            <td width="30" align="center"><input name='ArticleID' type='checkbox' onclick="unselectall()" id="ArticleID" value='<%=cstr(Rs("articleID"))%>'></td>
            <td width="25" align="center"><%=Rs("articleid")%></td>
            <td> <%
			if Rs("IncludePic")=true then
				response.write "<font color=blue>[ͼ��]</font>"
			end if
			response.write  Rs("title") 
			%></td>
            <td width="53" align="center"><%= Rs("Hits") %></td>
            <td width="75" align="center"> <%
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
            <td width="50" align="center"> <%
			if Rs("Passed")=true then
				response.write "��"
			else
				response.write "��"
			end if%></td>
            <td width="234" align="center"> <%
            	if Rs("Passed")=False then	
					response.write "<a href='Admin_ArticleProperty.asp?ArticleID=" & Rs("ArticleID") & "&Action=SetPassed'>ͨ�����</a>"
				else
					response.write "<a href='Admin_ArticleProperty.asp?ArticleID=" & Rs("ArticleID") & "&Action=CancelPassed'>ȡ�����</a>"
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
              ѡ�б�ҳ��ʾ�����в�Ʒ </td>
    <td><input name="submit" type='submit' value='<%if Passed="True" then response.write "ȡ��"%>���ѡ���Ĳ�Ʒ'>
             <input name="Action" type="hidden" id="Action" value="<% if Passed="False" then
			  	response.write "SetPassed"
			  else
			    response.write "CancelPassed"
			  end if%>">
               
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
	sqlStr="select * from BigClass"
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
end sub

call CloseConn()

%>