<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim Action
Action=request("Action")
%>

<%
dim strFileName,MaxPerPage,totalPut,CurrentPage,TotalPages
dim BeginTime,EndTime,founderr, errmsg,i
dim rs,sql,rsGuest,sqlGuest
dim PageTitle,strPath,strPageTitle
dim SkinID,ClassID,AnnounceCount
dim UserGuestName,UserType,UserSex,UserEmail,UserHomepage,UserOicq,UserIcq,UserMsn
dim WriteName,WriteType,WriteSex,WriteEmail,WriteOicq,WriteIcq,WriteMsn,WriteHomepage
dim WriteFace,WriteImages,WriteTitle,WriteContent,SaveEdit,SaveEditId
dim GuestType,LoginName,AdminReplyContent
dim SubmitType,GuestPath,TitleName,keyword
Set rsGuest= Server.CreateObject("ADODB.Recordset")


strFileName="Admin_Order.asp"
ComeUrl1="Admin_Order.asp"
MaxPerPage=5
SaveEdit=0							
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

SubmitType=request("SubmitType")
select case SubmitType
	case "ɾ��"
		Action="del" 
end select

%>

<html>
<title>��վ��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<br>
<%
call Guestbook()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="tdbg_leftall"> 
  <td> 
	<%
	call ShowGuestPage()
	%>
  </td>
</tr>
</table>
</body>
</html>

<%

'=================================================
'��������GuestBook()
'��  �ã����������ܵ���
'��  ������
'=================================================
sub GuestBook()
	select case Action
		case "del"
			call DelGuest()
		case else
			call GuestMain()
	end select
end sub

sub GuestMain()
	response.write "<form style=""margin:0;padding:0"">"
	call ShowAllGuest()
	call ShowGuestBottom()
	response.write "</form>"
end sub

'=================================================
'��������ShowAllGuest()
'��  �ã���ҳ��ʾ���ж���
'��  ����ShowType-----  0Ϊ��ʾ����
'						1Ϊ��ʾ��ͨ����˼��û��Լ�����Ķ���
'						2Ϊ��ʾ��ͨ����˵Ķ����������ο���ʾ��
'						3Ϊ��ʾ�û��Լ�����Ķ���
'=================================================
sub ShowAllGuest()
	sqlGuest="select * from [order] order by id"
	set rsGuest=server.createobject("adodb.recordset")
	rsGuest.open sqlGuest,conn,1,1
	if rsGuest.bof and rsGuest.eof then
		totalput=0
		response.write "<br><li>û���κζ���</li>"
	else
		totalput=rsGuest.recordcount
		if currentPage=1 then
			call ShowGuestList()
		else
			if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsGuest.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsGuest.bookmark
            	call ShowGuestList()
        	else
	        	currentPage=1
           		call ShowGuestList()
	    	end if
		end if
	end if
	rsGuest.close
	set rsGuest=nothing
end sub

'=================================================
'��������ShowGuestList()
'��  �ã���ʾ����
'��  ������
'=================================================
sub ShowGuestList()
	dim i,GuestTip,TipName,TipSex,TipEmail,TipOicq,TipHomepage,isdelUser
	i=0
	do while not rsGuest.eof
		%>
		<SCRIPT language=javascript>
		function CheckAll(form)
		{
		  for (var i=0;i<form.elements.length;i++)
			{
			var e = form.elements[i];
			if (e.Name != "chkAll"&&e.disabled!=true)
			   e.checked = form.chkAll.checked;
			}
		}
		</script>
		<table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
		  <tr>
			<td> 
			  <table width="100%" border="0" cellpadding="0" cellspacing="0" class="border">
				<tr > 
				  <td align="center" valign="top"> 
					<table width="100%" border="0" cellspacing="0" cellpadding="0" class="title">
					  <tr> 
						<td width="400"> </td>
						<td width="250">  
						  <font color="#006633">������ڣ� 
						  <% =rsGuest("addDatetime")%>
						  </font> <input name=guestid type=checkbox id=guestid value="<%=rsGuest("id")%>" style="BORDER:0px;"></td>
					  </tr>
					</table>
				  </td>
				</tr>
				<tr class="tdbg_leftall"> 
				  <td align="center" valign="top"> 
					<table width="100%" border="0" cellpadding="2" cellspacing="2">
					  <tr class="tdbg">
						<td width="50%">��Ʒ���ƣ�<%=rsGuest("product")%></td>
						<td width="50%" align="left">��˾���ƣ�<%=rsGuest("gsmc")%></td>
					  </tr>
					  <tr class="tdbg">
						<td>�ջ���������<%=rsGuest("shr")%> </td>
						<td align="left">�ջ���ַ��<%=rsGuest("Add")%></td>
					  </tr>

					  <tr class="tdbg">
						<td>��ϵ�绰��<%=rsGuest("tel")%></td>
						<td align="left">��ϵ���棺<%=rsGuest("Fox")%> </td>
					  </tr>
					  <tr class="tdbg">
						<td>�������䣺<%=rsGuest("Email")%> </td>
						<td align="left">�������룺<%=rsGuest("yb")%></td>
					  </tr>
					  	 <tr class="tdbg">
						<td colspan=2 align=left>��ѡ����˺ţ�<%=rsGuest("zh")%>  </td>
					  </tr>
					  <tr class="tdbg">
						<td colspan="2">��ע<br>
						<%=rsGuest("content")%></td>
					  </tr>
					</table>
				  </td>
				</tr>
			  </table>
			  </td>
		  </tr>
		</table>
		<%
		rsGuest.movenext
		i=i+1
		if i>=MaxPerPage then exit do
	loop
end sub

	%>
	<%

'**************************************************
'��������KeywordReplace
'��  �ã���ʾ�����ؼ���
'��  ����strChar-----Ҫת�����ַ�
'����ֵ��ת������ַ�
'**************************************************
function KeywordReplace(strChar)
	if strChar="" then
		KeywordReplace=""
	else
		KeywordReplace=	replace(strChar,""&keyword&"","<font color=red>"&keyword&"</font>")
	end if
end function

%>
<%


%>
<%


'=================================================
'��������DelGuest()
'��  �ã�ɾ������
'��  ������
'=================================================
sub DelGuest()
	dim delid
	delid=trim(Request("guestid"))
	if delid="" then
		call response.Write("<li>��ָ��Ҫɾ���Ķ���ID��</li>")
		exit sub
	end if
	if instr(delid,",")>0 then
		delid=replace(delid," ","")
		sql="Select * from [order] where ID in (" & delid & ")"
	else
		delid=clng(delid)
		sql="select * from [order] where ID=" & delid
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.bof and rs.eof then
		response.write "<br><li>�Ҳ�����ָ���Ķ�����</li>"
		exit sub
	end if

		do while not rs.eof
			rs.delete
			rs.update
			rs.movenext
		loop
		rs.close
		set rs=nothing
		response.redirect ComeUrl1
end sub



'=================================================
'��������ShowGuestBottom()
'��  �ã���ʾ�����ײ�������
'��  ������
'=================================================
sub ShowGuestBottom()
	dim strTemp
	if TotalPut>0 then
	 	strTemp= "<table align='center'><tr><td>"
		strTemp= strTemp & "&nbsp;&nbsp;������"
			strTemp= strTemp & "<input  type=""submit"" name=""SubmitType"" value=""ɾ��"" onClick=""return confirm('ȷ��Ҫɾ��ѡ�еĶ�����');"">&nbsp;&nbsp;"
		strTemp= strTemp & "<input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" style=""BORDER:0px;"">ѡ�б�ҳ��ʾ�����ж���"
		strTemp= strTemp & "</td></tr></table>"
		response.write strTemp
	end if
end sub

'=================================================
'��������ShowGuestPage()
'��  �ã���ʾ�����ײ���ҳ
'��  ������
'=================================================
sub ShowGuestPage()
	dim PageFileName
	PageFileName=strFileName
	if keyword<>"" then
		PageFileName=PageFileName&"?keyword="&keyword
	end if
	if action<>"" then
		PageFileName=PageFileName&"?action="&action
	end if
	if TotalPut>0 then
		call showpage(PageFileName,totalPut,MaxPerPage,true,true,"������")
	end if
end sub

%>