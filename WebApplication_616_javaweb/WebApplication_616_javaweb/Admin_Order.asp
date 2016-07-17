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
	case "删除"
		Action="del" 
end select

%>

<html>
<title>网站订单管理</title>
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
'过程名：GuestBook()
'作  用：订单本功能调用
'参  数：无
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
'过程名：ShowAllGuest()
'作  用：分页显示所有订单
'参  数：ShowType-----  0为显示所有
'						1为显示已通过审核及用户自己发表的订单
'						2为显示已通过审核的订单（用于游客显示）
'						3为显示用户自己发表的订单
'=================================================
sub ShowAllGuest()
	sqlGuest="select * from [order] order by id"
	set rsGuest=server.createobject("adodb.recordset")
	rsGuest.open sqlGuest,conn,1,1
	if rsGuest.bof and rsGuest.eof then
		totalput=0
		response.write "<br><li>没有任何订单</li>"
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
'过程名：ShowGuestList()
'作  用：显示订单
'参  数：无
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
						  <font color="#006633">添加日期： 
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
						<td width="50%">产品名称：<%=rsGuest("product")%></td>
						<td width="50%" align="left">公司名称：<%=rsGuest("gsmc")%></td>
					  </tr>
					  <tr class="tdbg">
						<td>收货人姓名：<%=rsGuest("shr")%> </td>
						<td align="left">收货地址：<%=rsGuest("Add")%></td>
					  </tr>

					  <tr class="tdbg">
						<td>联系电话：<%=rsGuest("tel")%></td>
						<td align="left">联系传真：<%=rsGuest("Fox")%> </td>
					  </tr>
					  <tr class="tdbg">
						<td>电子信箱：<%=rsGuest("Email")%> </td>
						<td align="left">邮政编码：<%=rsGuest("yb")%></td>
					  </tr>
					  	 <tr class="tdbg">
						<td colspan=2 align=left>所选汇款账号：<%=rsGuest("zh")%>  </td>
					  </tr>
					  <tr class="tdbg">
						<td colspan="2">备注<br>
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
'函数名：KeywordReplace
'作  用：标示搜索关键字
'参  数：strChar-----要转换的字符
'返回值：转换后的字符
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
'过程名：DelGuest()
'作  用：删除订单
'参  数：无
'=================================================
sub DelGuest()
	dim delid
	delid=trim(Request("guestid"))
	if delid="" then
		call response.Write("<li>请指定要删除的订单ID！</li>")
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
		response.write "<br><li>找不到您指定的订单！</li>"
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
'过程名：ShowGuestBottom()
'作  用：显示订单底部管理功能
'参  数：无
'=================================================
sub ShowGuestBottom()
	dim strTemp
	if TotalPut>0 then
	 	strTemp= "<table align='center'><tr><td>"
		strTemp= strTemp & "&nbsp;&nbsp;操作："
			strTemp= strTemp & "<input  type=""submit"" name=""SubmitType"" value=""删除"" onClick=""return confirm('确定要删除选中的订单吗？');"">&nbsp;&nbsp;"
		strTemp= strTemp & "<input name=""chkAll"" type=""checkbox"" id=""chkAll"" onclick=CheckAll(this.form) value=""checkbox"" style=""BORDER:0px;"">选中本页显示的所有订单"
		strTemp= strTemp & "</td></tr></table>"
		response.write strTemp
	end if
end sub

'=================================================
'过程名：ShowGuestPage()
'作  用：显示订单底部分页
'参  数：无
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
		call showpage(PageFileName,totalPut,MaxPerPage,true,true,"条订单")
	end if
end sub

%>