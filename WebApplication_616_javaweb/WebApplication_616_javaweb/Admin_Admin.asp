<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim rs, sql, strPurview,iCount
dim Action,FoundErr,ErrMsg
Action=Trim(request("Action"))
%>
<html>
<head>
<title>管理员管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
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
    if (e.Name != "chkAll"&&e.disabled!=true)
       e.checked = form.chkAll.checked;
    }
}

function CheckAdd()
{
  if(document.form1.username.value=="")
    {
      alert("用户名不能为空！");
	  document.form1.username.focus();
      return false;
    }

  if(document.form1.pname.value=="")
    {
      alert("真实姓名不能为空！");
	  document.form1.pname.focus();
      return false;
    }

  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}
function CheckModifyPwd()
{
  if(document.form1.Password.value=="")
    {
      alert("密码不能为空！");
	  document.form1.Password.focus();
      return false;
    }
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("初始密码与确认密码不同！");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}

</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>管 理 员 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_Admin.asp">管理员管理首页</a>&nbsp;|&nbsp;<a href="Admin_Admin.asp?Action=Add">新增管理员</a></td>
  </tr>
</table>
<%
if Action="Add" then
	call AddAdmin()
elseif Action="SaveAdd" then
	call SaveAdd()
elseif Action="ModifyPwd" then
	call ModifyPwd()
elseif Action="SaveModifyPwd" then
	call SaveModifyPwd()
elseif Action="Del" then
	call DelAdmin()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	Set rs=Server.CreateObject("Adodb.RecordSet")
	sql="select * from admin where username<>'root' order by ID"
	rs.Open sql,conn,1,1
	iCount=rs.recordcount
%>
<br>
<table width='80%' border="0" cellpadding="0" cellspacing="0" align=center>
  <tr>
  <form name="myform" method="Post" action="Admin_Admin.asp" onSubmit="return confirm('确定要删除选中的管理员吗？');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr align="center" class="title">
    <td  width="30"><strong>选中</strong></td>
    <td  width="30" height="22"><strong> 序号</strong></td>
    <td height="22"><strong> 用 户 名</strong></td>
	<td height="22"><strong> 真实姓名</strong></td>
	<td height="22"><strong> 登录次数</strong></td>
	<td height="22"><strong> 最后登录时间</strong></td>
	<td height="22"><strong> 管理权限</strong></td>
    <td  width="150" height="22"><strong> 操 作</strong></td>
  </tr>
  <%do while not rs.EOF %>
  <tr align="center" class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td width="30"><input name="ID" type="checkbox" id="ID" value="<%=rs("ID")%>" <%if rs("UserName")=AdminName then response.write " disabled"%> onClick="unselectall()"></td>
    <td width="30"><%=rs("ID")%></td>
    <td><%
	if rs("username")=session("AdminName") then
		response.write "<font color=red><b>" & rs("UserName") & "</b></font>"
	else
		response.write rs("UserName")
	end if
	%></td>
	<td height="22"><%=Rs("PName")%></td>
	<td height="22"><%=Rs("Hits")%></td>
	<td height="22"><%=Rs("LastDataTime")%></td>
	<td height="22"><%
	if Rs("flag")=1 then
		response.write "超级管理员"
	else
		response.write "普通管理员"
	end if
	%></td>
    <td width="150"><%
	response.write "<a href='Admin_Admin.asp?Action=ModifyPwd&ID=" & rs("ID") & "'>修改密码</a>&nbsp;&nbsp;"
	if iCount>1 and rs("UserName")<>session("AdminName") then
		response.write "<a href='Admin_Admin.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('确定要删除此管理员吗？');"">删除</a>"
	else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;"
	end if
	%> </td>
  </tr>
  <%
	rs.MoveNext
loop
  %>
</table>  
</td>
</form></tr></table>
<%
	rs.Close
	set rs=Nothing
end sub

sub AddAdmin()
%>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:return CheckAdd();">
  <table width="70%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><strong>新 增 管 理 员</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> 用 户 名：</strong></td>
      <td width="65%" class="tdbg"><input name="username" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> 真实姓名：</strong></td>
      <td width="65%" class="tdbg"><input name="pname" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> 初始密码： </strong></td>
      <td width="65%" class="tdbg"><font size="2"> 
        <input type="password" name="Password">
      </font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> 确认密码：</strong></td>
      <td width="65%" class="tdbg"><font size="2"> 
        <input type="password" name="PwdConfirm">
      </font></td>
    </tr>

    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input  type="submit" name="Submit" value=" 添 加 " style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
</form>
<%
end sub

sub ModifyPwd()
	dim UserID
	UserID=trim(Request("ID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的管理员ID</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	sql="Select * from Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>不存在此用户！</li>"
	else
%>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:return CheckModify();">
  <table width="70%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><font size="2"><strong>修 改 管 理 员 密 码</strong></font></div></td>
    </tr>
    <tr> 
      <td width="40%" class="tdbg"><strong>用 户 名：</strong></td>
      <td width="60%" class="tdbg"><%=rs("UserName")%> <input name="ID" type="hidden" value="<%=rs("ID")%>"></td>
    </tr>
    <tr> 
      <td width="40%" class="tdbg"><strong>新 密 码：</strong></td>
      <td width="60%" class="tdbg"><input type="password" name="Password">
      </td>
    </tr>
    <tr> 
      <td width="40%" class="tdbg"><strong>确认密码：</strong></td>
      <td width="60%" class="tdbg"><input type="password" name="PwdConfirm">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> 管理员权限：</strong></td>
      <td width="65%" class="tdbg"><input type="radio" name="flag" value="1" <%if Rs("flag")=1 then response.write "checked" end if %>>超级管理员 
      <input type="radio" name="flag" value="0" <%if Rs("flag")<>1 then response.write "checked" end if %>>普通管理员</td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveModifyPwd"> 
        <input  type="submit" name="Submit" value="保存修改结果" style="cursor:hand;">
        &nbsp;
        <input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;"></td>
    </tr>
  </table>
</form>
<%
	end if
	rs.close
	set rs=nothing
end sub
%>
</body>
</html>
<%
sub SaveAdd()
	dim username, password,PwdConfirm,pname,flag

	username=trim(Request("username"))
	pname=trim(Request("pname"))
	password=trim(Request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	flag=1

	if username="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>用户名不能为空！</li>"
	end if
	if pname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>真实姓名不能为空！</li>"
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>初始密码不能为空！</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>确认密码必须与初始密码相同！</li>"
	end if

	if FoundErr=True then
		exit sub
	end if
	sql="Select * from Admin where username='"&username&"'"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if not (rs.bof and rs.EOF) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>已经存在此管理员！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
   	rs.addnew
 	rs("username")=username
   	rs("password")=md5(password)
   	rs("pname")=pname
   	rs("flag")=flag
	rs.update
    rs.Close
	set rs=Nothing
	Call main()
end sub

sub SaveModifyPwd()
	dim UserID, UserName,password,PwdConfirm,flag
	UserID=trim(Request("ID"))
	password=trim(Request("Password"))
	PwdConfirm=trim(request("PwdConfirm"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要修改的管理员ID</li>"
	else
		UserID=Clng(UserID)
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>新密码不能为空！</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>确认密码必须与新密码相同！</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	
	sql="Select * from Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>不存在此管理员！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	rs("password")=md5(password)
	rs("flag")=flag
 	rs.update
	rs.Close
   	set rs=Nothing
    call main()
end sub


sub DelAdmin()
	dim UserID
	UserID=trim(Request("ID"))
	if UserID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要删除的管理员ID</li>"
		exit sub
	end if
	if instr(UserID,",")>0 then
		UserID=replace(UserID," ","")
		sql="Select * from Admin where ID in (" & UserID & ")"
	else
		UserID=clng(UserID)
		sql="select * from Admin where ID=" & UserID
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	do while not rs.eof
		rs.delete
		rs.update
		rs.movenext
	loop
	rs.close
	set rs=nothing
	call main()
end sub

%>