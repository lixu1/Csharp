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
<title>����Ա����</title>
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
      alert("�û�������Ϊ�գ�");
	  document.form1.username.focus();
      return false;
    }

  if(document.form1.pname.value=="")
    {
      alert("��ʵ��������Ϊ�գ�");
	  document.form1.pname.focus();
      return false;
    }

  if(document.form1.Password.value=="")
    {
      alert("���벻��Ϊ�գ�");
	  document.form1.Password.focus();
      return false;
    }
    
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("��ʼ������ȷ�����벻ͬ��");
	  document.form1.PwdConfirm.select();
	  document.form1.PwdConfirm.focus();	  
      return false;
    }
}
function CheckModifyPwd()
{
  if(document.form1.Password.value=="")
    {
      alert("���벻��Ϊ�գ�");
	  document.form1.Password.focus();
      return false;
    }
  if((document.form1.Password.value)!=(document.form1.PwdConfirm.value))
    {
      alert("��ʼ������ȷ�����벻ͬ��");
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
    <td height="22" colspan="2" align="center"><strong>�� �� Ա �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="Admin_Admin.asp">����Ա������ҳ</a>&nbsp;|&nbsp;<a href="Admin_Admin.asp?Action=Add">��������Ա</a></td>
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
  <form name="myform" method="Post" action="Admin_Admin.asp" onSubmit="return confirm('ȷ��Ҫɾ��ѡ�еĹ���Ա��');">
     <td>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr align="center" class="title">
    <td  width="30"><strong>ѡ��</strong></td>
    <td  width="30" height="22"><strong> ���</strong></td>
    <td height="22"><strong> �� �� ��</strong></td>
	<td height="22"><strong> ��ʵ����</strong></td>
	<td height="22"><strong> ��¼����</strong></td>
	<td height="22"><strong> ����¼ʱ��</strong></td>
	<td height="22"><strong> ����Ȩ��</strong></td>
    <td  width="150" height="22"><strong> �� ��</strong></td>
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
		response.write "��������Ա"
	else
		response.write "��ͨ����Ա"
	end if
	%></td>
    <td width="150"><%
	response.write "<a href='Admin_Admin.asp?Action=ModifyPwd&ID=" & rs("ID") & "'>�޸�����</a>&nbsp;&nbsp;"
	if iCount>1 and rs("UserName")<>session("AdminName") then
		response.write "<a href='Admin_Admin.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('ȷ��Ҫɾ���˹���Ա��');"">ɾ��</a>"
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
      <td height="22" colspan="2"> <div align="center"><strong>�� �� �� �� Ա</strong></div></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> �� �� ����</strong></td>
      <td width="65%" class="tdbg"><input name="username" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> ��ʵ������</strong></td>
      <td width="65%" class="tdbg"><input name="pname" type="text"> &nbsp;</td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> ��ʼ���룺 </strong></td>
      <td width="65%" class="tdbg"><font size="2"> 
        <input type="password" name="Password">
      </font></td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> ȷ�����룺</strong></td>
      <td width="65%" class="tdbg"><font size="2"> 
        <input type="password" name="PwdConfirm">
      </font></td>
    </tr>

    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input  type="submit" name="Submit" value=" �� �� " style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;"></td>
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
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
		exit sub
	else
		UserID=Clng(UserID)
	end if
	sql="Select * from Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����ڴ��û���</li>"
	else
%>
<form method="post" action="Admin_Admin.asp" name="form1" onSubmit="javascript:return CheckModify();">
  <table width="70%" border="0" align="center" cellpadding="2" cellspacing="1" class="border" >
    <tr class="title"> 
      <td height="22" colspan="2"> <div align="center"><font size="2"><strong>�� �� �� �� Ա �� ��</strong></font></div></td>
    </tr>
    <tr> 
      <td width="40%" class="tdbg"><strong>�� �� ����</strong></td>
      <td width="60%" class="tdbg"><%=rs("UserName")%> <input name="ID" type="hidden" value="<%=rs("ID")%>"></td>
    </tr>
    <tr> 
      <td width="40%" class="tdbg"><strong>�� �� �룺</strong></td>
      <td width="60%" class="tdbg"><input type="password" name="Password">
      </td>
    </tr>
    <tr> 
      <td width="40%" class="tdbg"><strong>ȷ�����룺</strong></td>
      <td width="60%" class="tdbg"><input type="password" name="PwdConfirm">
      </td>
    </tr>
    <tr class="tdbg"> 
      <td width="35%" class="tdbg"><strong> ����ԱȨ�ޣ�</strong></td>
      <td width="65%" class="tdbg"><input type="radio" name="flag" value="1" <%if Rs("flag")=1 then response.write "checked" end if %>>��������Ա 
      <input type="radio" name="flag" value="0" <%if Rs("flag")<>1 then response.write "checked" end if %>>��ͨ����Ա</td>
    </tr>
    <tr> 
      <td colspan="2" align="center" class="tdbg"><input name="Action" type="hidden" id="Action" value="SaveModifyPwd"> 
        <input  type="submit" name="Submit" value="�����޸Ľ��" style="cursor:hand;">
        &nbsp;
        <input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="window.location.href='Admin_Admin.asp'" style="cursor:hand;"></td>
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
		ErrMsg=ErrMsg & "<br><li>�û�������Ϊ�գ�</li>"
	end if
	if pname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ʵ��������Ϊ�գ�</li>"
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ʼ���벻��Ϊ�գ�</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ȷ������������ʼ������ͬ��</li>"
	end if

	if FoundErr=True then
		exit sub
	end if
	sql="Select * from Admin where username='"&username&"'"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if not (rs.bof and rs.EOF) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ѿ����ڴ˹���Ա��</li>"
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
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�޸ĵĹ���ԱID</li>"
	else
		UserID=Clng(UserID)
	end if
	if password="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����벻��Ϊ�գ�</li>"
	end if
	if PwdConfirm<>Password then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>ȷ�������������������ͬ��</li>"
	end if
	if FoundErr=True then
		exit sub
	end if
	
	sql="Select * from Admin where ID=" & UserID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.Bof and rs.EOF then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�����ڴ˹���Ա��</li>"
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
		ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ���Ĺ���ԱID</li>"
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