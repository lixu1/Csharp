<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim Rs,sqlStr
%>
<html>
<head>
<title>������Ŀ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("���´������Ʋ���Ϊ�գ�");
	document.form1.BigClassName.focus();
	return false;
  }
}
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>�� �� �� Ŀ �� ��</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>��������</strong></td>
    <td height="30"><a href="Admin_Class_Article.asp">������Ŀ������ҳ</a> | 
	<a href="Admin_Class_Article.asp?Action=Add">���������Ŀ</a></td>
  </tr>
</table>
<%
dim action,FoundErr,ErrMsg
action=request("action")
if Action="AddBig" then
	call AddBigClass()
elseif Action="AddSmallClass" then
	call AddSmallClass()
elseif Action="SaveBigAdd" then
	call SaveBigAdd()
elseif Action="SaveSmallAdd" then
	call SaveSmallAdd()
elseif Action="ModifyBig" then
	call ModifyBig()
elseif Action="ModifySmall" then
	call ModifySmall()
elseif Action="SaveModifyBig" then
	call SaveModifyBig()
elseif Action="SaveModifySmall" then
	call SaveModifySmall()
elseif Action="DelBig" then
	call DeleteBig()
elseif Action="DelSmall" then
	call DeleteSmall()
elseif Action="aa" then
	call aa()
elseif Action="bb" then
	call bb()
else
	call main()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()


sub main()
%>

      <br>
      
<table width="100%" height="51"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="28" align="center"><strong>�� �� �� �� �� ��</strong></td>
  </tr>
  <tr>
    <td align="center">����ѡ�<a href="?Action=AddBig">����������</a></td>
  </tr>
</table>
<br>
<table width="70%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">

  <tr class="title"> 
    <td height="22" colspan="2" align="center"><strong>��Ŀ����</strong></td>
    <td width="300" height="22" align="center"><strong>����ѡ��</strong></td>
  </tr>
  <% 
  dim RsSmall,sqlSmall
  set Rs=server.CreateObject("adodb.recordset")
  sqlStr="select * from BigClass order by Adddate asc"
  Rs.open sqlStr,conn,1,1
  if not Rs.eof then
  do while not Rs.eof
  %>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
    <td colspan="2" align="left">
	<img src="images/tree_folder4.gif" border="0"><%=Rs("BigName")%> </td>
    <td align="right" style="padding-right:50px; " > <a href="?action=AddSmallClass&BigID=<%=rs("bigID")%>&bigName=<%=rs("BigName")%>">�������</a> | <a href="?Action=ModifyBig&BigID=<%=Rs("BigID")%>">�޸�����</a> 
      | <a href="?Action=DelBig&BigID=<%=Rs("BigID")%>" onClick="return confirm('ɾ������Ŀ��ͬʱɾ������Ŀ�е��������£�ȷ��Ҫɾ������Ŀ��');">ɾ��</a> | 
       <a href="?Action=aa&BigID=<%=Rs("BigID")%>" >����</a>	  </td>
  </tr>
  <%
  	set rs1=server.CreateObject("adodb.recordset")
	sql1="select * from smallclass where BigID="&clng(rs("BigID"))&" order by Adddate asc" 
	rs1.open sql1,conn,1,1
	if not rs1.eof then
		do while not rs1.eof
  %>
  <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'">
    <td width="24" align="left">&nbsp;</td>
    <td width="207" align="left"><img src="images/tree_folder4.gif" border="0"><%=rs1("SmallName")%></td>
    <td align="right" style="padding-right:50px; " ><a href="?Action=ModifySmall&SmallID=<%=Rs1("SmallID")%>">�޸�����</a> | <a href="?Action=DelSmall&SmallID=<%=Rs1("SmallID")%>" onClick="return confirm('ɾ������Ŀ��ͬʱɾ������Ŀ�е��������£�ȷ��Ҫɾ������Ŀ��');">ɾ��</a> |  <a href="?Action=bb&SmallID=<%=Rs1("SmallID")%>" >����</a></td>
  </tr>
  <% 
  		rs1.movenext
		loop
	end if
	rs1.close
	set rs1=nothing
	  Rs.movenext
   loop
   else
   	    response.Write "<tr class=tdbg><td  colspan=2>��û���κ�������Ŀ����������������</td></tr>"
   end if
   Rs.close
   %>
</table> 
<br><br>
<%
end sub
'������´�����Ŀ
sub AddBigClass()
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
  <table width="50%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>�� �� �� ��  �� Ŀ</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���ƣ�</strong></td>
      <td><input name="BigClassName" type="text" size="37" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" value="SaveBigAdd"> 
        <input name="Add" type="submit" value=" �� �� "  style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="reset" id="Cancel" value=" ȡ �� " style="cursor:hand;"> 
      </td>
    </tr>
  </table>
</form>
<%
end sub

'�������С����Ŀ
sub AddSmallClass()
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
  <table width="50%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>�� �� �� ��  �� Ŀ</strong></td>
    </tr>
    <tr class="tdbg">
      <td><strong>��������Ŀ��</strong></td>
      <td>
	  <%
	  	bigID=request.QueryString("bigID")
		BigName=request.QueryString("BigName")
		if bigName="" or bigID="" then
		%>
		<script language="javascript">alert("ȱ�ٲ�����");this.history.go(-1)</script>
		<%else
		response.Write BigName
		%>
		<input type="hidden" value="<%=BigID%>" name="BigID" />
		<%end if%>
	  </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���ƣ�</strong></td>
      <td><input name="SmallClassName" type="text" size="37" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" value="SaveSmallAdd"> 
        <input name="Add" type="submit" value=" �� �� "  style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="reset" id="Cancel" value=" ȡ �� " style="cursor:hand;">      </td>
    </tr>
  </table>
</form>
<%
end sub

'�޸����´�����Ŀ
sub ModifyBig()
	dim BigID
	BigID=trim(request("BigID"))
	if BigID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		BigID=CLng(BigID)
	end if
	sqlStr="select * From BigClass where BigID=" & BigID
	set Rs=server.CreateObject ("Adodb.recordset")
	Rs.open sqlStr,conn,1,3
	if Rs.bof and Rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ��</li>"
	else
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
  <table width="50%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>�� �� �� ��  �� Ŀ</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���ƣ�</strong></td>
      <td><input name="BigName" type="text" value="<%=Rs("BigName")%>" size="37" maxlength="20"> 
        <input name="BigID" type="hidden" value="<%=Rs("BigID")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" value="SaveModifyBig"> 
        <input name="Submit" type="submit" value=" �����޸Ľ�� " style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="reset" id="Cancel" value=" ȡ �� " style="cursor:hand;"> 
      </td>
    </tr>
  </table>
</form>
<%
	end if
	Rs.close
	set Rs=nothing
end sub
'�޸����´�����Ŀ���ƽ���

'�޸�����С����Ŀ
sub ModifySmall()
	dim SmallID
	SmallID=trim(request("SmallID"))
	if SmallID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	else
		SmallID=CLng(SmallID)
	end if
	sqlStr="select * From SmallClass where SmallID=" & SmallID
	set Rs=server.CreateObject ("Adodb.recordset")
	Rs.open sqlStr,conn,1,3
	if Rs.bof and Rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ������Ŀ��</li>"
	else
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
  <table width="50%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>�� �� �� �� �� Ŀ</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>��Ŀ���ƣ�</strong></td>
      <td><input name="SmallName" type="text" value="<%=Rs("SmallName")%>" size="37" maxlength="20"> 
        <input name="SmallID" type="hidden" value="<%=Rs("SmallID")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" value="SaveModifySmall"> 
        <input name="Submit" type="submit" value=" �����޸Ľ�� " style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="reset" id="Cancel" value=" ȡ �� " style="cursor:hand;"> 
      </td>
    </tr>
  </table>
</form>
<%
	end if
	Rs.close
	set Rs=nothing
end sub
'�޸�����С����Ŀ���ƽ���

%>
</body> 
</html> 
<% 
'===============================
'����SaveBigAdd()����������´�������
'===============================
sub SaveBigAdd()
	dim BigClassName
	dim sql,rs,trs

	BigClassName=trim(request("BigClassName"))
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����������Ʋ���Ϊ�գ�</li>"
		exit sub
	end if
	'������ݿ��Ƿ��Ѵ��ڴ˴�������
	set Rs=server.createobject("adodb.recordset")
	sql="Select top 1 * From BigClass where BigName='" & BigClassName & "'"
	rs.open sql,conn,1,1
	if not(rs.bof and rs.eof) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��������������Ѿ�������Ŀ��" & BigClassName & "����</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.close
	'�����ݿ���Ӵ���
	set Rs=server.createobject("adodb.recordset")
	sql="Select top 1 * From BigClass"
	rs.open sql,conn,1,3
    rs.addnew
   	rs("BigName")=BigClassName
	rs.update
	rs.Close
    set rs=Nothing
	
	
    call CloseConn()
	Response.Redirect "Admin_Class_Article.asp"  
end sub

'===============================
'����SaveSmallAdd()����������´�������
'===============================
sub SaveSmallAdd()
	dim BigID,SmallClassName
	dim sql,rs,trs
	BigID=trim(request.Form("BigID"))
	SmallClassName=trim(request.Form("SmallClassName"))
	if SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����������Ʋ���Ϊ�գ�</li>"
		exit sub
	end if
	'������ݿ��Ƿ��Ѵ��ڴ˴�������
	set Rs=server.createobject("adodb.recordset")
	sql="Select top 1 * From SmallClass where SmallName='" & SmallClassName & "' and BigID="&clng(BidID)
	rs.open sql,conn,1,1
	if not(rs.bof and rs.eof) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��������������Ѿ�������Ŀ��" & SmallClassName & "����</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.close
	'�����ݿ���Ӵ���
	set Rs=server.createobject("adodb.recordset")
	sql="Select top 1 * From SmallClass"
	rs.open sql,conn,1,3
    rs.addnew
   	rs("SmallName")=SmallClassName
	rs("BigID")=BigID
	rs.update
	rs.Close
    set rs=Nothing
    call CloseConn()
	Response.Redirect "Admin_Class_Article.asp"  
end sub


'===============================
'����SaveModifyBig()�����޸����´�������
'===============================
sub SaveModifyBig()
	dim BigID,BigName
	
	BigID=trim(request.Form("BigID"))
	if BigID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	end if
	BigName=trim(request.Form("BigName"))
	if BigName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����������Ʋ���Ϊ�գ�</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	'�����ݿ���Ӵ���
	sqlStr="select * From BigClass where BigID=" & clng(BigID)
	set Rs=server.CreateObject ("Adodb.recordset")
	Rs.open sqlStr,conn,1,3
	
   	Rs("BigName")=BigName
	Rs.update
	Rs.close
	set Rs=nothing
	
    call CloseConn()
	Response.Redirect "Admin_Class_Article.asp"  
end sub


'===============================
'����SaveModifySmall()�����޸�����С������
'===============================
sub SaveModifySmall()
	dim SmallID,SmallName
	
	SmallID=trim(request.Form("SmallID"))
	if SmallID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
	end if
	SmallName=trim(request.Form("SmallName"))
	if SmallName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>����������Ʋ���Ϊ�գ�</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	'�����ݿ���Ӵ���
	sqlStr="select * From SmallClass where SmallID=" & clng(SmallID)
	set Rs=server.CreateObject ("Adodb.recordset")
	Rs.open sqlStr,conn,1,3
   	Rs("SmallName")=SmallName
	Rs.update
	Rs.close
	set Rs=nothing
	
    call CloseConn()
	Response.Redirect "Admin_Class_Article.asp"  
end sub

'===============================
'����DeleteBig()ɾ�����´��༰��Ӧ����С�༰����������
'===============================
sub DeleteBig()
	dim BigID
	BigID=trim(Request("BigID"))
	if BigID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	end if
	
	sql="select * From BigClass where BigID="& clng(BigID)
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��Ŀ�����ڣ������Ѿ���ɾ��</li>"
	end if
	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if
	
	rs.delete
	rs.update
	rs.close
	set rs=nothing
	'ɾ������Ŀ����������
	conn.execute("delete from Article where BigID=" & clng(BigID))
	
	call CloseConn()
	response.redirect "Admin_Class_Article.asp"
		
end sub

'===============================
'����DeleteSmall()ɾ������С�༰��Ӧ����
'===============================
sub DeleteSmall()
	dim SmallID
	SmallID=trim(Request("SmallID"))
	if SmallID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
		exit sub
	end if
	
	sql="select * From SmallClass where SmallID="& clng(SmallID)
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��Ŀ�����ڣ������Ѿ���ɾ��</li>"
	end if
	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if
	
	rs.delete
	rs.update
	rs.close
	set rs=nothing
	'ɾ������Ŀ����������
	conn.execute("delete from Article where SmallID=" & clng(SmallID))
	call CloseConn()
	response.redirect "Admin_Class_Article.asp"
		
end sub

sub aa()
	dim bigid
	bigid=request("bigid")
	set Rs=server.createobject("adodb.recordset")
	sqlStr="select * from BigClass where BigID=" & clng(bigid)
	Rs.open sqlStr,conn,1,3
	if not Rs.eof then
		Rs("adddate")=Now()
	end if
	Rs.update
	Rs.close
	call CloseConn()
	response.redirect "Admin_Class_Article.asp"
end sub

sub bb()
	dim SmallID
	SmallID=request("SmallID")
	set Rs=server.createobject("adodb.recordset")
	sqlStr="select * from SmallClass where SmallID=" & clng(SmallID)
	Rs.open sqlStr,conn,1,3
	if not Rs.eof then
		Rs("adddate")=Now()
	end if
	Rs.update
	Rs.close
	call CloseConn()
	response.redirect "Admin_Class_Article.asp"
end sub

%>