<!--#include file="inc/conn.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="inc/check.asp"-->
<%
dim Rs,sqlStr
%>
<html>
<head>
<title>文章栏目管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
function check()
{
  if (document.form1.BigClassName.value=="")
  {
    alert("文章大类名称不能为空！");
	document.form1.BigClassName.focus();
	return false;
  }
}
</script>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>文 章 栏 目 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Admin_Class_Article.asp">文章栏目管理首页</a> | 
	<a href="Admin_Class_Article.asp?Action=Add">添加文章栏目</a></td>
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
    <td height="28" align="center"><strong>文 章 类 别 设 置</strong></td>
  </tr>
  <tr>
    <td align="center">管理选项：<a href="?Action=AddBig">添加文章类别</a></td>
  </tr>
</table>
<br>
<table width="70%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">

  <tr class="title"> 
    <td height="22" colspan="2" align="center"><strong>栏目名称</strong></td>
    <td width="300" height="22" align="center"><strong>操作选项</strong></td>
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
    <td align="right" style="padding-right:50px; " > <a href="?action=AddSmallClass&BigID=<%=rs("bigID")%>&bigName=<%=rs("BigName")%>">添加子类</a> | <a href="?Action=ModifyBig&BigID=<%=Rs("BigID")%>">修改设置</a> 
      | <a href="?Action=DelBig&BigID=<%=Rs("BigID")%>" onClick="return confirm('删除此栏目将同时删除此栏目中的所有文章，确定要删除此栏目吗？');">删除</a> | 
       <a href="?Action=aa&BigID=<%=Rs("BigID")%>" >排序</a>	  </td>
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
    <td align="right" style="padding-right:50px; " ><a href="?Action=ModifySmall&SmallID=<%=Rs1("SmallID")%>">修改设置</a> | <a href="?Action=DelSmall&SmallID=<%=Rs1("SmallID")%>" onClick="return confirm('删除此栏目将同时删除此栏目中的所有文章，确定要删除此栏目吗？');">删除</a> |  <a href="?Action=bb&SmallID=<%=Rs1("SmallID")%>" >排序</a></td>
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
   	    response.Write "<tr class=tdbg><td  colspan=2>还没有任何文章栏目，请先添加文章类别。</td></tr>"
   end if
   Rs.close
   %>
</table> 
<br><br>
<%
end sub
'添加文章大类栏目
sub AddBigClass()
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
  <table width="50%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>添 加 文 章  栏 目</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目名称：</strong></td>
      <td><input name="BigClassName" type="text" size="37" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" value="SaveBigAdd"> 
        <input name="Add" type="submit" value=" 添 加 "  style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="reset" id="Cancel" value=" 取 消 " style="cursor:hand;"> 
      </td>
    </tr>
  </table>
</form>
<%
end sub

'添加文章小类栏目
sub AddSmallClass()
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
  <table width="50%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>添 加 文 章  栏 目</strong></td>
    </tr>
    <tr class="tdbg">
      <td><strong>所属大栏目：</strong></td>
      <td>
	  <%
	  	bigID=request.QueryString("bigID")
		BigName=request.QueryString("BigName")
		if bigName="" or bigID="" then
		%>
		<script language="javascript">alert("缺少参数！");this.history.go(-1)</script>
		<%else
		response.Write BigName
		%>
		<input type="hidden" value="<%=BigID%>" name="BigID" />
		<%end if%>
	  </td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目名称：</strong></td>
      <td><input name="SmallClassName" type="text" size="37" maxlength="20"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" value="SaveSmallAdd"> 
        <input name="Add" type="submit" value=" 添 加 "  style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="reset" id="Cancel" value=" 取 消 " style="cursor:hand;">      </td>
    </tr>
  </table>
</form>
<%
end sub

'修改文章大类栏目
sub ModifyBig()
	dim BigID
	BigID=trim(request("BigID"))
	if BigID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		BigID=CLng(BigID)
	end if
	sqlStr="select * From BigClass where BigID=" & BigID
	set Rs=server.CreateObject ("Adodb.recordset")
	Rs.open sqlStr,conn,1,3
	if Rs.bof and Rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目！</li>"
	else
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
  <table width="50%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>修 改 文 章  栏 目</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目名称：</strong></td>
      <td><input name="BigName" type="text" value="<%=Rs("BigName")%>" size="37" maxlength="20"> 
        <input name="BigID" type="hidden" value="<%=Rs("BigID")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" value="SaveModifyBig"> 
        <input name="Submit" type="submit" value=" 保存修改结果 " style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="reset" id="Cancel" value=" 取 消 " style="cursor:hand;"> 
      </td>
    </tr>
  </table>
</form>
<%
	end if
	Rs.close
	set Rs=nothing
end sub
'修改文章大类栏目名称结束

'修改文章小类栏目
sub ModifySmall()
	dim SmallID
	SmallID=trim(request("SmallID"))
	if SmallID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	else
		SmallID=CLng(SmallID)
	end if
	sqlStr="select * From SmallClass where SmallID=" & SmallID
	set Rs=server.CreateObject ("Adodb.recordset")
	Rs.open sqlStr,conn,1,3
	if Rs.bof and Rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的栏目！</li>"
	else
%>
<form name="form1" method="post" action="Admin_Class_Article.asp" onSubmit="return check()">
  <table width="50%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="title"> 
      <td height="22" colspan="2" align="center"><strong>修 改 文 章 栏 目</strong></td>
    </tr>
    <tr class="tdbg"> 
      <td width="350"><strong>栏目名称：</strong></td>
      <td><input name="SmallName" type="text" value="<%=Rs("SmallName")%>" size="37" maxlength="20"> 
        <input name="SmallID" type="hidden" value="<%=Rs("SmallID")%>"></td>
    </tr>
    <tr class="tdbg"> 
      <td height="40" colspan="2" align="center"><input name="Action" type="hidden" value="SaveModifySmall"> 
        <input name="Submit" type="submit" value=" 保存修改结果 " style="cursor:hand;"> 
        &nbsp; <input name="Cancel" type="reset" id="Cancel" value=" 取 消 " style="cursor:hand;"> 
      </td>
    </tr>
  </table>
</form>
<%
	end if
	Rs.close
	set Rs=nothing
end sub
'修改文章小类栏目名称结束

%>
</body> 
</html> 
<% 
'===============================
'函数SaveBigAdd()保存添加文章大类名称
'===============================
sub SaveBigAdd()
	dim BigClassName
	dim sql,rs,trs

	BigClassName=trim(request("BigClassName"))
	if BigClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章类别名称不能为空！</li>"
		exit sub
	end if
	'检查数据库是否已存在此大类名称
	set Rs=server.createobject("adodb.recordset")
	sql="Select top 1 * From BigClass where BigName='" & BigClassName & "'"
	rs.open sql,conn,1,1
	if not(rs.bof and rs.eof) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章类别名称中已经存在栏目“" & BigClassName & "”！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.close
	'往数据库添加大类
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
'函数SaveSmallAdd()保存添加文章大类名称
'===============================
sub SaveSmallAdd()
	dim BigID,SmallClassName
	dim sql,rs,trs
	BigID=trim(request.Form("BigID"))
	SmallClassName=trim(request.Form("SmallClassName"))
	if SmallClassName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章类别名称不能为空！</li>"
		exit sub
	end if
	'检查数据库是否已存在此大类名称
	set Rs=server.createobject("adodb.recordset")
	sql="Select top 1 * From SmallClass where SmallName='" & SmallClassName & "' and BigID="&clng(BidID)
	rs.open sql,conn,1,1
	if not(rs.bof and rs.eof) then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章类别名称中已经存在栏目“" & SmallClassName & "”！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.close
	'往数据库添加大类
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
'函数SaveModifyBig()保存修改文章大类名称
'===============================
sub SaveModifyBig()
	dim BigID,BigName
	
	BigID=trim(request.Form("BigID"))
	if BigID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	end if
	BigName=trim(request.Form("BigName"))
	if BigName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章类别名称不能为空！</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	'往数据库添加大类
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
'函数SaveModifySmall()保存修改文章小类名称
'===============================
sub SaveModifySmall()
	dim SmallID,SmallName
	
	SmallID=trim(request.Form("SmallID"))
	if SmallID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
	end if
	SmallName=trim(request.Form("SmallName"))
	if SmallName="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>文章类别名称不能为空！</li>"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	'往数据库添加大类
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
'函数DeleteBig()删除文章大类及相应文章小类及其所属文章
'===============================
sub DeleteBig()
	dim BigID
	BigID=trim(Request("BigID"))
	if BigID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	end if
	
	sql="select * From BigClass where BigID="& clng(BigID)
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>栏目不存在，或者已经被删除</li>"
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
	'删除本栏目的所有文章
	conn.execute("delete from Article where BigID=" & clng(BigID))
	
	call CloseConn()
	response.redirect "Admin_Class_Article.asp"
		
end sub

'===============================
'函数DeleteSmall()删除文章小类及相应文章
'===============================
sub DeleteSmall()
	dim SmallID
	SmallID=trim(Request("SmallID"))
	if SmallID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>参数不足！</li>"
		exit sub
	end if
	
	sql="select * From SmallClass where SmallID="& clng(SmallID)
	set rs=server.CreateObject ("Adodb.recordset")
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>栏目不存在，或者已经被删除</li>"
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
	'删除本栏目的所有文章
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