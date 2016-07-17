<!--#include file="inc/check.asp"-->
<!--#include file="connp.asp" -->
<link rel="stylesheet" href="Admin_Style.css">
  <script language="javascript">
  function Check(){
    if(myform.Title.value==""){
	alert("网站说明不能为空!");
	myform.Title.focus();
	return false;
	}
  	if(myform.BgPic3.value==myform.BgPic3.defaultValue){
		alert("链接地址不能为空！");
		myform.BgPic3.focus();
		return false;
	}
      if(myform.IsPic.checked==true){	
	 	if(myform.Pic.value==""){
		alert("图片地址不能为空!");
		return false;
		}
		}
	}
   function CheckLink(){
  	if(myform.Title.value==""){
		alert("链接标题不能为空！");
		myform.Title.focus();
		return false;
	}
	//if(myform.WebSite.value==myform.WebSite.defaultValue){
		////alert("链接地址不能为空！");
		//myform.WebSite.focus();
		//return false;
	//}
  }	 	
	
	
	</script>
	
<%
select case Request("action")
case "Link"  
	call Link()
case "LinkAdd"  
	call LinkAdd()
case "LinkAddSave"  
	call LinkAddSave()
case "LinkModify"  
	call LinkModify()
case "LinkModifySave"  
	call LinkModifySave()
case "LinkDel"  
	call LinkDel()
case else
	call link()	
	
end select		
%>
<%
Sub Link()
	set rsts=server.createobject("adodb.recordset")
	sqls="select * from Link"
	rsts.open sqls,conn1,1,1
	If rsts.eof Then 
		Response.write"<script>alert('暂时还没有记录!请添加友情链接!');document.location='link_man.asp?action=LinkAdd'</script>"
	Else
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" class="tableBorder">
  <tr bgcolor="#FFFFFF"> 
    <th height=25 colspan=4 align=left id=tabletitlelink  class="tdbg"><div align="center">友情链接<a href="?action=LinkAdd">[添加]</a></div></th>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="26%" class=Forumrow> <div align="center"><strong>说明</strong></div></td>
    <td width="43%" class=Forumrow> <div align="center"><strong>链接地址</strong></div></td>
    <td width="12%" align="center" class=Forumrow><strong>是否图片链接</strong></td>
    <td width="19%" class=Forumrow> <div align="center"><strong>操作</strong></div></td>
  </tr>
  <%
  	While not Rsts.eof
  %>
  <tr bgcolor="#FFFFFF"> 
    <td height="30" class=Forumrow> <div align="center"><%=Rsts("Title")%></div></td>
    <td class=Forumrow> <div align="center"><%=Rsts("WebSite")%></div></td>
    <td align="center" class=Forumrow><%If Rsts("Ispic")=0 Then response.write("否") else response.write("<img src="&rsts("pic")&" height=30 width=90>") end if%></td>
    <td class=Forumrow> <div align="center"><a href="?action=LinkModify&id=<%=Rsts("Id")%>">修改</a>　<a href="?action=LinkDel&id=<%=Rsts("Id")%>">删除</a></div></td>
  </tr>
  <%
 	Rsts.movenext
	Wend
 %>
</table>
<%
	End If
	Rsts.close
	Set Rsts=Nothing
End Sub

Sub LinkAdd()%>
<form method="POST" action="?action=LinkAddSave" name=myform>
  <table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
    <tr> 
      <td height=25 align=left class="tdbg"><B>说明</B>：这里的设置了所有系统中所有的友情链接。</td>
    </tr>
  </table>
  <table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
    <tr> 
      <th height=25 colspan=2 align=left id=tabletitlelink>&nbsp;</th>
    </tr>
    <tr> 
      <td width="28%" height="30" class=Forumrow>网站说明：</td>
      <td width="72%" class=Forumrow><input name="Title" type="text" size="50"></td>
    </tr>
    <tr> 
      <td height="30" class=Forumrow>链接地址：</td>
      <td class=Forumrow><input name="WebSite" type="text" id="BgPic3" value="http://" size="50"></td>
    </tr>
    <tr> 
      <td height="26" class=Forumrow>添加图片：</td>
      <td height="30" class=Forumrow> 
        <iframe name="frameHeadPic" width="500" height="30" frameborder="no" scrolling=no src="UploadLink.asp"></iframe></td>
    </tr>
    <!--tr> 
      <td height="30" class=Forumrow>是否是图片链接：</td>
      <td class=Forumrow><input name="IsPic" type="checkbox" id="IsPic3" value="1" checked="checked"> 
        <font color="#990000">（如果包含图片请选择） </font> 注意:如果是文字链接,取消勾选即可!此时网站说明即为首页链接显示内容!</td>
    </tr-->
    <tr>
      <td height="30" class=Forumrow>图片地址：</td>
      <td class=Forumrow><input name="Pic" type="text" size="50" readonly>
        (上传图片自动添加)</td>
    </tr>
    <tr> 
      <td colspan="2" class=Forumrow> <div align="center"> 
          <input type="submit" class=buttonmain name="Submit3" value=" 添 加  " onclick="return Check();">
        </div></td>
    </tr>
    <tr>
      <td colspan="2" class=Forumrow>注意:如果是文字链接,不用上传图片,网站说明即为链接显示内容!图片链接时,需要上传图片!</td>
    </tr>
  </table>
</form>
<%
End Sub

Sub LinkAddSave()
			set rsts=server.createobject("adodb.recordset")
			sqls="select * from Link"
			rsts.open sqls,conn1,2,3
				rsts.addnew
				rsts("Title")=request.form("Title")
				rsts("WebSite")=request.form("WebSite")
				rsts("Pic")=request.form("Pic")
				If request.form("IsPic") = 1 Then rsts("IsPic")=request.form("IsPic") Else rsts("IsPic")=0 End If
				rsts("AddDate")=now()
				rsts.update
				response.write "友情链接添加成功。<a href='link_man.asp'>返回</a>"
			rsts.close
			set rsts=nothing
End Sub

Sub LinkModify()
	sql1="select * from Link Where id="&request("Id")
	set rst1=server.createobject("adodb.recordset")
	rst1.open sql1,conn1,1,1
%>
<form method="POST" action="?action=LinkModifySave" name=myform>
  <table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center>
    <tr> 
      <td height=25 align=left class="tdbg"><B>说明</B>：这里的设置了所有系统中所有的友情链接。</td>
    </tr>
  </table>
  <table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
    <tr> 
      <th height=25 colspan=2 align=left id=tabletitlelink>&nbsp;</th>
    </tr>
    <tr> 
      <td width="28%" height="30" class=Forumrow>网站说明：</td>
      <td width="72%" class=Forumrow><input name="Title" type="text" id="说明" value="<%=rst1("Title")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="30" class=Forumrow>链接地址：</td>
      <td class=Forumrow><input name="WebSite" type="text" id="WebSite" value="<%=rst1("WebSite")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="30" class=Forumrow>添加图片：</td>
      <td class=Forumrow><iframe name="frameHeadPic" width="500" height="25" frameborder="no" scrolling=no src="UploadLink.asp"></iframe></td>
    </tr>
    <!--tr> 
      <td height="30" class=Forumrow>是否是图片链接：</td>
      <td class=Forumrow><input name="IsPic" type="checkbox" id="IsPic4" value="1"> 
        <font color="#990000">（如果包含图片请选择）</font></td>
    </tr-->
    <tr> 
      <td height="30" class=Forumrow>图片地址：</td>
      <td class=Forumrow><input name="Pic" type="text" id="Pic" size="50" readonly>
        (上传图片自动添加)</td>
    </tr>
    <tr> 
      <td colspan="2" class=Forumrow> <div align="center"> 
          <input type="hidden" name="Id" id="Id" value="<%=rst1("ID")%>">
          <input type="submit" class=buttonmain name="Submit" value=" 修 改   " onclick="return CheckLink();">
        </div></td>
    </tr>
  </table>
</form>
<%
rst1.close
set rst1=nothing
End Sub

Sub LinkModifySave()
			set rsts=server.createobject("adodb.recordset")
			sqls="select * from Link Where Id="&request.form("id")
			rsts.open sqls,conn1,2,3

				rsts("Title")=request.form("Title")
				rsts("WebSite")=request.form("WebSite")
				rsts("Pic")=request.form("Pic")
				If request.form("IsPic") = 1 Then rsts("IsPic")=request.form("IsPic") Else rsts("IsPic")=0 End If
				rsts.update
				response.write("友情链接修改成功。<a href='link_man.asp'>返回</a>")
			rsts.close
			set rsts=nothing
End Sub

Sub LinkDel()
	conn1.execute("Delete From Link Where Id="&Request("Id"))
	response.write("友情链接删除成功。<a href='link_man.asp'>返回</a>")
End Sub
%>