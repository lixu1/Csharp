<!--#include file="inc/check.asp"-->
<!--#include file="connp.asp" -->
<link rel="stylesheet" href="Admin_Style.css">
  <script language="javascript">
  function Check(){
    if(myform.Title.value==""){
	alert("��վ˵������Ϊ��!");
	myform.Title.focus();
	return false;
	}
  	if(myform.BgPic3.value==myform.BgPic3.defaultValue){
		alert("���ӵ�ַ����Ϊ�գ�");
		myform.BgPic3.focus();
		return false;
	}
      if(myform.IsPic.checked==true){	
	 	if(myform.Pic.value==""){
		alert("ͼƬ��ַ����Ϊ��!");
		return false;
		}
		}
	}
   function CheckLink(){
  	if(myform.Title.value==""){
		alert("���ӱ��ⲻ��Ϊ�գ�");
		myform.Title.focus();
		return false;
	}
	//if(myform.WebSite.value==myform.WebSite.defaultValue){
		////alert("���ӵ�ַ����Ϊ�գ�");
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
		Response.write"<script>alert('��ʱ��û�м�¼!�������������!');document.location='link_man.asp?action=LinkAdd'</script>"
	Else
%>
<table width="95%" border="0"  align=center cellpadding="3" cellspacing="1" bgcolor="#CCCCCC" class="tableBorder">
  <tr bgcolor="#FFFFFF"> 
    <th height=25 colspan=4 align=left id=tabletitlelink  class="tdbg"><div align="center">��������<a href="?action=LinkAdd">[���]</a></div></th>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td width="26%" class=Forumrow> <div align="center"><strong>˵��</strong></div></td>
    <td width="43%" class=Forumrow> <div align="center"><strong>���ӵ�ַ</strong></div></td>
    <td width="12%" align="center" class=Forumrow><strong>�Ƿ�ͼƬ����</strong></td>
    <td width="19%" class=Forumrow> <div align="center"><strong>����</strong></div></td>
  </tr>
  <%
  	While not Rsts.eof
  %>
  <tr bgcolor="#FFFFFF"> 
    <td height="30" class=Forumrow> <div align="center"><%=Rsts("Title")%></div></td>
    <td class=Forumrow> <div align="center"><%=Rsts("WebSite")%></div></td>
    <td align="center" class=Forumrow><%If Rsts("Ispic")=0 Then response.write("��") else response.write("<img src="&rsts("pic")&" height=30 width=90>") end if%></td>
    <td class=Forumrow> <div align="center"><a href="?action=LinkModify&id=<%=Rsts("Id")%>">�޸�</a>��<a href="?action=LinkDel&id=<%=Rsts("Id")%>">ɾ��</a></div></td>
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
      <td height=25 align=left class="tdbg"><B>˵��</B>�����������������ϵͳ�����е��������ӡ�</td>
    </tr>
  </table>
  <table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
    <tr> 
      <th height=25 colspan=2 align=left id=tabletitlelink>&nbsp;</th>
    </tr>
    <tr> 
      <td width="28%" height="30" class=Forumrow>��վ˵����</td>
      <td width="72%" class=Forumrow><input name="Title" type="text" size="50"></td>
    </tr>
    <tr> 
      <td height="30" class=Forumrow>���ӵ�ַ��</td>
      <td class=Forumrow><input name="WebSite" type="text" id="BgPic3" value="http://" size="50"></td>
    </tr>
    <tr> 
      <td height="26" class=Forumrow>���ͼƬ��</td>
      <td height="30" class=Forumrow> 
        <iframe name="frameHeadPic" width="500" height="30" frameborder="no" scrolling=no src="UploadLink.asp"></iframe></td>
    </tr>
    <!--tr> 
      <td height="30" class=Forumrow>�Ƿ���ͼƬ���ӣ�</td>
      <td class=Forumrow><input name="IsPic" type="checkbox" id="IsPic3" value="1" checked="checked"> 
        <font color="#990000">���������ͼƬ��ѡ�� </font> ע��:�������������,ȡ����ѡ����!��ʱ��վ˵����Ϊ��ҳ������ʾ����!</td>
    </tr-->
    <tr>
      <td height="30" class=Forumrow>ͼƬ��ַ��</td>
      <td class=Forumrow><input name="Pic" type="text" size="50" readonly>
        (�ϴ�ͼƬ�Զ����)</td>
    </tr>
    <tr> 
      <td colspan="2" class=Forumrow> <div align="center"> 
          <input type="submit" class=buttonmain name="Submit3" value=" �� ��  " onclick="return Check();">
        </div></td>
    </tr>
    <tr>
      <td colspan="2" class=Forumrow>ע��:�������������,�����ϴ�ͼƬ,��վ˵����Ϊ������ʾ����!ͼƬ����ʱ,��Ҫ�ϴ�ͼƬ!</td>
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
				response.write "����������ӳɹ���<a href='link_man.asp'>����</a>"
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
      <td height=25 align=left class="tdbg"><B>˵��</B>�����������������ϵͳ�����е��������ӡ�</td>
    </tr>
  </table>
  <table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
    <tr> 
      <th height=25 colspan=2 align=left id=tabletitlelink>&nbsp;</th>
    </tr>
    <tr> 
      <td width="28%" height="30" class=Forumrow>��վ˵����</td>
      <td width="72%" class=Forumrow><input name="Title" type="text" id="˵��" value="<%=rst1("Title")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="30" class=Forumrow>���ӵ�ַ��</td>
      <td class=Forumrow><input name="WebSite" type="text" id="WebSite" value="<%=rst1("WebSite")%>" size="50"></td>
    </tr>
    <tr> 
      <td height="30" class=Forumrow>���ͼƬ��</td>
      <td class=Forumrow><iframe name="frameHeadPic" width="500" height="25" frameborder="no" scrolling=no src="UploadLink.asp"></iframe></td>
    </tr>
    <!--tr> 
      <td height="30" class=Forumrow>�Ƿ���ͼƬ���ӣ�</td>
      <td class=Forumrow><input name="IsPic" type="checkbox" id="IsPic4" value="1"> 
        <font color="#990000">���������ͼƬ��ѡ��</font></td>
    </tr-->
    <tr> 
      <td height="30" class=Forumrow>ͼƬ��ַ��</td>
      <td class=Forumrow><input name="Pic" type="text" id="Pic" size="50" readonly>
        (�ϴ�ͼƬ�Զ����)</td>
    </tr>
    <tr> 
      <td colspan="2" class=Forumrow> <div align="center"> 
          <input type="hidden" name="Id" id="Id" value="<%=rst1("ID")%>">
          <input type="submit" class=buttonmain name="Submit" value=" �� ��   " onclick="return CheckLink();">
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
				response.write("���������޸ĳɹ���<a href='link_man.asp'>����</a>")
			rsts.close
			set rsts=nothing
End Sub

Sub LinkDel()
	conn1.execute("Delete From Link Where Id="&Request("Id"))
	response.write("��������ɾ���ɹ���<a href='link_man.asp'>����</a>")
End Sub
%>