<!--#include file="inc/check.asp"-->
<!--#include file="inc/conn.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="Admin_Style.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE4 {color: #FF3300}
-->
</style>
<body>
<script language="JavaScript" type="text/JavaScript">
//������Ϣ��д���
function checkFormData(theForm){
	if(theForm.imgname.value==""){ alert("����дͼƬ���ƣ�");theForm.imgname.focus();return false; }
	if(theForm.imgpath.value==""){ alert("���ϴ�ͼƬ��");theForm.imgpath.focus();return false; }
	if(theForm.sort.value==""){ alert("����д�����ţ�");theForm.sort.focus();return false; }
	return true;
}
</script>
<%
Dim action, id, rs, sql
action = LCase(Trim(request.QueryString("action"))) '����LCase �����Ѵ�д��ĸת��ΪСд��ĸ
id = CInt(request.QueryString("id"))
Select Case action
    Case ""
        Call showlist() '��ʾ��Ϣ�б�
    Case "showlist"
        Call showlist() '��ʾ��Ϣ�б�
    Case "show"
        Call show(id) '��ʾ��Ϣ����ϸ���
    Case "showadd"
        Call showadd() '��ʾ����Ѽ���Ϣ
    Case "showedit"
        Call showedit(id) '��ʾ�޸��Ѽ���Ϣ
    Case "showdel"
        Call showdel(id) '��ʾɾ���Ѽ���Ϣ
    Case "dbadd"
        Call dbadd() '�����ݿ��������Ϣ
    Case "dbedit"
        Call dbedit(id) '�����ݿ��и�����Ϣ
    Case "dbdel"
        Call dbdel(id) '�����ݿ���ɾ����Ϣ
End Select
'ʵ����ʾ�б���

Function showlist()
    Dim currentpage, endpage, page_count, Pcount, totalrec, i
    sql = "SELECT * FROM img"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open sql, Conn, 1, 1
    If rs.bof And rs.EOF Then
        response.Write "���ݿ��еļ�¼��Ϊ�գ�����Ӽ�¼��"
        Exit Function
    End If

    totalrec = rs.recordcount '��¼����
    If totalrec = 0 Then
        response.Write("��ʱû��ͼƬ��")
        response.End()
    End If
    If request("page") = "" Or Not IsNumeric(request("page")) Then 'isNumeric���������ֵΪTrue�����е���˼�ǣ�Ϊ�ջ�����������һ��ֵΪ1
        currentPage = 1
    Else
        currentPage = CInt(request("page")) '�ѻ�ȡ��ֵת��Ϊ����
    End If

    rs.PageSize = 20 'ÿҳ����ʾ�ļ�¼����
    rs.AbsolutePage = currentpage '����Ŀǰ��ҳ��
    page_count = 0 '�Ѿ���ʾ�ļ�¼��
%>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="2">
  <tr class="tdbg">
    <td height="30" colspan="5" align="center" class="title"><strong>��ʾͼƬ�б�</strong></td>
  </tr>
  <tr class="tdbg">
    <td colspan="5" align="left" valign="top"><%call Pages(currentpage,totalrec)%></td>
  </tr>
  <tr class="tdbg">
    <td width="12%" align="left">���</td>
    <td width="38%" align="left">ͼƬ����</td>
    <td width="16%" align="left">����</td>
    <td width="16%" align="left">�Ƿ���ʾ</td>
    <td width="18%" align="left">����</td>
  </tr>
  <%
While (Not rs.EOF) And (Not page_count = rs.PageSize)
    page_count = page_count + 1
    rsid = rs("imgid") 'Ϊ�˱�������޷��ع���ȱ���ȴ���������
%>
  <tr class="tdbg">
    <td align="left"><%=rs("sort")%></td>
    <td align="left"><%=rs("imgname")%></td>
    <td align="left">
	<%
	if rs("Role")=1 then
	response.Write("��ҳ����")
	end if
	if rs("Role")=2 then
	response.Write("��Ʒչʾ")
	end if
	%></td>
    <td align="left">
	<%
	if rs("show")=true then
	response.Write("��")
	else
	response.Write("��")
	end if
	%></td>
    <td align="left"><a href="?action=show&id=<%=rsid%>">�鿴</a>||<a href="?action=showedit&id=<%=rsid%>">�޸�</a>||<a href="?action=showdel
&id=<%=rsid%>">ɾ��</a></td>
  </tr>
  <%
rs.movenext
Wend
%>
  <tr class="tdbg">
    <td colspan="5" align="left"><%call Pages(currentpage,totalrec)%></td>
  </tr>
</table>
<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
End Function

Function show(id) '��ʾ��Ϣ����ϸ���
    sql = "SELECT * FROM img WHERE imgid="&id
    Set rs = conn.Execute(sql)
	imgname=rs("imgname")
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="2">
  <tr class="tdbg">
    <td height="30" colspan="2" align="center" class="title"><strong>��ʾ��ϸ��Ϣ</strong></td>
  </tr>
  <tr class="tdbg">
    <td align="right">����</td>
    <td align="left">
	<%
	if rs("Role")=1 then
	response.Write("��ҳ����")
	end if
	if rs("Role")=2 then
	response.Write("��Ʒչʾ")
	end if
	%></td>
  </tr>
  <tr class="tdbg">
    <td width="12%" align="right">ͼƬ����</td>
    <td width="88%" align="left"><%=rs("imgname")%></td>
  </tr>
  <tr class="tdbg">
    <td align="right" valign="top">ͼƬ</td>
    <td align="left" valign="top"><a href="<%=rs("imgpath")%>" title="<%=imgname%>" target="_blank"><img src="<%=rs("imgpath")%>" alt="<%=imgname%>" border="0"></a></td>
  </tr>
  <tr class="tdbg">
    <td align="right" valign="top">���ӵ�ַ</td>
    <td align="left"><%=rs("imglink")%></td>
  </tr>
  <tr class="tdbg">
    <td align="right" valign="top">�Ƿ���ʾ:</td>
    <td align="left"><span class="STYLE4">
      <%
	s=rs("show")
	if s="True" then
	response.Write("��")
	else
	response.Write("��")
	end if
	%>
    </span></td>
  </tr>
  <tr class="tdbg">
    <td align="right" valign="top">����:</td>
    <td align="left"><%=rs("sort")%></td>
  </tr>
</table>
<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
End Function

Function showadd() '��ʾ����Ѽ���Ϣ
%>
<form action="?action=dbadd" method="post" name="form1" id="form1">
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="2">
    <tr class="tdbg">
      <td height="30" colspan="2" align="center" valign="top" class="title"><strong>���ͼƬ</strong></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">����</td>
      <td align="left" valign="top"><select name="Role" id="Role">
        <option value="1" selected>��ҳ����</option>
            </select></td>
    </tr>
    <tr class="tdbg">
      <td width="12%" align="right" valign="top">ͼƬ����:</td>
      <td width="88%" align="left" valign="top"><input name="imgname" type="text" id="imgname" size="50">      </td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">ͼƬ·��:</td>
      <td align="left" valign="top"><input name="imgpath" type="text" id="imgpath" value="<%=pathimg%>" size="50">
      ��Ѵ�С150px*100px</td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="middle">�ϴ�ͼƬ:</td>
      <td height="25" align="left" valign="middle"><iframe src="uploadimg.asp" width="400" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">���ӵ�ַ</td>
      <td align="left" valign="top"><input name="imglink" type="text" id="imglink" size="50"></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">����:</td>
      <td align="left" valign="top"><input name="sort" type="text" id="sort" size="5">
      *����</td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">��ʾ:</td>
      <td align="left" valign="top"><input name="show" type="checkbox" id="show" value="1"></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">&nbsp;</td>
      <td align="left" valign="top"><input type="submit" name="Submit" value="ȷ���ύ" onClick="return checkFormData(this.form)">
        <input type="reset" name="Submit2" value="������д">      </td>
    </tr>
  </table>
</form>
<%
End Function

Function showedit(id) '��ʾ�޸��Ѽ���Ϣ
    sql = "SELECT * FROM img WHERE  imgid ="&id
    Set rs = conn.Execute(sql)
  'Ϊ�˱�������޷��ع���ȱ���ȴ���������
    pathimg = rs("imgpath")
%>
<form id="form1" name="form1" method="post" action="?action=dbedit&id=<%=rs("imgid")%>">
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="2">
    <tr class="tdbg">
      <td height="30" colspan="2" align="center" class="title"><strong>��ʾ�޸���Ϣ</strong></td>
    </tr>
    <tr class="tdbg">
      <td align="right">����</td>
      <td><select name="Role" id="Role">
        <option value="1" <%if rs("Role")=1 then response.Write("selected")%>>��ҳ����</option>
      </select></td>
    </tr>
    <tr class="tdbg">
      <td align="right">ͼƬ����:</td>
      <td><input name="imgname" type="text" id="imgname" value="<%=rs("imgname")%>" size="50">      </td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">ͼƬ·��:</td>
      <td><input name="imgpath" type="text" id="imgpath" value="<%=pathimg%>" size="50">
      ��Ѵ�С150px*100px</td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">�ϴ�ͼƬ:</td>
      <td><iframe src="uploadimg.asp" width="400" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">���ӵ�ַ</td>
      <td><input name="imglink" type="text" id="imglink" value="<%=rs("imglink")%>" size="50"></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">����:</td>
      <td><input name="sort" type="text" id="sort" value="<%=rs("sort")%>" size="5">
      *����</td>
    </tr>
    <tr class="tdbg">
      <td align="right">��ʾ:</td>
      <td><input name="show" type="checkbox" id="show" value="1" 
		<%
		s=rs("show")
		if s="True" then
		response.Write("checked")
		end if
		%>></td>
    </tr>
    <tr class="tdbg">
      <td width="13%" align="right">&nbsp;</td>
      <td width="87%"><input type="submit" name="Submit3" value="ȷ���޸�" onClick="return checkFormData(this.form)">
      <input type="reset" name="Submit22" value="������д"></td>
    </tr>
  </table>
</form>
<%
End Function

Function showdel(id) '��ʾɾ���Ѽ���Ϣ
%>
<form name="form1" action="?action=dbdel&id=<%=id%>" method="post">
</form>
<script language="JavaScript" type="text/JavaScript">
if(confirm("ȷ��Ҫɾ����"))
{
	document.form1.submit();
}
else
{
	window.history.go(-1)
}
</script>
<%
End Function

Function dbadd() '�����ݿ��������Ϣ
    Dim imgname, imgpath, ssort, show 
	 Role = Trim(request("Role"))
    imgname = Trim(request("imgname"))
    imgpath = Trim(request.Form("imgpath"))
	imglink = Trim(request.Form("imglink"))
    ssort = Trim(request.Form("sort"))
    show = Trim(request.Form("show"))
    sql = "INSERT INTO img(imgname,imgpath,imglink,sort,show,Role) VALUES ('"&imgname&"','"&imgpath&"','"&imglink&"','"&ssort&"','"&show&"','"&Role&"')"
    Set rs = conn.Execute(sql)
    response.Write("���ͼƬ�ɹ�!")
    response.Redirect("img.asp")
End Function

Function dbedit(id) '�����ݿ��и�����Ϣ
    Dim imgname, imgpath, ssort, show
	Role = Trim(request("Role"))
    imgname = Trim(request("imgname"))
    imgpath = Trim(request.Form("imgpath"))
	imglink = Trim(request.Form("imglink"))
    ssort = Trim(request.Form("sort"))
    show = Trim(request.Form("show"))
	if show="1" then
	else
	show=0
	end if
    sql = "UPDATE img SET imgname='"&imgname&"',imglink='"&imglink&"',imgpath='"&imgpath&"',sort='"&ssort&"',show='"&show&"',Role='"&Role&"' WHERE imgid="&id
    Set rs = conn.Execute(sql)
    response.Write("�޸�ͼƬ�ɹ�!")
    response.Redirect("img.asp")
End Function

Function dbdel(id) '�����ݿ���ɾ����Ϣ
    sql = "DELETE FROM img WHERE imgid = "&id
    Set rs = conn.Execute(sql)
    response.Write("ɾ��ͼƬ�ɹ�!")
    response.Redirect("img.asp")
End Function

Function Pages(currentpage, totalrec)
    Pcount = rs.PageCount 'ͳ�Ƽ�¼������
    response.Write "ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"&_
    " ÿҳ<b>"&rs.pagesize&"</b>,  ����:<b>"&totalrec&"</b>,  "&_
    " ��ҳ��"

    If currentpage > 3 Then '��ǰҳ������3
        response.Write " <a href=?page=1>[1]</a> ..."
    End If

    If Pcount>currentpage+3 Then '��ǰҳ����������ǰ3ҳ
        endpage = currentpage+3
    Else
        endpage = Pcount
    End If

    For i = currentpage -2 To endpage
        If Not i<1 Then
            If i = CLng(currentpage) Then
                response.Write " ["&i&"]"
            Else
                response.Write " <a href=?page="&i&">["&i&"]</a>"
            End If
        End If
    Next

    If currentpage+3 < Pcount Then
        response.Write " ...<a href=?page="&Pcount&">["&Pcount&"]</a>"
    End If
End Function
%>
</body>
