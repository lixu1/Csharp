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
//检查表信息填写情况
function checkFormData(theForm){
	if(theForm.imgname.value==""){ alert("请填写图片名称！");theForm.imgname.focus();return false; }
	if(theForm.imgpath.value==""){ alert("请上传图片！");theForm.imgpath.focus();return false; }
	if(theForm.sort.value==""){ alert("请填写排序编号！");theForm.sort.focus();return false; }
	return true;
}
</script>
<%
Dim action, id, rs, sql
action = LCase(Trim(request.QueryString("action"))) '利用LCase 函数把大写字母转换为小写字母
id = CInt(request.QueryString("id"))
Select Case action
    Case ""
        Call showlist() '显示信息列表
    Case "showlist"
        Call showlist() '显示信息列表
    Case "show"
        Call show(id) '显示信息的详细情况
    Case "showadd"
        Call showadd() '显示添加搜集信息
    Case "showedit"
        Call showedit(id) '显示修改搜集信息
    Case "showdel"
        Call showdel(id) '显示删除搜集信息
    Case "dbadd"
        Call dbadd() '在数据库中添加信息
    Case "dbedit"
        Call dbedit(id) '在数据库中更新信息
    Case "dbdel"
        Call dbdel(id) '在数据库中删除信息
End Select
'实现显示列表功能

Function showlist()
    Dim currentpage, endpage, page_count, Pcount, totalrec, i
    sql = "SELECT * FROM img"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3
    rs.Open sql, Conn, 1, 1
    If rs.bof And rs.EOF Then
        response.Write "数据库中的记录数为空！请添加记录！"
        Exit Function
    End If

    totalrec = rs.recordcount '记录总数
    If totalrec = 0 Then
        response.Write("暂时没有图片！")
        response.End()
    End If
    If request("page") = "" Or Not IsNumeric(request("page")) Then 'isNumeric如果是数字值为True，整行的意思是，为空或不是数字令下一行值为1
        currentPage = 1
    Else
        currentPage = CInt(request("page")) '把获取的值转化为整型
    End If

    rs.PageSize = 20 '每页的显示的记录总数
    rs.AbsolutePage = currentpage '返回目前的页码
    page_count = 0 '已经显示的记录数
%>
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="2">
  <tr class="tdbg">
    <td height="30" colspan="5" align="center" class="title"><strong>显示图片列表</strong></td>
  </tr>
  <tr class="tdbg">
    <td colspan="5" align="left" valign="top"><%call Pages(currentpage,totalrec)%></td>
  </tr>
  <tr class="tdbg">
    <td width="12%" align="left">序号</td>
    <td width="38%" align="left">图片名称</td>
    <td width="16%" align="left">作用</td>
    <td width="16%" align="left">是否显示</td>
    <td width="18%" align="left">操作</td>
  </tr>
  <%
While (Not rs.EOF) And (Not page_count = rs.PageSize)
    page_count = page_count + 1
    rsid = rs("imgid") '为了避免出现无法回滚的缺陷先传到变量里
%>
  <tr class="tdbg">
    <td align="left"><%=rs("sort")%></td>
    <td align="left"><%=rs("imgname")%></td>
    <td align="left">
	<%
	if rs("Role")=1 then
	response.Write("首页滚动")
	end if
	if rs("Role")=2 then
	response.Write("产品展示")
	end if
	%></td>
    <td align="left">
	<%
	if rs("show")=true then
	response.Write("是")
	else
	response.Write("否")
	end if
	%></td>
    <td align="left"><a href="?action=show&id=<%=rsid%>">查看</a>||<a href="?action=showedit&id=<%=rsid%>">修改</a>||<a href="?action=showdel
&id=<%=rsid%>">删除</a></td>
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

Function show(id) '显示信息的详细情况
    sql = "SELECT * FROM img WHERE imgid="&id
    Set rs = conn.Execute(sql)
	imgname=rs("imgname")
%>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="2">
  <tr class="tdbg">
    <td height="30" colspan="2" align="center" class="title"><strong>显示详细信息</strong></td>
  </tr>
  <tr class="tdbg">
    <td align="right">作用</td>
    <td align="left">
	<%
	if rs("Role")=1 then
	response.Write("首页滚动")
	end if
	if rs("Role")=2 then
	response.Write("产品展示")
	end if
	%></td>
  </tr>
  <tr class="tdbg">
    <td width="12%" align="right">图片名称</td>
    <td width="88%" align="left"><%=rs("imgname")%></td>
  </tr>
  <tr class="tdbg">
    <td align="right" valign="top">图片</td>
    <td align="left" valign="top"><a href="<%=rs("imgpath")%>" title="<%=imgname%>" target="_blank"><img src="<%=rs("imgpath")%>" alt="<%=imgname%>" border="0"></a></td>
  </tr>
  <tr class="tdbg">
    <td align="right" valign="top">链接地址</td>
    <td align="left"><%=rs("imglink")%></td>
  </tr>
  <tr class="tdbg">
    <td align="right" valign="top">是否显示:</td>
    <td align="left"><span class="STYLE4">
      <%
	s=rs("show")
	if s="True" then
	response.Write("是")
	else
	response.Write("否")
	end if
	%>
    </span></td>
  </tr>
  <tr class="tdbg">
    <td align="right" valign="top">排序:</td>
    <td align="left"><%=rs("sort")%></td>
  </tr>
</table>
<%
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
End Function

Function showadd() '显示添加搜集信息
%>
<form action="?action=dbadd" method="post" name="form1" id="form1">
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="2">
    <tr class="tdbg">
      <td height="30" colspan="2" align="center" valign="top" class="title"><strong>添加图片</strong></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">作用</td>
      <td align="left" valign="top"><select name="Role" id="Role">
        <option value="1" selected>首页滚动</option>
            </select></td>
    </tr>
    <tr class="tdbg">
      <td width="12%" align="right" valign="top">图片名称:</td>
      <td width="88%" align="left" valign="top"><input name="imgname" type="text" id="imgname" size="50">      </td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">图片路径:</td>
      <td align="left" valign="top"><input name="imgpath" type="text" id="imgpath" value="<%=pathimg%>" size="50">
      最佳大小150px*100px</td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="middle">上传图片:</td>
      <td height="25" align="left" valign="middle"><iframe src="uploadimg.asp" width="400" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">链接地址</td>
      <td align="left" valign="top"><input name="imglink" type="text" id="imglink" size="50"></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">排序:</td>
      <td align="left" valign="top"><input name="sort" type="text" id="sort" size="5">
      *整数</td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">显示:</td>
      <td align="left" valign="top"><input name="show" type="checkbox" id="show" value="1"></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">&nbsp;</td>
      <td align="left" valign="top"><input type="submit" name="Submit" value="确认提交" onClick="return checkFormData(this.form)">
        <input type="reset" name="Submit2" value="重新填写">      </td>
    </tr>
  </table>
</form>
<%
End Function

Function showedit(id) '显示修改搜集信息
    sql = "SELECT * FROM img WHERE  imgid ="&id
    Set rs = conn.Execute(sql)
  '为了避免出现无法回滚的缺陷先传到变量里
    pathimg = rs("imgpath")
%>
<form id="form1" name="form1" method="post" action="?action=dbedit&id=<%=rs("imgid")%>">
  <table width="98%" border="0" align="center" cellpadding="3" cellspacing="2">
    <tr class="tdbg">
      <td height="30" colspan="2" align="center" class="title"><strong>显示修改信息</strong></td>
    </tr>
    <tr class="tdbg">
      <td align="right">作用</td>
      <td><select name="Role" id="Role">
        <option value="1" <%if rs("Role")=1 then response.Write("selected")%>>首页滚动</option>
      </select></td>
    </tr>
    <tr class="tdbg">
      <td align="right">图片名称:</td>
      <td><input name="imgname" type="text" id="imgname" value="<%=rs("imgname")%>" size="50">      </td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">图片路径:</td>
      <td><input name="imgpath" type="text" id="imgpath" value="<%=pathimg%>" size="50">
      最佳大小150px*100px</td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">上传图片:</td>
      <td><iframe src="uploadimg.asp" width="400" marginwidth="0" height="25" marginheight="0" scrolling="no" frameborder="0"></iframe></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">链接地址</td>
      <td><input name="imglink" type="text" id="imglink" value="<%=rs("imglink")%>" size="50"></td>
    </tr>
    <tr class="tdbg">
      <td align="right" valign="top">排序:</td>
      <td><input name="sort" type="text" id="sort" value="<%=rs("sort")%>" size="5">
      *整数</td>
    </tr>
    <tr class="tdbg">
      <td align="right">显示:</td>
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
      <td width="87%"><input type="submit" name="Submit3" value="确认修改" onClick="return checkFormData(this.form)">
      <input type="reset" name="Submit22" value="重新填写"></td>
    </tr>
  </table>
</form>
<%
End Function

Function showdel(id) '显示删除搜集信息
%>
<form name="form1" action="?action=dbdel&id=<%=id%>" method="post">
</form>
<script language="JavaScript" type="text/JavaScript">
if(confirm("确认要删除吗？"))
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

Function dbadd() '在数据库中添加信息
    Dim imgname, imgpath, ssort, show 
	 Role = Trim(request("Role"))
    imgname = Trim(request("imgname"))
    imgpath = Trim(request.Form("imgpath"))
	imglink = Trim(request.Form("imglink"))
    ssort = Trim(request.Form("sort"))
    show = Trim(request.Form("show"))
    sql = "INSERT INTO img(imgname,imgpath,imglink,sort,show,Role) VALUES ('"&imgname&"','"&imgpath&"','"&imglink&"','"&ssort&"','"&show&"','"&Role&"')"
    Set rs = conn.Execute(sql)
    response.Write("添加图片成功!")
    response.Redirect("img.asp")
End Function

Function dbedit(id) '在数据库中更新信息
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
    response.Write("修改图片成功!")
    response.Redirect("img.asp")
End Function

Function dbdel(id) '在数据库中删除信息
    sql = "DELETE FROM img WHERE imgid = "&id
    Set rs = conn.Execute(sql)
    response.Write("删除图片成功!")
    response.Redirect("img.asp")
End Function

Function Pages(currentpage, totalrec)
    Pcount = rs.PageCount '统计记录的总数
    response.Write "页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
    " 每页<b>"&rs.pagesize&"</b>,  总数:<b>"&totalrec&"</b>,  "&_
    " 分页："

    If currentpage > 3 Then '当前页数大于3
        response.Write " <a href=?page=1>[1]</a> ..."
    End If

    If Pcount>currentpage+3 Then '当前页数在总数的前3页
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
