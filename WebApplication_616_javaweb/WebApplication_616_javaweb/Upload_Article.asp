<!--#include file="inc/conn.asp"-->
<!--#include file="inc/check.asp"-->
<!--#include file="inc/Function.asp"-->
<%
	action=request("action")
	classID=request.Form("classID")
	picurl=request.Form("picurl")
	if action="save" then
	if picurl="" then Call ShowErr("请先点上传再提交!",-1)
	set rs=server.CreateObject("adodb.recordset")
	sql="select * from QDpic where ClassID="&ClassID
	rs.open sql,conn,1,3
	if rs.eof then
		rs.addnew
	end if
	rs("ClassID")=ClassID
	rs("PicUrl")=PicUrl
	rs.update
	rs.close
	set rs=nothing
	Call ShowErr("添加成功！","Upload_Article.asp")
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="Admin_Style.css">
<style type="text/css">
<!--
BODY{
BACKGROUND-COLOR: #E1F4EE;
font-size:9pt
}
.tx1 { height: 20px;font-size: 9pt; border: 1px solid; border-color: #000000; color: #0000FF}
.STYLE1 {color: #FF0000}
-->
</style>

<SCRIPT language=javascript>
function AddItem(strFileName){
  document.form1.PicUrl.value=strFileName;
}
</SCRIPT>
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg">
    <td height="22"  align="center"><strong> 广 告 管 理</strong></td>
  </tr>
  <tr class="tdbg">
    <td height="30" align="center" ><a href="Upload_Article.asp">专题头题广告图片</a> | <a href="Upload_Articlepf.asp">网站首页漂浮广告</a> </td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
<form action="upload_article.asp" method="post" name="form1">
  <tr>
    <td height="36" colspan="2"><h4>庆典广告图片上传</h4></td>
    </tr>
  <tr>
    <td colspan="2"><%
            dim Rs,sqlStr
            set Rs=server.createobject("adodb.recordset")
            sqlStr="select * from QDclass"
            Rs.open sqlStr,conn,1,1
			if not Rs.eof then
			response.write "<select name='ClassID' size=1>"
            do while not Rs.eof
            	response.write "<option value="&Rs("ClassID")&">"&Rs("ClassName")&"</option>"
            	Rs.movenext
            loop
			response.write "</select>"
			else
				response.write "还没有任何栏目，请先新闻栏目。"
			end if
            Rs.close
            %></td>
    </tr>
  <tr>
    <td width="35%"><input name="PicUrl" type="text" id="PicUrl" size="40"></td>
    <td width="65%"><iframe src="upload_form.asp" width="260" height="25" frameborder="0"></iframe>
      <span class="STYLE1">&nbsp;大小:565×200px</span></td>
  </tr>
  <tr>
    <td height="27" colspan="2" style="padding-left:"><input type="submit" name="Submit2" value="提交"><input type="hidden" name="action" id="action" value="save"></td>
    </tr>
  </form>
</table>
</body>
</html>
